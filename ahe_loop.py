"""
AHE 세 기둥 구현:
  ❶ Component Observability — harness/ 파일들이 편집 대상
  ❷ Experience Observability — 트레이스 → digest 압축
  ❸ Decision Observability — Evolve Agent가 harness 편집 + manifest 예측 기록
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

SKILL_DIR = Path.home() / ".ppt-skill"


def _load_json(path: Path) -> dict:
    if path.exists():
        return json.loads(path.read_text())
    return {}


def _save_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2))


def _load_jsonl(path: Path) -> list[dict]:
    """change_manifest.jsonl 원장을 1줄=1엔트리 리스트로 로드한다.

    빈 줄·파싱 불가 줄은 건너뛴다(데이터 손실 방지: 원본은 그대로 둠).
    """
    if not path.exists():
        return []
    entries: list[dict] = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            entries.append(json.loads(ln))
        except json.JSONDecodeError:
            # 손상 줄은 보존을 위해 무시하지 않고 그대로 둘 수 없으므로
            # 안전하게 건너뛴다(write-back 시 재기록 안 함 → 호출부에서 미사용).
            continue
    return entries


def _save_jsonl(path: Path, entries: list[dict]) -> None:
    """엔트리 리스트를 1줄=1엔트리 JSONL로 원자적 재기록한다."""
    lines = [json.dumps(e, ensure_ascii=False) for e in entries]
    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


# ── 독립 QA verdict 스키마 (#13 independent-qa-not-implemented) ─────────────
# AHE_PRINCIPLES §2(생성≠판정) + §1(명시적·채점 가능 기준).
# 독립 격리 QA 에이전트(SKILL.md 11단계)는 지금까지 prose 프롬프트로만 존재했고
# 코드화된 스키마·결과 파싱이 없어, 그 종합 판정이 F1 루프(qa_ok)로 흘러갈 정형
# 인터페이스가 없었다. 아래가 그 인터페이스다: verdict를 채점 가능한 구조로 코드화하고,
# 임의의 에이전트 출력(JSON 또는 prose)을 그 구조로 관용 파싱한다.
#
# 이 모듈은 *인터페이스/스키마 우선*이다. 실제 에이전트 spawn은 optional이며
# (run_independent_qa의 spawn 콜백 주입), 비대화형/headless에서는 죽지 않고
# qa_ok=None(판정 보류) verdict로 안전 폴백한다.

VERDICT_PASS = "pass"          # 이슈 없음 → qa_ok=True
VERDICT_NEEDS_FIX = "needs_fix"  # 수정 필요 → qa_ok=False
VERDICT_DEFERRED = "deferred"    # 독립 판정 보류(에이전트 미가용) → qa_ok=None

_VERDICT_TO_QA_OK = {
    VERDICT_PASS: True,
    VERDICT_NEEDS_FIX: False,
    VERDICT_DEFERRED: None,
}

# 에이전트 prose 종합 판정 → 정식 verdict 토큰 (관용 매핑, 한/영)
_PASS_TOKENS = ("pass", "통과", "이슈 없음", "이슈없음", "ok", "approved", "no issues")
_FAIL_TOKENS = ("needs_fix", "needs fix", "수정 필요", "수정필요", "fail", "이슈 있음",
                "rejected", "issues found")


@dataclass
class QaVerdict:
    """독립 QA 에이전트의 채점 가능한 종합 판정 (§1 concrete/gradable terms).

    `verdict`는 VERDICT_PASS / VERDICT_NEEDS_FIX / VERDICT_DEFERRED 중 하나.
    `issues`는 슬라이드별 발견 목록(없으면 빈 리스트). `summary`는 1~2줄 요약.
    `independent`는 이 판정이 **생성 컨텍스트를 모르는 분리 에이전트**에서 왔는지
    (True)를 표시한다 — §2 확증편향 차단의 런타임 표식. 인라인 자기-QA가
    잘못 verdict로 승격되는 것을 막기 위해 기본 False.
    """

    verdict: str = VERDICT_DEFERRED
    issues: list[dict] = field(default_factory=list)
    summary: str = ""
    independent: bool = False

    def __post_init__(self) -> None:
        if self.verdict not in _VERDICT_TO_QA_OK:
            self.verdict = VERDICT_DEFERRED

    @property
    def qa_ok(self) -> bool | None:
        """F1 루프(update_last_run_qa / maybe_run_evolve_loop)에 넘길 3-값.

        독립 판정이 아니면(self.independent=False) PASS여도 qa_ok로 승격하지
        않는다 — §2: 생성자 자기-QA는 '통과'를 주장할 자격이 없다(deferred 취급).
        NEEDS_FIX(결함 보고)는 분리 여부와 무관하게 False로 신뢰한다.
        """
        if self.verdict == VERDICT_NEEDS_FIX:
            return False
        if not self.independent:
            return None
        return _VERDICT_TO_QA_OK[self.verdict]

    def to_dict(self) -> dict:
        return {
            "verdict": self.verdict,
            "issues": self.issues,
            "summary": self.summary,
            "independent": self.independent,
            "qa_ok": self.qa_ok,
        }


def parse_verdict(raw: object, *, independent: bool = False) -> QaVerdict:
    """독립 QA 에이전트의 임의 출력 → QaVerdict로 관용 파싱한다 (결과 파싱 코드화).

    허용 입력:
      - dict: {"verdict": "...", "issues": [...], "summary": "..."} (정식)
      - dict: SKILL.md 11단계 리포트형 {"slides": [{"issues": [...]}], "summary": ...}
      - str:  ```json 펜스 포함 가능 → JSON 시도 후 실패 시 prose 토큰 매칭
      - None/빈 값: VERDICT_DEFERRED (에이전트 미가용)

    파싱 불가/모호하면 결함을 *날조하지 않고* DEFERRED로 떨어진다(보수적).
    """
    if raw is None or raw == "" or raw == {}:
        return QaVerdict(verdict=VERDICT_DEFERRED, independent=independent)

    data: object = raw
    if isinstance(raw, str):
        s = raw.strip()
        if s.startswith("```"):
            import re
            s = re.sub(r"^```[a-z]*\n?", "", s)
            s = re.sub(r"\n?```$", "", s).strip()
        try:
            data = json.loads(s)
        except (json.JSONDecodeError, ValueError):
            # prose: 종합 판정 토큰만 본다. FAIL 신호 우선(보수적).
            low = s.lower()
            if any(t in low for t in _FAIL_TOKENS):
                return QaVerdict(VERDICT_NEEDS_FIX, summary=s[:200],
                                 independent=independent)
            if any(t in low for t in _PASS_TOKENS):
                return QaVerdict(VERDICT_PASS, summary=s[:200],
                                 independent=independent)
            return QaVerdict(VERDICT_DEFERRED, summary=s[:200],
                             independent=independent)

    if not isinstance(data, dict):
        return QaVerdict(verdict=VERDICT_DEFERRED, independent=independent)

    # 정식 스키마
    if "verdict" in data:
        v = str(data.get("verdict", "")).strip().lower()
        if v not in _VERDICT_TO_QA_OK:
            # 자유 토큰 → 정식 토큰 매핑
            if any(t in v for t in _FAIL_TOKENS):
                v = VERDICT_NEEDS_FIX
            elif any(t in v for t in _PASS_TOKENS):
                v = VERDICT_PASS
            else:
                v = VERDICT_DEFERRED
        return QaVerdict(
            verdict=v,
            issues=list(data.get("issues", []) or []),
            summary=str(data.get("summary", "")),
            independent=independent,
        )

    # 리포트형(slides[].issues) → 종합 verdict 도출
    issues: list[dict] = []
    for slide in data.get("slides", []) or []:
        for iss in slide.get("issues", []) or []:
            issues.append({"slide": slide.get("index"),
                           "role": slide.get("role"), **iss})
    if "slides" in data:
        v = VERDICT_NEEDS_FIX if issues else VERDICT_PASS
        return QaVerdict(verdict=v, issues=issues,
                         summary=str(data.get("summary", "")),
                         independent=independent)

    return QaVerdict(verdict=VERDICT_DEFERRED,
                     summary=str(data.get("summary", "")),
                     independent=independent)


def _is_headless() -> bool:
    """현재 실행이 headless(오케스트레이터 없는 비대화형)인지 추정한다.

    PPT_SKILL_HEADLESS=1 이 명시되면 그 값을 신뢰한다(main.py --evolve 경로가 설정).
    그 외에는 stdin이 tty가 아니면 headless로 본다(파이프/cron/CI).
    오케스트레이터(Claude Code 세션)에서 inline_vision_qa를 켜는 것을 차단하는 게 목적.
    """
    explicit = os.environ.get("PPT_SKILL_HEADLESS")
    if explicit is not None:
        return explicit not in ("0", "", "false", "False")
    try:
        return not sys.stdin.isatty()
    except (ValueError, AttributeError):
        return True


def assert_inline_qa_headless_only() -> None:
    """인라인 vision QA(analyze_qa_images)가 headless 폴백에서만 돌도록 런타임 보장.

    AHE_PRINCIPLES §2: 인라인 vision QA는 생성과 동일 프로세스의 자기-검증이라
    확증편향 위험이 있다. 오케스트레이터 세션(대화형)에서 호출되면 §2 위반이므로
    RuntimeError로 막는다. PPT_SKILL_ALLOW_INLINE_QA=1 로 명시 오버라이드 가능
    (테스트/특수 상황용).
    """
    if os.environ.get("PPT_SKILL_ALLOW_INLINE_QA") == "1":
        return
    if not _is_headless():
        raise RuntimeError(
            "인라인 vision QA(analyze_qa_images)는 headless 폴백 전용입니다 "
            "(AHE_PRINCIPLES §2 생성≠판정). 오케스트레이터 세션에서는 SKILL.md "
            "11단계의 독립 격리 QA 에이전트를 사용하세요. "
            "강제하려면 PPT_SKILL_ALLOW_INLINE_QA=1."
        )


def run_independent_qa(pptx_path: Path, topic: str,
                       spawn=None) -> QaVerdict:
    """독립 격리 QA 에이전트 호출의 코드화된 진입점 (#13, §2 분리).

    `spawn`은 (pptx_path, topic) → 에이전트 raw 출력(str|dict)을 반환하는 콜백이다.
    오케스트레이터(Claude Code 세션)가 Agent tool로 독립 에이전트를 띄워 그 출력을
    여기로 넘긴다 — 그 콜백은 plan·생성 근거를 전달받지 않는다(독립=True).

    spawn=None(비대화형/headless, 또는 콜백 미주입)이면 **죽지 않고** DEFERRED
    verdict(qa_ok=None)로 폴백한다 — 닫히지 않은 run은 진화 대상에서 빠진다(§1·#12).
    """
    if spawn is None:
        return QaVerdict(verdict=VERDICT_DEFERRED,
                         summary="독립 QA 에이전트 미주입 — 판정 보류(headless 폴백)",
                         independent=False)
    try:
        raw = spawn(pptx_path, topic)
    except Exception as e:  # 에이전트 실패가 파이프라인을 죽이지 않게
        return QaVerdict(verdict=VERDICT_DEFERRED,
                         summary=f"독립 QA spawn 실패: {e}", independent=False)
    return parse_verdict(raw, independent=True)


# ── ❷ Experience Observability ───────────────────────────────

def collect_trace(work_dir: Path, topic: str) -> dict:
    """실행 결과를 구조화된 트레이스 dict로 반환한다."""
    plan_path   = work_dir / "plan.json"
    output_path = work_dir / "output.pptx"
    qa_report   = work_dir / "qa_report.json"

    plan = _load_json(plan_path)
    trace = {
        "run_id":           work_dir.name,
        "topic":            topic,
        "timestamp":        datetime.utcnow().isoformat(),
        "n_slides_planned": plan.get("n_slides", 0),
        "n_slides_actual":  len(plan.get("slides", [])),
        "output_exists":    output_path.exists(),
        "issues":           [],
        "qa": {},
    }

    # 플레이스홀더 잔여 검사
    if output_path.exists():
        extract = (
            "import sys; from pptx import Presentation;"
            "prs=Presentation(sys.argv[1]);"
            "[print(p.text.strip()) for s in prs.slides"
            " for sh in s.shapes if sh.has_text_frame"
            " for p in sh.text_frame.paragraphs if p.text.strip()]"
        )
        result = subprocess.run(
            [sys.executable, "-c", extract, str(output_path)],
            capture_output=True, text=True,
        )
        trace["issues"] = [
            ln.strip() for ln in result.stdout.splitlines()
            if any(kw in ln.lower() for kw in ["lorem", "todo", "placeholder", "작성해주세요"])
        ]

    # QA 리포트 (이미지 수, 소스)
    if qa_report.exists():
        qa = _load_json(qa_report)
        images = [Path(p) for p in qa.get("images", []) if Path(p).exists()]
        trace["qa"] = {
            "image_count": len(images),
            "source":      qa.get("source", "unknown"),
            "image_paths": [str(p) for p in images],
        }

    return trace


# ── Vision 이미지 분석 ────────────────────────────────────────

def analyze_qa_images(image_paths: list[str], plan: dict) -> dict:
    """
    QA 이미지를 Claude Vision으로 분석해 디자인 이슈를 반환한다.

    반환 구조:
    {
      "slides": [
        {
          "index": 1,
          "role": "cover",
          "issues": [
            {"type": "overflow", "severity": "HIGH", "detail": "대제목 텍스트 오른쪽 잘림"},
            {"type": "spacing",  "severity": "MED",  "detail": "날짜와 제목 간격 너무 좁음"}
          ]
        }
      ],
      "summary": "전체 디자인 요약",
      "design_patterns": ["긴 제목은 48pt에서 줄바꿈 발생"]
    }
    """
    if not image_paths:
        return {"slides": [], "summary": "QA 이미지 없음", "design_patterns": []}

    # AHE_PRINCIPLES §2(생성≠판정): 이 인라인 self-QA는 plan(slide_roles)을 그대로
    # 받는 동일-프로세스 검증이라 확증편향 위험이 있다. headless 폴백에서만 허용한다
    # (#13 independent-qa-not-implemented). 오케스트레이터 세션이면 RuntimeError.
    assert_inline_qa_headless_only()

    import base64

    # 슬라이드별 역할 매핑
    slide_roles = {s["index"]: s.get("role", "content") for s in plan.get("slides", [])}

    system = """당신은 PowerPoint 슬라이드 디자인 QA 전문가입니다.
슬라이드 이미지를 보고 디자인 문제를 정확하게 감지합니다.

아래 JSON 형식으로만 응답하세요:
{
  "slides": [
    {
      "index": 1,
      "role": "cover",
      "issues": [
        {
          "type": "overflow|spacing|overlap|alignment|font|color|placeholder|other",
          "severity": "CRITICAL|HIGH|MEDIUM|LOW",
          "detail": "구체적인 문제 설명 (한국어)"
        }
      ],
      "ok": true
    }
  ],
  "summary": "전체 디자인 품질 요약 (한국어, 2줄 이내)",
  "design_patterns": ["반복 패턴이나 학습할 내용 (있을 때만)"]
}

감지 대상:
- overflow: 텍스트/요소가 박스나 슬라이드 경계 밖으로 넘침
- spacing: 요소 간 여백이 너무 좁거나 넓음
- overlap: 요소들이 겹침
- alignment: 정렬 불일치
- font: 폰트 크기 과소/과대, 가독성 문제
- color: 배경과 텍스트 대비 부족
- placeholder: 작성해주세요, ?? 등 미교체 텍스트
- other: 기타 디자인 문제"""

    # 이미지를 base64로 인코딩해 API에 전달
    content = []
    loaded_count = 0
    for i, img_path in enumerate(image_paths[:10]):  # 최대 10장
        p = Path(img_path)
        if not p.exists():
            continue
        b64 = base64.standard_b64encode(p.read_bytes()).decode()
        ext = p.suffix.lower().lstrip(".")
        media = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
        slide_num = i + 1
        role = slide_roles.get(slide_num, "content")
        content.append({
            "type": "text",
            "text": f"=== 슬라이드 {slide_num} (role: {role}) ===",
        })
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media, "data": b64},
        })
        loaded_count += 1

    if loaded_count == 0:
        return {"slides": [], "summary": "이미지 로드 실패", "design_patterns": []}

    content.append({
        "type": "text",
        "text": f"위 {loaded_count}장의 슬라이드를 분석해 디자인 이슈를 JSON으로 반환하세요.",
    })

    # Claude API 호출
    raw = _call_claude_vision(system, content)
    if not raw:
        return {"slides": [], "summary": "Vision API 호출 실패", "design_patterns": []}

    import re
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"slides": [], "summary": raw[:200], "design_patterns": []}


def _select_backend(vertex_proj, vertex_region, aws_region):
    """백엔드 선택 — ppt_generator.generate_plan_with_claude 와 동일한 우선순위.
    PPT_SKILL_BACKEND > CLAUDE_CODE_USE_* > 자동 감지. (use_vertex, use_bedrock) 반환."""
    import os
    explicit = os.environ.get("PPT_SKILL_BACKEND", "auto")  # auto|vertex|bedrock|anthropic
    cc_bedrock = os.environ.get("CLAUDE_CODE_USE_BEDROCK", "0")
    cc_vertex  = os.environ.get("CLAUDE_CODE_USE_VERTEX", "0")
    if explicit == "vertex":
        return True, False
    if explicit == "bedrock":
        return False, True
    if explicit == "anthropic":
        return False, False
    if cc_bedrock == "1":
        return False, True
    if cc_vertex == "1":
        return True, False
    if cc_bedrock == "0" and cc_vertex == "0":
        # switch-provider.sh direct — 두 플래그가 명시적으로 0이면 직접 API
        return False, False
    # 환경 변수 미설정 시 자동 감지
    use_vertex = bool(vertex_proj or (vertex_region and not aws_region))
    use_bedrock = bool(aws_region) and not use_vertex
    return use_vertex, use_bedrock


def _call_claude_vision(system: str, content: list) -> str | None:
    """멀티모달(Vision) Claude API 호출."""
    import os

    api_key      = os.environ.get("ANTHROPIC_API_KEY")
    vertex_proj  = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region   = os.environ.get("AWS_REGION")

    use_vertex, use_bedrock = _select_backend(vertex_proj, vertex_region, aws_region)

    if not api_key and not use_vertex and not use_bedrock:
        print("  ⚠ Vision API 없음 — 이미지 분석 건너뜀")
        return None

    model = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")
    messages = [{"role": "user", "content": content}]

    try:
        if use_bedrock:
            # Bedrock은 base64 이미지를 다르게 처리
            import boto3, json as _json
            client = boto3.client("bedrock-runtime", region_name=aws_region or "us-east-1")
            bedrock_model = model if (model.startswith("us.anthropic.") or model.startswith("anthropic.")) else "us.anthropic.claude-sonnet-4-6"
            body = _json.dumps({
                "anthropic_version": "bedrock-2023-05-31",
                "max_tokens": 4096,
                "system": system,
                "messages": messages,
            })
            resp = client.invoke_model(
                modelId=bedrock_model, body=body)
            return _json.loads(resp["body"].read())["content"][0]["text"].strip()

        import anthropic
        sdk_model = model[len("us.anthropic."):] if model.startswith("us.anthropic.") else model
        client = (
            anthropic.AnthropicVertex(project_id=vertex_proj or "", region=vertex_region or "us-east5")
            if use_vertex else anthropic.Anthropic(api_key=api_key)
        )
        resp = client.messages.create(
            model=sdk_model,
            max_tokens=4096,
            thinking={"type": "adaptive"},
            system=system, messages=messages,
        )
        return resp.content[0].text.strip()

    except Exception as e:
        print(f"  ⚠ Vision API 호출 실패: {e}")
        return None


def distill_digest(trace: dict, vision_result: dict | None = None) -> dict:
    """트레이스 + Vision 분석 결과 → 핵심 패턴 digest."""
    # 주의: 이것은 "QA 이미지가 존재하는가"일 뿐, 실제 QA 판정(QaVerdict.qa_ok)이 아니다.
    # 이전 이름(qa_ok/qa_done)이 Evolve Agent에게 "QA 통과"로 오해됐다 → 명확히 rename.
    qa_images_exist = trace["qa"].get("image_count", 0) > 0
    slide_match = trace["n_slides_planned"] == trace["n_slides_actual"]

    # Vision 이슈 집계
    vision_issues: list[dict] = []
    vision_summary = ""
    design_patterns: list[str] = []
    if vision_result:
        vision_summary = vision_result.get("summary", "")
        design_patterns = vision_result.get("design_patterns", [])
        for slide in vision_result.get("slides", []):
            for issue in slide.get("issues", []):
                vision_issues.append({
                    "slide": slide.get("index"),
                    "role":  slide.get("role"),
                    **issue,
                })

    critical_vision = [v for v in vision_issues if v.get("severity") in ("CRITICAL", "HIGH")]

    return {
        "run_id":          trace["run_id"],
        "topic":           trace["topic"],
        "success":         trace["output_exists"]
                           and len(trace["issues"]) == 0
                           and len(critical_vision) == 0,
        "issue_count":     len(trace["issues"]),
        "patterns":        trace["issues"][:5],
        "qa_images_exist": qa_images_exist,
        "slide_match":     slide_match,
        "vision": {
            "total_issues":    len(vision_issues),
            "critical_issues": critical_vision,
            "summary":         vision_summary,
            "design_patterns": design_patterns,
        },
        "flags": {
            "has_placeholder":   len(trace["issues"]) > 0,
            "has_design_issues": len(critical_vision) > 0,
            "headless_qa_images_exist": qa_images_exist,
            "slide_count_ok":    slide_match,
        },
    }


# ── ❸ Decision Observability — Evolve Agent ──────────────────

def _call_claude(system: str, user: str) -> str | None:
    """Claude API 호출. 백엔드 선택은 _select_backend (PPT_SKILL_BACKEND > CLAUDE_CODE_USE_* > 자동 감지)."""
    import os
    api_key      = os.environ.get("ANTHROPIC_API_KEY")
    vertex_proj  = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region   = os.environ.get("AWS_REGION")

    use_vertex, use_bedrock = _select_backend(vertex_proj, vertex_region, aws_region)

    if not api_key and not use_vertex and not use_bedrock:
        return None

    model = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")
    messages = [{"role": "user", "content": user}]

    try:
        if use_bedrock:
            import boto3, json as _json
            client = boto3.client("bedrock-runtime", region_name=aws_region or "us-east-1")
            bedrock_model = model if (model.startswith("us.anthropic.") or model.startswith("anthropic.")) else "us.anthropic.claude-sonnet-4-6"
            body = _json.dumps({
                "anthropic_version": "bedrock-2023-05-31",
                "max_tokens": 4096,
                "system": system,
                "messages": messages,
            })
            resp = client.invoke_model(
                modelId=bedrock_model, body=body)
            return _json.loads(resp["body"].read())["content"][0]["text"].strip()

        import anthropic
        sdk_model = model[len("us.anthropic."):] if model.startswith("us.anthropic.") else model
        client = (
            anthropic.AnthropicVertex(project_id=vertex_proj or "", region=vertex_region or "us-east5")
            if use_vertex else anthropic.Anthropic(api_key=api_key)
        )
        resp = client.messages.create(
            model=sdk_model,
            max_tokens=4096,
            thinking={"type": "adaptive"},
            system=system, messages=messages,
        )
        return resp.content[0].text.strip()

    except Exception as e:
        print(f"  ⚠ Claude API 호출 실패: {e}")
        return None


def run_evolve_agent(digest: dict, harness_dir: Path,
                     vision_result: dict | None = None) -> dict:
    """
    Claude API를 호출해 harness 파일을 실제로 편집하고
    change_manifest(예측 포함)를 반환한다.

    편집 대상:
      - long_term_memory.json : 이번 실행 경험 누적
      - verifier_rules.json   : 새 검증 규칙 추가
      - CLAUDE.md             : 행동 규칙 보완 (텍스트 추가)
    """
    memory_path   = harness_dir / "long_term_memory.json"
    verifier_path = harness_dir / "verifier_rules.json"
    claude_md     = harness_dir / "CLAUDE.md"

    memory   = _load_json(memory_path)
    verifier = _load_json(verifier_path)
    claude_system_prompt = claude_md.read_text() if claude_md.exists() else ""

    system = """당신은 AHE(Agentic Harness Engineering) Evolve Agent입니다.
PPT 생성 시스템의 실행 결과(digest)를 분석하고 하네스 파일을 개선합니다.

반드시 아래 JSON 형식으로만 응답하세요 (다른 텍스트 없이):
{
  "analysis": "이번 실행에서 발견한 문제 요약 (1~3줄)",
  "memory_patch": {
    "key": "추가할 최상위 키 (없으면 null)",
    "value": "추가할 값 (없으면 null)"
  },
  "verifier_patch": {
    "rule_name": "추가할 규칙 이름 (없으면 null)",
    "rule": { "description": "...", "severity": "HIGH|MEDIUM|LOW", "action": "FAIL|WARN" }
  },
  "claude_md_append": "CLAUDE.md 끝에 추가할 텍스트 (없으면 null)",
  "predictions": [
    { "change": "변경 내용", "expected": "기대 효과", "metric": "검증 방법",
      "manifest_id": "연결된 change_manifest.jsonl 엔트리 id (없으면 null)" }
  ]
}

manifest_id는 이 예측이 검증하는 change_manifest.jsonl 원장 엔트리의 id(예: "2026-06-10-08")를
가리킨다. 해당 엔트리가 verification:'pending'이면 다음 라운드에 자동으로 verified/refuted로 해소된다.
연결 대상이 없으면 null."""

    vision = digest.get("vision", {})
    vision_block = ""
    if vision.get("critical_issues"):
        vision_block = f"""
## Vision 분석 결과 (디자인 이슈)
요약: {vision.get('summary', '')}
CRITICAL/HIGH 이슈:
{json.dumps(vision['critical_issues'], ensure_ascii=False, indent=2)}
학습할 디자인 패턴: {vision.get('design_patterns', [])}
"""

    user = f"""## 이번 실행 digest
```json
{json.dumps(digest, ensure_ascii=False, indent=2)}
```
{vision_block}
## 현재 long_term_memory.json
```json
{json.dumps(memory, ensure_ascii=False, indent=2)}
```

## 현재 verifier_rules.json (rules 키 목록)
{list(verifier.get('rules', {}).keys())}

## 현재 CLAUDE.md (마지막 5줄)
{chr(10).join(claude_system_prompt.splitlines()[-5:])}

위 정보를 바탕으로 harness 파일 개선안을 JSON으로 제시하세요.
Vision 분석에서 발견된 디자인 이슈는 long_term_memory의 known_failure_fixes나
verifier_rules에 반드시 반영하세요.
이슈가 없는 성공 실행이라면 모든 patch를 null로 반환하세요."""

    actual_changes: list[dict] = []
    predictions: list[dict] = []

    raw = _call_claude(system, user)
    if raw:
        # JSON 블록 추출
        import re
        if raw.startswith("```"):
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw)
        try:
            proposal = json.loads(raw)
        except json.JSONDecodeError:
            proposal = {}

        print(f"  [Evolve] 분석: {proposal.get('analysis', '(없음)')[:80]}")

        # long_term_memory 패치
        mp = proposal.get("memory_patch", {}) or {}
        if mp.get("key") and mp.get("value") is not None:
            memory[mp["key"]] = mp["value"]
            memory["last_evolved"] = datetime.utcnow().isoformat()
            _save_json(memory_path, memory)
            actual_changes.append({"file": "long_term_memory.json", "key": mp["key"]})
            print(f"  ✓ long_term_memory.json 패치: {mp['key']}")

        # verifier_rules 패치
        vp = proposal.get("verifier_patch", {}) or {}
        if vp.get("rule_name") and vp.get("rule"):
            verifier.setdefault("rules", {})[vp["rule_name"]] = vp["rule"]
            verifier["last_evolved"] = datetime.utcnow().isoformat()
            _save_json(verifier_path, verifier)
            actual_changes.append({"file": "verifier_rules.json", "rule": vp["rule_name"]})
            print(f"  ✓ verifier_rules.json 패치: {vp['rule_name']}")

        # CLAUDE.md append
        append_text = proposal.get("claude_md_append")
        if append_text:
            with open(claude_md, "a", encoding="utf-8") as f:
                f.write(f"\n\n## AHE 추가 규칙 ({digest['run_id'][:12]})\n{append_text}")
            actual_changes.append({"file": "CLAUDE.md", "appended": True})
            print("  ✓ CLAUDE.md 규칙 추가")

        predictions = proposal.get("predictions", [])

    # ── Vision → Planning constraints 저장 ─────────────────────
    if vision_result:
        new_constraints = extract_planning_constraints(vision_result, digest)
        if new_constraints:
            memory_path = harness_dir / "long_term_memory.json"
            memory = _load_json(memory_path)
            # 기존 constraints와 병합 (최근 20개 유지)
            existing = memory.get("planning_constraints", [])
            merged = list(dict.fromkeys(existing + new_constraints))[-20:]
            memory["planning_constraints"] = merged
            memory["last_evolved"] = datetime.utcnow().isoformat()
            _save_json(memory_path, memory)
            actual_changes.append({
                "file": "long_term_memory.json",
                "key": "planning_constraints",
                "count": len(merged),
            })
            print(f"  ✓ planning_constraints 업데이트: {len(merged)}개")

    else:
        # Claude API 없음 → 규칙 기반 폴백
        if digest["issue_count"] > 0:
            key = f"fix_{datetime.utcnow().strftime('%Y%m%d')}"
            memory[key] = {
                "patterns": digest["patterns"],
                "action": "플레이스홀더 제거 필요",
            }
            memory["last_evolved"] = datetime.utcnow().isoformat()
            _save_json(memory_path, memory)
            actual_changes.append({"file": "long_term_memory.json", "key": key})
            predictions.append({
                "change": "placeholder 패턴 기록",
                "expected": "다음 실행에서 동일 패턴 감지 시 경고",
                "metric": "issue_count == 0",
            })
            print("  ✓ 폴백: long_term_memory.json 패치 (플레이스홀더 패턴)")
        else:
            print("  ✓ 성공 실행 — 하네스 변경 없음")

    return {
        "actual_changes": actual_changes,
        "predictions": predictions,
    }


# ── Prediction Verification ───────────────────────────────────

def verify_predictions(current_digest: dict, evolution_dir: Path) -> None:
    """
    이전 실행의 manifest 예측 vs 현재 실행 결과를 비교해
    검증 결과를 manifest에 기록한다.
    """
    manifests = sorted(evolution_dir.glob("iteration_*_manifest.json"))
    if len(manifests) < 2:
        return  # 비교 대상 없음

    prev_manifest_path = manifests[-2]   # 직전 manifest
    prev = _load_json(prev_manifest_path)

    if not prev.get("predictions"):
        return

    verified: list[dict] = []
    current_success = current_digest["success"]

    for pred in prev["predictions"]:
        metric = pred.get("metric", "")
        # 간단한 지표 검증: "issue_count == 0" 형태
        if "issue_count == 0" in metric:
            result = "PASS" if current_success else "FAIL"
        else:
            result = "PASS" if current_success else "UNVERIFIED"

        verified.append({**pred, "verification": result})

    prev["prediction_results"] = verified
    prev["verified_at"] = datetime.utcnow().isoformat()
    _save_json(prev_manifest_path, prev)

    n_pass = sum(1 for v in verified if v["verification"] == "PASS")
    print(f"  ✓ 예측 검증: {n_pass}/{len(verified)} PASS (이전 manifest: {prev_manifest_path.name})")

    # ── 단방향 브리지: 검증 결과 → change_manifest.jsonl 원장 ──
    # iteration_*_manifest.json(휘발성 per-run)의 검증 verdict를
    # authoritative 원장(change_manifest.jsonl)의 pending 엔트리로 흘려보낸다.
    # 역방향(원장→manifest) 금지. 스키마는 분리 유지하고 verdict만 매핑.
    bridge_verdicts_to_ledger(verified, SKILL_DIR / "harness")


# ── ❸ Decision Observability — 원장 브리지 ────────────────────

# iteration manifest의 검증 결과(PASS/FAIL) → 원장 verification 매핑
_VERDICT_TO_LEDGER = {"PASS": "verified", "FAIL": "refuted"}


def bridge_verdicts_to_ledger(verified_predictions: list[dict],
                              harness_dir: Path) -> int:
    """단방향 브리지: 검증된 per-run 예측의 verdict를 change_manifest.jsonl로 옮긴다.

    `change_manifest.jsonl`이 authoritative 결정 관찰성 원장(AHE_PRINCIPLES §4/§5③).
    이 원장에 write하는 코드가 0건이라 `verification:'pending'` 엔트리는
    영영 갱신되지 않았다. 이 브리지가 그 유일한 write 경로다.

    매칭은 **명시적 링크**(prediction['manifest_id'] == ledger entry['id'])로만 한다.
    텍스트 휴리스틱 매칭은 큐레이션된 엔트리를 잘못 뒤집을 수 있어 금지.
    PASS→verified, FAIL→refuted. UNVERIFIED 등 결정 불가 verdict는 건너뛴다.

    기존 원장 데이터는 보존(필드 손실 없음): 추가로
    `verified_at`/`verified_by_run` 메타만 덧붙이고, 기존 `verification`
    prose가 'pending' prefix일 때만 상태 토큰을 교체한다.

    Returns: 갱신된 원장 엔트리 수.
    """
    ledger_path = harness_dir / "change_manifest.jsonl"
    entries = _load_jsonl(ledger_path)
    if not entries:
        return 0

    # manifest_id → verdict(PASS/FAIL) 매핑 (결정 가능한 것만)
    verdict_by_id: dict[str, str] = {}
    for pred in verified_predictions:
        mid = pred.get("manifest_id")
        verdict = pred.get("verification")
        if mid and verdict in _VERDICT_TO_LEDGER:
            verdict_by_id[mid] = verdict  # 같은 id 다건이면 마지막 verdict 채택

    if not verdict_by_id:
        return 0

    now = datetime.utcnow().isoformat()
    updated = 0
    for entry in entries:
        eid = entry.get("id")
        if eid not in verdict_by_id:
            continue
        current = str(entry.get("verification", ""))
        # pending 상태만 해소 (이미 verified/refuted면 큐레이션 결과 존중)
        if not current.lower().startswith("pending"):
            continue
        new_state = _VERDICT_TO_LEDGER[verdict_by_id[eid]]
        # 기존 prose 보존: 'pending …' 꼬리말은 유지하되 상태 토큰만 교체
        tail = current[len("pending"):].lstrip(" :—-")
        entry["verification"] = f"{new_state} (auto-bridge)" + (f" — {tail}" if tail else "")
        entry["verified_at"] = now
        entry["verified_by_run"] = "verify_predictions"
        updated += 1

    if updated:
        _save_jsonl(ledger_path, entries)
        print(f"  ✓ 원장 브리지: change_manifest.jsonl {updated}건 pending→해소")
    return updated


# ── 장기 기억 통계 업데이트 ───────────────────────────────────

def extract_planning_constraints(vision_result: dict, digest: dict) -> list[str]:
    """
    Vision 분석 결과와 digest에서 다음 실행의 plan 생성에 주입할
    구체적 제약 조건 문자열 목록을 추출한다.

    예: "슬라이드 7(toc): 텍스트 오버플로우 — 항목당 20자 이내 권장"
    """
    constraints: list[str] = []
    if not vision_result:
        return constraints

    for slide in vision_result.get("slides", []):
        idx  = slide.get("index", "?")
        role = slide.get("role", "content")
        for issue in slide.get("issues", []):
            sev    = issue.get("severity", "LOW")
            itype  = issue.get("type", "other")
            detail = issue.get("detail", "")

            if sev not in ("CRITICAL", "HIGH"):
                continue

            if itype == "overflow":
                constraints.append(
                    f"슬라이드 {idx}({role}) 오버플로우 방지: {detail[:60]} "
                    "— 제목 25자 이내, 항목 30자 이내로 제한"
                )
            elif itype == "placeholder":
                constraints.append(
                    f"슬라이드 {idx}({role}): 플레이스홀더 미교체 발생 "
                    "— 반드시 실제 콘텐츠로 채울 것 (작성해주세요 금지)"
                )
            elif itype == "font":
                constraints.append(
                    f"슬라이드 {idx}({role}) 폰트 조정 필요: {detail[:60]} "
                    "— 텍스트를 더 짧게 작성할 것"
                )
            elif itype == "spacing":
                constraints.append(
                    f"슬라이드 {idx}({role}) 여백 문제: {detail[:60]}"
                )

    # 전역 패턴
    for pattern in vision_result.get("design_patterns", []):
        if len(pattern) < 100:
            constraints.append(f"[반복 패턴] {pattern}")

    return constraints[:10]  # 최대 10개


def update_long_term_memory(digest: dict, harness_dir: Path) -> None:
    """long_term_memory.json에 실행 통계를 누적한다."""
    memory_path = harness_dir / "long_term_memory.json"
    memory = _load_json(memory_path)

    runs: list[dict] = memory.get("runs", [])
    runs.append({
        "run_id":      digest["run_id"],
        "success":     digest["success"],
        "issue_count": digest["issue_count"],
        "patterns":    digest["patterns"],
    })
    memory["runs"]          = runs[-20:]   # 최근 20회
    memory["total_runs"]    = memory.get("total_runs", 0) + 1
    memory["success_rate"]  = round(
        sum(1 for r in memory["runs"] if r["success"]) / len(memory["runs"]), 2
    )
    memory["last_updated"]  = datetime.utcnow().isoformat()

    _save_json(memory_path, memory)
    print(f"  ✓ long_term_memory: 총 {memory['total_runs']}회, "
          f"성공률 {memory['success_rate']*100:.0f}%")


# ── Git 커밋 ─────────────────────────────────────────────────

def git_commit_harness(skill_dir: Path, message: str) -> None:
    """harness/ 변경 사항을 git commit으로 추적한다."""
    harness_dir = skill_dir / "harness"
    r1 = subprocess.run(["git", "add", str(harness_dir)],
                        cwd=str(skill_dir), capture_output=True)
    if r1.returncode != 0:
        return
    subprocess.run(["git", "commit", "-m", message],
                   cwd=str(skill_dir), capture_output=True)
    print(f"  ✓ git commit: {message}")


# ── Evolve 트리거 게이트 (#12 evolve-manual-trigger-gap) ──────────
# skill 경로는 inline_vision_qa=False라 _vision_critical_total=0이 고정 반환되고
# main.py를 경유하지 않아 run_evolve_loop에 구조적으로 도달할 수 없었다.
# 독립 QA가 qa_ok=False를 내면 evolve를 *제안*하는 코드 경로를 추가하되,
# AHE_PRINCIPLES §1(human-in-loop: 고위험 결정엔 사람을 둔다)을 보존하기 위해
# 실제 실행은 명시적 사람 승인(approved=True) 없이는 절대 일어나지 않는다.

def should_trigger_evolve(qa_ok: bool | None,
                          vision_issues: int = 0,
                          explicit: bool = False) -> bool:
    """evolve가 '제안되어야 하는' 조건인지 판정하는 순수 술어 (채점 가능, §3).

    True이면 '진화할 가치가 있는 신호가 있다' — 이때도 실행 여부는 사람이 결정한다
    (maybe_run_evolve_loop의 approved 게이트). 실제 자동 실행을 의미하지 않는다.

      - explicit=True            → 사람이 명시적으로 --evolve 요청 (항상 제안)
      - qa_ok is False           → 독립 QA가 결함 판정 (skill 경로 환류 경로)
      - vision_issues > 0        → 인라인 vision 이슈 발견 (headless 경로)
    qa_ok=None(판정 보류)는 신호가 아니다 — 닫히지 않은 run을 진화시키지 않는다.
    """
    if explicit:
        return True
    if qa_ok is False:
        return True
    return int(vision_issues or 0) > 0


def maybe_run_evolve_loop(work_dir: Path, topic: str,
                          qa_ok: bool | None = None,
                          vision_issues: int = 0,
                          explicit: bool = False,
                          approved: bool = False) -> bool:
    """게이트가 걸린 evolve 진입점 — 독립 QA fail → evolve 환류 경로 (#12).

    self-healing이 --evolve/인라인 vision 이슈에만 의존하던 갭을 닫는다:
    독립 QA가 qa_ok=False를 기록한 run도 이 함수를 통해 evolve로 연결된다.

    그러나 자동 진화는 강제하지 않는다 (§1 human-in-loop, §4 회귀 위험).
    실제 run_evolve_loop 호출은 `approved=True`일 때만 일어난다:
      - explicit=True(사람이 --evolve 직접 지정)는 그 자체로 승인으로 본다.
      - 그 외(qa_ok=False 등)는 오케스트레이터/사용자가 승인 게이트를 통과시켜야 한다.

    반환: 실제로 run_evolve_loop를 실행했으면 True, 제안만 하고 멈췄으면 False.
    """
    if not should_trigger_evolve(qa_ok, vision_issues, explicit):
        return False

    # explicit --evolve 요청은 사람 의도가 이미 명시된 것이므로 승인으로 간주.
    human_approved = approved or explicit
    if not human_approved:
        reason = ("독립 QA 결함(qa_ok=False)"
                  if qa_ok is False else f"vision 이슈 {vision_issues}건")
        print(f"\n[Evolve 제안] {reason} 감지 — 진화 루프 실행을 권장합니다.")
        print("  ⏸ human-in-loop 게이트(AHE_PRINCIPLES §1): 자동 실행하지 않음.")
        print("  → 승인하려면 maybe_run_evolve_loop(..., approved=True) 또는 "
              "run_evolve_loop를 직접 호출하세요.")
        return False

    run_evolve_loop(work_dir=work_dir, topic=topic)
    return True


# ── 메인 진화 루프 ────────────────────────────────────────────

def run_evolve_loop(work_dir: Path, topic: str) -> None:
    """AHE 세 기둥을 순서대로 실행한다."""
    print("\n[AHE Evolve Loop] 시작")
    harness_dir   = SKILL_DIR / "harness"
    evolution_dir = SKILL_DIR / "evolution"
    traces_dir    = SKILL_DIR / "traces"
    evolution_dir.mkdir(exist_ok=True)
    traces_dir.mkdir(exist_ok=True)

    # ── ❷ 트레이스 수집 ──────────────────────────────────────
    trace = collect_trace(work_dir, topic)
    _save_json(traces_dir / f"{trace['run_id']}.json", trace)

    # ── Vision 이미지 분석 ────────────────────────────────────
    vision_result: dict | None = None
    image_paths = trace["qa"].get("image_paths", [])
    if image_paths:
        print(f"  [Vision] QA 이미지 {len(image_paths)}장 분석 중...")
        plan = _load_json(work_dir / "plan.json")
        vision_result = analyze_qa_images(image_paths, plan)
        n_issues = len(vision_result.get("slides", []))
        critical = sum(
            1 for s in vision_result.get("slides", [])
            for iss in s.get("issues", [])
            if iss.get("severity") in ("CRITICAL", "HIGH")
        )
        print(f"  ✓ Vision 분석 완료 — CRITICAL/HIGH 이슈: {critical}건")
        print(f"    요약: {vision_result.get('summary', '')[:80]}")
        _save_json(traces_dir / f"{trace['run_id']}_vision.json", vision_result)
    else:
        print("  ⚠ QA 이미지 없음 — Vision 분석 건너뜀")

    # ── digest (Vision 포함) ──────────────────────────────────
    digest = distill_digest(trace, vision_result)
    _save_json(traces_dir / f"{trace['run_id']}_digest.json", digest)
    print(f"  ✓ digest 저장 (issues={digest['issue_count']}, "
          f"vision_critical={len(digest['vision']['critical_issues'])}, "
          f"qa_images_exist={digest['qa_images_exist']})")

    # ── 이전 예측 검증 ──────────────────────────────────────
    verify_predictions(digest, evolution_dir)

    # ── ❸ Evolve Agent — harness 실제 편집 ─────────────────
    print("  [Evolve Agent] 하네스 분석 및 편집 중...")
    evolve_result = run_evolve_agent(digest, harness_dir, vision_result=vision_result)

    # 장기 기억 통계 업데이트 (Evolve Agent와 별도로 항상 실행)
    update_long_term_memory(digest, harness_dir)

    # ── manifest 저장 (편집 내역 + 예측) ────────────────────
    manifest = {
        "run_id":          trace["run_id"],
        "timestamp":       datetime.utcnow().isoformat(),
        "digest_summary":  {
            "success":       digest["success"],
            "issue_count":   digest["issue_count"],
            "qa_images_exist": digest["qa_images_exist"],
        },
        "actual_changes":  evolve_result["actual_changes"],
        "predictions":     evolve_result["predictions"],
    }
    manifest_path = evolution_dir / f"iteration_{trace['run_id']}_manifest.json"
    _save_json(manifest_path, manifest)
    print(f"  ✓ manifest 저장: {manifest_path.name}")

    # ── harness 변경이 있으면 git commit ────────────────────
    if evolve_result["actual_changes"]:
        changed = [c["file"] for c in evolve_result["actual_changes"]]
        git_commit_harness(
            SKILL_DIR,
            f"ahe: evolve {trace['run_id'][:12]} — edited {', '.join(changed)}",
        )
    else:
        print("  ✓ 하네스 변경 없음 (성공 실행 또는 이슈 없음)")

    print("[AHE Evolve Loop] 완료")
