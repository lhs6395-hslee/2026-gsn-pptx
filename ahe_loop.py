"""
AHE 세 기둥 구현:
  ❶ Component Observability — harness/ 파일들이 편집 대상
  ❷ Experience Observability — 트레이스 → digest 압축
  ❸ Decision Observability — Evolve Agent가 harness 편집 + manifest 예측 기록
"""
from __future__ import annotations

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent


def _load_json(path: Path) -> dict:
    if path.exists():
        return json.loads(path.read_text())
    return {}


def _save_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2))


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

    import base64, os

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


def _call_claude_vision(system: str, content: list) -> str | None:
    """멀티모달(Vision) Claude API 호출."""
    import os

    api_key      = os.environ.get("ANTHROPIC_API_KEY")
    vertex_proj  = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region   = os.environ.get("AWS_REGION")

    use_vertex  = bool(vertex_proj or (vertex_region and not aws_region))
    use_bedrock = bool(aws_region) and not use_vertex

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
            body = _json.dumps({
                "anthropic_version": "bedrock-2023-05-31",
                "max_tokens": 4096,
                "system": system,
                "messages": messages,
            })
            resp = client.invoke_model(
                modelId="us.anthropic.claude-sonnet-4-5-20250929-v1:0", body=body)
            return _json.loads(resp["body"].read())["content"][0]["text"].strip()

        import anthropic
        client = (
            anthropic.AnthropicVertex(project_id=vertex_proj or "", region=vertex_region or "us-east5")
            if use_vertex else anthropic.Anthropic(api_key=api_key)
        )
        resp = client.messages.create(
            model=model, max_tokens=4096,
            system=system, messages=messages,
        )
        return resp.content[0].text.strip()

    except Exception as e:
        print(f"  ⚠ Vision API 호출 실패: {e}")
        return None


def distill_digest(trace: dict, vision_result: dict | None = None) -> dict:
    """트레이스 + Vision 분석 결과 → 핵심 패턴 digest."""
    qa_ok = trace["qa"].get("image_count", 0) > 0
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
        "qa_ok":           qa_ok,
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
            "qa_done":           qa_ok,
            "slide_count_ok":    slide_match,
        },
    }


# ── ❸ Decision Observability — Evolve Agent ──────────────────

def _call_claude(system: str, user: str) -> str | None:
    """Claude API 호출 (Vertex > Bedrock > Anthropic 순 자동 감지)."""
    import os
    api_key      = os.environ.get("ANTHROPIC_API_KEY")
    vertex_proj  = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region   = os.environ.get("AWS_REGION")

    use_vertex  = bool(vertex_proj or (vertex_region and not aws_region))
    use_bedrock = bool(aws_region) and not use_vertex

    if not api_key and not use_vertex and not use_bedrock:
        return None

    model = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")
    messages = [{"role": "user", "content": user}]

    try:
        if use_bedrock:
            import boto3, json as _json
            client = boto3.client("bedrock-runtime", region_name=aws_region or "us-east-1")
            body = _json.dumps({
                "anthropic_version": "bedrock-2023-05-31",
                "max_tokens": 4096,
                "system": system,
                "messages": messages,
            })
            resp = client.invoke_model(
                modelId="us.anthropic.claude-sonnet-4-5-20250929-v1:0", body=body)
            return _json.loads(resp["body"].read())["content"][0]["text"].strip()

        import anthropic
        client = (
            anthropic.AnthropicVertex(project_id=vertex_proj or "", region=vertex_region or "us-east5")
            if use_vertex else anthropic.Anthropic(api_key=api_key)
        )
        resp = client.messages.create(
            model=model, max_tokens=4096,
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
    { "change": "변경 내용", "expected": "기대 효과", "metric": "검증 방법" }
  ]
}"""

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
          f"qa_ok={digest['qa_ok']})")

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
            "qa_ok":         digest["qa_ok"],
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
