"""
PPT 생성 로직: analyze_template → generate_plan → edit_slide 루프
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path


SKILL_DIR = Path.home() / ".ppt-skill"
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# ── Plan 콘텐츠 스키마 ────────────────────────────────────────
# role별 content 필드 요구사항 정의
PLAN_CONTENT_SCHEMA: dict[str, dict] = {
    "cover":     {"required": ["subtitle"], "optional": ["date"]},
    "toc":       {"required": ["items"]},
    "closing":   {"required": [], "optional": []},
    # 모든 본문 슬라이드 공통 필드:
    #   section_no: "1.1" 형식 번호
    #   section_title: 사이드바 본문제목 (20자 이내)
    #   section_desc: 사이드바 본문설명 (3줄 이내)
    #   body: 슬라이드 유형별 상세 내용
    "content":   {"required": ["section_no","section_title"], "optional": ["section_desc","body"]},
    "timeline":  {"required": ["section_no","section_title"], "recommended": ["periods"]},
    "quarterly": {"required": ["section_no","section_title"], "recommended": ["quarters"]},
    "steps":     {"required": ["section_no","section_title"], "recommended": ["steps"]},
    "flow":      {"required": ["section_no","section_title"], "recommended": ["keywords","solutions"]},
    "comparison":{"required": ["section_no","section_title"], "recommended": ["before","after"]},
}

# 슬라이드별 사이드바 shape ID 매핑
_SIDEBAR_LABEL_ID: dict[str, str] = {
    "slide13.xml": "14",  "slide15.xml": "14",
    "slide29.xml": "18",
    "slide30.xml": "17",  "slide31.xml": "17",
    "slide33.xml": "25",
    "slide35.xml": "22",  "slide36.xml": "17",
    "slide38.xml": "22",  "slide39.xml": "22",
    "slide21.xml": "16",  "slide22.xml": "14",
}
_SIDEBAR_DESC_ID: dict[str, str] = {
    "slide13.xml": "17",  "slide15.xml": "17",
    "slide29.xml": "19",
    "slide30.xml": "18",  "slide31.xml": "18",
    "slide33.xml": "26",
    "slide35.xml": "12",  "slide36.xml": "18",
    "slide38.xml": "12",  "slide39.xml": "12",
    "slide21.xml": "14",  "slide22.xml": "16",
}

# ── 레이아웃 레지스트리 ────────────────────────────────────────
# CLAUDE.md 카탈로그 기반: 역할별 최적 슬라이드 + 편집기 함수명
# shape 구조:
#   slide24: ID=8(제목), ID=7(본문1/bullets), ID=10(본문2/body)        — 순수 텍스트 2블록
#   slide30: ID=8(제목), ID=9(부제목), ID=28-31(Step1~4)               — 4단계 프로세스
#   slide32: ID=8(제목), ID=9(부제목), ID=26(본문/wide text)            — 텍스트+배너
#   slide38: ID=8(제목), ID=13-15(keyword×3), ID=7/10/11(solution×3)  — 3행 흐름도
#   slide8:  ID=2(섹션제목/40pt), ID=4(서브항목 목록)                    — 섹션 구분
LAYOUT_REGISTRY: dict[str, dict] = {
    "slide24.xml": {"label": "2블록 텍스트", "best_for": ["content", "body"],
                    "editor": "_edit_slide24"},
    "slide30.xml": {"label": "4단계 스텝",   "best_for": ["steps", "process"],
                    "editor": "_edit_slide30"},
    "slide32.xml": {"label": "텍스트+배너",  "best_for": ["content", "description"],
                    "editor": "_edit_slide32"},
    "slide38.xml": {"label": "3행 흐름도",   "best_for": ["flow", "architecture"],
                    "editor": "_edit_slide38"},
    "slide8.xml":  {"label": "섹션 구분",    "best_for": ["section"],
                    "editor": "_edit_slide8"},
    "slide6.xml":  {"label": "표지",         "best_for": ["cover"],
                    "editor": "edit_cover_slide"},
    "slide7.xml":  {"label": "목차",         "best_for": ["toc"],
                    "editor": "edit_toc_slide"},
    "slide46.xml": {"label": "감사합니다",   "best_for": ["closing"],
                    "editor": "noop"},
}


# ── 1. 템플릿 분석 ────────────────────────────────────────────

def analyze_template(work_dir: Path) -> list[dict]:
    """언팩된 슬라이드 XML을 읽어 슬라이드 목록과 텍스트를 반환한다."""
    unpacked = work_dir / "unpacked"
    slides_dir = unpacked / "ppt" / "slides"
    if not slides_dir.exists():
        raise FileNotFoundError(f"unpacked slides not found: {slides_dir}")

    slides = []
    for xml_path in sorted(slides_dir.glob("slide*.xml")):
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
        except ET.ParseError as e:
            print(f"  ⚠ {xml_path.name} parse error: {e}", file=sys.stderr)
            continue

        texts: list[str] = []
        for t_elem in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
            if t_elem.text and t_elem.text.strip():
                texts.append(t_elem.text.strip())

        slides.append({
            "file": xml_path.name,
            "path": str(xml_path),
            "texts": texts,
        })

    return slides


# ── 2. 계획 수립 ──────────────────────────────────────────────

def generate_plan_with_claude(
    topic: str,
    audience: str,
    n_slides: int,
    slide_info: list[dict],
    memory: dict,
) -> dict:
    """
    Claude API를 사용해 주제에 맞는 실제 콘텐츠 계획을 생성한다.
    ANTHROPIC_API_KEY 없으면 None을 반환해 폴백으로 넘긴다.
    """
    import os
    import json as _json

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    vertex_project = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region = os.environ.get("AWS_REGION")
    explicit = os.environ.get("PPT_SKILL_BACKEND", "auto")  # auto|vertex|bedrock|anthropic

    # 백엔드 결정 — 명시 우선, 그 다음 자동 감지
    if explicit == "vertex":
        use_vertex, use_bedrock = True, False
    elif explicit == "bedrock":
        use_vertex, use_bedrock = False, True
    elif explicit == "anthropic":
        use_vertex, use_bedrock = False, False
    else:
        # auto: Vertex > Bedrock > Anthropic
        # CLOUD_ML_REGION과 AWS_REGION 둘 다 있으면 ANTHROPIC_VERTEX_PROJECT_ID 유무로 판단
        use_vertex = bool(vertex_project or (vertex_region and not aws_region))
        use_bedrock = bool(aws_region) and not use_vertex

    if not api_key and not use_vertex and not use_bedrock:
        return None

    available_layouts = {s["file"]: s["texts"][:3] for s in slide_info}
    layout_hints = memory.get("slide_layout_hints", {})
    known_issues = []
    for tmpl_data in memory.get("template_patterns", {}).values():
        known_issues = [i["issue"] for i in tmpl_data.get("known_issues", [])]

    system = (
        "당신은 프로 컨설턴트 수준의 PowerPoint 슬라이드 기획자입니다.\n"
        "콘텐츠를 먼저 설계하고, 그 콘텐츠에 가장 적합한 template_file을 선택합니다.\n"
        "반드시 JSON만 출력하고 다른 텍스트는 포함하지 않습니다.\n\n"
        "=== 사용 가능한 template_file 카탈로그 ===\n"
        "slide6.xml   → cover    : 표지 (전용, 변경 불가)\n"
        "slide7.xml   → toc      : 목차 (전용, 변경 불가)\n"
        "slide8.xml   → section  : 챕터 구분 (사용자 명시 요청 시에만 — 자동 선택 금지)\n"
        "slide29.xml  → timeline : 연도별 타임라인(2023→2026) 또는 월별(2026.1→2026.6)\n"
        "slide31.xml  → quarterly: 분기별 Q1→Q2→Q3→Q4 레이아웃\n"
        "slide33.xml  → quarterly: slide31 변형 — 상단 설명 + Q1-Q4열\n"
        "slide13.xml  → content  : 3가지 기능/특징/장점 카드 (제목+설명 3열)\n"
        "slide15.xml  → content  : 3가지 핵심 가치 (제목+설명 3열)\n"
        "slide14.xml  → content  : 4가지 기능/구성요소 (제목+설명+Insight 4열)\n"
        "slide16.xml  → content  : 4가지 구성요소 컴팩트 (제목+설명 4열)\n"
        "slide9.xml   → content  : 3가지 사례/제품 (이미지+제목+설명 3열)\n"
        "slide10.xml  → content  : 3가지 사례 + 하단 Insight 배너\n"
        "slide12.xml  → content  : 3가지 도구/기술 (이미지+우측 텍스트)\n"
        "slide30.xml  → steps    : 4단계 프로세스 (Step1→Step2→Step3→Step4)\n"
        "slide32.xml  → content  : 텍스트+배너 (순수 텍스트, 긴 설명/개요/서술형/2~3가지 요점)\n"
        "slide36.xml  → content  : As-is/To-be 벤다이어그램 (현황→목표 비교)\n"
        "slide38.xml  → flow     : 3행 흐름도 keyword→solution→service (아키텍처/파이프라인)\n"
        "slide46.xml  → closing  : 감사합니다 (전용, 변경 불가)\n\n"
        "=== template 선택 기준 (반드시 콘텐츠 형태와 일치시킬 것) ===\n"
        "- 서술형 설명·시장동향·개요·배경 등 줄글 → slide32 (절대 타임라인 금지)\n"
        "- 3가지 기능/장점 나열 → slide13 (items 3개 + descriptions 3개 필수)\n"
        "- 4가지 기능/구성요소 → slide14 또는 slide16 (items 4개 + descriptions 4개)\n"
        "- 3가지 사례·제품·도구(시각 강조) → slide9/slide10/slide12 "
        "(items + descriptions + image_descriptions 3개)\n"
        "- 4단계 프로세스 → slide30 (steps 4개 필수)\n"
        "- 아키텍처/파이프라인 흐름 → slide38 (keywords + solutions + services/details 필수)\n"
        "- 현황↔목표·이전↔이후·비교 → slide36 (as_is/to_be 각 키워드 필수)\n"
        "- 짧은 2~3가지 요점 → slide32 (body + bullets)\n"
        "- 연도별 로드맵/연혁 → slide29 (periods=[{label,content}] 필수, 시계열일 때만)\n"
        "- 분기별 계획 → slide31/slide33 (quarters 필수, 시계열일 때만)\n"
        "- 챕터 시작 구분 → slide8 (사용자가 '섹션 구분' 명시 요청 시에만)\n"
        "★ 중요: 실사진/아이콘 이미지는 제공되지 않는다. 이미지가 필수인 레이아웃은 쓰지 말 것.\n"
        "★ 중요: slide29/31/33(타임라인)은 연도·분기 등 '시간 흐름' 데이터일 때만. "
        "그 외 모든 설명은 slide32를 기본으로 사용.\n"
        "★ 레이아웃별 필수 필드가 없으면 자동으로 slide32(텍스트)로 교체되니, "
        "선택한 레이아웃에 맞는 content 필드를 반드시 채울 것.\n\n"
        "=== role별 content 필드 ===\n"
        "cover  : subtitle(30자 이내 1줄), date\n"
        "toc    : items(문자열 배열)\n"
        "steps  : section_no(예 '4'), section_title(사이드바 제목 2줄 이내), "
        "section_desc(사이드바 설명 3줄 이내), "
        "steps(4개, 각 25자 이내 짧은 한 줄 — 예 '1단계: PoC 검증'. 긴 설명 금지, 잘림)\n"
        "flow   : keywords(3개), solutions(3개), details(3개, 선택)\n"
        "content(slide13/15): items(3개, 각 16자 이내), descriptions(3개, 각 45자 이내 — 길면 잘림)\n"
        "content(slide14/16): items(4개, 16자 이내), descriptions(4개, 40자 이내)\n"
        "content(slide9/10/12): items(제목 3개), descriptions(설명 3개), "
        "image_descriptions(각 칸에 '어떤 이미지를 넣어야 하는지' 구체적 설명 3개 — 실제 이미지 아님)\n"
        "content(slide10): items/descriptions/image_descriptions(3개) + insight(하단 배너 1개)\n"
        "content(slide35): before(3개 키워드), after(4개 키워드)\n"
        "content(slide36): as_is(4개 키워드), to_be(4개 키워드), body\n"
        "content(slide32): body(긴 텍스트), bullets(선택)\n"
        "closing: 없음\n\n"
        "출력 형식:\n"
        '{"title":"...","topic":"...","audience":"...","n_slides":N,'
        '"slides":[{"index":1,"template_file":"slideN.xml",'
        '"role":"cover|toc|section|content|steps|flow|closing",'
        '"title":"...","content":{...}}]}'
    )

    layout_summary = "\n".join(f"  {f}: {t}" for f, t in list(available_layouts.items())[:15])
    hints_summary = "\n".join(f"  {k}: {v}" for k, v in layout_hints.items())

    user_prompt = (
        f"주제: {topic}\n"
        f"대상 청중: {audience}\n"
        f"슬라이드 수: {n_slides}\n\n"
        f"사용 가능한 슬라이드 파일 (파일명: 샘플 텍스트):\n{layout_summary}\n\n"
        f"레이아웃 힌트:\n{hints_summary}\n\n"
        f"주의할 이슈: {', '.join(known_issues) if known_issues else '없음'}\n\n"
        f"위 조건에 맞는 {n_slides}장 슬라이드 계획을 JSON으로 작성하세요. "
        f"각 슬라이드의 콘텐츠는 {topic} 전문 지식을 반영해 구체적으로 작성하세요."
    )

    messages = [{"role": "user", "content": user_prompt}]
    model = os.environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")

    if use_bedrock:
        # project-steer 패턴: boto3 직접 사용
        import boto3 as _boto3
        bedrock = _boto3.client("bedrock-runtime", region_name=aws_region or "us-east-1")
        bedrock_model = model
        if not (bedrock_model.startswith("us.anthropic.") or bedrock_model.startswith("anthropic.")):
            bedrock_model = "us.anthropic.claude-sonnet-4-5-20250929-v1:0"
        body = _json.dumps({
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": 4096,
            "system": system,
            "messages": messages,
        })
        resp = bedrock.invoke_model(modelId=bedrock_model, body=body)
        raw = _json.loads(resp["body"].read())["content"][0]["text"].strip()
    else:
        try:
            import anthropic
        except ImportError:
            return None

        if use_vertex:
            client = anthropic.AnthropicVertex(
                project_id=vertex_project or "",
                region=vertex_region or "us-east5",
            )
        else:
            client = anthropic.Anthropic(api_key=api_key)

        response = client.messages.create(
            model=model,
            max_tokens=4096,
            system=system,
            messages=messages,
        )
        raw = response.content[0].text.strip()

    # JSON 블록 추출 (```json ... ``` 형식 대응)
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)

    plan = _json.loads(raw)
    plan["_generated_by"] = "vertex" if use_vertex else ("bedrock" if use_bedrock else "anthropic")
    return plan


def generate_plan(topic: str, audience: str, n_slides: int, slide_info: list[dict]) -> dict:
    """
    슬라이드 계획(plan.json)을 생성한다.
    Claude Code 스킬 모드에서는 Claude 자신이 이 함수를 대체하지만,
    폴백 모드에서는 단순 규칙 기반 계획을 만든다.
    """
    available = [s["file"] for s in slide_info]

    # 사용 가능한 슬라이드 파일을 역할에 매핑 (GS Neotek 템플릿 기준)
    role_map: dict[str, str] = {}
    for f in available:
        name = Path(f).stem  # e.g. "slide6"
        idx = int(re.search(r"\d+", name).group())
        if idx == 6:
            role_map["cover"] = f
        elif idx == 7:
            role_map["toc"] = f
        elif idx == 8:
            role_map["section"] = f
        elif idx == 14:
            role_map["three_col"] = f
        elif idx == 30:
            role_map["steps"] = f
        elif idx == 46:
            role_map["closing"] = f

    def pick(preferred: str, fallback_idx: int) -> str:
        if preferred in role_map:
            return role_map[preferred]
        if available:
            return available[min(fallback_idx, len(available) - 1)]
        return "slide1.xml"

    slides_plan: list[dict] = []

    # 표지
    slides_plan.append({
        "index": 1,
        "template_file": pick("cover", 0),
        "role": "cover",
        "title": topic,
        "content": {"subtitle": f"대상: {audience}"},
    })

    # 목차
    if n_slides >= 3:
        slides_plan.append({
            "index": 2,
            "template_file": pick("toc", 1),
            "role": "toc",
            "title": "목차",
            "content": {"items": [f"Chapter {i}" for i in range(1, min(n_slides - 1, 7) + 1)]},
        })

    # 본문 슬라이드
    for i in range(3, n_slides):
        slides_plan.append({
            "index": i,
            "template_file": pick("three_col", 2),
            "role": "content",
            "title": f"{topic} — 주요 내용 {i - 2}",
            "content": {"body": f"{topic}에 대한 심층 내용 {i - 2}"},
        })

    # 마지막: 감사합니다
    slides_plan.append({
        "index": n_slides,
        "template_file": pick("closing", -1),
        "role": "closing",
        "title": "감사합니다",
        "content": {},
    })

    return {
        "title": topic,
        "topic": topic,
        "audience": audience,
        "n_slides": n_slides,
        "slides": slides_plan,
    }


# ── 3. 슬라이드 XML 편집 ─────────────────────────────────────

_PLACEHOLDER_RE = re.compile(
    r"lorem\s+ipsum|작성해주세요|TODO|\[insert|placeholder",
    re.IGNORECASE,
)

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# ── 슬라이드 존 맵 (analyze_zones.py로 템플릿에서 자동 생성) ──────────
# 각 본문 슬라이드의 5개 존 + 본문구역 하위 존(shape ID 리스트)을 정의.
# 단일 진실 공급원(single source of truth) — 편집기는 하드코딩 ID 대신 이걸 참조.
_ZONE_MAP_CACHE: dict | None = None


def _load_zone_map() -> dict:
    """harness/layout_zone_map.json 로드 (캐시). 없으면 빈 dict."""
    global _ZONE_MAP_CACHE
    if _ZONE_MAP_CACHE is not None:
        return _ZONE_MAP_CACHE
    for p in (SKILL_DIR / "harness" / "layout_zone_map.json",
              Path(__file__).parent / "harness" / "layout_zone_map.json"):
        try:
            if p.exists():
                _ZONE_MAP_CACHE = json.loads(p.read_text())
                return _ZONE_MAP_CACHE
        except Exception:
            pass
    _ZONE_MAP_CACHE = {}
    return _ZONE_MAP_CACHE


def _zone(template_file: str) -> dict:
    """주어진 slideN.xml의 존 정의 반환 (base 이름으로도 조회)."""
    zm = _load_zone_map()
    if template_file in zm:
        return zm[template_file]
    base = re.sub(r"_c\d+\.xml$", ".xml", template_file)  # 중복 복사본(slide38_c1.xml) 대응
    return zm.get(base, {})

# 표지 레이아웃 상수 (GS Neotek 2026 템플릿 기준, EMU 단위)
_COVER_TITLE_ID = "12"       # 대제목 shape id
_COVER_SUBTITLE_ID = "6"     # 소제목 shape id
_COVER_DATE_ID = "15"        # 날짜 shape id
_COVER_TITLE_FONT_PT = 48
_COVER_SUBTITLE_FONT_PT = 32
_EMU_PER_PT = 12700
# 가시성 최적 자간 권장값 (OOXML spcPct):
#   대제목: 100% (lnSpc) — 빽빽하지 않고 깔끔한 밀도
#   소제목: 100% (lnSpc) — 단일 줄이므로 줄간격 자체 의미 없으나 일관성 유지
_COVER_TITLE_LNSPC = 100000   # 100 % × 1000
# 동적 여백 — 타이포그래피 권장: 대제목 폰트의 0.5×, 0.4× 배율
_GAP_DATE_TITLE_RATIO = 0.5   # date → title 간격 = 0.5 × title_font_height
_GAP_TITLE_SUBTITLE_RATIO = 0.4  # title → subtitle 간격


def _make_cover_run_xml(text: str, font_pt: int, font_name: str) -> str:
    """한국어 혼합 텍스트용 단일 run XML 조각 반환 (latin + ea + cs 모두 지정)."""
    sz = font_pt * 100  # hundredths of pt
    # lang=ko-KR 으로 설정해야 LibreOffice가 ea 폰트로 한국어를 올바르게 렌더링
    return (
        f'<a:r xmlns:a="{_NS_A}">'
        f'<a:rPr lang="ko-KR" altLang="en-US" sz="{sz}" dirty="0">'
        f'<a:latin typeface="{font_name}" pitchFamily="2" charset="-127"/>'
        f'<a:ea typeface="{font_name}" pitchFamily="2" charset="-127"/>'
        f'<a:cs typeface="{font_name}" pitchFamily="2" charset="-127"/>'
        f'</a:rPr>'
        f'<a:t>{text}</a:t>'
        f'</a:r>'
    )


def _text_width_emu(text: str, font_pt: int) -> int:
    """
    한글/영문 혼합 텍스트의 렌더 폭을 EMU로 추정한다.
    - 한글(Hangul, CJK): font_pt × 0.85 (거의 전각)
    - 영문/숫자/기호: font_pt × 0.52 (절반 정도)
    - 공백: font_pt × 0.28
    Pretendard 폰트 기준 경험치.
    """
    width_pt = 0.0
    for ch in text:
        cp = ord(ch)
        if (0xAC00 <= cp <= 0xD7A3   # 한글 음절
                or 0x3131 <= cp <= 0x318E   # 한글 자모
                or 0x4E00 <= cp <= 0x9FFF): # CJK
            width_pt += font_pt * 0.85
        elif ch == ' ':
            width_pt += font_pt * 0.28
        else:
            width_pt += font_pt * 0.52
    return int(width_pt * _EMU_PER_PT)


def _estimate_lines(text: str, cx_emu: int, font_pt: int) -> int:
    """텍스트가 실제로 몇 줄로 렌더될지 추정."""
    if not text:
        return 1
    line_width = 0
    lines = 1
    for ch in text:
        cp = ord(ch)
        if (0xAC00 <= cp <= 0xD7A3 or 0x3131 <= cp <= 0x318E or 0x4E00 <= cp <= 0x9FFF):
            cw = int(font_pt * 0.85 * _EMU_PER_PT)
        elif ch == ' ':
            cw = int(font_pt * 0.28 * _EMU_PER_PT)
        else:
            cw = int(font_pt * 0.52 * _EMU_PER_PT)
        if line_width + cw > cx_emu:
            lines += 1
            line_width = cw
        else:
            line_width += cw
    return lines


def _truncate_to_lines(text: str, cx_emu: int, font_pt: int, max_lines: int) -> str:
    """
    실제 렌더 폭 기반으로 max_lines 초과 시 말줄임표로 자른다.
    한글/영문 혼합 텍스트의 문자별 폭을 개별 계산해 정확하게 처리.
    """
    if not text:
        return text
    ellipsis_w = int(font_pt * 0.52 * _EMU_PER_PT)  # "…" 폭
    line_width = 0
    lines = 1
    for i, ch in enumerate(text):
        cp = ord(ch)
        if (0xAC00 <= cp <= 0xD7A3 or 0x3131 <= cp <= 0x318E or 0x4E00 <= cp <= 0x9FFF):
            cw = int(font_pt * 0.85 * _EMU_PER_PT)
        elif ch == ' ':
            cw = int(font_pt * 0.28 * _EMU_PER_PT)
        else:
            cw = int(font_pt * 0.52 * _EMU_PER_PT)

        if line_width + cw > cx_emu:
            lines += 1
            if lines > max_lines:
                # 현재 위치에서 잘라야 함 — 말줄임표 공간 확보
                cut = i
                # 앞으로 당겨서 "…" 넣을 자리 확보
                while cut > 0 and line_width > cx_emu - ellipsis_w:
                    cut -= 1
                    pc = text[cut]
                    pcp = ord(pc)
                    if (0xAC00 <= pcp <= 0xD7A3 or 0x3131 <= pcp <= 0x318E
                            or 0x4E00 <= pcp <= 0x9FFF):
                        line_width -= int(font_pt * 0.85 * _EMU_PER_PT)
                    elif pc == ' ':
                        line_width -= int(font_pt * 0.28 * _EMU_PER_PT)
                    else:
                        line_width -= int(font_pt * 0.52 * _EMU_PER_PT)
                return text[:cut] + "…"
            line_width = cw
        else:
            line_width += cw
    return text


def _write_xml(root: ET.Element, xml_path: Path) -> None:
    """네임스페이스 처리를 안전하게 유지하며 XML을 저장한다."""
    xml_str = ET.tostring(root, encoding="unicode")
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write(xml_str)


def _find_shape_by_id(root: ET.Element, shape_id: str):
    for sp in root.iter(f"{{{_NS_P}}}sp"):
        cpr = sp.find(f"{{{_NS_P}}}nvSpPr/{{{_NS_P}}}cNvPr")
        if cpr is not None and cpr.get("id") == shape_id:
            return sp
    return None


def _set_shape_text_single_run(sp: ET.Element, text: str, font_pt: int, font_name: str,
                                lnspc_pct: int = None) -> None:
    """
    shape의 기존 rPr(색상·효과 포함)을 보존하면서 텍스트만 교체한다.
    - 첫 번째 run의 rPr을 복사해 lang만 ko-KR로 변경
    - 나머지 run 제거, lnSpc 추가
    """
    ns_a = _NS_A
    ns_p = _NS_P
    txBody = sp.find(f"{{{ns_p}}}txBody")
    if txBody is None:
        return

    # 첫 번째 p의 첫 번째 r에서 rPr 복사 (색상·폰트 스타일 보존용)
    original_rPr = None
    for p in txBody.findall(f"{{{ns_a}}}p"):
        r = p.find(f"{{{ns_a}}}r")
        if r is not None:
            rPr_elem = r.find(f"{{{ns_a}}}rPr")
            if rPr_elem is not None:
                import copy
                original_rPr = copy.deepcopy(rPr_elem)
            break

    # 기존 <a:p> 모두 제거
    for p in txBody.findall(f"{{{ns_a}}}p"):
        txBody.remove(p)

    # 새 <a:p> 구성
    p_elem = ET.SubElement(txBody, f"{{{ns_a}}}p")

    # pPr (줄간격)
    pPr = ET.SubElement(p_elem, f"{{{ns_a}}}pPr")
    if lnspc_pct is not None:
        lnSpc = ET.SubElement(pPr, f"{{{ns_a}}}lnSpc")
        ET.SubElement(lnSpc, f"{{{ns_a}}}spcPct", val=str(lnspc_pct))

    # run — 원본 rPr 보존 + lang만 ko-KR로 변경
    r_elem = ET.SubElement(p_elem, f"{{{ns_a}}}r")
    if original_rPr is not None:
        original_rPr.set("lang", "ko-KR")
        original_rPr.set("altLang", "en-US")
        original_rPr.set("dirty", "0")
        r_elem.append(original_rPr)
    else:
        # 원본 없을 때만 새로 생성 (폴백)
        rPr = ET.SubElement(r_elem, f"{{{ns_a}}}rPr",
                            lang="ko-KR", altLang="en-US",
                            sz=str(font_pt * 100), dirty="0")
        ET.SubElement(rPr, f"{{{ns_a}}}latin",
                      typeface=font_name, pitchFamily="2", charset="-127")
        ET.SubElement(rPr, f"{{{ns_a}}}ea",
                      typeface=font_name, pitchFamily="2", charset="-127")
        ET.SubElement(rPr, f"{{{ns_a}}}cs",
                      typeface=font_name, pitchFamily="2", charset="-127")

    t_elem = ET.SubElement(r_elem, f"{{{ns_a}}}t")
    t_elem.text = text


def _set_shape_cy(sp: ET.Element, cy: int) -> None:
    """shape의 xfrm ext cy 값을 교체한다."""
    xfrm = sp.find(f".//{{{_NS_A}}}xfrm")
    if xfrm is None:
        return
    ext = xfrm.find(f"{{{_NS_A}}}ext")
    if ext is not None:
        ext.set("cy", str(cy))


def _set_shape_cx(sp: ET.Element, cx: int) -> None:
    """shape의 xfrm ext cx 값을 교체한다."""
    xfrm = sp.find(f".//{{{_NS_A}}}xfrm")
    if xfrm is None:
        return
    ext = xfrm.find(f"{{{_NS_A}}}ext")
    if ext is not None:
        ext.set("cx", str(cx))


def _set_shape_y(sp: ET.Element, y: int) -> None:
    """shape의 xfrm off y 값을 교체한다."""
    xfrm = sp.find(f".//{{{_NS_A}}}xfrm")
    if xfrm is None:
        return
    off = xfrm.find(f"{{{_NS_A}}}off")
    if off is not None:
        off.set("y", str(y))


def _get_shape_geometry(sp: ET.Element) -> tuple[int, int, int, int]:
    """(x, y, cx, cy) 반환. 없으면 (0,0,0,0)."""
    xfrm = sp.find(f".//{{{_NS_A}}}xfrm")
    if xfrm is None:
        return (0, 0, 0, 0)
    off = xfrm.find(f"{{{_NS_A}}}off")
    ext = xfrm.find(f"{{{_NS_A}}}ext")
    x = int(off.get("x", 0)) if off is not None else 0
    y = int(off.get("y", 0)) if off is not None else 0
    cx = int(ext.get("cx", 0)) if ext is not None else 0
    cy = int(ext.get("cy", 0)) if ext is not None else 0
    return (x, y, cx, cy)


def _enable_autofit(sp: ET.Element) -> None:
    """bodyPr의 noAutofit을 normAutofit으로 전환한다."""
    ns_a = _NS_A
    bodyPr = sp.find(f".//{{{ns_a}}}bodyPr")
    if bodyPr is None:
        return
    for child in list(bodyPr):
        tag = child.tag.split("}")[-1]
        if tag in ("noAutofit", "spAutoFit", "normAutofit"):
            bodyPr.remove(child)
    ET.SubElement(bodyPr, f"{{{ns_a}}}normAutofit")


def edit_cover_slide(xml_path: Path, title: str, subtitle: str, date_str: str) -> None:
    """
    표지 슬라이드(slide6) 전용 편집:
    - 대제목 (shape 12): Pretendard SemiBold 48pt, lnSpc=100%, 최대 2줄
    - 소제목 (shape 6 ): Pretendard 32pt, 1줄, 대제목과 같은 cx 폭
    - 날짜   (shape 15): 그대로 삽입
    - 날짜→대제목→소제목 Y 위치를 동적으로 재배치

    타이포그래피 권장값 (발표 자료 가시성 기준):
      - 줄간격 1.0× (100%): 대형 제목은 1.0~1.15가 최적 — 지나치게 넓으면 연결감 소실
      - Date→Title 간격: title_font_height × 0.5 (여백이 확보되면서 묶임감 유지)
      - Title→Subtitle 간격: title_font_height × 0.4 (소제목은 제목에 가깝게)
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    sp_title = _find_shape_by_id(root, _COVER_TITLE_ID)
    sp_sub = _find_shape_by_id(root, _COVER_SUBTITLE_ID)
    sp_date = _find_shape_by_id(root, _COVER_DATE_ID)

    if sp_title is None or sp_sub is None or sp_date is None:
        # shape ID가 맞지 않는 템플릿 → 폴백: 첫 번째 텍스트만 교체
        _replace_first_text_legacy(xml_path, title)
        return

    # ── 대제목 cx 가져오기 (소제목 동기화 기준) ──────────────────
    _, date_y, _, date_cy = _get_shape_geometry(sp_date)
    _, _, title_cx, _ = _get_shape_geometry(sp_title)

    # ── 1. 날짜 ────────────────────────────────────────────────
    _set_shape_text_single_run(sp_date, date_str,
                               font_pt=14, font_name="Pretendard SemiBold")

    # ── 2. 대제목 ───────────────────────────────────────────────
    title_text = _truncate_to_lines(title, title_cx, _COVER_TITLE_FONT_PT, max_lines=2)
    title_lines = _estimate_lines(title_text, title_cx, _COVER_TITLE_FONT_PT)
    title_cy = title_lines * _COVER_TITLE_FONT_PT * _EMU_PER_PT
    _enable_autofit(sp_title)
    _set_shape_cy(sp_title, title_cy)
    _set_shape_text_single_run(sp_title, title_text,
                               font_pt=_COVER_TITLE_FONT_PT,
                               font_name="Pretendard SemiBold",
                               lnspc_pct=_COVER_TITLE_LNSPC)

    # ── 3. 소제목 ───────────────────────────────────────────────
    # 정확한 렌더 폭으로 1줄 강제 잘라냄
    sub_text = _truncate_to_lines(subtitle, title_cx, _COVER_SUBTITLE_FONT_PT, max_lines=1)
    sub_cy = _COVER_SUBTITLE_FONT_PT * _EMU_PER_PT
    _set_shape_cx(sp_sub, title_cx)
    _set_shape_cy(sp_sub, sub_cy)
    # wrap="none" + noAutofit: 잘림 방지 (텍스트가 박스 밖으로 보이지 않게)
    ns_a = _NS_A
    bodyPr_sub = sp_sub.find(f".//{{{ns_a}}}bodyPr")
    if bodyPr_sub is not None:
        bodyPr_sub.set("wrap", "none")
        for child in list(bodyPr_sub):
            if child.tag.split("}")[-1] in ("spAutoFit", "normAutofit", "noAutofit"):
                bodyPr_sub.remove(child)
        ET.SubElement(bodyPr_sub, f"{{{ns_a}}}noAutofit")
    _set_shape_text_single_run(sp_sub, sub_text,
                               font_pt=_COVER_SUBTITLE_FONT_PT,
                               font_name="Pretendard",
                               lnspc_pct=_COVER_TITLE_LNSPC)

    # ── 4. 동적 Y 위치 재배치 ────────────────────────────────────
    # 기준: 날짜 위치는 고정, 아래로 title→subtitle 배치
    title_font_height = _COVER_TITLE_FONT_PT * _EMU_PER_PT
    gap_date_title = int(title_font_height * _GAP_DATE_TITLE_RATIO)
    gap_title_sub = int(title_font_height * _GAP_TITLE_SUBTITLE_RATIO)

    title_y = date_y + date_cy + gap_date_title
    sub_y = title_y + title_cy + gap_title_sub

    _set_shape_y(sp_title, title_y)
    _set_shape_y(sp_sub, sub_y)

    _write_xml(root, xml_path)


def _replace_first_text_legacy(xml_path: Path, new_text: str) -> None:
    """폴백: 첫 번째 <a:t>만 교체 (cover 전용 shape ID가 없는 구형 템플릿용)."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns_a = _NS_A
    t_elems = list(root.iter(f"{{{ns_a}}}t"))
    if not t_elems:
        return
    t_elems[0].text = new_text
    for t in t_elems[1:]:
        t.text = ""
    _write_xml(root, xml_path)


def _replace_first_text(xml_path: Path, new_text: str) -> None:
    """슬라이드 XML의 첫 번째 텍스트 run을 new_text로 교체하고 나머지 run은 비운다."""
    _replace_first_text_legacy(xml_path, new_text)


def _toc_row_xml(row_idx: int, row_y: int, number: str, item_text: str, page_num: str,
                 row_h: int = 674674, line_offset: int = 431324) -> str:
    """
    목차 슬라이드의 한 행(Row)을 구성하는 4개 shape XML 반환.
    번호·텍스트·선·페이지가 동일 Y를 공유하여 항상 정렬됨.
    wrap=none + noAutofit으로 자동줄바꿈 완전 차단.
    """
    y = row_y + row_idx * row_h
    line_y = y + line_offset
    base_id = 200 + row_idx * 4

    # 색상/폰트는 원본 템플릿 그대로 사용
    # 번호: bg2 lumMod=75000 / 텍스트: tx1 lumMod=95000,lumOff=5000 / 페이지: bg2 lumMod=75000
    NUM_X, NUM_CX = 3851538, 536801
    TXT_X, TXT_CX = 4541178, 4383747   # 선 시작점(8924925)까지 확장
    LINE_X, LINE_CX = 8924925, 3267075
    PAGE_X, PAGE_CX = 11431475, 760525

    # 비어있는 행이면 빈 텍스트로 (선도 투명하게)
    is_empty = not number and not item_text and not page_num
    line_alpha = "0" if is_empty else "100000"

    return f"""
  <p:sp>
    <p:nvSpPr><p:cNvPr id="{base_id}" name="toc_num_{row_idx}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="{NUM_X}" y="{y}"/><a:ext cx="{NUM_CX}" cy="{row_h}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="ko-KR" altLang="en-US" sz="3000" b="1" dirty="0">
            <a:solidFill><a:schemeClr val="bg2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
            <a:latin typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
            <a:ea typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
            <a:cs typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
          </a:rPr>
          <a:t>{number}</a:t>
        </a:r>
      </a:p>
    </p:txBody>
  </p:sp>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="{base_id+1}" name="toc_text_{row_idx}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="{TXT_X}" y="{y}"/><a:ext cx="{TXT_CX}" cy="{row_h}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:r>
          <a:rPr lang="ko-KR" altLang="en-US" sz="3000" dirty="0">
            <a:solidFill><a:schemeClr val="tx1"><a:lumMod val="95000"/><a:lumOff val="5000"/><a:alpha val="99000"/></a:schemeClr></a:solidFill>
            <a:latin typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
            <a:ea typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
            <a:cs typeface="Pretendard SemiBold" panose="02000703000000020004" pitchFamily="2" charset="-127"/>
          </a:rPr>
          <a:t>{item_text}</a:t>
        </a:r>
      </a:p>
    </p:txBody>
  </p:sp>
  <p:cxnSp>
    <p:nvCxnSpPr><p:cNvPr id="{base_id+2}" name="toc_line_{row_idx}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
    <p:spPr><a:xfrm><a:off x="{LINE_X}" y="{line_y}"/><a:ext cx="{LINE_CX}" cy="0"/></a:xfrm>
      <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
      <a:ln><a:solidFill><a:schemeClr val="bg2"><a:lumMod val="90000"/><a:alpha val="{line_alpha}"/></a:schemeClr></a:solidFill></a:ln>
    </p:spPr>
  </p:cxnSp>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="{base_id+3}" name="toc_page_{row_idx}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="{PAGE_X}" y="{y}"/><a:ext cx="{PAGE_CX}" cy="{row_h}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"><a:noAutofit/></a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r>
          <a:rPr lang="ko-KR" altLang="en-US" sz="1400" dirty="0">
            <a:solidFill><a:schemeClr val="bg2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
            <a:latin typeface="Pretendard" panose="02000503000000020004" pitchFamily="2" charset="-127"/>
            <a:ea typeface="Pretendard" panose="02000503000000020004" pitchFamily="2" charset="-127"/>
            <a:cs typeface="Pretendard" panose="02000503000000020004" pitchFamily="2" charset="-127"/>
          </a:rPr>
          <a:t>{page_num}</a:t>
        </a:r>
      </a:p>
    </p:txBody>
  </p:sp>"""


def edit_toc_slide(xml_path: Path, prs_title: str, items: list[str],
                   page_nums: list[str]) -> None:
    """
    목차 슬라이드(slide7) 편집.
    기존 shape 구조를 그대로 유지하면서 내용만 교체한다:
    - ID=10 (항목 텍스트): wrap=none + noAutofit, 각 paragraph에 항목 삽입
    - ID=11 (번호):        wrap=none + noAutofit, N개 초과분 비움
    - ID=14 (페이지번호):  wrap=none + noAutofit, N개 초과분 비움
    - ID=12~41 (선 7개):  N개 초과 행의 선은 투명(alpha=0) 처리
    - ID=42 (제목 ph):    제목 텍스트 교체
    원본 shape를 유지해야 슬라이드 레이아웃 placeholder가 노출되지 않는다.
    """
    ns_a = _NS_A
    ns_p = _NS_P
    import copy

    tree = ET.parse(xml_path)
    root = tree.getroot()

    MAX_ROWS = 7
    items_padded   = list(items)    + [""] * (MAX_ROWS - len(items))
    pages_padded   = list(page_nums) + [""] * (MAX_ROWS - len(page_nums))

    # ── 헬퍼: 텍스트박스 너비(cx) 추출 ──────────────────────────────
    def get_cx(sp: ET.Element) -> int:
        xfrm = sp.find(f".//{{{ns_a}}}xfrm")
        if xfrm is None:
            return 0
        ext = xfrm.find(f"{{{ns_a}}}ext")
        return int(ext.get("cx", 0)) if ext is not None else 0

    # ── 헬퍼: paragraph 내 모든 <a:t>를 합쳐 하나의 run으로 정리 ──
    def set_para_text(para: ET.Element, text: str) -> None:
        """기존 첫 번째 run의 rPr를 보존하고 텍스트만 교체, 나머지 run 제거.
        새 run은 endParaRPr 앞에 삽입 — 뒤에 넣으면 PowerPoint가 무시함."""
        # 첫 run rPr 복사
        first_r = para.find(f"{{{ns_a}}}r")
        orig_rPr = None
        if first_r is not None:
            rPr_e = first_r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = copy.deepcopy(rPr_e)
                orig_rPr.set("lang", "ko-KR")
                orig_rPr.set("dirty", "0")

        # 기존 run/fld 전체 제거
        for r in para.findall(f"{{{ns_a}}}r"):
            para.remove(r)
        for fld in para.findall(f"{{{ns_a}}}fld"):
            para.remove(fld)

        # endParaRPr 위치 파악 — 새 run은 반드시 그 앞에 삽입
        end_rpr = para.find(f"{{{ns_a}}}endParaRPr")
        insert_idx = list(para).index(end_rpr) if end_rpr is not None else len(para)

        r_new = ET.Element(f"{{{ns_a}}}r")
        if orig_rPr is not None:
            r_new.append(orig_rPr)
        else:
            ET.SubElement(r_new, f"{{{ns_a}}}rPr", lang="ko-KR", dirty="0")
        t_new = ET.SubElement(r_new, f"{{{ns_a}}}t")
        t_new.text = text
        para.insert(insert_idx, r_new)

    # ── 헬퍼: connector 선 투명 처리 ───────────────────────────────
    def set_line_alpha(cxn: ET.Element, alpha: str) -> None:
        spPr = cxn.find(f"{{{ns_p}}}spPr")
        if spPr is None:
            return
        ln = spPr.find(f"{{{ns_a}}}ln")
        if ln is None:
            return
        for schemeClr in ln.findall(f".//{{{ns_a}}}schemeClr"):
            # 기존 alpha 제거 후 새 값 설정
            for alp in schemeClr.findall(f"{{{ns_a}}}alpha"):
                schemeClr.remove(alp)
            ET.SubElement(schemeClr, f"{{{ns_a}}}alpha", val=alpha)

    # ── 1. 제목 placeholder (ID=42) ────────────────────────────────
    for sp in root.findall(f".//{{{ns_p}}}sp"):
        cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
        if cpr is None or cpr.get("id") != "42":
            continue
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            break
        paras = txBody.findall(f"{{{ns_a}}}p")
        if paras:
            set_para_text(paras[0], prs_title)
            for p in paras[1:]:
                set_para_text(p, "")
        break

    # ── 2. 항목 텍스트 (ID=10) ─────────────────────────────────────
    # cx를 선(x=8924925) 직전까지 확장해 1줄로 충분히 수용
    # bodyPr는 원본 그대로 유지 — paragraph 높이가 바뀌면 layout placeholder 노출됨
    TOC_TEXT_X  = 4541178
    TOC_PAGE_X  = 11431475        # 페이지번호 박스 시작 x
    GAP         = 100000          # 페이지번호 앞 여백
    TOC_TEXT_CX = TOC_PAGE_X - TOC_TEXT_X - GAP  # ≈ 6790297 EMU ≈ 7.43"

    for sp in root.findall(f".//{{{ns_p}}}sp"):
        cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
        if cpr is None or cpr.get("id") != "10":
            continue
        # cx 확장 (4.08" → 4.79")
        _set_shape_cx(sp, TOC_TEXT_CX)
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            break
        paras = txBody.findall(f"{{{ns_a}}}p")
        for i, para in enumerate(paras):
            raw_text = items_padded[i] if i < MAX_ROWS else ""
            # 1줄 수렴 보장: 넓어진 cx 기준으로 자름 (font_pt=30)
            safe = _truncate_to_lines(raw_text, TOC_TEXT_CX, 30, max_lines=1) if raw_text else ""
            set_para_text(para, safe)
        break

    # ── 3. 번호 열 (ID=11) — bodyPr 유지, 텍스트만 교체 ────────────
    for sp in root.findall(f".//{{{ns_p}}}sp"):
        cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
        if cpr is None or cpr.get("id") != "11":
            continue
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            break
        paras = txBody.findall(f"{{{ns_a}}}p")
        for i, para in enumerate(paras):
            num = f"{i+1}." if (i < len(items) and items[i]) else ""
            set_para_text(para, num)
        break

    # ── 4. 페이지 번호 열 (ID=14) — bodyPr 유지, 텍스트만 교체 ──────
    for sp in root.findall(f".//{{{ns_p}}}sp"):
        cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
        if cpr is None or cpr.get("id") != "14":
            continue
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            break
        paras = txBody.findall(f"{{{ns_a}}}p")
        for i, para in enumerate(paras):
            set_para_text(para, pages_padded[i] if i < MAX_ROWS else "")
        break

    # ── 5. 수평선 7개 (ID=12,13,37,38,39,40,41) ────────────────────
    # 가장 긴 항목 기준으로 모든 선의 x 시작점을 통일 (텍스트 겹침 방지)
    TOC_FONT_PT = 30
    LINE_GAP    = 200000   # 텍스트 끝 ~ 선 시작 여백 (≈0.22")
    TOC_PAGE_X  = 11431475
    SLIDE_RIGHT = 12192000  # 슬라이드 우측 끝 (16:9 표준 = 13.33")

    def _est_width_emu(text: str, font_pt: int) -> int:
        """한글/영문 혼합 텍스트의 렌더 폭 추정 (EMU)."""
        ko = sum(1 for c in text if ord(c) > 0x1000)
        en = len(text) - ko
        return int((ko * font_pt * 0.9 + en * font_pt * 0.5) * _EMU_PER_PT)

    max_text_emu = max(
        (_est_width_emu(it, TOC_FONT_PT) for it in items if it),
        default=0,
    )
    # 밑줄: 가장 긴 항목 텍스트 끝 + 여백 → 슬라이드 우측 끝까지
    # 페이지번호는 밑줄 위에 우측 정렬 (템플릿 동일 구조)
    line_x_start = TOC_TEXT_X + max_text_emu + LINE_GAP
    line_cx_len  = SLIDE_RIGHT - line_x_start  # 슬라이드 끝까지

    LINE_IDS = {"12", "13", "37", "38", "39", "40", "41"}
    line_cxns = []
    for cxn in root.findall(f".//{{{ns_p}}}cxnSp"):
        cpr = cxn.find(f"{{{ns_p}}}nvCxnSpPr/{{{ns_p}}}cNvPr")
        if cpr is not None and cpr.get("id") in LINE_IDS:
            line_cxns.append(cxn)

    line_cxns.sort(key=lambda c: int(
        (c.find(f".//{{{ns_a}}}off") or ET.Element("x")).get("y", "0")))

    for i, cxn in enumerate(line_cxns):
        # x, cx 업데이트
        off = cxn.find(f".//{{{ns_a}}}off")
        ext = cxn.find(f".//{{{ns_a}}}ext")
        if off is not None:
            off.set("x", str(line_x_start))
        if ext is not None:
            ext.set("cx", str(max(line_cx_len, 0)))
        # 빈 행 투명 처리
        alpha = "100000" if i < len(items) else "0"
        set_line_alpha(cxn, alpha)

    _write_xml(root, xml_path)
    ET.parse(xml_path)  # 유효성 검증


def edit_slide(work_dir: Path, slide_plan: dict) -> bool:
    """
    단일 슬라이드를 계획에 따라 편집한다.
    성공하면 True, 실패하면 False를 반환한다.
    """
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    xml_path = slides_dir / slide_plan["template_file"]

    if not xml_path.exists():
        print(f"  ⚠ 슬라이드 파일 없음: {xml_path}", file=sys.stderr)
        return False

    # 원본 백업
    backup = xml_path.with_suffix(".xml.bak")
    shutil.copy2(xml_path, backup)

    try:
        role = slide_plan.get("role", "")
        content = slide_plan.get("content", {})

        if role == "cover":
            edit_cover_slide(
                xml_path,
                title=slide_plan["title"],
                subtitle=content.get("subtitle", ""),
                date_str=content.get("date", "2026.00.00"),
            )
        elif role == "toc":
            # items 필드가 없으면 toc_items, bullets, body 순으로 폴백
            items = (
                content.get("items")
                or [i["text"] if isinstance(i, dict) else i
                    for i in content.get("toc_items", [])]
                or content.get("bullets", [])
                or []
            )
            page_nums = content.get("page_nums", [])
            edit_toc_slide(xml_path, slide_plan["title"], items, page_nums)
        elif role == "closing":
            # 마무리 슬라이드(slide46): 일체 편집하지 않음
            # slide46은 레이아웃에서 "감사합니다." 텍스트를 제공
            # — 편집하면 레이아웃 placeholder가 덮어씌워져 텍스트가 사라짐
            # placeholder 잔여 텍스트만 클리어
            try:
                tree_c = ET.parse(xml_path)
                _clear_residual_placeholders(tree_c.getroot())
                _write_xml(tree_c.getroot(), xml_path)
            except ET.ParseError:
                pass

        else:
            # template_file 기반 전용 편집기 디스패치
            tmpl_name = slide_plan.get("template_file", "")
            _SLIDE_EDITORS = {
                "slide8.xml":  _edit_slide8,
                "slide13.xml": _edit_slide13,
                "slide15.xml": _edit_slide13,
                "slide21.xml": _edit_slide21,
                "slide22.xml": _edit_slide24,
                "slide24.xml": _edit_slide24,
                "slide29.xml": _edit_slide29,
                "slide30.xml": _edit_slide30,
                "slide31.xml": _edit_slide31,
                "slide32.xml": _edit_slide32,
                "slide33.xml": _edit_slide33,
                "slide35.xml": _edit_slide35,
                "slide36.xml": _edit_slide36,
                "slide38.xml": _edit_slide38,
                "slide39.xml": _edit_slide39,
            }
            editor = _SLIDE_EDITORS.get(tmpl_name)
            if editor:
                editor(xml_path, slide_plan)
            else:
                # 전용 편집기 없으면 존 맵 기반 제너릭 편집기 (이미지 그리드 등)
                _edit_zonemap_slide(xml_path, slide_plan)

        # 편집 후 남은 placeholder 자동 제거
        try:
            tree = ET.parse(xml_path)
            if _clear_residual_placeholders(tree.getroot()):
                _write_xml(tree.getroot(), xml_path)
        except ET.ParseError:
            pass

        # 유효성 검증
        ET.parse(xml_path)
        print(f"  ✓ slide {slide_plan['index']} ({xml_path.name}) 편집 완료")
        backup.unlink(missing_ok=True)
        return True

    except ET.ParseError as e:
        print(f"  ✗ XML 오류, 원본 복원: {e}", file=sys.stderr)
        shutil.copy2(backup, xml_path)
        backup.unlink(missing_ok=True)
        return False


# ── 4. 슬라이드 트리밍 + 클린 & 패킹 ────────────────────────

def generate_excel_for_charts(work_dir: Path, plan: dict, output_dir: Path) -> Path | None:
    """
    차트 슬라이드(slide40~43)가 있을 때 Excel 데이터 파일을 생성한다.
    생성된 xlsx를 PPT와 같은 폴더에 저장 — 사용자가 PPT에서 '데이터 편집'으로 열 수 있다.
    """
    chart_slides = [s for s in plan.get("slides", [])
                    if s.get("template_file","") in
                    {"slide40.xml","slide41.xml","slide42.xml","slide43.xml"}]
    if not chart_slides:
        return None
    try:
        import openpyxl as _xl
    except ImportError:
        print("  ⚠ openpyxl 없음 — Excel 생성 건너뜀 (pip install openpyxl)")
        return None

    wb = _xl.Workbook()
    for slide in chart_slides:
        tmpl = slide.get("template_file","")
        chart_data = slide.get("content", {}).get("chart_data", {})
        sheet_name = f"{slide['index']}p_{tmpl.replace('.xml','')}"[:31]
        ws = wb.create_sheet(sheet_name)

        # 기본 헤더
        ws.append(["항목", "값", "비고"])
        if chart_data:
            for k, v in chart_data.items():
                ws.append([k, v, ""])
        else:
            # 템플릿별 기본 데이터 예시
            if "40" in tmpl:
                for item in [("KPI 1", 20), ("KPI 2", 35), ("KPI 3", 45)]:
                    ws.append(list(item) + [""])
            elif "41" in tmpl:
                for item in [("항목A", 83), ("항목B", 76), ("항목C", 65),
                              ("항목D", 58), ("항목E", 57)]:
                    ws.append(list(item) + [""])
            elif "42" in tmpl:
                for item in [("메인 KPI", 65), ("보조1", 75),
                              ("보조2", 75), ("보조3", 75)]:
                    ws.append(list(item) + [""])
            elif "43" in tmpl:
                for month, val in enumerate([2, 4, -2, 11, 4, 5], 1):
                    ws.append([f"{month}월", val, ""])

    if wb.sheetnames:
        wb.remove(wb["Sheet"])  # 기본 시트 제거
    xlsx_path = output_dir / f"{plan.get('topic','chart').replace(' ','_')}_data.xlsx"
    wb.save(str(xlsx_path))
    print(f"  ✓ Excel 데이터 파일 생성: {xlsx_path.name}")
    return xlsx_path


def _add_slide_to_presentation(work_dir: Path, slide_filename: str,
                                before_filename: str | None = None) -> None:
    """복사된 슬라이드 파일을 presentation.xml과 rels에 등록한다.
    before_filename: 이 파일의 rId 앞에 삽입 (None이면 맨 뒤).
    올바른 위치에 삽입해야 감사합니다 슬라이드가 마지막에 유지된다.
    """
    import re as _re
    prs_path  = work_dir / "unpacked" / "ppt" / "presentation.xml"
    rels_path = work_dir / "unpacked" / "ppt" / "_rels" / "presentation.xml.rels"
    if not prs_path.exists() or not rels_path.exists():
        return

    rels_raw = rels_path.read_text(encoding="utf-8")
    prs_raw  = prs_path.read_text(encoding="utf-8")

    # 새 rId 생성
    existing_ids = [int(m) for m in _re.findall(r'Id="rId(\d+)"', rels_raw)]
    new_rid_num  = max(existing_ids, default=0) + 1
    new_rid      = f"rId{new_rid_num}"

    # rels에 추가 (순서 무관)
    slide_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    new_rel    = f'<Relationship Id="{new_rid}" Type="{slide_type}" Target="slides/{slide_filename}"/>'
    rels_new   = rels_raw.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
    rels_path.write_text(rels_new, encoding="utf-8")

    # presentation.xml sldIdLst에 올바른 위치에 삽입
    existing_sld_ids = [int(m) for m in _re.findall(r'<p:sldId\b[^>]*\bid="(\d+)"', prs_raw)]
    new_sld_id = max(existing_sld_ids, default=255) + 1
    new_sld    = f'<p:sldId id="{new_sld_id}" r:id="{new_rid}"/>'

    if before_filename:
        # before_filename의 rId를 찾아서 그 앞에 삽입
        before_rid = None
        for m in _re.finditer(
            r'<Relationship\b([^>]+)/>', rels_raw
        ):
            attrs = m.group(1)
            if before_filename in attrs:
                rid_m = _re.search(r'\bId="(rId\d+)"', attrs)
                if rid_m:
                    before_rid = rid_m.group(1)
                    break
        if before_rid:
            # before_rid를 가진 sldId 앞에 삽입
            prs_new = _re.sub(
                rf'(\s*<p:sldId\b[^>]*\br:id="{_re.escape(before_rid)}"[^/]*/>\s*)',
                f'\n  {new_sld}\\1',
                prs_raw, count=1
            )
            prs_path.write_text(prs_new, encoding="utf-8")
            return

    # before_filename 없거나 못 찾으면 맨 뒤
    prs_new = _re.sub(r'(</p:sldIdLst>)', f'  {new_sld}\n  \\1', prs_raw, count=1)
    prs_path.write_text(prs_new, encoding="utf-8")


def _register_slide_content_type(work_dir: Path, slide_filename: str) -> None:
    """복사된 슬라이드를 [Content_Types].xml에 등록한다."""
    import re as _re
    ct_path = work_dir / "unpacked" / "[Content_Types].xml"
    if not ct_path.exists():
        return
    ct_raw = ct_path.read_text(encoding="utf-8")
    slide_ct = (
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
    )
    new_entry = (
        f'<Override PartName="/ppt/slides/{slide_filename}" '
        f'ContentType="{slide_ct}"/>'
    )
    if slide_filename not in ct_raw:
        ct_new = ct_raw.replace("</Types>", f"  {new_entry}\n</Types>")
        ct_path.write_text(ct_new, encoding="utf-8")


def _trim_to_plan_slides(work_dir: Path, plan: dict) -> None:
    """
    plan.json에 명시된 슬라이드만 남기고, plan의 index 순서대로 재정렬한다.
    순서 보장이 핵심 — 기존에는 템플릿 원본 순서를 따라 plan 순서와 불일치했음.
    """
    prs_path = work_dir / "unpacked" / "ppt" / "presentation.xml"
    rels_path = work_dir / "unpacked" / "ppt" / "_rels" / "presentation.xml.rels"
    if not prs_path.exists() or not rels_path.exists():
        return

    # plan의 index 순서대로 파일명 목록 (순서 보장용)
    ordered_files = [s.get("template_file", "") for s in
                     sorted(plan.get("slides", []), key=lambda x: x.get("index", 0))]
    plan_files = set(ordered_files)
    if not plan_files:
        return

    prs_raw  = prs_path.read_text(encoding="utf-8")
    rels_raw = rels_path.read_text(encoding="utf-8")

    # rId → 파일명 매핑
    rid_to_file: dict[str, str] = {}
    for m in re.finditer(r'<Relationship\b([^>]+)/>', rels_raw):
        attrs = m.group(1)
        rid = re.search(r'\bId="(rId\d+)"', attrs)
        tgt = re.search(r'\bTarget="[^"]*/(slide[\w]+\.xml)"', attrs)
        typ = re.search(r'\bType="([^"]+)"', attrs)
        if rid and tgt and typ and 'slide' in typ.group(1) and 'Layout' not in typ.group(1):
            rid_to_file[rid.group(1)] = tgt.group(1)

    file_to_rid = {v: k for k, v in rid_to_file.items()}

    # plan에 없는 sldId 수집
    keep_rids = {file_to_rid[f] for f in plan_files if f in file_to_rid}
    remove_rids: set[str] = set()
    for rid in rid_to_file:
        if rid not in keep_rids:
            remove_rids.add(rid)

    # 불필요한 슬라이드 제거
    new_prs = re.sub(
        r'\s*<p:sldId\b[^/]*/>',
        lambda m: '' if (r := re.search(r'\br:id="(rId\d+)"', m.group(0)))
                        and r.group(1) in remove_rids else m.group(0),
        prs_raw,
    )
    # ── plan의 index 순서대로 sldIdLst 재정렬 ──────────────────
    # 기존 순서는 템플릿 원본 순서를 따름 → plan 순서와 불일치
    existing_slds: list[str] = re.findall(r'\s*<p:sldId\b[^/]*/>', new_prs)
    rid_to_sld = {}
    for sld in existing_slds:
        r_m = re.search(r'\br:id="(rId\d+)"', sld)
        if r_m:
            rid_to_sld[r_m.group(1)] = sld

    # plan 순서대로 sldId 태그를 재배열
    ordered_slds = []
    for fname in ordered_files:
        rid = file_to_rid.get(fname)
        if rid and rid in rid_to_sld:
            ordered_slds.append(rid_to_sld[rid].strip())

    # sldIdLst 내부를 ordered_slds로 교체
    new_sld_block = "\n    " + "\n    ".join(ordered_slds) + "\n  "
    new_prs2 = re.sub(
        r'(<p:sldIdLst>).*?(</p:sldIdLst>)',
        rf'\1{new_sld_block}\2',
        new_prs, flags=re.DOTALL, count=1
    )
    prs_path.write_text(new_prs2, encoding="utf-8")
    print(f"  ✓ 슬라이드 트리밍 + 순서 재정렬: {len(plan_files)}장 (plan index 순서 보장)")


# ── 5. 클린 & 패킹 ───────────────────────────────────────────

def pack_output(work_dir: Path, output_path: Path,
                skip_validation: bool = False) -> bool:
    """clean.py → pack.py 순서로 실행해 pptx를 생성한다."""
    scripts = SKILL_DIR / "scripts"
    template = work_dir / "template.pptx"
    unpacked = work_dir / "unpacked"

    result = subprocess.run(
        [sys.executable, str(scripts / "clean.py"), str(unpacked)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"  ⚠ clean.py stderr: {result.stderr}", file=sys.stderr)

    cmd = [
        sys.executable, str(scripts / "office" / "pack.py"),
        str(unpacked), str(output_path),
        "--original", str(template),
    ]
    if skip_validation:
        cmd += ["--validate", "false"]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ✗ pack.py 실패:\n{result.stderr}", file=sys.stderr)
        return False

    print(f"  ✓ 패킹 완료: {output_path}")
    return True


# ── 5. 시각 QA ───────────────────────────────────────────────

def _pdf_via_powerpoint(pptx_path: Path, pdf_path: Path) -> bool:
    """
    macOS AppleScript으로 PowerPoint에 PDF 내보내기를 요청한다.
    실패 시 2초 대기 후 1회 재시도.
    PowerPoint 없으면 False 반환.

    실패 원인 분석 (2026-06-02):
    - PowerPoint 자체는 연속 호출에도 정상 동작 (5초/회)
    - 간헐적 실패는 Vision Fix 루프에서 pack.py 직후 즉시 호출 시
      이전 PowerPoint 프로세스가 완전히 정리되기 전에 새 요청이 들어오는 경우
    - 해결책: 실패 시 2초 대기 후 재시도
    """
    import shutil as _shutil, time as _time
    if not _shutil.which("osascript"):
        return False

    # 문제: PowerPoint에 블로킹 다이얼로그가 있으면 AppleScript 전체가 차단됨
    # 해결: pkill로 완전 종료 → 새 인스턴스로 열기 → PDF 저장
    # open 명령은 로딩 완료를 기다리지 않으므로 동적 대기 필요 (최대 30초)

    # macOS 샌드박스 제한 해결:
    # PowerPoint는 ~/.ppt-skill/runs/... 같은 숨김 경로에 접근 시
    # "Grant File Access" 다이얼로그를 띄움 → 무인 파이프라인에서 타임아웃
    # 해결: result 폴더 안 tmp/ 에 복사 후 변환 (사용자가 PowerPoint 접근 허용한 위치)
    import shutil as _shutil2

    result_dir = Path.home() / "Documents" / "claude" / "ppt-skill" / "result"
    tmp_dir  = result_dir / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_pptx = tmp_dir / "_ppt_qa_tmp.pptx"
    tmp_pdf  = tmp_dir / "_ppt_qa_tmp.pdf"

    try:
        _shutil2.copy2(str(pptx_path), str(tmp_pptx))
    except Exception as e:
        print(f"  ⚠ 임시 파일 복사 실패: {e}")
        return False

    script = f"""
on run
    tell application "Microsoft PowerPoint"
        repeat while (count of presentations) > 0
            close presentation 1 saving no
        end repeat
        open (POSIX file "{tmp_pptx}")
        set maxWait to 30
        set waited to 0
        repeat while (count of presentations) = 0 and waited < maxWait
            delay 1
            set waited to waited + 1
        end repeat
        if (count of presentations) = 0 then
            error "프레젠테이션 열기 실패"
        end if
        save active presentation in (POSIX file "{tmp_pdf}") as save as PDF
        close active presentation saving no
    end tell
end run
"""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=180,
        )
        if result.returncode == 0 and tmp_pdf.exists():
            # 결과를 원래 경로로 이동
            _shutil2.move(str(tmp_pdf), str(pdf_path))
            return True
        if result.stderr:
            print(f"  ⚠ PowerPoint 오류: {result.stderr[:100]}")
    except subprocess.TimeoutExpired:
        print("  ⚠ PowerPoint 타임아웃 (180s)")
    finally:
        # 임시 파일 정리
        tmp_pptx.unlink(missing_ok=True)
        tmp_pdf.unlink(missing_ok=True)

    return False


def visual_qa(work_dir: Path, output_path: Path) -> list[str]:
    """
    PowerPoint로 PDF 변환 후 pdftoppm으로 슬라이드 이미지를 생성한다.
    LibreOffice 폴백 없음 — PowerPoint만 사용 (정확한 렌더링 보장).
    PowerPoint 실패 시 빈 리스트 반환 (QA 건너뜀).
    """
    import shutil as _shutil

    pdf_path = work_dir / "output.pdf"

    # PowerPoint (유일한 QA 소스)
    if not _shutil.which("osascript"):
        print("  ⚠ osascript 없음 (macOS 전용) — 시각 QA 건너뜀")
        return []

    ok = _pdf_via_powerpoint(output_path, pdf_path)
    if not ok:
        print("  ⚠ PowerPoint PDF 변환 실패 — 시각 QA 건너뜀")
        print("    (원인: PowerPoint 미설치, 타임아웃, 또는 파일 접근 오류)")
        return []

    print("  ✓ PowerPoint PDF 변환 완료")
    qa_source = "powerpoint"

    # 이미지 변환
    slides_img_dir = work_dir / "qa_images"
    slides_img_dir.mkdir(exist_ok=True)
    prefix = str(slides_img_dir / "slide")
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", "120", str(pdf_path), prefix],
        capture_output=True, text=True,
    )
    images = sorted(slides_img_dir.glob("*.jpg"))
    print(f"  ✓ QA 이미지 {len(images)}장 ({qa_source}): {slides_img_dir}")

    qa_report = {
        "images": [str(p) for p in images],
        "pdf": str(pdf_path),
        "source": qa_source,
        "note": "텍스트 오버플로우·요소 겹침·플레이스홀더 잔여를 확인하세요.",
    }
    (work_dir / "qa_report.json").write_text(
        __import__("json").dumps(qa_report, ensure_ascii=False, indent=2)
    )
    return [str(p) for p in images]


# ── 6. 섹션 정리 ─────────────────────────────────────────────

def restructure_sections(output_path: Path) -> None:
    """
    생성된 pptx의 섹션을 표지/목차·간지/본문/마무리 4개로 정리한다.
    - 가이드·변화·타임라인·흐름 섹션: 제거 (슬라이드는 이미 없음)
    - 그래프·차트 등 나머지 섹션: 슬라이드를 본문으로 이동 후 제거
    - 매번 생성되는 PPT에 자동 적용된다.
    """
    import re as _re
    import shutil as _shutil

    KEEP = {'표지', '목차/간지', '본문', '마무리'}
    MOVE_TO_BODY = set()   # 본문으로 이동
    DELETE = set()          # 그냥 삭제

    tmp = output_path.with_suffix(".tmp.pptx")
    try:
        with zipfile.ZipFile(output_path) as zin:
            if 'ppt/presentation.xml' not in zin.namelist():
                return
            prs_raw = zin.read('ppt/presentation.xml').decode('utf-8')

        # 섹션 이름 수집
        all_names = _re.findall(r'<p14:section\b[^>]*name="([^"]+)"', prs_raw)
        for name in all_names:
            if name not in KEEP:
                MOVE_TO_BODY.add(name)

        if not MOVE_TO_BODY:
            return   # 이미 정리된 상태

        def get_section(xml, name):
            pat = rf'<p14:section\b[^>]*name="{_re.escape(name)}"[^>]*>.*?</p14:section>'
            m = _re.search(pat, xml, _re.DOTALL)
            return m.group(0) if m else ''

        prs_new = prs_raw

        # 비정규 섹션의 sldId를 본문으로 이동
        move_tags = []
        for name in MOVE_TO_BODY:
            sec = get_section(prs_new, name)
            move_tags += _re.findall(r'<p14:sldId\b[^/]*/>', sec)

        if move_tags:
            body_sec = get_section(prs_new, '본문')
            if body_sec:
                body_updated = _re.sub(
                    r'(</p14:sldIdLst>)',
                    ''.join(move_tags) + r'\1',
                    body_sec, count=1,
                )
                prs_new = prs_new.replace(body_sec, body_updated, 1)

        # 비정규 섹션 제거
        for name in MOVE_TO_BODY:
            sec = get_section(prs_new, name)
            if sec:
                prs_new = prs_new.replace(sec, '')

        with zipfile.ZipFile(output_path) as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'ppt/presentation.xml':
                    zout.writestr(item, prs_new.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))

        _shutil.move(str(tmp), str(output_path))
        remaining = _re.findall(r'<p14:section\b[^>]*name="([^"]+)"', prs_new)
        print(f"  ✓ 섹션 정리 완료: {remaining}")

    except Exception as e:
        print(f"  ⚠ 섹션 정리 실패 (무시): {e}")
        if tmp.exists():
            tmp.unlink()


# ── 7. 콘텐츠 검증 ───────────────────────────────────────────

def verify_content(output_path: Path) -> list[str]:
    """pptx 텍스트를 추출해 플레이스홀더 잔여 여부를 확인한다."""
    extract_script = """
import sys
from pptx import Presentation
prs = Presentation(sys.argv[1])
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t:
                    print(t)
"""
    result = subprocess.run(
        [sys.executable, "-c", extract_script, str(output_path)],
        capture_output=True, text=True,
    )
    issues = [
        line.strip()
        for line in result.stdout.splitlines()
        if _PLACEHOLDER_RE.search(line)
    ]
    return issues


# ── Plan 검증 ────────────────────────────────────────────────

# planning_constraints에서 파싱한 금지/허용 슬라이드 목록
# 사용 불가 슬라이드 (이미지 전용, 차트 고정, 구조적 편집 불가)
_BANNED_SLIDES: set[str] = {
    # 섹션 구분: 명시적 요청 시에만 (section role 자동 선택 금지)
    "slide8.xml",
    # 이미지가 화면 대부분을 차지해 텍스트 설명으로 대체가 어려운 레이아웃
    "slide18.xml", "slide19.xml", "slide20.xml", "slide23.xml",
    "slide21.xml", "slide22.xml",
    # 버블 개념도 (복잡 — 전용 편집기 필요)
    "slide34.xml", "slide35.xml",
    # 3행 상세 흐름도 — service(4번째) 컬럼 미충전 이슈 → flow는 slide38로 통일
    "slide39.xml",
    # 레이아웃에 영상(media) placeholder가 있어 ▶ 아이콘 노출 → 텍스트는 slide32 사용
    "slide24.xml", "slide25.xml", "slide26.xml", "slide27.xml", "slide28.xml",
    # 차트 고정 (Excel 데이터 임베딩 미완성 — 추후 과제)
    "slide40.xml", "slide41.xml", "slide42.xml", "slide43.xml",
}

# ── 콘텐츠 유형별 슬라이드 카탈로그 ───────────────────────────────
# Claude가 콘텐츠를 분석한 후 아래 목록에서 template_file을 선택
_ALLOWED_CONTENT_SLIDES: list[str] = [
    # 텍스트 위주 (가장 안전 — 이미지/도형 불필요, 미배정 시 기본 폴백)
    "slide32.xml",  # ✅ 텍스트+배너 (긴 설명/개요/서술형)
    # 프로세스/흐름
    "slide38.xml",  # ✅ 3행 흐름도 (keyword→Solution→Service) — flow 표준
    "slide30.xml",  # ✅ 4단계 스텝 프로세스
    # 시간축 레이아웃 (반드시 연도/분기 시계열 콘텐츠일 때만)
    "slide29.xml",  # ✅ 연도별/월별 타임라인 (2023→2026)
    "slide31.xml",  # ✅ 분기별 Q1→Q2→Q3→Q4
    "slide33.xml",  # ✅ 분기별 (상단 설명+Q4열, slide31 변형)
    # 비교/분석
    "slide36.xml",  # 🟡 As-is/To-be 벤다이어그램
    # 카드/아이콘 (아이콘 영역은 비워둠 — placeholder 텍스트 미노출)
    "slide13.xml",  # 🟡 3열 아이콘카드 (3가지 기능/특징/장점)
    "slide15.xml",  # 🟡 3열 대형아이콘 (3가지 핵심 가치)
    "slide14.xml",  # 🟡 4열 아이콘카드 + Insight
    "slide16.xml",  # 🟡 4열 아이콘 컴팩트
    # 이미지+텍스트 (이미지 자리는 '어떤 이미지' 설명 텍스트로 채움)
    "slide9.xml",   # 3열 이미지+제목+설명
    "slide10.xml",  # 3열 이미지+제목+설명 + Insight 배너
    "slide11.xml",  # 3행 이미지(좌)+설명
    "slide12.xml",  # 3열 이미지+우측 텍스트
    "slide17.xml",  # 2x2 이미지+설명
    # slide24~28(영상 placeholder), slide34/35(버블), slide39(4컬럼), slide40~43(차트)
    # 는 _BANNED — 추후 전용 처리 후 활성화
]

# 레이아웃별 필수 콘텐츠 필드 — 하나도 없으면 텍스트 배너(slide32)로 리맵.
# 빈 타임라인 막대·빈 카드·빈 흐름도가 배포되는 것을 코드 레벨에서 차단한다.
_LAYOUT_CONTENT_REQ: dict[str, list[str]] = {
    "slide29.xml": ["periods"],
    "slide31.xml": ["quarters"],
    "slide33.xml": ["quarters"],
    "slide30.xml": ["steps"],
    "slide38.xml": ["keywords"],
    "slide35.xml": ["before", "after"],
    "slide36.xml": ["as_is", "to_be"],
    "slide13.xml": ["items"],
    "slide15.xml": ["items"],
    "slide14.xml": ["items"],
    "slide16.xml": ["items"],
    "slide9.xml":  ["items"],
    "slide10.xml": ["items"],
    "slide11.xml": ["items"],
    "slide12.xml": ["items"],
    "slide17.xml": ["items"],
}


def _has_content_field(content: dict, fields: list[str]) -> bool:
    """content(또는 content.body) 안에 주어진 필드 중 하나라도 값이 있으면 True."""
    body = content.get("body") if isinstance(content.get("body"), dict) else {}
    for f in fields:
        if content.get(f) or (isinstance(body, dict) and body.get(f)):
            return True
    return False

_CLOSING_SLIDES: list[str] = ["slide46.xml", "slide44.xml", "slide45.xml"]
_CONTENT_TITLE_ID   = "8"
_CONTENT_BULLETS_ID = "11"  # 호환성 유지


def enforce_plan_constraints(plan: dict, slide_info: list[dict]) -> tuple[dict, list[str]]:
    """
    생성된 plan의 template_file이 금지 목록에 있거나 중복되면 자동 교체한다.
    planning_constraints를 Claude에게만 맡기지 않고 코드 레벨에서 강제.
    반환: (수정된 plan, 변경 로그)
    """
    available_files = {s["file"] for s in slide_info}
    allowed = [f for f in _ALLOWED_CONTENT_SLIDES if f in available_files]
    if not allowed:
        allowed = [f for f in available_files
                   if f not in _BANNED_SLIDES
                   and f not in {"slide6.xml", "slide7.xml", "slide9.xml"}]

    used_files: set[str] = set()
    changes: list[str] = []
    allowed_pool = list(allowed)
    pool_idx = 0

    for slide in plan.get("slides", []):
        role = slide.get("role", "")
        tmpl = slide.get("template_file", "")

        # cover/toc는 전용 파일 유지
        if role in ("cover", "toc"):
            used_files.add(tmpl)
            continue

        # closing은 전용 마무리 슬라이드로 강제
        if role == "closing":
            for cs in _CLOSING_SLIDES:
                if cs in available_files and cs not in used_files:
                    if cs != tmpl:
                        changes.append(f"slide {slide['index']}: {tmpl} → {cs} (closing 전용)")
                        slide["template_file"] = cs
                        tmpl = cs
                    break
            used_files.add(tmpl)
            continue

        # 금지 목록이거나 허용 목록에 없으면 교체
        # Claude가 직접 template을 선택했으면 존중, 금지된 경우만 교체
        # section role은 사용자 명시 요청 없이는 일반 content로 변환
        if role == "section" and tmpl == "slide8.xml":
            role = "content"
            slide["role"] = "content"

        needs_replace = (tmpl in _BANNED_SLIDES) or (
            tmpl not in _ALLOWED_CONTENT_SLIDES
            and role not in ("cover", "toc", "closing"))

        if needs_replace:
            replacement = None
            for candidate in _ALLOWED_CONTENT_SLIDES:
                if candidate in available_files and candidate not in used_files:
                    replacement = candidate
                    break
            if replacement is None:
                # 중복 허용 (pool이 부족한 경우)
                for candidate in _ALLOWED_CONTENT_SLIDES:
                    if candidate in available_files:
                        replacement = candidate
                        break
            if replacement:
                changes.append(f"slide {slide['index']}: {tmpl} → {replacement}"
                                + f" (금지/미허용 → 교체)")
                slide["template_file"] = replacement
                tmpl = replacement
            else:
                changes.append(f"slide {slide['index']}: 대체 슬라이드 없음 ({tmpl} 유지)")

        # ── 레이아웃-콘텐츠 적합성 가드 ──
        # 선택된 레이아웃의 필수 콘텐츠가 없으면 텍스트 배너(slide32)로 리맵.
        # 예) 서술형 내용에 연도 타임라인(slide29) 배정 → 빈 막대 방지.
        req = _LAYOUT_CONTENT_REQ.get(tmpl)
        if req and not _has_content_field(slide.get("content", {}), req):
            fallback = "slide32.xml" if "slide32.xml" in available_files else (
                "slide24.xml" if "slide24.xml" in available_files else None)
            if fallback and fallback != tmpl:
                changes.append(
                    f"slide {slide['index']}: {tmpl} → {fallback} "
                    f"(필수 콘텐츠 {req} 없음 → 텍스트 레이아웃)")
                slide["template_file"] = fallback
                tmpl = fallback

        used_files.add(tmpl)

    return plan, changes


def validate_plan(plan: dict) -> tuple[bool, list[str]]:
    """
    plan.json의 각 슬라이드가 role별 필수 content 필드를 가졌는지 검증.
    반환: (ok, warnings)
    """
    warnings: list[str] = []
    for slide in plan.get("slides", []):
        role = slide.get("role", "content")
        content = slide.get("content", {})
        schema = PLAN_CONTENT_SCHEMA.get(role, {})

        for field in schema.get("required", []):
            if field not in content or not content[field]:
                warnings.append(
                    f"slide {slide['index']} ({role}): 필수 필드 '{field}' 없음"
                )
        for field in schema.get("recommended", []):
            if field not in content or not content[field]:
                warnings.append(
                    f"slide {slide['index']} ({role}): 권장 필드 '{field}' 없음 — 폴백 사용"
                )

    ok = not any("필수" in w for w in warnings)
    return ok, warnings


# ── known_fixes 자동 적용 ────────────────────────────────────

def apply_known_fixes(xml_path: Path, slide_plan: dict, known_fixes: list) -> None:
    """
    long_term_memory의 known_failure_fixes 패턴을 슬라이드 편집 전에 검사·적용.
    현재 적용 가능한 픽스:
      - endParaRPr 순서 보정 (run_after_endpararpr)
      - lang=ko-KR 강제 (korean_in_en_us_run)
    """
    if not xml_path.exists() or not known_fixes:
        return

    ns_a = _NS_A
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        fixed = False

        for para in root.iter(f"{{{ns_a}}}p"):
            children = list(para)
            tags = [c.tag.split("}")[-1] for c in children]

            # Fix 1: <a:r>가 <a:endParaRPr> 뒤에 있으면 앞으로 이동
            if "endParaRPr" in tags and "r" in tags:
                end_idx = tags.index("endParaRPr")
                runs_after = [i for i, t in enumerate(tags) if t == "r" and i > end_idx]
                if runs_after:
                    end_elem = children[end_idx]
                    for i in sorted(runs_after, reverse=True):
                        r_elem = children[i]
                        para.remove(r_elem)
                        para.insert(end_idx, r_elem)
                    fixed = True

            # Fix 2: lang=en-US run에 한국어 텍스트 → lang=ko-KR
            for r in para.findall(f"{{{ns_a}}}r"):
                rPr = r.find(f"{{{ns_a}}}rPr")
                t   = r.find(f"{{{ns_a}}}t")
                if rPr is not None and t is not None and t.text:
                    if rPr.get("lang") == "en-US" and re.search(r"[가-힣]", t.text):
                        rPr.set("lang", "ko-KR")
                        fixed = True

        if fixed:
            _write_xml(root, xml_path)

    except ET.ParseError:
        pass  # 편집 함수가 따로 처리


# ── Verifier 실행 파이프라인 ─────────────────────────────────

def _vfy_xml_validity(output_path: Path, work_dir: Path, **_) -> list[dict]:
    issues = []
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    for xml_f in sorted(slides_dir.glob("slide*.xml")):
        try:
            ET.parse(xml_f)
        except ET.ParseError as e:
            issues.append({"rule": "xml_validity", "severity": "CRITICAL",
                           "file": xml_f.name, "detail": str(e)})
    return issues


def _vfy_placeholder(output_path: Path, work_dir: Path, **_) -> list[dict]:
    """plan.json에 명시된 슬라이드(실제 편집된 것)만 placeholder 검사."""
    issues = []
    if not output_path.exists():
        return issues

    # plan.json에서 편집된 슬라이드 파일명 집합
    plan_path = work_dir / "plan.json"
    edited_files: set[str] = set()
    if plan_path.exists():
        plan = json.loads(plan_path.read_text())
        edited_files = {s.get("template_file", "") for s in plan.get("slides", [])}

    extract = (
        "import sys, json; from pptx import Presentation; prs=Presentation(sys.argv[1]);"
        "[print(p.text.strip()) for s in prs.slides for sh in s.shapes"
        " if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip()]"
    )
    result = subprocess.run([sys.executable, "-c", extract, str(output_path)],
                            capture_output=True, text=True)

    # plan이 없으면 전체 검사, 있으면 편집된 슬라이드만 검사
    # pptx 슬라이드 텍스트를 파일 단위로 분리하기 어려우므로
    # 실제 편집 슬라이드에서만 발생할 수 있는 패턴을 검사
    PATTERNS = ["lorem", "작성해주세요", "todo", "[insert", "placeholder"]
    seen: set[str] = set()
    for line in result.stdout.splitlines():
        text = line.strip()
        if not text or text in seen:
            continue
        seen.add(text)
        if any(p in text.lower() for p in PATTERNS):
            # 편집 슬라이드에서만 실제 문제가 되는 패턴 (중복 제거)
            issues.append({"rule": "placeholder_check", "severity": "CRITICAL",
                           "detail": text})
            if len(issues) >= 5:  # 최대 5건만 보고 (템플릿 잔여 노이즈 방지)
                break
    return issues


def _vfy_run_order(output_path: Path, work_dir: Path, **_) -> list[dict]:
    issues = []
    ns_a = _NS_A
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    for xml_f in sorted(slides_dir.glob("slide*.xml")):
        try:
            tree = ET.parse(xml_f)
            for para in tree.getroot().iter(f"{{{ns_a}}}p"):
                tags = [c.tag.split("}")[-1] for c in para]
                if "endParaRPr" in tags:
                    ei = tags.index("endParaRPr")
                    if any(t == "r" for t in tags[ei+1:]):
                        issues.append({"rule": "run_before_endpararpr",
                                       "severity": "CRITICAL", "file": xml_f.name,
                                       "detail": "<a:r> after <a:endParaRPr>"})
                        break
        except ET.ParseError:
            pass
    return issues


def _vfy_section_structure(output_path: Path, **_) -> list[dict]:
    if not output_path.exists():
        return []
    EXPECTED = {"표지", "목차/간지", "본문", "마무리"}
    try:
        with zipfile.ZipFile(output_path) as z:
            prs = z.read("ppt/presentation.xml").decode("utf-8")
        found = set(re.findall(r'<p14:section\b[^>]*name="([^"]+)"', prs))
        if found != EXPECTED:
            return [{"rule": "section_structure", "severity": "MEDIUM",
                     "detail": f"섹션 불일치: {found} ≠ {EXPECTED}"}]
    except Exception:
        pass
    return []


def _vfy_qa_done(output_path: Path, work_dir: Path, **_) -> list[dict]:
    qa = work_dir / "qa_report.json"
    if not qa.exists() or not json.loads(qa.read_text()).get("images"):
        return [{"rule": "qa_completion_check", "severity": "MEDIUM",
                 "detail": "qa_ok=false — QA 이미지 없음"}]
    return []


def _vfy_toc_paragraph_count(output_path: Path, work_dir: Path, **_) -> list[dict]:
    """slide7(목차) ID=10의 paragraph 수가 7개인지 검증.
    Vision Fix Agent가 set_paragraphs로 줄이면 번호·텍스트·페이지 높이 불일치 발생."""
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    slide7 = slides_dir / "slide7.xml"
    if not slide7.exists():
        return []
    try:
        tree = ET.parse(slide7)
        ns_p = _NS_P
        ns_a = _NS_A
        for sp in tree.getroot().findall(f".//{{{ns_p}}}sp"):
            cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
            if cpr is None or cpr.get("id") != "10":
                continue
            paras = sp.findall(f".//{{{ns_a}}}p")
            if len(paras) != 7:
                return [{"rule": "toc_paragraph_count", "severity": "CRITICAL",
                         "detail": f"TOC ID=10 paragraph 수={len(paras)} (7 필요) — "
                                   "Vision Fix Agent가 구조 파괴했을 수 있음"}]
    except ET.ParseError:
        pass
    return []


_VERIFIER_REGISTRY: dict[str, object] = {
    "xml_validity":          _vfy_xml_validity,
    "placeholder_check":     _vfy_placeholder,
    "run_before_endpararpr": _vfy_run_order,
    "section_structure":     _vfy_section_structure,
    "qa_completion_check":   _vfy_qa_done,
    "toc_paragraph_count":   _vfy_toc_paragraph_count,
}


def execute_verifier_rules(output_path: Path, work_dir: Path) -> list[dict]:
    """
    ~/.ppt-skill/harness/verifier_rules.json 에 정의된 규칙을 실행하고
    위반 목록을 반환한다. CRITICAL 위반이 있으면 호출자가 재시도 트리거.
    """
    rules_path = SKILL_DIR / "harness" / "verifier_rules.json"
    if not rules_path.exists():
        return []

    rules_cfg = json.loads(rules_path.read_text()).get("rules", {})
    all_issues: list[dict] = []

    for rule_name, cfg in rules_cfg.items():
        fn = _VERIFIER_REGISTRY.get(rule_name)
        if fn is None:
            continue
        try:
            issues = fn(output_path=output_path, work_dir=work_dir)
            # 규칙 설정의 severity 우선 (없으면 함수 기본값)
            for iss in issues:
                iss.setdefault("severity", cfg.get("severity", "MEDIUM"))
            all_issues.extend(issues)
        except Exception as e:
            print(f"  ⚠ verifier [{rule_name}] 실행 오류: {e}")

    critical = [i for i in all_issues if i["severity"] == "CRITICAL"]
    if critical:
        print(f"  ✗ CRITICAL 위반 {len(critical)}건:")
        for i in critical[:3]:
            print(f"    [{i['rule']}] {i['detail']}")
    elif all_issues:
        print(f"  ⚠ verifier 경고 {len(all_issues)}건 (non-critical)")
    else:
        print("  ✓ 모든 verifier 규칙 통과")

    return all_issues


# ── 슬라이드별 전용 편집기 ────────────────────────────────────────

def _edit_slide8(xml_path: Path, slide_plan: dict) -> None:
    """slide8 (섹션 구분): ID=2(섹션 제목 40pt), ID=4(서브항목 목록), ID=5(번호)."""
    import copy as _copy
    ns_p, ns_a = _NS_P, _NS_A
    title   = slide_plan.get("title", "")
    content = slide_plan.get("content", {})
    items   = content.get("items") or content.get("bullets") or []

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return

    def _set(sp, text):
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        first_r = sp.find(f".//{{{ns_a}}}r")
        orig_rPr = None
        if first_r is not None:
            rPr_e = first_r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang", "ko-KR"); orig_rPr.set("dirty", "0")
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr is not None: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
            p.insert(idx, r_new)
            break

    sp2 = _find_shape_by_id(root, "2")
    if sp2: _set(sp2, title)

    # ID=4: 서브항목 (여러 paragraph)
    sp4 = _find_shape_by_id(root, "4")
    if sp4 and items:
        txBody = sp4.find(f"{{{ns_p}}}txBody")
        if txBody:
            orig_rPr = None
            for r in sp4.findall(f".//{{{ns_a}}}r"):
                rPr_e = r.find(f"{{{ns_a}}}rPr")
                if rPr_e is not None:
                    orig_rPr = _copy.deepcopy(rPr_e)
                    orig_rPr.set("lang", "ko-KR"); orig_rPr.set("dirty", "0")
                    break
            paras = txBody.findall(f"{{{ns_a}}}p")
            for i, para in enumerate(paras):
                for r in para.findall(f"{{{ns_a}}}r"): para.remove(r)
                text = items[i] if i < len(items) else ""
                end = para.find(f"{{{ns_a}}}endParaRPr")
                idx = list(para).index(end) if end is not None else len(para)
                r_new = ET.Element(f"{{{ns_a}}}r")
                if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
                ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
                para.insert(idx, r_new)

    _clear_residual_placeholders(root)
    _write_xml(root, xml_path)


def _apply_common_zones(root, slide_plan: dict, template_file: str) -> None:
    """
    모든 본문 슬라이드 공통 3-Zone 처리:
      Zone 1 (헤더 바): ID=8(대제목), ID=9(중제목)
      Zone 2 (사이드바): 슬라이드별 label/desc ID → 본문제목, 본문설명글
      Zone 3 (본문 구역): 각 편집기가 직접 처리

    사이드바 형식:
      label: "1.1\\n본문제목" (번호 + 제목 분리)
      desc:  "본문설명글 (3줄 이내)"
    """
    import copy as _copy
    ns_p, ns_a = _NS_P, _NS_A

    title    = slide_plan.get("title", "")
    subtitle = slide_plan.get("subtitle", "")
    content  = slide_plan.get("content", {})
    sec_no   = content.get("section_no", "")
    # section_title 없으면 비워둠 — 슬라이드 제목(대제목)을 본문제목에 중복 삽입하지 않음
    sec_title= content.get("section_title", "")
    sec_desc = content.get("section_desc", "")

    def _set_shape(sid: str, text: str) -> None:
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = None
        for r in sp.findall(f".//{{{ns_a}}}r"):
            rPr_e = r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang", "ko-KR"); orig_rPr.set("dirty", "0"); break
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
            p.insert(idx, r_new); break

    # Zone 1: 헤더 바 — 긴 제목은 autofit으로 축소(오버플로우/잘림 방지)
    _set_shape("8", title)
    hdr = _find_shape_by_id(root, "8")
    if hdr is not None:
        _enable_autofit(hdr)

    if subtitle:
        _set_shape("9", _truncate_to_lines(subtitle, 8_000_000, 14, 1))

    # Zone 2: 사이드바 — 존 맵(body_title/body_desc) 우선, 없으면 레거시 dict
    z = _zone(template_file)
    label_id = z.get("body_title") or _SIDEBAR_LABEL_ID.get(template_file)
    desc_id  = z.get("body_desc")  or _SIDEBAR_DESC_ID.get(template_file)
    if label_id:
        # section_title 미지정 시 슬라이드 제목으로 폴백 (사이드바 공백 방지)
        effective_title = sec_title or title
        label_text = f"{sec_no}\n{effective_title}" if sec_no else effective_title
        _set_shape(label_id, label_text)
        lbl = _find_shape_by_id(root, label_id)
        if lbl is not None:
            _enable_autofit(lbl)
    if desc_id and sec_desc:
        _set_shape(desc_id, _truncate_to_lines(sec_desc, 2_200_000, 12, 5))


def _edit_slide29(xml_path: Path, slide_plan: dict) -> None:
    """slide29 (연도별/월별 타임라인):
    Zone1: 헤더(ID=8,9) | Zone2: 사이드바(ID=18,19) | Zone3: 연도별 본문(ID=16/17/22/23 레이블 + ID=24/12 내용)
    periods: [{"label":"2024","content":"..."}]"""
    content = slide_plan.get("content", {})
    periods = content.get("periods", [])
    body    = content.get("body", {})
    if isinstance(body, dict):
        periods = periods or body.get("periods", [])

    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return

    # Zone 1+2: 공통 처리
    _apply_common_zones(root, slide_plan, "slide29.xml")

    # Zone 3: 연도별 레이블 + 본문 내용
    import copy as _copy
    ns_p, ns_a = _NS_P, _NS_A

    def _set(sid, text):
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = None
        for r in sp.findall(f".//{{{ns_a}}}r"):
            rPr_e = r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0"); break
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
            p.insert(idx, r_new); break

    label_ids   = ["16","17","22","23"]   # 2026, 2025, 2024, 2023
    content_ids = ["24","12", None, None] # 2026 내용, 2024 내용
    for i, period in enumerate(periods[:4]):
        if isinstance(period, dict):
            if i < len(label_ids): _set(label_ids[i], period.get("label",""))
            if i < len(content_ids) and content_ids[i]:
                body_text = period.get("content", period.get("items",""))
                if isinstance(body_text, list): body_text = "\n".join(body_text)
                _set(content_ids[i], _truncate_to_lines(str(body_text), 2_700_000, 12, 8))

    _clear_residual_placeholders(root)
    _write_xml(root, xml_path)


def _edit_slide31(xml_path: Path, slide_plan: dict) -> None:
    """slide31 (분기별 Q1→Q4): Zone1+2 공통 + Zone3 Q레이블."""
    content  = slide_plan.get("content", {})
    quarters = content.get("quarters", content.get("body", {}).get("quarters", []) if isinstance(content.get("body"), dict) else [])
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide31.xml")
    import copy as _copy; ns_p, ns_a = _NS_P, _NS_A
    def _set(sid, text):
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = next((_copy.deepcopy(r.find(f"{{{ns_a}}}rPr")) for r in sp.findall(f".//{{{ns_a}}}r") if r.find(f"{{{ns_a}}}rPr") is not None), None)
        if orig_rPr is not None: orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr"); idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text; p.insert(idx, r_new); break
    for i, (lid, q) in enumerate(zip(["23","24","27","28"], quarters)):
        _set(lid, q.get("label", f"Q{i+1}") if isinstance(q, dict) else str(q))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide33(xml_path: Path, slide_plan: dict) -> None:
    """slide33 (분기별 변형): Zone1+2 공통 + Zone3 Q레이블."""
    content  = slide_plan.get("content", {})
    quarters = content.get("quarters", content.get("body", {}).get("quarters", []) if isinstance(content.get("body"), dict) else [])
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide33.xml")
    import copy as _copy; ns_p, ns_a = _NS_P, _NS_A
    def _set(sid, text):
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = next((_copy.deepcopy(r.find(f"{{{ns_a}}}rPr")) for r in sp.findall(f".//{{{ns_a}}}r") if r.find(f"{{{ns_a}}}rPr") is not None), None)
        if orig_rPr is not None: orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr"); idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text; p.insert(idx, r_new); break
    for i, (lid, q) in enumerate(zip(["35","36","37","38"], quarters)):
        _set(lid, q.get("label", f"Q{i+1}") if isinstance(q, dict) else str(q))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide13(xml_path: Path, slide_plan: dict) -> None:
    """slide13 (3열 아이콘카드): Zone1+2 공통 + Zone3 3열(제목/설명/아이콘안내)."""
    content = slide_plan.get("content", {})
    body    = content.get("body", {})
    items   = (content.get("items") or content.get("bullets")
               or (body.get("items") if isinstance(body, dict) else None) or [])[:3]
    descs   = content.get("descriptions", body.get("descriptions", []) if isinstance(body, dict) else [])
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide13.xml")
    import copy as _copy; ns_p, ns_a = _NS_P, _NS_A
    def _set(sid, text):
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = next((_copy.deepcopy(r.find(f"{{{ns_a}}}rPr")) for r in sp.findall(f".//{{{ns_a}}}r") if r.find(f"{{{ns_a}}}rPr") is not None), None)
        if orig_rPr is not None: orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr"); idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text; p.insert(idx, r_new); break
    # 설명 shape ID는 x좌표 기준 컬럼 순서로 매핑: 62=col1, 38=col2, 69=col3
    # (기존 ["38","62","69"]는 col2/col1이 뒤바뀌어 설명이 한 칸씩 밀렸음)
    for i, (t_id, d_id, ic_id) in enumerate(zip(["18","19","21"],["62","38","69"],["20","24","25"])):
        label = items[i] if i < len(items) else ""
        _set(t_id, _truncate_to_lines(label, 2_000_000, 14, 2))
        _set(d_id, _truncate_to_lines(descs[i] if i < len(descs) else "", 2_000_000, 12, 4))
        # 아이콘 칸: placeholder 텍스트를 쓰지 않고 비워둠 (템플릿 아이콘 영역 유지)
        # — 과거 "[아이콘: ...]" 텍스트가 슬라이드에 그대로 노출되는 버그 수정
        _set(ic_id, "")
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide21(xml_path: Path, slide_plan: dict) -> None:
    """slide21 (이미지+2블록): Zone1+2 공통 + Zone3 bullet+이미지안내."""
    content = slide_plan.get("content", {})
    body    = content.get("body", {})
    bullets = (content.get("bullets") or content.get("items")
               or (body.get("bullets") if isinstance(body, dict) else None) or [])
    body_text = content.get("section_desc", "")
    title     = slide_plan.get("title", "")
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide21.xml")
    import copy as _copy; ns_p, ns_a = _NS_P, _NS_A
    def _set(sid, text):
        sp = _find_shape_by_id(root, sid)
        if sp is None: return
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None: return
        orig_rPr = next((_copy.deepcopy(r.find(f"{{{ns_a}}}rPr")) for r in sp.findall(f".//{{{ns_a}}}r") if r.find(f"{{{ns_a}}}rPr") is not None), None)
        if orig_rPr is not None: orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr"); idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text; p.insert(idx, r_new); break
    _set("12", _truncate_to_lines(bullets[0] if bullets else body_text, 3_000_000, 14, 2))
    _set("14", _truncate_to_lines(bullets[1] if len(bullets) > 1 else body_text, 2_500_000, 12, 3))
    _set("15", _truncate_to_lines(bullets[2] if len(bullets) > 2 else "", 2_500_000, 12, 3))
    # 이미지 칸: placeholder 텍스트를 쓰지 않고 비워둠 (노출 버그 수정)
    _set("18", "")
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _slide_set_helper(root, ns_p, ns_a, sid, text):
    """편집기 공통 shape 텍스트 설정 헬퍼."""
    import copy as _copy
    sp = _find_shape_by_id(root, sid)
    if sp is None: return
    txBody = sp.find(f"{{{ns_p}}}txBody")
    if txBody is None: return
    orig_rPr = next((_copy.deepcopy(r.find(f"{{{ns_a}}}rPr")) for r in sp.findall(f".//{{{ns_a}}}r")
                     if r.find(f"{{{ns_a}}}rPr") is not None), None)
    if orig_rPr is not None: orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
    import copy as _copy2
    for p in txBody.findall(f"{{{ns_a}}}p"):
        for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
        end = p.find(f"{{{ns_a}}}endParaRPr"); idx = list(p).index(end) if end is not None else len(p)
        r_new = ET.Element(f"{{{ns_a}}}r")
        if orig_rPr: r_new.append(_copy2.deepcopy(orig_rPr))
        ET.SubElement(r_new, f"{{{ns_a}}}t").text = text; p.insert(idx, r_new); break


def _edit_slide35(xml_path: Path, slide_plan: dict) -> None:
    """slide35 (Before→After 버블): Zone1+2 공통 + Zone3 버블 키워드."""
    content = slide_plan.get("content", {})
    body    = content.get("body", {})
    before  = (content.get("before") or (body.get("before") if isinstance(body,dict) else None) or content.get("as_is",[]))[:3]
    after   = (content.get("after")  or (body.get("after")  if isinstance(body,dict) else None) or content.get("to_be",[]))[:4]
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide35.xml")
    ns_p, ns_a = _NS_P, _NS_A
    _slide_set_helper(root, ns_p, ns_a, "39", content.get("before_label","Before"))
    _slide_set_helper(root, ns_p, ns_a, "38", content.get("after_label","After"))
    for i, sid in enumerate(["13","16","17"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(before[i] if i<len(before) else "",1_500_000,16,2))
    for i, sid in enumerate(["18","19","21","35"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(after[i] if i<len(after) else "",1_500_000,16,2))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide36(xml_path: Path, slide_plan: dict) -> None:
    """slide36 (As-is/To-be 벤다이어그램): Zone1+2 공통 + Zone3 벤다이어그램 키워드."""
    content = slide_plan.get("content", {})
    body    = content.get("body", {})
    as_is   = (content.get("as_is")  or (body.get("as_is")  if isinstance(body,dict) else None) or content.get("before",[]))[:4]
    to_be   = (content.get("to_be")  or (body.get("to_be")  if isinstance(body,dict) else None) or content.get("after", []))[:4]
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide36.xml")
    ns_p, ns_a = _NS_P, _NS_A
    for i, sid in enumerate(["13","15","16","21"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(as_is[i] if i<len(as_is) else "",1_500_000,16,2))
    for i, sid in enumerate(["22","23","7","27"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(to_be[i] if i<len(to_be) else "",1_500_000,16,2))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide39(xml_path: Path, slide_plan: dict) -> None:
    """slide39 (3행 상세 흐름도): Zone1+2 공통 + Zone3 keyword/solution/detail."""
    content   = slide_plan.get("content", {})
    body      = content.get("body", {})
    keywords  = (content.get("keywords")  or (body.get("keywords")  if isinstance(body,dict) else None) or [])[:3]
    solutions = (content.get("solutions") or (body.get("solutions") if isinstance(body,dict) else None) or [])[:3]
    details   = (content.get("details")   or (body.get("details")   if isinstance(body,dict) else None) or [])[:3]
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide39.xml")
    ns_p, ns_a = _NS_P, _NS_A
    for i, sid in enumerate(["13","14","15"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(keywords[i] if i<len(keywords) else "",1_500_000,12,2))
    for i, sid in enumerate(["7","10","11"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(solutions[i] if i<len(solutions) else "",2_000_000,12,3))
    for i, sid in enumerate(["16","17","18"]):
        _slide_set_helper(root, ns_p, ns_a, sid, _truncate_to_lines(details[i] if i<len(details) else "",2_000_000,12,3))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide24(xml_path: Path, slide_plan: dict) -> None:
    """slide24 (2블록 텍스트, 이미지 없음):
    ID=8(제목), ID=7(본문1/bullets), ID=10(본문2/body).
    이미지 없는 순수 텍스트 슬라이드 — 가장 범용적."""
    import copy as _copy
    ns_p, ns_a = _NS_P, _NS_A
    title   = slide_plan.get("title", "")
    content = slide_plan.get("content", {})
    bullets = content.get("bullets") or content.get("items") or []
    body    = content.get("body", "")

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return

    def _set_text(sp, text):
        txBody = sp.find(f"{{{ns_p}}}txBody") if sp is not None else None
        if txBody is None: return
        orig_rPr = None
        for r in sp.findall(f".//{{{ns_a}}}r"):
            rPr_e = r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
                break
        for p in txBody.findall(f"{{{ns_a}}}p"):
            for r in p.findall(f"{{{ns_a}}}r"): p.remove(r)
            end = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end) if end is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
            p.insert(idx, r_new)
            break

    def _set_bullets(sp, items):
        txBody = sp.find(f"{{{ns_p}}}txBody") if sp is not None else None
        if txBody is None: return
        orig_rPr = None
        for r in sp.findall(f".//{{{ns_a}}}r"):
            rPr_e = r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang","ko-KR"); orig_rPr.set("dirty","0")
                break
        paras = txBody.findall(f"{{{ns_a}}}p")
        for i, para in enumerate(paras):
            for r in para.findall(f"{{{ns_a}}}r"): para.remove(r)
            text = items[i] if i < len(items) else ""
            end = para.find(f"{{{ns_a}}}endParaRPr")
            idx = list(para).index(end) if end is not None else len(para)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr: r_new.append(_copy.deepcopy(orig_rPr))
            ET.SubElement(r_new, f"{{{ns_a}}}t").text = text
            para.insert(idx, r_new)

    sp8  = _find_shape_by_id(root, "8")
    sp7  = _find_shape_by_id(root, "7")
    sp10 = _find_shape_by_id(root, "10")

    _set_text(sp8, title)
    if bullets:
        _set_bullets(sp7, bullets)
    elif body:
        _set_text(sp7, body)
    if body and sp10:
        _set_text(sp10, body)

    _clear_residual_placeholders(root)
    _write_xml(root, xml_path)


def _edit_slide30(xml_path: Path, slide_plan: dict) -> None:
    """slide30 (4단계 스텝): Zone1+2 공통 + Zone3 Step1~4 박스."""
    content = slide_plan.get("content", {})
    body    = content.get("body", {})
    steps   = (content.get("steps") or (body.get("steps") if isinstance(body,dict) else None)
               or content.get("bullets") or content.get("items") or [])
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide30.xml")
    ns_p, ns_a = _NS_P, _NS_A
    for i, sid in enumerate(["28","29","30","31"]):
        _slide_set_helper(root, ns_p, ns_a, sid,
                          _truncate_to_lines(steps[i] if i<len(steps) else "",2_000_000,16,2))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide32(xml_path: Path, slide_plan: dict) -> None:
    """slide32 (타임라인/텍스트+배너): Zone1+2 공통 + Zone3 body wide text."""
    content = slide_plan.get("content", {})
    body    = content.get("body") or "\n".join(content.get("bullets", []))
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide32.xml")
    if body:
        _slide_set_helper(root, _NS_P, _NS_A, "26",
                          _truncate_to_lines(body, 4_000_000, 12, 8))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


def _edit_slide38(xml_path: Path, slide_plan: dict) -> None:
    """slide38 (3행 흐름도 keyword→Solution→Service): Zone1+2 공통 + Zone3 흐름 박스."""
    content   = slide_plan.get("content", {})
    body      = content.get("body", {})
    keywords  = (content.get("keywords")  or (body.get("keywords")  if isinstance(body,dict) else None) or content.get("bullets",[]))[:3]
    solutions = (content.get("solutions") or (body.get("solutions") if isinstance(body,dict) else None) or content.get("items",  []))[:3]
    # 플랜이 'details'로 생성하는 경우가 많아 services 폴백으로 받음 (빈 3번째 컬럼 방지)
    services  = (content.get("services")  or content.get("details")
                 or (body.get("services") if isinstance(body,dict) else None)
                 or (body.get("details")  if isinstance(body,dict) else None) or [])[:3]
    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError: return
    _apply_common_zones(root, slide_plan, "slide38.xml")
    ns_p, ns_a = _NS_P, _NS_A
    for i, sid in enumerate(["13","14","15"]):
        _slide_set_helper(root, ns_p, ns_a, sid,
                          _truncate_to_lines(keywords[i] if i<len(keywords) else "",1_500_000,12,2))
    for i, sid in enumerate(["7","10","11"]):
        _slide_set_helper(root, ns_p, ns_a, sid,
                          _truncate_to_lines(solutions[i] if i<len(solutions) else "",2_000_000,12,3))
    for i, sid in enumerate(["19","24","25"]):
        _slide_set_helper(root, ns_p, ns_a, sid,
                          _truncate_to_lines(services[i] if i<len(services) else "",2_000_000,12,2))
    _clear_residual_placeholders(root); _write_xml(root, xml_path)


# ── 존 맵 기반 제너릭 편집기 ──────────────────────────────────────

# 본문구역 하위 존 role → plan content 필드 + 채움 규칙
#   key      : content에서 리스트를 꺼낼 필드(우선순위)
#   maxlines : 셀 줄 수
#   empty    : True면 비워둠(아이콘 자리)
_ZONE_FILL_RULES: dict[str, dict] = {
    "item_titles":   {"keys": ["items", "item_titles", "keywords"], "cx": 2_400_000, "pt": 14, "lines": 2},
    "item_descs":    {"keys": ["descriptions", "descs", "bullets", "details"], "cx": 2_400_000, "pt": 12, "lines": 4},
    "image_slots":   {"keys": ["image_descriptions", "images"], "cx": 2_800_000, "pt": 11, "lines": 4,
                      "prefix": "🖼 "},   # 이미지 자리 → 어떤 이미지인지 설명 텍스트
    "icon_slots":    {"empty": True},
    "insights":      {"keys": ["insights", "insight"], "cx": 2_400_000, "pt": 11, "lines": 3},
    "keywords":      {"keys": ["keywords", "items"], "cx": 1_800_000, "pt": 12, "lines": 2},
    "flow_keyword":  {"keys": ["keywords"], "cx": 2_400_000, "pt": 12, "lines": 2},
    "flow_solution": {"keys": ["solutions"], "cx": 2_700_000, "pt": 12, "lines": 3},
    "flow_service":  {"keys": ["services", "details"], "cx": 3_500_000, "pt": 12, "lines": 2},
    "detail_items":  {"keys": ["details", "sub_items"], "cx": 1_600_000, "pt": 11, "lines": 5},
    "steps":         {"keys": ["steps"], "cx": 2_200_000, "pt": 12, "lines": 2},
    "sub_titles":    {"keys": ["sub_titles", "callouts", "labels"], "cx": 2_500_000, "pt": 12, "lines": 2},
    "explains":      {"keys": ["explains", "descriptions"], "cx": 2_500_000, "pt": 11, "lines": 3},
    "banner":        {"keys": ["banner", "insight", "body"], "cx": 9_500_000, "pt": 14, "lines": 2},
    "sub_heading":   {"keys": ["sub_heading", "subtitle"], "cx": 9_000_000, "pt": 14, "lines": 1},
}


def _coerce_list(val) -> list[str]:
    if val is None:
        return []
    if isinstance(val, str):
        return [val]
    if isinstance(val, list):
        out = []
        for v in val:
            if isinstance(v, dict):
                out.append(v.get("content") or v.get("text") or v.get("label") or "")
            else:
                out.append(str(v))
        return out
    return [str(val)]


def _edit_zonemap_slide(xml_path: Path, slide_plan: dict) -> None:
    """존 맵(layout_zone_map.json)을 읽어 본문구역 하위 존을 채우는 제너릭 편집기.
    전용 편집기가 없는 모든 본문 슬라이드(이미지 그리드 등)에 사용."""
    template_file = slide_plan.get("template_file", "")
    z = _zone(template_file)
    if not z:
        return _edit_layout_slide(xml_path, slide_plan)
    content = slide_plan.get("content", {})
    body = content.get("body", {})
    bdict = body if isinstance(body, dict) else {}

    try:
        tree = ET.parse(xml_path); root = tree.getroot()
    except ET.ParseError:
        return
    _apply_common_zones(root, slide_plan, template_file)
    ns_p, ns_a = _NS_P, _NS_A

    def fetch(keys: list[str]) -> list[str]:
        for k in keys:
            v = content.get(k)
            if not v and isinstance(bdict, dict):
                v = bdict.get(k)
            if v:
                return _coerce_list(v)
        return []

    body_zones = z.get("body", {})
    for role, ids in body_zones.items():
        if not ids:
            continue
        rule = _ZONE_FILL_RULES.get(role)
        if rule is None:
            continue
        if rule.get("empty"):
            for sid in ids:
                _slide_set_helper(root, ns_p, ns_a, sid, "")
            continue
        if role == "charts":
            continue  # 차트는 별도 처리(Excel)
        values = fetch(rule["keys"])
        prefix = rule.get("prefix", "")
        for i, sid in enumerate(ids):
            txt = values[i] if i < len(values) else ""
            if txt:
                txt = prefix + str(txt)
                txt = _truncate_to_lines(txt, rule["cx"], rule["pt"], rule["lines"])
            _slide_set_helper(root, ns_p, ns_a, sid, txt)

    _clear_residual_placeholders(root)
    _write_xml(root, xml_path)


# ── 레이아웃 placeholder 기반 본문 편집기 (레거시 폴백) ──────────────

def _edit_layout_slide(xml_path: Path, slide_plan: dict) -> None:
    """
    GS Neotek 템플릿 slide22 전용 편집기.

    slide22 레이아웃:
      ID=8  — 대제목 헤더 (상단 왼쪽)
      ID=9  — 중제목 헤더 (상단 중앙)
      ID=10 — 이미지 영역 (회색 박스, 편집 안 함)
      ID=11 — 핵심 설명 bullet 3개 (하단, sz=1600) ← 메인 콘텐츠
      ID=14 — 좌측 사이드바 레이블 (짧은 키워드)
      ID=16 — 좌측 사이드바 설명 (한 문장)

    편집 전략:
      - ID=8  → slide title (대제목)
      - ID=11 → bullets (최대 3개, 이후는 말줄임)
      - ID=14 → bullets[0] 키워드 요약 (30자 이내)
      - ID=16 → body 한 줄 요약
      - 나머지 → placeholder 클리어
    """
    import copy as _copy
    ns_p = _NS_P
    ns_a = _NS_A
    title   = slide_plan.get("title", "")
    content = slide_plan.get("content", {})
    bullets = (content.get("bullets") or content.get("items") or
               content.get("steps") or [])
    body    = content.get("body", "")

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return

    def _set_placeholder(sp: ET.Element, text: str) -> None:
        """shape의 첫 번째 paragraph를 text로 교체 (rPr 보존)."""
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            return
        for p in txBody.findall(f"{{{ns_a}}}p"):
            first_r = p.find(f"{{{ns_a}}}r")
            orig_rPr = None
            if first_r is not None:
                rPr_e = first_r.find(f"{{{ns_a}}}rPr")
                if rPr_e is not None:
                    orig_rPr = _copy.deepcopy(rPr_e)
                    orig_rPr.set("lang", "ko-KR")
                    orig_rPr.set("dirty", "0")
            for r in p.findall(f"{{{ns_a}}}r"):
                p.remove(r)
            end_rpr = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end_rpr) if end_rpr is not None else len(p)
            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr is not None:
                r_new.append(orig_rPr)
            else:
                ET.SubElement(r_new, f"{{{ns_a}}}rPr", lang="ko-KR", dirty="0")
            t = ET.SubElement(r_new, f"{{{ns_a}}}t")
            t.text = text
            p.insert(idx, r_new)
            break  # 첫 paragraph만

    def _set_bullets(sp: ET.Element, items: list[str]) -> None:
        """shape의 paragraph들을 bullet 항목으로 채운다."""
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            return
        # 기존 rPr 보존
        orig_rPr = None
        for r in sp.findall(f".//{{{ns_a}}}r"):
            rPr_e = r.find(f"{{{ns_a}}}rPr")
            if rPr_e is not None:
                orig_rPr = _copy.deepcopy(rPr_e)
                orig_rPr.set("lang", "ko-KR")
                orig_rPr.set("dirty", "0")
                break

        paras = txBody.findall(f"{{{ns_a}}}p")
        for i, item in enumerate(items):
            if i < len(paras):
                p = paras[i]
                for r in p.findall(f"{{{ns_a}}}r"):
                    p.remove(r)
                end_rpr = p.find(f"{{{ns_a}}}endParaRPr")
                idx = list(p).index(end_rpr) if end_rpr is not None else len(p)
                r_new = ET.Element(f"{{{ns_a}}}r")
                if orig_rPr is not None:
                    r_new.append(_copy.deepcopy(orig_rPr))
                t = ET.SubElement(r_new, f"{{{ns_a}}}t")
                t.text = item
                p.insert(idx, r_new)
            else:
                # 새 paragraph 추가
                p_new = ET.SubElement(txBody, f"{{{ns_a}}}p")
                r_new = ET.SubElement(p_new, f"{{{ns_a}}}r")
                if orig_rPr is not None:
                    r_new.append(_copy.deepcopy(orig_rPr))
                t = ET.SubElement(r_new, f"{{{ns_a}}}t")
                t.text = item
                ET.SubElement(p_new, f"{{{ns_a}}}endParaRPr", lang="ko-KR")
        # 초과 paragraph 비우기
        for p in paras[len(items):]:
            for r in p.findall(f"{{{ns_a}}}r"):
                for t in r.findall(f"{{{ns_a}}}t"):
                    t.text = ""

    # slide22 실제 폭 (slideLayout5에서 측정):
    #   ID=14 사이드바 레이블: 173pt → 2,197,100 EMU (한글 ~12자/줄)
    #   ID=16 사이드바 설명:   173pt → 2,197,100 EMU (12pt, 한글 ~17자/줄)
    #   ID=11 하단 bullet:     750pt → 9,525,000 EMU (16pt, 한글 ~50자/줄)
    _CX_SIDEBAR = 2_197_100
    _CX_BULLETS = 9_525_000

    # ── ID=8: 대제목 헤더 (넓은 영역 → 전체 제목) ──────────────
    sp8 = _find_shape_by_id(root, _CONTENT_TITLE_ID)
    if sp8 is not None:
        _set_placeholder(sp8, title)

    # ── ID=11: 하단 bullet (넓은 영역 → 각 bullet 1줄 요약) ─────
    sp11 = _find_shape_by_id(root, _CONTENT_BULLETS_ID)
    if sp11 is not None:
        content_list = bullets[:3] if bullets else (
            [s.strip() for s in body.split("·") if s.strip()][:3] if body else []
        )
        # 각 bullet을 750pt 폭에 맞게 1줄로 정리
        trimmed = [_truncate_to_lines(b, _CX_BULLETS, 16, max_lines=1)
                   for b in content_list]
        if trimmed:
            _set_bullets(sp11, trimmed)

    # ── ID=14: 사이드바 레이블 (좁은 영역 → 핵심 키워드 2줄) ────
    sp14 = _find_shape_by_id(root, "14")
    if sp14 is not None:
        # bullets[0]에서 핵심 키워드만 뽑아서 2줄 이내로 요약
        raw = bullets[0] if bullets else title
        label = _truncate_to_lines(raw, _CX_SIDEBAR, 20, max_lines=2)
        _set_placeholder(sp14, label)

    # ── ID=16: 사이드바 설명 (좁은 영역 → body 3줄 요약) ─────────
    sp16 = _find_shape_by_id(root, "16")
    if sp16 is not None:
        raw16 = body or (bullets[1] if len(bullets) > 1 else "")
        desc = _truncate_to_lines(raw16, _CX_SIDEBAR, 12, max_lines=3)
        _set_placeholder(sp16, desc)

    # ── 나머지: placeholder 텍스트만 클리어 ─────────────────────
    _clear_residual_placeholders(root)

    _write_xml(root, xml_path)


# ── 범용 본문 슬라이드 편집기 ─────────────────────────────────

_PLACEHOLDER_TEXTS = re.compile(
    r"(lorem|작성해주세요|중제목을|대제목을|제목을|설명을|설명 타이틀|"
    r"상세 설명|이미지를|01 제목|insight / definition|image|"
    r"\[insert|\bTODO\b|ipsum dolor|"
    r"이미지/\s*영상|이미지/영상|image placeholder|"
    r"nibh euismod|tincidunt ut|elit,\s*sed diam|nonummy|"
    r"짧은 텍스트에 사용|텍스트에 사용|1\.1\s|1\.2\s|1\.3\s|"
    r"insight /|conclusion /|definition|\bicon\b|image\b|"
    r"핵심 설명을 작성|부분 설명|설명 타이틀|01 핵심|02 핵심|03 핵심|"
    r"ppt 대제목|01 중제목|01 컨텐츠|중제목 작성|대제목 작성|컨텐츠 작성|"
    r"solution 0[123]|sevice 0[123]|keyword\b|step[1-4]|"
    r"1\.4\s|텍스트 작성해주세요|텍스트 길어질|작성하여 사용|"
    r"\[아이콘|\[이미지|관련 도식|아이콘 이미지|"
    r"아래 확장|최대 [0-9]+\s*줄|여기에 입력|내용을 입력|"
    r"항목 0[1-9]|항목0[1-9]|세부 항목|detail 0[1-9])",
    re.IGNORECASE,
)


def _clear_residual_placeholders(root: ET.Element) -> bool:
    """
    편집 후 남아있는 placeholder 텍스트를 빈 문자열로 교체한다.
    편집된 콘텐츠(한국어 실제 내용)는 건드리지 않는다.
    수정이 일어나면 True 반환.
    """
    ns_p = _NS_P
    ns_a = _NS_A
    modified = False

    for sp in root.findall(f".//{{{ns_p}}}sp"):
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            continue
        for p in txBody.findall(f".//{{{ns_a}}}p"):
            # paragraph 내 모든 run의 텍스트를 합쳐서 판단
            full_text = "".join(
                t.text or "" for t in p.findall(f".//{{{ns_a}}}t")
            )
            if _PLACEHOLDER_TEXTS.search(full_text):
                for t in p.findall(f".//{{{ns_a}}}t"):
                    t.text = ""
                modified = True
            # 개별 run도 검사 (분산 lorem 대응)
            else:
                for r in p.findall(f"{{{ns_a}}}r"):
                    t = r.find(f"{{{ns_a}}}t")
                    if t is not None and t.text and _PLACEHOLDER_TEXTS.search(t.text):
                        t.text = ""
                        modified = True

    return modified


def edit_content_slide(xml_path: Path, slide_plan: dict) -> None:
    """
    cover/toc 외 모든 role의 슬라이드를 편집한다.
    텍스트 shape을 Y 좌표 순으로 정렬한 뒤 content 필드를 매핑.

    매핑 우선순위:
      1번 shape(최상단): title
      2번 shape: subtitle / body / description
      이후 shape: bullets / items / steps (항목 순서대로)
    """
    import copy as _copy

    ns_a = _NS_A
    ns_p = _NS_P

    title   = slide_plan.get("title", "")
    content = slide_plan.get("content", {})

    # 본문 콘텐츠 우선순위 추출
    body_text = (
        content.get("body")
        or content.get("description")
        or content.get("subtitle")
        or ""
    )
    list_items: list[str] = (
        content.get("bullets")
        or content.get("items")
        or content.get("steps")
        or []
    )

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return

    # 텍스트 shape 수집 → Y 좌표 정렬
    shapes_with_y: list[tuple[int, ET.Element]] = []
    for sp in root.findall(f".//{{{ns_p}}}sp"):
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            continue
        xfrm = sp.find(f".//{{{ns_a}}}xfrm")
        off  = xfrm.find(f"{{{ns_a}}}off") if xfrm is not None else None
        y    = int(off.get("y", 0)) if off is not None else 0
        shapes_with_y.append((y, sp))

    shapes_with_y.sort(key=lambda x: x[0])
    shapes = [sp for _, sp in shapes_with_y]

    def _set_text(sp: ET.Element, text: str) -> None:
        """shape의 첫 번째 paragraph의 run 교체 (rPr 보존)."""
        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            return
        for p in txBody.findall(f"{{{ns_a}}}p"):
            # 기존 rPr 복사
            first_r = p.find(f"{{{ns_a}}}r")
            orig_rPr = None
            if first_r is not None:
                rPr_e = first_r.find(f"{{{ns_a}}}rPr")
                if rPr_e is not None:
                    orig_rPr = _copy.deepcopy(rPr_e)
                    orig_rPr.set("lang", "ko-KR")
                    orig_rPr.set("dirty", "0")

            for r in p.findall(f"{{{ns_a}}}r"):
                p.remove(r)

            end_rpr = p.find(f"{{{ns_a}}}endParaRPr")
            idx = list(p).index(end_rpr) if end_rpr is not None else len(p)

            r_new = ET.Element(f"{{{ns_a}}}r")
            if orig_rPr is not None:
                r_new.append(orig_rPr)
            else:
                ET.SubElement(r_new, f"{{{ns_a}}}rPr", lang="ko-KR", dirty="0")
            t_new = ET.SubElement(r_new, f"{{{ns_a}}}t")
            t_new.text = text
            p.insert(idx, r_new)
            break  # 첫 번째 paragraph만

    # shape 0: 제목
    if shapes:
        _set_text(shapes[0], title)

    # shape 1: 본문 단일 텍스트
    if len(shapes) > 1 and body_text:
        _set_text(shapes[1], body_text)

    # shape 2+: 리스트 항목
    if list_items:
        start = 2 if body_text and len(shapes) > 2 else 1
        for idx, item in enumerate(list_items):
            si = start + idx
            if si < len(shapes):
                _set_text(shapes[si], item)
            else:
                break  # shape 부족 시 중단

    _write_xml(root, xml_path)


# ── 병렬 Plan 생성 (2-Phase) ─────────────────────────────────

def _build_claude_client():
    """
    Claude 클라이언트 빌드 (generate_plan_with_claude와 동일한 우선순위 로직).
    Vertex > Bedrock > Anthropic 순서, ANTHROPIC_VERTEX_PROJECT_ID가 있으면 Vertex 우선.
    """
    import os, anthropic
    vertex_proj   = os.environ.get("ANTHROPIC_VERTEX_PROJECT_ID")
    vertex_region = os.environ.get("CLOUD_ML_REGION")
    aws_region    = os.environ.get("AWS_REGION")
    api_key       = os.environ.get("ANTHROPIC_API_KEY")
    explicit      = os.environ.get("PPT_SKILL_BACKEND", "auto")

    if explicit == "vertex":
        use_vertex, use_bedrock = True, False
    elif explicit == "bedrock":
        use_vertex, use_bedrock = False, True
    elif explicit == "anthropic":
        use_vertex, use_bedrock = False, False
    else:
        # generate_plan_with_claude와 동일한 auto 감지 로직
        use_vertex  = bool(vertex_proj or (vertex_region and not aws_region))
        use_bedrock = bool(aws_region) and not use_vertex

    if use_vertex:
        return anthropic.AnthropicVertex(
            project_id=vertex_proj or "",
            region=vertex_region or "us-east5"
        ), None
    if use_bedrock:
        try:
            import boto3
            return None, boto3.client("bedrock-runtime", region_name=aws_region or "us-east-1")
        except ImportError:
            pass  # boto3 없으면 Anthropic 폴백
    if api_key:
        return anthropic.Anthropic(api_key=api_key), None
    return None, None


def _call_api(client, bedrock, system: str, user: str, max_tokens: int = 4096) -> str:
    """백엔드 무관 단일 API 호출."""
    import json as _json
    model = __import__("os").environ.get("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")
    messages = [{"role": "user", "content": user}]
    if bedrock:
        body = _json.dumps({"anthropic_version": "bedrock-2023-05-31",
                            "max_tokens": max_tokens, "system": system, "messages": messages})
        resp = bedrock.invoke_model(
            modelId="us.anthropic.claude-sonnet-4-5-20250929-v1:0", body=body)
        return _json.loads(resp["body"].read())["content"][0]["text"].strip()
    resp = client.messages.create(model=model, max_tokens=max_tokens,
                                   system=system, messages=messages)
    return resp.content[0].text.strip()


def _parse_json_response(raw: str) -> dict | list | None:
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def generate_outline(topic: str, audience: str, n_slides: int,
                     slide_info: list[dict], memory: dict,
                     constraints: list[str]) -> list[dict] | None:
    """
    Phase 1: 슬라이드 구조(아웃라인)만 생성 — 빠른 1회 호출.
    각 슬라이드의 role, template_file, title, key_points만 반환.
    """
    try:
        client, bedrock = _build_claude_client()
    except Exception:
        return None
    if client is None and bedrock is None:
        return None

    available = {s["file"]: s["texts"][:2] for s in slide_info}
    layout_hints = memory.get("slide_layout_hints", {})
    constraint_str = "\n".join(f"- {c}" for c in constraints) if constraints else "없음"

    system = (
        "당신은 PPT 구조 설계자입니다. 슬라이드 아웃라인만 JSON 배열로 출력합니다.\n"
        "각 항목: {index, role, template_file, title, key_points: [str]}\n"
        "role 목록: cover|toc|section|content|three_col|steps|timeline|closing\n"
        "JSON 배열만 출력, 다른 텍스트 없음."
    )
    user = (
        f"주제: {topic}\n청중: {audience}\n슬라이드 수: {n_slides}\n\n"
        f"사용 가능 템플릿: {list(available.keys())[:10]}\n"
        f"레이아웃 힌트: {layout_hints}\n"
        f"품질 제약 조건:\n{constraint_str}\n\n"
        f"{n_slides}장 아웃라인을 JSON 배열로 작성하세요."
    )

    try:
        raw = _call_api(client, bedrock, system, user, max_tokens=2048)
        result = _parse_json_response(raw)
        if isinstance(result, list) and len(result) > 0:
            return result
    except Exception as e:
        print(f"  ⚠ outline 생성 실패: {e}")
    return None


def _generate_single_slide_content(args: tuple) -> dict:
    """
    Phase 2 단일 슬라이드 콘텐츠 생성 — ThreadPoolExecutor에서 호출.
    args: (slide_outline, topic, audience, constraints, client, bedrock)
    """
    slide_outline, topic, audience, constraints, client, bedrock = args
    role  = slide_outline.get("role", "content")
    title = slide_outline.get("title", "")
    keys  = slide_outline.get("key_points", [])

    schema = PLAN_CONTENT_SCHEMA.get(role, {})
    req_fields = schema.get("required", []) + schema.get("recommended", [])
    constraint_str = "\n".join(f"- {c}" for c in constraints) if constraints else "없음"

    system = (
        "당신은 PPT 콘텐츠 전문가입니다. 슬라이드 1장의 content 필드만 JSON으로 출력합니다.\n"
        "JSON 객체만 출력, 다른 텍스트 없음."
    )
    user = (
        f"주제: {topic}\n청중: {audience}\n"
        f"슬라이드 역할: {role}\n제목: {title}\n핵심 키워드: {keys}\n"
        f"필요 필드: {req_fields}\n"
        f"품질 제약:\n{constraint_str}\n\n"
        "위 슬라이드의 content JSON을 작성하세요. "
        "텍스트는 발표 수준으로 구체적이고 전문적이어야 합니다."
    )

    try:
        raw = _call_api(client, bedrock, system, user, max_tokens=1024)
        content = _parse_json_response(raw)
        if isinstance(content, dict):
            return {**slide_outline, "content": content}
    except Exception as e:
        pass  # 폴백: key_points를 bullets로

    # 폴백
    return {**slide_outline, "content": {"bullets": keys, "body": title}}


def generate_plan_parallel(topic: str, audience: str, n_slides: int,
                            slide_info: list[dict], memory: dict,
                            constraints: list[str]) -> dict | None:
    """
    2-Phase 병렬 Plan 생성:
      Phase 1: 아웃라인 1회 호출 (role/title/key_points)
      Phase 2: 슬라이드별 콘텐츠 병렬 호출 (ThreadPoolExecutor)
    슬라이드 수와 무관하게 일정한 응답 시간 유지.
    """
    try:
        client, bedrock = _build_claude_client()
    except Exception:
        return None
    if client is None and bedrock is None:
        return None

    print("  [병렬 생성] Phase 1: 아웃라인 생성...")
    outline = generate_outline(topic, audience, n_slides, slide_info, memory, constraints)
    if not outline:
        return None
    print(f"  ✓ 아웃라인 {len(outline)}장 완료")

    print(f"  [병렬 생성] Phase 2: {len(outline)}장 콘텐츠 병렬 생성...")
    args_list = [
        (slide, topic, audience, constraints, client, bedrock)
        for slide in outline
    ]
    slides: list[dict] = [None] * len(outline)
    max_workers = min(8, len(outline))

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(_generate_single_slide_content, args): i
            for i, args in enumerate(args_list)
        }
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                slides[idx] = future.result()
            except Exception:
                slides[idx] = {**outline[idx], "content": {}}

    slides = [s for s in slides if s is not None]
    print(f"  ✓ 병렬 콘텐츠 생성 완료 ({len(slides)}장)")

    return {
        "title":    topic,
        "topic":    topic,
        "audience": audience,
        "n_slides": n_slides,
        "slides":   slides,
        "_generated_by": "parallel",
    }


# ── Vision Fix Agent (독립 검증 에이전트) ─────────────────────

def _gather_shape_info(xml_path: Path) -> list[dict]:
    """
    슬라이드 XML에서 text shape 목록을 추출한다.
    반환: [{id, y, cx, current_texts}]  (Y 오름차순 정렬)
    """
    ns_p = _NS_P
    ns_a = _NS_A
    shapes = []
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError:
        return shapes

    for sp in root.findall(f".//{{{ns_p}}}sp"):
        cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
        if cpr is None:
            continue
        xfrm = sp.find(f".//{{{ns_a}}}xfrm")
        off  = xfrm.find(f"{{{ns_a}}}off") if xfrm is not None else None
        ext  = xfrm.find(f"{{{ns_a}}}ext") if xfrm is not None else None
        y  = int(off.get("y", 0)) if off is not None else 0
        cx = int(ext.get("cx", 0)) if ext is not None else 0
        x  = int(off.get("x", 0)) if off is not None else 0

        txBody = sp.find(f"{{{ns_p}}}txBody")
        if txBody is None:
            continue
        texts = [
            "".join(t.text or "" for t in p.findall(f".//{{{ns_a}}}t"))
            for p in txBody.findall(f".//{{{ns_a}}}p")
        ]
        texts = [t for t in texts if t.strip()]

        shapes.append({
            "id":   cpr.get("id", "?"),
            "name": cpr.get("name", ""),
            "y":    y,
            "x":    x,
            "cx":   cx,
            "current_texts": texts[:3],  # 최대 3개 미리보기
        })

    return sorted(shapes, key=lambda s: s["y"])


def _call_vision_api(system: str, content: list) -> str | None:
    """Vision(멀티모달) Claude API 호출."""
    try:
        client, bedrock = _build_claude_client()
    except Exception:
        return None
    if client is None and bedrock is None:
        return None
    model = __import__("os").environ.get(
        "ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-sonnet-4-6")
    messages = [{"role": "user", "content": content}]
    try:
        if bedrock:
            import json as _j, boto3
            body = _j.dumps({
                "anthropic_version": "bedrock-2023-05-31",
                "max_tokens": 4096, "system": system, "messages": messages,
            })
            resp = bedrock.invoke_model(
                modelId="us.anthropic.claude-sonnet-4-5-20250929-v1:0", body=body)
            return _j.loads(resp["body"].read())["content"][0]["text"].strip()
        resp = client.messages.create(
            model=model, max_tokens=4096, system=system, messages=messages)
        return resp.content[0].text.strip()
    except Exception as e:
        print(f"  ⚠ Vision API 호출 실패: {e}")
        return None


def _run_vision_fix_agent(
    work_dir: Path,
    qa_images: list[str],
    plan: dict,
    slides_dir: Path,
) -> list[dict]:
    """
    독립 Vision 검증 에이전트.
    QA 이미지 + plan 의도 + shape 구조를 독립 컨텍스트에서 분석해
    슬라이드별 수술적(shape ID 기반) 수정 지시를 생성한다.

    반환 스키마:
    [
      {
        "slide_file": "slide8.xml",
        "slide_index": 3,
        "fixes": [
          {"shape_id": "4", "action": "set_paragraphs",
           "texts": ["항목1", "항목2", "항목3"]},
          {"shape_id": "2", "action": "clear"}
        ]
      }
    ]
    action 종류:
      set_text       - 단일 텍스트 설정
      set_paragraphs - 다중 paragraph 설정
      clear          - 모든 텍스트 제거
    """
    import base64, os

    if not qa_images:
        return []

    try:
        client, bedrock = _build_claude_client()
    except Exception:
        return []
    if client is None and bedrock is None:
        print("  ⚠ Vision Fix Agent: API 없음 — 건너뜀")
        return []

    # 슬라이드 인덱스 → plan 매핑
    plan_by_idx = {s["index"]: s for s in plan.get("slides", [])}
    # 이미지 순서 → slide index 매핑 (이미지는 1-indexed)
    image_paths = [p for p in qa_images if Path(p).exists()]

    system = """당신은 독립적인 PPT 슬라이드 품질 검증 에이전트입니다.
슬라이드 이미지, 계획된 콘텐츠, shape 구조를 보고 정확한 수정 지시를 생성합니다.

반드시 아래 JSON 배열 형식으로만 응답하세요:
[
  {
    "slide_file": "slideN.xml",
    "slide_index": N,
    "has_issues": true/false,
    "issue_summary": "문제 요약",
    "fixes": [
      {
        "shape_id": "ID",
        "action": "set_text|set_paragraphs|clear",
        "text": "단일 텍스트 (set_text용)",
        "texts": ["항목1", "항목2"] // set_paragraphs용
      }
    ]
  }
]

수정 원칙:
- placeholder(작성해주세요, 중제목, 대제목, lorem ipsum 등)는 반드시 제거 또는 교체
- 실제 콘텐츠(plan에 있는 내용)로만 교체
- 이미 올바른 슬라이드는 has_issues=false, fixes=[]
- shape_id는 반드시 제공된 shape 목록의 실제 ID 사용
- shape ID=14(사이드바 레이블, 173pt 폭): 텍스트 overflow가 보여도 수정 금지 — 이미 너비에 맞게 요약됨
- shape ID=16(사이드바 설명, 173pt 폭): 동일하게 수정 금지
- 사이드바는 의도적으로 짧은 요약 텍스트를 담는 좁은 영역임
- shape ID=9("01 중제목 작성", "01 컨텐츠 작성" 등이 보이면): 레이아웃 placeholder — set_text("")로 비우거나 무시할 것, 슬라이드 콘텐츠로 교체 금지
- shape ID=8("PPT 대제목 작성"이 보이면): 레이아웃 placeholder — 비울 것
- "Solution 01", "keyword", "Sevice 01", "Step1"~"Step4" 등 템플릿 안내 텍스트는 placeholder — 비울 것
- 이미 실제 콘텐츠(한국어 설명문)가 있는 shape는 수정 금지"""

    # 슬라이드별 정보 구성
    content_parts = []
    slide_contexts = []

    for i, img_path in enumerate(image_paths[:10]):
        slide_idx = i + 1
        slide_plan = plan_by_idx.get(slide_idx, {})
        slide_file = slide_plan.get("template_file", "")
        xml_path   = slides_dir / slide_file if slide_file else None
        shape_info = _gather_shape_info(xml_path) if xml_path and xml_path.exists() else []

        slide_contexts.append({
            "slide_index": slide_idx,
            "slide_file":  slide_file,
            "role":        slide_plan.get("role", "content"),
            "title":       slide_plan.get("title", ""),
            "content":     slide_plan.get("content", {}),
            "shapes":      shape_info,
        })

        b64 = base64.standard_b64encode(Path(img_path).read_bytes()).decode()
        content_parts.append({
            "type": "text",
            "text": f"=== 슬라이드 {slide_idx} ({slide_plan.get('role','')}) ===\n"
                    f"의도 제목: {slide_plan.get('title','')}\n"
                    f"의도 콘텐츠: {json.dumps(slide_plan.get('content',{}), ensure_ascii=False)[:300]}\n"
                    f"shape 구조: {json.dumps(shape_info, ensure_ascii=False)[:400]}",
        })
        content_parts.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg", "data": b64},
        })

    content_parts.append({
        "type": "text",
        "text": "각 슬라이드를 분석해 수정 지시 JSON을 반환하세요. "
                "placeholder가 없고 의도한 콘텐츠가 올바르게 표시된 슬라이드는 has_issues=false."
    })

    # 슬라이드별 개별 분석 (토큰 한도 초과 방지)
    # 전용 편집기가 처리하는 슬라이드는 Vision Fix 대상 제외
    # — toc: edit_toc_slide가 7-paragraph 구조를 완전 관리, Vision Fix가 수정하면 구조 파괴
    # — cover: edit_cover_slide가 담당
    # — closing: 편집 금지 (레이아웃에 "감사합니다." 내장)
    SKIP_ROLES = {"cover", "closing", "toc"}
    all_results = []
    for ctx in slide_contexts:
        if ctx.get("role") in SKIP_ROLES:
            continue
        img_path = image_paths[ctx["slide_index"] - 1] \
            if ctx["slide_index"] - 1 < len(image_paths) else None
        if not img_path or not Path(img_path).exists():
            continue

        b64 = base64.standard_b64encode(Path(img_path).read_bytes()).decode()
        per_slide_content = [
            {
                "type": "text",
                "text": (
                    f"슬라이드 {ctx['slide_index']} ({ctx['role']})\n"
                    f"의도 제목: {ctx['title']}\n"
                    f"의도 콘텐츠: {json.dumps(ctx['content'], ensure_ascii=False)[:300]}\n"
                    f"shape 구조: {json.dumps(ctx['shapes'], ensure_ascii=False)[:400]}"
                ),
            },
            {
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": b64},
            },
            {
                "type": "text",
                "text": (
                    "이 슬라이드에 placeholder(작성해주세요, lorem ipsum 등) 또는 레이아웃 문제가 있으면 "
                    "JSON 객체로 수정 지시를 반환하세요. 문제 없으면 {\"has_issues\": false}."
                ),
            },
        ]

        try:
            raw = _call_vision_api(system, per_slide_content)
            if not raw or not raw.strip():
                continue
            # JSON 블록 추출 (```json ... ```, { ... }, 또는 전체 텍스트)
            text = raw.strip()
            # 방법 1: ```json ... ``` 코드 블록
            m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
            if m:
                text = m.group(1)
            else:
                # 방법 2: 첫 번째 { ... } 객체 추출
                m2 = re.search(r"\{.*\}", text, re.DOTALL)
                if m2:
                    text = m2.group(0)
            parsed = json.loads(text)
            if isinstance(parsed, dict):
                # ctx의 정확한 파일명으로 강제 설정 (에이전트가 잘못 추측할 수 있음)
                parsed["slide_file"]  = ctx["slide_file"]
                parsed["slide_index"] = ctx["slide_index"]
                all_results.append(parsed)
        except (json.JSONDecodeError, Exception) as e:
            print(f"  ⚠ slide {ctx['slide_index']} Vision 분석 실패: {e}")

    issues_found = [r for r in all_results if r.get("has_issues")]
    if issues_found:
        print(f"  [Vision Fix] {len(issues_found)}/{len(all_results)}장 수정 필요")
        for r in issues_found:
            print(f"    slide {r.get('slide_index')} ({r.get('slide_file')}): "
                  f"{r.get('issue_summary', '')[:60]}")
    else:
        print(f"  [Vision Fix] {len(all_results)}장 분석 — 수정 불필요")
    return all_results


def _apply_fix_instructions(work_dir: Path, fix_instructions: list[dict]) -> bool:
    """
    Vision Fix Agent의 수정 지시를 슬라이드 XML에 적용한다.
    수정이 일어나면 True 반환.
    """
    import copy as _copy
    ns_p = _NS_P
    ns_a = _NS_A
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    any_modified = False

    for instruction in fix_instructions:
        if not instruction.get("has_issues"):
            continue
        fixes = instruction.get("fixes", [])
        if not fixes:
            continue

        slide_file = instruction.get("slide_file", "")
        xml_path   = slides_dir / slide_file
        if not xml_path.exists():
            continue

        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
        except ET.ParseError:
            continue

        modified = False
        shape_map = {}
        for sp in root.findall(f".//{{{ns_p}}}sp"):
            cpr = sp.find(f"{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr")
            if cpr is not None:
                shape_map[cpr.get("id", "")] = sp

        for fix in fixes:
            shape_id = str(fix.get("shape_id", ""))
            action   = fix.get("action", "")
            sp = shape_map.get(shape_id)
            if sp is None:
                continue

            txBody = sp.find(f"{{{ns_p}}}txBody")
            if txBody is None:
                continue

            # 기존 rPr 보존용
            first_rPr = None
            for r in sp.findall(f".//{{{ns_a}}}r"):
                rPr_e = r.find(f"{{{ns_a}}}rPr")
                if rPr_e is not None:
                    first_rPr = _copy.deepcopy(rPr_e)
                    first_rPr.set("lang", "ko-KR")
                    first_rPr.set("dirty", "0")
                    break

            def _make_para(text_str: str) -> ET.Element:
                p = ET.Element(f"{{{ns_a}}}p")
                pPr = ET.SubElement(p, f"{{{ns_a}}}pPr")
                r_new = ET.SubElement(p, f"{{{ns_a}}}r")
                if first_rPr is not None:
                    r_new.append(_copy.deepcopy(first_rPr))
                else:
                    ET.SubElement(r_new, f"{{{ns_a}}}rPr",
                                  lang="ko-KR", dirty="0")
                t = ET.SubElement(r_new, f"{{{ns_a}}}t")
                t.text = text_str
                ET.SubElement(p, f"{{{ns_a}}}endParaRPr",
                              lang="ko-KR", dirty="0")
                return p

            if action == "clear":
                for p in txBody.findall(f"{{{ns_a}}}p"):
                    for t in p.findall(f".//{{{ns_a}}}t"):
                        t.text = ""
                modified = True

            elif action == "set_text":
                text = fix.get("text", "")
                for p in txBody.findall(f"{{{ns_a}}}p"):
                    txBody.remove(p)
                txBody.append(_make_para(text))
                modified = True

            elif action == "set_paragraphs":
                texts = fix.get("texts", [])
                if not texts:
                    continue
                for p in txBody.findall(f"{{{ns_a}}}p"):
                    txBody.remove(p)
                for t in texts:
                    txBody.append(_make_para(t))
                modified = True

        if modified:
            _write_xml(root, xml_path)
            any_modified = True
            print(f"  ✓ Vision Fix 적용: {slide_file} ({len(fixes)}개 지시)")

    return any_modified


# ── 메인 루프 ────────────────────────────────────────────────

def run_ppt_generation(
    topic: str,
    template_path: Path,
    work_dir: Path,
    audience: str = "전문가",
    n_slides: int = 10,
) -> Path:
    """
    analyze_template → generate_plan → edit_slide 루프 → pack → verify
    최종 output.pptx 경로를 반환한다.
    """
    print(f"\n[PPT Generator] topic={topic!r}, slides={n_slides}, audience={audience!r}")

    # ── 언팩 ─────────────────────────────────────
    unpacked = work_dir / "unpacked"
    subprocess.run(
        [sys.executable, str(SKILL_DIR / "scripts" / "office" / "unpack.py"),
         str(template_path), str(unpacked)],
        check=True, capture_output=True,
    )
    print("  ✓ 언팩 완료")

    # ── 분석 ─────────────────────────────────────
    slide_info = analyze_template(work_dir)
    print(f"  ✓ 슬라이드 {len(slide_info)}개 분석 완료")

    # ── 하네스 로드 ───────────────────────────────
    memory: dict = {}
    memory_path = SKILL_DIR / "harness" / "long_term_memory.json"
    if memory_path.exists():
        memory = json.loads(memory_path.read_text())

    # Vision 분석에서 축적된 planning constraints 로드
    constraints: list[str] = memory.get("planning_constraints", [])
    known_fixes: list[dict] = memory.get("known_failure_fixes", [])
    if constraints:
        print(f"  ✓ planning constraints {len(constraints)}개 로드")
    if known_fixes:
        print(f"  ✓ known_failure_fixes {len(known_fixes)}개 로드")

    # ── 계획 생성 ─────────────────────────────────
    # 15장 초과: 병렬 2-phase 생성 / 이하: 단일 호출
    plan = None
    if n_slides > 14:
        print(f"  [병렬 모드] {n_slides}장 → 2-phase 병렬 생성")
        plan = generate_plan_parallel(topic, audience, n_slides,
                                       slide_info, memory, constraints)

    if plan is None:
        plan = generate_plan_with_claude(topic, audience, n_slides,
                                          slide_info, memory)
        if plan:
            print("  ✓ Claude API 단일 호출로 계획 생성")
        else:
            plan = generate_plan(topic, audience, n_slides, slide_info)
            print("  ✓ 규칙 기반 계획 생성 (폴백)")

    # ── Closing 슬라이드 보장 ─────────────────────
    # 총 N장 = 표지(1) + 목차(1) + 본문(N-3) + 마무리(1)
    # Claude가 closing을 빠뜨리면 자동 추가
    has_closing = any(s.get("role") == "closing" for s in plan.get("slides", []))
    if not has_closing:
        closing_idx = len(plan["slides"]) + 1
        plan["slides"].append({
            "index": closing_idx,
            "template_file": "slide46.xml",
            "role": "closing",
            "title": "감사합니다",
            "content": {}
        })
        # 초과 content 슬라이드 제거 (총 n_slides 유지)
        content_slides = [s for s in plan["slides"] if s.get("role") not in ("cover","toc","closing")]
        max_content = n_slides - 3  # cover+toc+closing
        if len(content_slides) > max_content:
            to_remove = len(content_slides) - max_content
            removed = 0
            new_slides = []
            for s in reversed(plan["slides"]):
                if s.get("role") not in ("cover","toc","closing") and removed < to_remove:
                    removed += 1
                else:
                    new_slides.insert(0, s)
            plan["slides"] = new_slides
        # index 재부여
        for i, s in enumerate(plan["slides"], 1):
            s["index"] = i
        print(f"  ✓ 마무리 슬라이드 자동 추가 (총 {len(plan['slides'])}장, 본문 {n_slides-3}장)")

    # ── Plan 제약 강제 (banned slides 자동 교체) ──
    plan, constraint_changes = enforce_plan_constraints(plan, slide_info)

    # ── 표지 날짜를 오늘 날짜로 자동 설정 ───────────
    from datetime import date as _date
    today_str = _date.today().strftime("%Y.%m.%d")
    for s in plan["slides"]:
        if s.get("role") == "cover":
            s.setdefault("content", {})["date"] = today_str

    # ── 중복 template_file 복사 ───────────────────
    # 같은 template_file을 여러 슬라이드가 쓰면 편집이 덮어써짐
    # → 복사본을 만들어 각 슬라이드가 독립 파일을 갖게 함
    seen_templates: dict[str, int] = {}
    slides_dir_tmp = work_dir / "unpacked" / "ppt" / "slides"
    for s in plan["slides"]:
        tmpl = s.get("template_file", "")
        if not tmpl or s.get("role") in ("cover", "toc", "closing"):
            continue
        if tmpl in seen_templates:
            # 복사본 생성
            seen_templates[tmpl] += 1
            suffix = seen_templates[tmpl]
            new_name = tmpl.replace(".xml", f"_c{suffix}.xml")
            src = slides_dir_tmp / tmpl
            dst = slides_dir_tmp / new_name
            if src.exists() and not dst.exists():
                shutil.copy2(src, dst)
                # rels 파일도 복사
                src_rels = slides_dir_tmp / "_rels" / f"{tmpl}.rels"
                dst_rels = slides_dir_tmp / "_rels" / f"{new_name}.rels"
                if src_rels.exists():
                    shutil.copy2(src_rels, dst_rels)
                # 올바른 위치(closing 슬라이드 앞)에 삽입 → 순서 보장
                closing_file = next(
                    (sv.get("template_file") for sv in plan.get("slides", [])
                     if sv.get("role") == "closing"), None)
                _add_slide_to_presentation(work_dir, new_name,
                                           before_filename=closing_file)
                _register_slide_content_type(work_dir, new_name)
            s["template_file"] = new_name
            print(f"  ✓ 중복 방지: slide {s['index']} {tmpl} → {new_name}")
        else:
            seen_templates[tmpl] = 1

    # ── TOC 항목 및 페이지번호 설정 ─────────────────
    # 목차 구조: 1개 TOC 항목 = 1개 챕터, 챕터당 N개 슬라이드 가능
    # page_nums = 각 챕터의 첫 번째 슬라이드 번호
    # TOC 항목 수 ≤ content 슬라이드 수 보장 (항목당 최소 1개 슬라이드)
    content_slides = [s for s in plan["slides"]
                      if s.get("role") not in ("cover", "toc", "closing")]
    n_content = len(content_slides)

    for s in plan["slides"]:
        if s.get("role") == "toc":
            c = s.setdefault("content", {})
            items = (c.get("items")
                     or [i["text"] if isinstance(i, dict) else i
                         for i in c.get("toc_items", [])]
                     or c.get("bullets", []))

            # TOC 항목 수가 content 슬라이드 수를 초과하면 자름
            # (항목당 최소 1개 슬라이드 보장)
            if len(items) > n_content:
                items = items[:n_content]

            # 챕터별 첫 번째 슬라이드 번호 계산
            # content slide를 TOC 항목 수로 균등 분배하고 첫 페이지만 표시
            n_items = len(items)
            if n_items > 0:
                # 각 챕터의 시작 슬라이드 인덱스 (균등 분배)
                slides_per_chapter = n_content / n_items
                page_nums = [
                    f"{content_slides[min(int(i * slides_per_chapter), n_content - 1)]['index']:02d}"
                    for i in range(n_items)
                ]
            else:
                page_nums = []

            c["items"]     = items
            c["page_nums"] = page_nums
            print(f"  ✓ TOC: {n_items}개 챕터, 페이지번호={page_nums} (content {n_content}장)")
    for ch in constraint_changes:
        print(f"  ⚠ 제약 교체: {ch}")

    # ── Plan 검증 ─────────────────────────────────
    ok, warnings = validate_plan(plan)
    for w in warnings[:5]:
        print(f"  ⚠ plan 검증: {w}")
    if not ok:
        print("  ⚠ 필수 content 필드 누락 — 편집 시 폴백 처리됨")

    plan_path = work_dir / "plan.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2))
    print(f"  ✓ plan.json 저장 ({len(plan['slides'])}개 슬라이드)")

    # ── 편집 루프 ────────────────────────────────
    slides_dir = work_dir / "unpacked" / "ppt" / "slides"
    for slide_plan in plan["slides"]:
        xml_path = slides_dir / slide_plan.get("template_file", "")

        # known_fixes 자동 적용 (편집 전)
        if xml_path.exists():
            apply_known_fixes(xml_path, slide_plan, known_fixes)

        success = edit_slide(work_dir, slide_plan)
        if not success:
            print(f"  → slide {slide_plan['index']} 재시도...")
            edit_slide(work_dir, slide_plan)

    # ── 사용 슬라이드만 남기기 (plan에 없는 슬라이드 제거) ────
    _trim_to_plan_slides(work_dir, plan)

    # ── 패킹 ─────────────────────────────────────
    output = work_dir / "output.pptx"
    # 복사본 슬라이드가 있으면 스키마 검증 우회 (원본에 없는 파일이라 오탐)
    has_copies = any("_c" in s.get("template_file", "") for s in plan["slides"])
    if not pack_output(work_dir, output, skip_validation=has_copies):
        raise RuntimeError("패킹 실패")

    # ── 섹션 정리 ─────────────────────────────────
    restructure_sections(output)

    # ── Verifier 실행 ─────────────────────────────
    print("\n  [Verifier] 규칙 검증 중...")
    violations = execute_verifier_rules(output, work_dir)
    critical = [v for v in violations if v["severity"] == "CRITICAL"]

    # CRITICAL 위반이 있으면 해당 슬라이드 재편집 (최대 1회)
    if critical:
        print(f"  → CRITICAL 위반 {len(critical)}건 감지 — 영향 슬라이드 재편집")
        retry_files = {v.get("file") for v in critical if v.get("file")}
        for slide_plan in plan["slides"]:
            if slide_plan.get("template_file") in retry_files or not retry_files:
                xml_path = slides_dir / slide_plan.get("template_file", "")
                if xml_path.exists():
                    apply_known_fixes(xml_path, slide_plan, known_fixes)
                edit_slide(work_dir, slide_plan)
        # 재패킹 후 재검증
        pack_output(work_dir, output)
        restructure_sections(output)
        violations2 = execute_verifier_rules(output, work_dir)
        remaining = [v for v in violations2 if v["severity"] == "CRITICAL"]
        if remaining:
            print(f"  ⚠ 재시도 후 CRITICAL {len(remaining)}건 잔존 (best-effort 출력)")

    # ── 시각 QA ──────────────────────────────────
    images = visual_qa(work_dir, output)
    if images:
        print(f"  → QA 이미지 {len(images)}장 생성")

    # ── Vision Fix Agent 루프 (최대 3회) ────────
    MAX_VISION_ITER = 3
    if images:
        for vision_iter in range(1, MAX_VISION_ITER + 1):
            print(f"\n  [Vision Fix Agent] 검증 {vision_iter}/{MAX_VISION_ITER}회...")
            fix_instructions = _run_vision_fix_agent(
                work_dir   = work_dir,
                qa_images  = images,
                plan       = plan,
                slides_dir = slides_dir,
            )

            has_issues = any(f.get("has_issues") for f in fix_instructions)
            if not has_issues:
                print("  ✓ Vision 검증 통과 — 모든 슬라이드 정상")
                break

            applied = _apply_fix_instructions(work_dir, fix_instructions)
            if not applied:
                print("  ⚠ 적용 가능한 수정 없음 — 구조적 한계로 best-effort 종료")
                break

            print(f"  → 수정 적용 후 재패킹 (이터레이션 {vision_iter})...")
            _trim_to_plan_slides(work_dir, plan)
            pack_output(work_dir, output)
            restructure_sections(output)

            # 재검증 후 QA 이미지 갱신
            violations = execute_verifier_rules(output, work_dir)
            critical   = [v for v in violations if v["severity"] == "CRITICAL"]
            if critical:
                print(f"  ⚠ CRITICAL {len(critical)}건 잔존")
            else:
                print(f"  ✓ Verifier 통과 (이터레이션 {vision_iter})")

            images = visual_qa(work_dir, output)  # QA 이미지 갱신

            if vision_iter == MAX_VISION_ITER:
                print(f"  ⚠ 최대 반복({MAX_VISION_ITER}회) 도달 — best-effort 출력")

    # ── 콘텐츠 검증 ──────────────────────────────
    issues = verify_content(output)
    if issues:
        print(f"  ⚠ 플레이스홀더 잔여 ({len(issues)}건):")
        for issue in issues[:3]:
            print(f"    - {issue}")
    else:
        print("  ✓ 콘텐츠 검증 통과")

    # Vision 이슈 수를 meta로 반환 (조건부 auto-evolve용)
    _vision_critical_total = sum(
        1 for slide in (fix_instructions if "fix_instructions" in dir() else [])
        if slide.get("has_issues")
    )
    # ── Excel 차트 데이터 파일 생성 ─────────────────
    dest_dir_for_excel = work_dir.parent  # runs 상위 디렉토리가 아닌 work_dir 사용
    generate_excel_for_charts(work_dir, plan, work_dir)

    return output, _vision_critical_total
