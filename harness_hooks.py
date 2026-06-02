"""
PPT 하네스 훅 — XML 편집 전후 검사
스킬이 슬라이드 편집 시 직접 import해서 사용
"""
import re, xml.etree.ElementTree as ET


def pre_xml_edit(content: str) -> dict:
    """Pre-hook: 편집할 내용에 날 & < 가 있는지 검사"""
    issues = []
    for m in re.finditer(r'&(?!amp;|lt;|gt;|quot;|apos;|#[\w]+;)', content):
        issues.append({"type": "unescaped_ampersand", "pos": m.start()})
    for m in re.finditer(r'<(?![a-zA-Z/!?])', content):
        issues.append({"type": "unescaped_lt", "pos": m.start()})
    return {"pass": not issues, "issues": issues}


def post_xml_edit(filepath: str) -> dict:
    """Post-hook: 편집 후 XML 유효성 검증"""
    try:
        ET.parse(filepath)
        return {"pass": True, "issues": []}
    except ET.ParseError as e:
        return {"pass": False, "issues": [{"type": "xml_parse_error", "detail": str(e)}]}


def check_placeholders(text: str) -> dict:
    """콘텐츠 검증: 플레이스홀더 잔여 탐지"""
    patterns = [
        (r'lorem\s+ipsum',      "lorem_ipsum"),
        (r'작성해\s*주세요',     "korean_placeholder"),
        (r'\bTODO\b',           "todo"),
        (r'\[insert',           "insert_bracket"),
        (r'x{3,}',              "xxx_placeholder"),
        (r'대제목을\s*작성',    "title_placeholder"),
        (r'중제목을\s*작성',    "subtitle_placeholder"),
    ]
    issues = []
    for pattern, ptype in patterns:
        if re.search(pattern, text, re.IGNORECASE):
            issues.append({"type": "placeholder_remaining", "pattern": ptype})
    return {"pass": not issues, "issues": issues}
