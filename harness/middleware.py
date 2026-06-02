"""
PPT 하네스 미들웨어 — v0.0
PreToolUse 훅: XML 편집 전 이스케이프 검사
PostToolUse 훅: 편집 후 ET.parse() 자동 실행
AHE 진화 대상: 이 파일에 규칙을 추가·수정하면 하네스가 개선된다
"""

import re
import xml.etree.ElementTree as ET

MIDDLEWARE_VERSION = "0.0"
EVOLUTION_NOTES = "시드 버전 - 기본 XML 검증만 있음"


def pre_xml_edit(replacement: str) -> dict:
    """XML 편집 전 이스케이프 검사"""
    issues = []
    # 날 & 탐지 (엔티티가 아닌 것)
    raw_amp = re.findall(r'&(?!amp;|lt;|gt;|quot;|apos;|#\w+;)', replacement)
    if raw_amp:
        issues.append({"type": "unescaped_ampersand", "detail": str(raw_amp)})
    # 날 < 탐지 (태그가 아닌 것)
    raw_lt = re.findall(r'<(?![a-zA-Z/!?])', replacement)
    if raw_lt:
        issues.append({"type": "unescaped_lt", "detail": str(raw_lt)})
    return {"pass": len(issues) == 0, "issues": issues}


def post_xml_edit(filepath: str) -> dict:
    """XML 편집 후 유효성 검증"""
    try:
        ET.parse(filepath)
        return {"pass": True, "issues": []}
    except ET.ParseError as e:
        return {"pass": False, "issues": [{"type": "xml_parse_error", "detail": str(e)}]}


def check_placeholder_remaining(text: str) -> dict:
    """플레이스홀더 잔여 검사"""
    PATTERNS = [
        (r'lorem\s+ipsum', "lorem_ipsum"),
        (r'작성해\s*주세요', "korean_placeholder"),
        (r'\bTODO\b', "todo"),
        (r'\[insert', "insert_placeholder"),
        (r'x{3,}', "xxx_placeholder"),
    ]
    issues = []
    for pattern, ptype in PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            issues.append({"type": "placeholder_remaining", "pattern_type": ptype})
    return {"pass": len(issues) == 0, "issues": issues}
