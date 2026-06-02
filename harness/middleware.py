"""
PPT 하네스 미들웨어 — v0.2 (2026-06-02)
PreToolUse 훅: XML 편집 전 검사
PostToolUse 훅: 편집 후 ET.parse() 자동 실행
"""

import re
import xml.etree.ElementTree as ET

MIDDLEWARE_VERSION = "0.2"
EVOLUTION_NOTES = "v0.2: run/endParaRPr 순서 검사, lang=ko-KR 검사 추가"


def pre_xml_edit(replacement: str) -> dict:
    """XML 편집 전 이스케이프 검사"""
    issues = []
    raw_amp = re.findall(r'&(?!amp;|lt;|gt;|quot;|apos;|#\w+;)', replacement)
    if raw_amp:
        issues.append({"type": "unescaped_ampersand", "detail": str(raw_amp)})
    raw_lt = re.findall(r'<(?![a-zA-Z/!?])', replacement)
    if raw_lt:
        issues.append({"type": "unescaped_lt", "detail": str(raw_lt)})
    return {"pass": len(issues) == 0, "issues": issues}


def post_xml_edit(filepath: str) -> dict:
    """XML 편집 후 유효성 검증 + run/endParaRPr 순서 검사"""
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    try:
        tree = ET.parse(filepath)
    except ET.ParseError as e:
        return {"pass": False, "issues": [{"type": "xml_parse_error", "detail": str(e)}]}

    issues = []
    for para in tree.getroot().iter(f"{{{NS_A}}}p"):
        children = list(para)
        tags = [c.tag.split("}")[-1] for c in children]
        if "endParaRPr" in tags and "r" in tags:
            end_idx = tags.index("endParaRPr")
            # endParaRPr 뒤에 r이 있으면 경고
            if any(t == "r" for t in tags[end_idx + 1:]):
                issues.append({
                    "type": "run_after_endpararpr",
                    "detail": "PowerPoint ignores <a:r> after <a:endParaRPr>",
                    "file": filepath,
                })
                break  # 첫 번째 발견만 보고

    return {"pass": len(issues) == 0, "issues": issues}


def check_lang_attribute(xml_str: str, has_korean: bool = True) -> dict:
    """한국어 텍스트 run의 lang 속성 검사"""
    issues = []
    if has_korean:
        # lang=en-US인 run에 한국어가 있으면 경고
        pattern = r'lang="en-US"[^>]*>\s*<a:t>[가-힣]+'
        if re.search(pattern, xml_str):
            issues.append({
                "type": "korean_in_en_us_run",
                "detail": "lang=ko-KR 필요. en-US run의 한국어 텍스트는 ?? 렌더링될 수 있음",
            })
    return {"pass": len(issues) == 0, "issues": issues}


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
