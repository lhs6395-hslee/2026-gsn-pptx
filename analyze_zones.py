#!/usr/bin/env python3
"""
analyze_zones.py — 템플릿 본문 슬라이드의 5개 존을 자동 분류해 SLIDE_ZONE_MAP 생성.

존 분류 체계 (사용자 정의):
  title        대제목      ("PPT 대제목 작성")
  subtitle     중제목      ("01 중제목 작성" / "01 컨텐츠 작성")
  body_title   본문제목     ("1.x ...")
  body_desc    본문제목설명  (좌측 "상세 설명을 작성해주세요 / 12pt")
  본문구역(body sub-zones):
    image_slots     이미지 자리        ("image"/"Image") → 이미지 설명 텍스트로 대체
    icon_slots      아이콘 자리        ("icon")
    item_titles     항목/이미지 제목    ("01 제목/14pt", "01 설명 타이틀…")
    item_descs      항목/이미지 설명    (본문 영역의 "01 상세 설명…")
    insights        인사이트/결론/정의   ("Insight"/"Conclusion"/"Definition")
    keywords        키워드 버블         ("Keyword")
    steps           스텝               ("Step1".."Step4")
    timeline        연/분기 라벨        ("2023".."2026","Q1".."Q4")
    flow_solution   흐름 솔루션         ("Solution …")
    flow_service    흐름 서비스         ("Sevice …")
    detail_items    세부 항목           ("항목 01 …")
    sub_titles      부분 설명 타이틀
    explains        설명 박스           ("explain")
    banner          핵심 요약 배너       ("핵심 설명")
    sub_heading     소제목 라인          ("1.1.1 …")

각 sub-zone은 (행 y, 열 x) 순으로 정렬된 shape ID 리스트.
"""
import json, re, sys
import xml.etree.ElementTree as ET
from pathlib import Path

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _gi(v):
    try: return int(v)
    except Exception: return 0


def _geom(el):
    off = el.find(f".//{{{A}}}off"); ext = el.find(f".//{{{A}}}ext")
    return (_gi(off.get("x")) if off is not None else 0,
            _gi(off.get("y")) if off is not None else 0,
            _gi(ext.get("cx")) if ext is not None else 0,
            _gi(ext.get("cy")) if ext is not None else 0)


def _collect(parent, out, depth=0):
    for el in parent:
        tag = el.tag.split("}")[-1]
        nv = el.find(f"./{{{P}}}nvSpPr/{{{P}}}cNvPr")
        if tag == "sp":
            txt = " ".join((e.text or "").strip()
                           for e in el.iter(f"{{{A}}}t") if (e.text or "").strip())
            x, y, cx, cy = _geom(el)
            sid = nv.get("id") if nv is not None else None
            out.append({"id": sid, "kind": "sp", "x": x, "y": y, "cx": cx, "cy": cy, "txt": txt})
        elif tag == "pic":
            nv = el.find(f".//{{{P}}}cNvPr")
            x, y, cx, cy = _geom(el)
            out.append({"id": nv.get("id") if nv is not None else None,
                        "kind": "pic", "x": x, "y": y, "cx": cx, "cy": cy, "txt": ""})
        elif tag == "graphicFrame":
            nv = el.find(f".//{{{P}}}cNvPr")
            x, y, cx, cy = _geom(el)
            out.append({"id": nv.get("id") if nv is not None else None,
                        "kind": "chart", "x": x, "y": y, "cx": cx, "cy": cy, "txt": ""})
        elif tag == "grpSp":
            _collect(el, out, depth + 1)


def classify(shapes):
    """각 shape을 존 role로 분류."""
    zones = {"title": None, "subtitle": None, "body_title": None, "body_desc": None}
    buckets = {}  # role -> list of shapes

    def add(role, sh):
        buckets.setdefault(role, []).append(sh)

    for sh in shapes:
        if sh["kind"] == "chart":
            add("charts", sh); continue
        if sh["kind"] == "pic":
            add("pics", sh); continue
        t = sh["txt"].strip()
        tl = t.lower()
        x = sh["x"]
        if t == "PPT 대제목 작성":
            zones["title"] = sh["id"]; continue
        if re.match(r"^01 (중제목|컨텐츠) 작성", t):
            zones["subtitle"] = sh["id"]; continue
        if re.match(r"^1\.\d", t):
            zones["body_title"] = sh["id"]; continue
        # 좌측(x<800000) 상세설명 = 본문제목 설명글
        if t.startswith("상세 설명을 작성해주세요") and x < 900000:
            if zones["body_desc"] is None:
                zones["body_desc"] = sh["id"]; continue
        # 본문구역 sub-zones
        if tl == "image":
            add("image_slots", sh); continue
        if tl == "icon":
            add("icon_slots", sh); continue
        if re.match(r"^01 제목", t) or "설명 타이틀을 작성" in t:
            add("item_titles", sh); continue
        if t.startswith("01 상세 설명을 작성") or (t.startswith("상세 설명을 작성") and x >= 900000):
            # explain 박스(우측 keyword 설명)과 구분: cx 작은 우측은 explain
            add("item_descs", sh); continue
        if re.search(r"insight|conclusion|definition", tl):
            add("insights", sh); continue
        if t == "Keyword":            # 대문자 = 버블 키워드 (slide34/35/36)
            add("keywords", sh); continue
        if tl == "keyword":           # 소문자 = 흐름도 키워드 (slide38/39)
            add("flow_keyword", sh); continue
        if re.match(r"^Step ?[1-4]", t):
            add("steps", sh); continue
        if re.match(r"^Q[1-4]$", t):
            add("timeline_q", sh); continue
        if re.match(r"^20\d\d$", t):
            add("timeline_y", sh); continue
        if re.match(r"^Solution", t, re.I):
            add("flow_solution", sh); continue
        if re.match(r"^Se[rv]vice|^Service", t, re.I):
            add("flow_service", sh); continue
        if re.match(r"^항목 ?0?1", t):
            add("detail_items", sh); continue
        if "부분 설명 타이틀" in t:
            add("sub_titles", sh); continue
        if tl == "explain":
            add("explains", sh); continue
        if t.startswith("핵심 설명"):
            add("banner", sh); continue
        if re.match(r"^\|?\s*1\.\d\.\d", t):
            add("sub_heading", sh); continue
        if re.match(r"^Before$|^After$|^As-is$|^To-be$", t, re.I):
            add("compare_label", sh); continue

    # 정렬: 행(y, 80000 EMU 버킷) → 열(x)
    def order(lst):
        return [s["id"] for s in sorted(lst, key=lambda s: (round(s["y"] / 400000), s["x"]))]

    body = {role: order(lst) for role, lst in buckets.items()}
    return zones, body


def main():
    base = Path(sys.argv[1]) if len(sys.argv) > 1 else \
        Path("/tmp/tmpl_inspect/unpacked/ppt/slides")
    result = {}
    for n in range(6, 47):
        f = base / f"slide{n}.xml"
        if not f.exists():
            continue
        t = ET.parse(f)
        spTree = t.getroot().find(f".//{{{P}}}cSld/{{{P}}}spTree")
        shapes = []
        _collect(spTree, shapes)
        zones, body = classify(shapes)
        entry = {k: v for k, v in zones.items() if v}
        entry["body"] = body
        result[f"slide{n}.xml"] = entry
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
