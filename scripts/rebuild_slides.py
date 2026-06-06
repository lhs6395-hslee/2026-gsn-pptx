#!/usr/bin/env python3
"""
재현검증 파일 슬라이드 3-13 재생성 스크립트.

원본 PPTX에서 슬라이드 3-13의 content/layout을 추출하여 plan.json을 생성한 후,
edit_slide 시스템으로 재생성하고 시각 QA를 수행한다.
"""
import sys, json, re, shutil, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

# ppt_generator 임포트
PROJECT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT))
import ppt_generator as ppt

ORIG_PPTX   = PROJECT / "result" / "머신러닝_심층_기술_가이드.pptx"
TARGET_PPTX = PROJECT / "result" / "머신러닝_심층_기술_가이드_재현검증.pptx"
WORK_DIR    = PROJECT / "result" / "tmp" / "regen_work"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ── 헬퍼: shape ID로 텍스트 추출 ─────────────────────────────────

def get_all_text(root: ET.Element, shape_id: str) -> str:
    """shape ID로 shape를 찾아 모든 <a:t>를 합쳐 반환."""
    for sp in root.iter(f"{{{NS_P}}}sp"):
        cpr = sp.find(f"{{{NS_P}}}nvSpPr/{{{NS_P}}}cNvPr")
        if cpr is not None and cpr.get("id") == shape_id:
            return "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()
    return ""


def get_para_texts(root: ET.Element, shape_id: str) -> list[str]:
    """shape ID로 shape를 찾아 paragraph별 텍스트 리스트 반환 (빈 paragraph 포함)."""
    for sp in root.iter(f"{{{NS_P}}}sp"):
        cpr = sp.find(f"{{{NS_P}}}nvSpPr/{{{NS_P}}}cNvPr")
        if cpr is not None and cpr.get("id") == shape_id:
            result = []
            for p in sp.findall(f".//{{{NS_A}}}p"):
                texts = [t.text or "" for t in p.findall(f".//{{{NS_A}}}t")]
                result.append("".join(texts).strip())
            return result
    return []


def strip_num_prefix(text: str) -> str:
    """'01 Transformer' → 'Transformer', '1.1. 제목' → '제목' 등 번호 prefix 제거."""
    # "| N.N.N 텍스트" 패턴
    m = re.match(r'^\|\s*[\d.]+\s+(.+)$', text)
    if m:
        return m.group(1).strip()
    # "N.N. 텍스트" 패턴
    m = re.match(r'^[\d]+\.[\d]*\.?\s+(.+)$', text)
    if m:
        return m.group(1).strip()
    # "01 텍스트" 패턴 (2자리 숫자 + 공백)
    m = re.match(r'^\d{2}\s+(.+)$', text)
    if m:
        return m.group(1).strip()
    return text


def strip_image_format(text: str) -> str:
    """'[이미지: 설명]' → '설명'."""
    m = re.match(r'^\[이미지:\s*(.+)\]$', text, re.DOTALL)
    if m:
        return m.group(1).strip()
    return text


# ── zone_map 기반 content 추출 ───────────────────────────────────

def extract_content_by_zone(root: ET.Element, template_file: str, zone_map: dict) -> dict:
    """zone_map을 참조해 각 zone의 텍스트를 content dict로 반환."""
    z = zone_map.get(template_file, {})
    body = z.get("body", {})
    content = {}

    # Zone 1 헤더: title(ID=8), subtitle(ID=9)
    title_id    = z.get("title", "8")
    subtitle_id = z.get("subtitle", "9")
    content["title"]    = get_all_text(root, title_id)
    content["subtitle"] = strip_num_prefix(get_all_text(root, subtitle_id))

    # Zone 2 사이드바: body_title, body_desc
    bt_id = z.get("body_title")
    bd_id = z.get("body_desc")
    if bt_id:
        content["section_title"] = get_all_text(root, bt_id)  # prefix 보존: _apply_common_zones에서 그대로 사용
    if bd_id:
        content["section_desc"] = get_all_text(root, bd_id)

    # Zone 3 본문: item_titles, item_descs, image_slots, sub_heading 등
    item_title_ids = body.get("item_titles", [])
    item_desc_ids  = body.get("item_descs", [])
    image_slot_ids = body.get("image_slots", [])
    sub_heading_id = body.get("sub_heading", [None])[0] if body.get("sub_heading") else None
    insights_ids   = body.get("insights", [])
    bullets_ids    = body.get("bullets", [])
    body_body_ids  = body.get("body", [])
    overview_ids   = body.get("overview", [])
    main_desc_ids  = body.get("main_desc", [])
    steps_ids      = body.get("steps", [])
    keywords_ids   = body.get("keywords", [])
    solutions_ids  = body.get("solutions", [])

    # item_titles: 각 shape의 paragraph 리스트 → item{i}_title, item{i}_subtitle
    items_list = []
    for i, tid in enumerate(item_title_ids, 1):
        paras = get_para_texts(root, tid)
        if paras:
            t = strip_num_prefix(paras[0]) if paras else ""
            s = paras[1] if len(paras) > 1 else ""
            content[f"item{i}_title"]    = t
            content[f"item{i}_subtitle"] = s
            if t:
                items_list.append(t)

    # item_descs
    descs_list = []
    for i, did in enumerate(item_desc_ids, 1):
        d = get_all_text(root, did)
        content[f"item{i}_desc"] = d
        if d:
            descs_list.append(d)

    # image_slots: strip [이미지: ...] wrapper
    img_descs = []
    for i, iid in enumerate(image_slot_ids, 1):
        txt = get_all_text(root, iid)
        desc = strip_image_format(txt)
        content[f"item{i}_image_desc"] = desc
        img_descs.append(desc)

    # sub_heading — prefix 보존: _edit_zonemap_slide에서 그대로 사용
    if sub_heading_id:
        content["sub_heading"] = get_all_text(root, sub_heading_id)

    # insights
    insights_list = []
    for i, iid in enumerate(insights_ids, 1):
        val = get_all_text(root, iid)
        content[f"item{i}_insight"] = val
        if val:
            insights_list.append(val)
    if insights_list:
        content["insights"] = insights_list

    # slide24: bullets (ID=7) — paragraph 리스트로 추출 (_edit_slide24가 리스트로 처리)
    if bullets_ids:
        paras = get_para_texts(root, bullets_ids[0])
        content["bullets"] = [p for p in paras if p.strip()]
    if body_body_ids:
        content["body"] = get_all_text(root, body_body_ids[0])

    # slide25/26: overview, main_desc
    if overview_ids:
        content["overview"] = get_all_text(root, overview_ids[0])
    if main_desc_ids:
        content["main_desc"] = get_all_text(root, main_desc_ids[0])

    # 범용: items, descriptions, image_descriptions 리스트 필드
    if items_list:
        content["items"] = items_list
    if descs_list:
        content["descriptions"] = descs_list
    if img_descs and any(img_descs):
        content["image_descriptions"] = img_descs

    return content


# ── 슬라이드별 특수 추출 ─────────────────────────────────────────

def extract_slide29_content(root: ET.Element) -> dict:
    """slide29 (연도별 타임라인): 시계열 레이블 추출."""
    # label IDs: 16, 17, 22, 23 (연도)
    label_ids = ["16", "17", "22", "23"]
    periods = []
    for lid in label_ids:
        label = get_all_text(root, lid)
        if label:
            periods.append({"label": label, "items": []})
    return {"periods": periods} if periods else {}


def extract_slide31_content(root: ET.Element) -> dict:
    """slide31 (분기별 Q1-Q4): Q 레이블 추출."""
    # Q 레이블들은 zone_map에 없으므로 텍스트로 Q 패턴 검색
    quarters = []
    for sp in root.iter(f"{{{NS_P}}}sp"):
        txt = "".join(t.text or "" for t in sp.iter(f"{{{NS_A}}}t")).strip()
        if re.match(r'^Q[1-4]$', txt):
            quarters.append({"label": txt, "items": []})
    return {"quarters": quarters} if quarters else {}


def extract_slide33_content(root: ET.Element) -> dict:
    """slide33 (분기별 변형): Q 레이블 추출."""
    return extract_slide31_content(root)


def extract_slide35_content(root: ET.Element, zone_map: dict) -> dict:
    """slide35 (Before/After): before/after 키워드 추출."""
    z = zone_map.get("slide35.xml", {})
    body = z.get("body", {})
    # zone_map 키: before_items, after_items (before/after는 구버전 키)
    before_ids = body.get("before_items", body.get("before", []))
    after_ids  = body.get("after_items",  body.get("after", []))
    before = [get_all_text(root, bid) for bid in before_ids if get_all_text(root, bid)]
    after  = [get_all_text(root, aid) for aid in after_ids  if get_all_text(root, aid)]
    return {"before": before or ["기존 방식"], "after": after or ["개선 방식"]}


# ── 원본 슬라이드 슬라이드 순서 파악 ─────────────────────────────

def get_slide_order(pptx_path: Path) -> list[str]:
    """원본 PPTX의 슬라이드 순서 (파일명 목록) 반환."""
    with zipfile.ZipFile(pptx_path) as z:
        prs  = z.read("ppt/presentation.xml").decode("utf-8")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")

    rid_file = {}
    for m in re.finditer(
        r'Id="(rId\d+)"[^>]*Type="[^"]*slide[^"]*"[^>]*Target="slides/(slide[\w]+\.xml)"', rels
    ):
        rid_file[m.group(1)] = m.group(2)

    order = []
    for m in re.finditer(r'<p:sldId\b[^>]*\br:id="(rId\d+)"', prs):
        fname = rid_file.get(m.group(1))
        if fname:
            order.append(fname)
    return order


# ── 메인 실행 ─────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("재현검증 재생성 시작")
    print("=" * 60)

    # zone_map 로드
    zone_map = ppt._load_zone_map()

    # ── Step 1: 원본 언팩 ─────────────────────────────────────────
    print("\n[1] 원본 PPTX 언팩...")
    orig_unpack = WORK_DIR / "orig_unpacked"
    if orig_unpack.exists():
        shutil.rmtree(orig_unpack)
    orig_unpack.mkdir(parents=True)

    with zipfile.ZipFile(ORIG_PPTX) as z:
        z.extractall(orig_unpack)

    slide_order = get_slide_order(ORIG_PPTX)
    print(f"  원본 슬라이드 순서: {slide_order}")

    # ── Step 2: 슬라이드 3-13 content 추출 ──────────────────────
    print("\n[2] 슬라이드 3-13 content 추출...")
    slides_dir = orig_unpack / "ppt" / "slides"

    # 원본 슬라이드 3-13 = index 2..12 (0-based)
    body_slides = slide_order[2:13]  # slide15, slide29, ..., slide12
    print(f"  대상: {body_slides}")

    plan_slides = []

    # 표지 (index=1) — 원본에서 subtitle(ID=6), title(ID=12), date(ID=15) 추출
    cover_root = ET.parse(slides_dir / slide_order[0]).getroot()
    cover_title    = get_all_text(cover_root, "12") or "머신러닝 심층 기술 가이드"
    cover_subtitle = get_all_text(cover_root, "6")  or ""
    cover_date     = get_all_text(cover_root, "15") or "2026.06"
    plan_slides.append({
        "index": 1,
        "template_file": slide_order[0],
        "role": "cover",
        "title": cover_title,
        "content": {"subtitle": cover_subtitle, "date": cover_date},
    })

    # 목차 (index=2) — 원본에서 텍스트 추출
    toc_root = ET.parse(slides_dir / slide_order[1]).getroot()
    toc_items = []
    # TOC ID=10에서 paragraph별 텍스트 추출
    toc_texts = get_para_texts(toc_root, "10")
    for t in toc_texts:
        if t.strip():
            toc_items.append(t.strip())
    # ID=14에서 페이지번호 추출
    toc_page_texts = get_para_texts(toc_root, "14")
    toc_page_nums = [t.strip() for t in toc_page_texts if t.strip()]
    # ID=42에서 제목 추출
    toc_title = get_all_text(toc_root, "42") or "목차"

    plan_slides.append({
        "index": 2,
        "template_file": slide_order[1],
        "role": "toc",
        "title": toc_title,
        "content": {"items": toc_items, "page_nums": toc_page_nums},
    })

    # 본문 슬라이드 3-13
    for slide_idx, template_file in enumerate(body_slides, 3):
        xml_path = slides_dir / template_file
        if not xml_path.exists():
            print(f"  ⚠ {template_file} 없음, 건너뜀")
            continue

        root = ET.parse(xml_path).getroot()

        # 기본 content 추출 (zone_map 기반)
        content = extract_content_by_zone(root, template_file, zone_map)

        # 슬라이드별 특수 추출
        if template_file == "slide29.xml":
            extra = extract_slide29_content(root)
            content.update(extra)
        elif template_file in ("slide31.xml", "slide33.xml"):
            extra = extract_slide31_content(root)
            content.update(extra)
        elif template_file == "slide35.xml":
            extra = extract_slide35_content(root, zone_map)
            content.update(extra)

        # slide15: chapter/section/subsection 번호 추론
        chapter = section = subsection = 1
        sec_title = content.get("section_title", "")
        m = re.match(r'^(\d+)\.(\d+)\.?(\d*)', sec_title)
        if m:
            chapter    = int(m.group(1))
            section    = int(m.group(2))
            subsection = int(m.group(3)) if m.group(3) else 1

        # body_title/body_desc 복원 (slide15 전용 content key)
        bt_raw = get_all_text(root, zone_map.get(template_file, {}).get("body_title", "14"))
        bd_raw = get_all_text(root, zone_map.get(template_file, {}).get("body_desc",  "17"))
        if bt_raw:
            if template_file == "slide15.xml":
                # edit_slide15가 chapter.section. prefix를 직접 추가하므로 여기선 strip
                content["body_title"] = strip_num_prefix(bt_raw)
            else:
                content["body_title"] = bt_raw  # 다른 슬라이드는 직접 사용 안 함 (section_title 우선)
        if bd_raw:
            content["body_desc"]  = bd_raw

        plan_slide = {
            "index": slide_idx,
            "template_file": template_file,
            "role": "content",
            "title": content.pop("title", ""),
            "subtitle": content.pop("subtitle", ""),
            "content": content,
        }
        if template_file == "slide15.xml":
            plan_slide["content"]["chapter"]    = chapter
            plan_slide["content"]["section"]    = section
            plan_slide["content"]["subsection"] = subsection

        plan_slides.append(plan_slide)
        print(f"  ✓ page {slide_idx} ({template_file}): section_title={content.get('section_title','?')!r}")

    # 감사합니다 (마지막)
    plan_slides.append({
        "index": 14,
        "template_file": slide_order[13],
        "role": "closing",
        "title": "감사합니다",
        "content": {},
    })

    plan = {
        "title": "머신러닝 심층 기술 가이드",
        "topic": "머신러닝 심층 기술 가이드",
        "audience": "기술 전문가",
        "n_slides": 14,
        "slides": plan_slides,
    }

    plan_path = WORK_DIR / "plan.json"
    plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2))
    print(f"\n  ✓ plan.json 저장: {plan_path}")

    # ── Step 3: 원본 PPTX를 template로 복사 + 언팩 ───────────────
    print("\n[3] 재생성용 작업 디렉토리 준비...")
    unpacked = WORK_DIR / "unpacked"
    if unpacked.exists():
        shutil.rmtree(unpacked)
    template_dest = WORK_DIR / "template.pptx"
    shutil.copy2(ORIG_PPTX, template_dest)

    with zipfile.ZipFile(template_dest) as z:
        z.extractall(unpacked)
    print(f"  ✓ 원본 언팩 완료: {unpacked}")

    # ── Step 4: 각 슬라이드 edit_slide 실행 ──────────────────────
    print("\n[4] edit_slide 실행 중...")
    success_count = 0
    for slide_plan in plan["slides"]:
        result = ppt.edit_slide(WORK_DIR, slide_plan)
        if result:
            success_count += 1
        else:
            print(f"  ✗ 실패: slide {slide_plan['index']} ({slide_plan['template_file']})")

    print(f"\n  ✓ edit_slide 완료: {success_count}/{len(plan['slides'])}장")

    # ── Step 5: 슬라이드 순서 정렬 ────────────────────────────────
    print("\n[5] 슬라이드 순서 정렬...")
    ppt._trim_to_plan_slides(WORK_DIR, plan)

    # ── Step 6: 팩킹 ──────────────────────────────────────────────
    print("\n[6] PPTX 팩킹...")
    ok = ppt.pack_output(WORK_DIR, TARGET_PPTX, skip_validation=False)
    if not ok:
        print("  ✗ 팩킹 실패!")
        sys.exit(1)
    print(f"  ✓ 출력: {TARGET_PPTX}")

    # ── Step 7: 비주얼 QA ─────────────────────────────────────────
    print("\n[7] 비주얼 QA (PowerPoint → PDF → 이미지)...")
    images = ppt.visual_qa(WORK_DIR, TARGET_PPTX)
    if images:
        print(f"  ✓ QA 이미지 {len(images)}장 생성")
        for img in images:
            print(f"    {img}")
    else:
        print("  ⚠ QA 이미지 생성 실패")

    print("\n" + "=" * 60)
    print("재생성 완료!")
    print(f"결과 파일: {TARGET_PPTX}")
    print("=" * 60)


if __name__ == "__main__":
    main()
