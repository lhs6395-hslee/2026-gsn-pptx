"""
format_qa.py — 텍스트 + 폰트 서식 비교 QA

원본 vs 재생성 PPTX의 모든 shape에 대해:
  - 텍스트 내용
  - 폰트명, 크기(pt), bold, 색상, 정렬
를 비교하고 차이를 보고한다.
"""
import zipfile, sys, json
from pathlib import Path
from lxml import etree

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ORIG_PPTX  = Path("result/머신러닝_심층_기술_가이드.pptx")
REGEN_PPTX = Path("result/머신러닝_심층_기술_가이드_재현검증.pptx")


# ── 헬퍼 ────────────────────────────────────────────────────────────────────

def slide_order(pptx_path: Path) -> list[str]:
    with zipfile.ZipFile(pptx_path) as z:
        prs  = etree.fromstring(z.read("ppt/presentation.xml"))
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rid_map = {
        r.get("Id"): f"ppt/{r.get('Target')}"
        for r in rels
        if r.get("Target", "").startswith("slides/slide")
    }
    return [
        rid_map[s.get(f"{{{NS_R}}}id")]
        for s in prs.iter(f"{{{NS_P}}}sldId")
        if s.get(f"{{{NS_R}}}id") in rid_map
    ]


def get_color(rPr):
    """rPr에서 색상 문자열 추출 (solidFill srgbClr 우선, 없으면 scheme)"""
    if rPr is None:
        return None
    sf = rPr.find(f"{{{NS_A}}}solidFill")
    if sf is None:
        return None
    srgb = sf.find(f"{{{NS_A}}}srgbClr")
    if srgb is not None:
        return f"#{srgb.get('val', '').upper()}"
    schm = sf.find(f"{{{NS_A}}}schemeClr")
    if schm is not None:
        lm = schm.find(f"{{{NS_A}}}lumMod")
        lo = schm.find(f"{{{NS_A}}}lumOff")
        parts = [f"scheme:{schm.get('val')}"]
        if lm is not None:
            parts.append(f"lumMod={lm.get('val')}")
        if lo is not None:
            parts.append(f"lumOff={lo.get('val')}")
        return " ".join(parts)
    return None


def run_fmt(rPr) -> dict:
    """단일 run의 서식 dict 반환"""
    if rPr is None:
        return {}
    latin = rPr.find(f"{{{NS_A}}}latin")
    sz_raw = rPr.get("sz")
    return {
        "font": latin.get("typeface") if latin is not None else None,
        "sz_pt": round(int(sz_raw) / 100) if sz_raw else None,
        "bold": rPr.get("b"),
        "italic": rPr.get("i"),
        "color": get_color(rPr),
    }


def para_fmt(para: etree._Element) -> dict:
    """paragraph 정렬 등"""
    pPr = para.find(f"{{{NS_A}}}pPr")
    if pPr is None:
        return {}
    return {"algn": pPr.get("algn")}


def extract_shapes(pptx_path: Path, slide_name: str) -> dict[str, dict]:
    """shape_id → {text, runs:[{text, fmt}], para_fmts:[{algn}]} 매핑"""
    with zipfile.ZipFile(pptx_path) as z:
        root = etree.fromstring(z.read(slide_name))

    shapes = {}
    for sp in root.iter(f"{{{NS_P}}}sp"):
        for el in sp.iter():
            if el.tag.endswith("}cNvPr"):
                sid  = el.get("id", "?")
                name = el.get("name", "")
                break
        else:
            continue

        txBody = sp.find(f".//{{{NS_P}}}txBody")
        if txBody is None:
            continue

        all_text = "".join(
            t.text for t in txBody.iter(f"{{{NS_A}}}t") if t.text
        )
        if not all_text.strip():
            continue

        runs = []
        pfmts = []
        for para in txBody.findall(f"{{{NS_A}}}p"):
            pfmts.append(para_fmt(para))
            for r in para.findall(f"{{{NS_A}}}r"):
                t_el = r.find(f"{{{NS_A}}}t")
                if t_el is None or not t_el.text:
                    continue
                rPr = r.find(f"{{{NS_A}}}rPr")
                runs.append({"text": t_el.text, "fmt": run_fmt(rPr)})

        shapes[sid] = {
            "name":       name,
            "text":       all_text,
            "runs":       runs,
            "para_fmts":  pfmts,
        }
    return shapes


# ── 비교 ────────────────────────────────────────────────────────────────────

def compare_fmt(a: dict, b: dict) -> list[str]:
    diffs = []
    for key in ("font", "sz_pt", "bold", "italic", "color"):
        va, vb = a.get(key), b.get(key)
        if va != vb:
            diffs.append(f"{key}: 원본={va!r} / 재생성={vb!r}")
    return diffs


def compare_shape(orig: dict, regen: dict) -> list[str]:
    issues = []

    # 텍스트 전체
    if orig["text"] != regen["text"]:
        issues.append(f"텍스트: {orig['text'][:60]!r} ≠ {regen['text'][:60]!r}")

    # run 서식 비교 (run 개수가 다를 수 있으므로 합쳐서 대표 서식끼리 비교)
    # 같은 위치 run끼리 비교
    for i, (ro, rr) in enumerate(zip(orig["runs"], regen["runs"])):
        fdiffs = compare_fmt(ro["fmt"], rr["fmt"])
        for d in fdiffs:
            # 텍스트가 비어있는 run의 서식 차이는 무시
            if ro["text"].strip() or rr["text"].strip():
                issues.append(f"  run[{i}] '{ro['text'][:30]}': {d}")

    # paragraph 정렬
    for i, (po, pr) in enumerate(zip(orig["para_fmts"], regen["para_fmts"])):
        if po.get("algn") != pr.get("algn"):
            issues.append(f"  para[{i}] 정렬: {po.get('algn')!r} → {pr.get('algn')!r}")

    return issues


# ── main ────────────────────────────────────────────────────────────────────

def main():
    orig_order  = slide_order(ORIG_PPTX)
    regen_order = slide_order(REGEN_PPTX)

    total_shape_diffs = 0
    slide_diffs = 0

    for sn, (on, rn) in enumerate(zip(orig_order, regen_order), 1):
        orig_shapes  = extract_shapes(ORIG_PPTX,  on)
        regen_shapes = extract_shapes(REGEN_PPTX, rn)

        slide_issues = []

        for sid, orig_s in orig_shapes.items():
            if sid not in regen_shapes:
                slide_issues.append(f"  [ID={sid} '{orig_s['name']}'] 재생성에 없음")
                continue
            regen_s = regen_shapes[sid]
            issues = compare_shape(orig_s, regen_s)
            if issues:
                total_shape_diffs += 1
                slide_issues.append(f"  [ID={sid} '{orig_s['name']}']")
                for iss in issues:
                    slide_issues.append(f"    {iss}")

        if slide_issues:
            slide_diffs += 1
            print(f"\n슬라이드 {sn} ({on}):")
            for line in slide_issues:
                print(line)

    print()
    if slide_diffs == 0:
        print("✅ 모든 슬라이드 텍스트+서식 일치!")
    else:
        print(f"⚠  {slide_diffs}개 슬라이드, {total_shape_diffs}개 shape에 차이 있음")


if __name__ == "__main__":
    main()
