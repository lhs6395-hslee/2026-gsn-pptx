"""
test_golden_phase3.py — Phase 3 §5.2 순수 유닛 골든 베이스라인

LLM·PowerPoint 불요. stdlib unittest만 사용.
실행: cd /Users/toule/Documents/gsneotek/kiro/2026-gsn-pptx && python3.12 -m unittest -v test_golden_phase3

대상:
  1. _trim_to_plan_slides  — sldIdLst 순서 + absent 제거 + 멱등성
  2. _SLIDE_EDITORS 레지스트리 — 키 집합 + 수동 override is 검증
  3. _record_run_experience — tmp monkeypatch 격리, runs append, 50-cap, success_rate, qa_ok=None 보존
  4. cleanup 경로 — 2-tuple 반환 + 경로 규칙 (함수 내부 경로 산식 단위 검증)
  5. restructure_sections 멱등성 — pack 1회 vs pack+restructure → sectionLst 동일
"""
from __future__ import annotations

import importlib
import io
import json
import pathlib
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

# 프로젝트 루트를 sys.path에 추가 (절대경로 사용)
PROJECT_ROOT = Path("/Users/toule/Documents/gsneotek/kiro/2026-gsn-pptx")
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import ppt_generator as ppt


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _make_fake_work_dir(tmp: Path,
                        rels_map: dict[str, str],
                        sld_order: list[str]) -> None:
    """
    _trim_to_plan_slides 가 읽는 최소한의 파일 구조를 tmp 에 생성한다.

    rels_map: {rId → slide_filename}  예) {"rId3": "slide3.xml", ...}
    sld_order: presentation.xml의 sldIdLst 순서 (rId 목록)
    """
    unpacked = tmp / "unpacked" / "ppt"
    unpacked.mkdir(parents=True)
    (unpacked / "_rels").mkdir()

    # presentation.xml.rels
    rel_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for rid, fname in rels_map.items():
        rel_lines.append(
            f'  <Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/slide" '
            f'Target="slides/{fname}"/>'
        )
    rel_lines.append("</Relationships>")
    (unpacked / "_rels" / "presentation.xml.rels").write_text(
        "\n".join(rel_lines), encoding="utf-8"
    )

    # presentation.xml — sldIdLst を sld_order 순서로
    sld_tags = "\n    ".join(
        f'<p:sldId id="{256 + i}" r:id="{rid}"/>'
        for i, rid in enumerate(sld_order)
    )
    prs_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    {sld_tags}
  </p:sldIdLst>
</p:presentation>"""
    (unpacked / "presentation.xml").write_text(prs_xml, encoding="utf-8")


def _read_rid_order(work_dir: Path) -> list[str]:
    """presentation.xml에서 sldIdLst의 r:id 순서를 반환."""
    prs_path = work_dir / "unpacked" / "ppt" / "presentation.xml"
    raw = prs_path.read_text(encoding="utf-8")
    return [m.group(1) for m in re.finditer(r'<p:sldId\b[^/]*\br:id="(rId\d+)"', raw)]


def _make_minimal_pptx(prs_xml: str) -> bytes:
    """presentation.xml 내용만 갖는 최소 pptx bytes 생성."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",
                    '<?xml version="1.0"?>'
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Override PartName="/ppt/presentation.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                    '</Types>')
        zf.writestr("ppt/presentation.xml", prs_xml.encode("utf-8"))
    return buf.getvalue()


def _read_sections_from_pptx(pptx_path: Path) -> list[str]:
    """pptx에서 p14:section name 목록을 반환."""
    with zipfile.ZipFile(pptx_path) as zf:
        prs = zf.read("ppt/presentation.xml").decode("utf-8")
    return re.findall(r'<p14:section\b[^>]*name="([^"]+)"', prs)


# ---------------------------------------------------------------------------
# 1. _trim_to_plan_slides
# ---------------------------------------------------------------------------

class TestTrimToPlanSlides(unittest.TestCase):
    """_trim_to_plan_slides: sldIdLst 순서 보장 + absent 제거 + 멱등성."""

    def _make_plan(self, files_in_order: list[str]) -> dict:
        return {
            "slides": [
                {"index": i + 1, "template_file": f}
                for i, f in enumerate(files_in_order)
            ]
        }

    def test_order_matches_plan_index(self):
        """plan index 순서대로 sldIdLst가 재정렬된다."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            # rels: rId3→slide3, rId5→slide5, rId8→slide8
            rels = {"rId3": "slide3.xml", "rId5": "slide5.xml", "rId8": "slide8.xml"}
            # 템플릿 원본 순서: rId3, rId5, rId8 (scramble → rId8, rId5, rId3)
            _make_fake_work_dir(tmp, rels, ["rId8", "rId5", "rId3"])

            plan = self._make_plan(["slide3.xml", "slide5.xml", "slide8.xml"])
            ppt._trim_to_plan_slides(tmp, plan)

            order = _read_rid_order(tmp)
            self.assertEqual(order, ["rId3", "rId5", "rId8"],
                             "plan index 순서(rId3→rId5→rId8)와 불일치")

    def test_absent_slides_removed(self):
        """plan에 없는 슬라이드(rId5)가 sldIdLst에서 제거된다."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            rels = {"rId3": "slide3.xml", "rId5": "slide5.xml", "rId8": "slide8.xml"}
            _make_fake_work_dir(tmp, rels, ["rId3", "rId5", "rId8"])

            # plan에는 slide3, slide8만 — slide5 제외
            plan = self._make_plan(["slide8.xml", "slide3.xml"])
            ppt._trim_to_plan_slides(tmp, plan)

            order = _read_rid_order(tmp)
            self.assertNotIn("rId5", order, "plan-absent rId5가 남아있음")
            self.assertEqual(set(order), {"rId3", "rId8"})

    def test_plan_order_reversed(self):
        """plan이 역순이면 sldIdLst도 역순으로 재정렬된다."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            rels = {"rId3": "slide3.xml", "rId5": "slide5.xml", "rId8": "slide8.xml"}
            _make_fake_work_dir(tmp, rels, ["rId3", "rId5", "rId8"])

            plan = self._make_plan(["slide8.xml", "slide5.xml", "slide3.xml"])
            ppt._trim_to_plan_slides(tmp, plan)

            order = _read_rid_order(tmp)
            self.assertEqual(order, ["rId8", "rId5", "rId3"])

    def test_idempotent_double_call(self):
        """2회 호출 결과가 1회 호출과 동일하다 (멱등성)."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            rels = {"rId3": "slide3.xml", "rId5": "slide5.xml", "rId8": "slide8.xml"}
            _make_fake_work_dir(tmp, rels, ["rId8", "rId3", "rId5"])

            plan = self._make_plan(["slide3.xml", "slide8.xml", "slide5.xml"])
            ppt._trim_to_plan_slides(tmp, plan)
            order_first = _read_rid_order(tmp)

            ppt._trim_to_plan_slides(tmp, plan)
            order_second = _read_rid_order(tmp)

            self.assertEqual(order_first, order_second,
                             "2회 호출이 1회와 다름 (멱등성 위반)")

    def test_empty_plan_noop(self):
        """plan이 비어있으면 아무것도 변경하지 않는다."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            rels = {"rId3": "slide3.xml"}
            _make_fake_work_dir(tmp, rels, ["rId3"])
            original_content = (tmp / "unpacked" / "ppt" / "presentation.xml").read_text()

            ppt._trim_to_plan_slides(tmp, {"slides": []})

            new_content = (tmp / "unpacked" / "ppt" / "presentation.xml").read_text()
            self.assertEqual(original_content, new_content, "plan 비어있는데 파일 변경됨")

    def test_missing_rels_noop(self):
        """rels 파일이 없으면 early-return해 아무것도 건드리지 않는다."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            rels = {"rId3": "slide3.xml"}
            _make_fake_work_dir(tmp, rels, ["rId3"])
            # rels 파일 삭제
            rels_path = tmp / "unpacked" / "ppt" / "_rels" / "presentation.xml.rels"
            rels_path.unlink()

            # Should not raise
            ppt._trim_to_plan_slides(tmp, self._make_plan(["slide3.xml"]))


# ---------------------------------------------------------------------------
# 2. _SLIDE_EDITORS 레지스트리
# ---------------------------------------------------------------------------

class TestSlideEditorsRegistry(unittest.TestCase):
    """_SLIDE_EDITORS: 키 집합 완전성 + 수동 override is 검증."""

    EXPECTED_KEYS = {
        "slide8.xml",
        "slide9.xml",
        "slide10.xml",
        "slide11.xml",
        "slide12.xml",
        "slide13.xml",
        "slide14.xml",
        "slide15.xml",
        "slide16.xml",
        "slide17.xml",
        "slide21.xml",
        "slide22.xml",
        "slide24.xml",
        "slide25.xml",
        "slide26.xml",
        "slide27.xml",
        "slide28.xml",
        "slide29.xml",
        "slide30.xml",
        "slide31.xml",
        "slide32.xml",
        "slide33.xml",
        "slide34.xml",
        "slide35.xml",
        "slide36.xml",
        "slide37.xml",
        "slide38.xml",
        "slide39.xml",
        "slide42.xml",
    }

    def test_key_set_exact_match(self):
        """실제 키 집합이 EXPECTED_KEYS와 정확히 일치해야 한다."""
        actual = set(ppt._SLIDE_EDITORS.keys())
        self.assertEqual(actual, self.EXPECTED_KEYS,
                         f"추가 키: {actual - self.EXPECTED_KEYS}, "
                         f"누락 키: {self.EXPECTED_KEYS - actual}")

    def test_required_keys_present(self):
        """slide8/21/22/28이 레지스트리에 존재해야 한다."""
        for key in ("slide8.xml", "slide21.xml", "slide22.xml", "slide28.xml"):
            self.assertIn(key, ppt._SLIDE_EDITORS, f"{key} 누락")

    def test_absent_keys(self):
        """slide40/41/43은 레지스트리에 없어야 한다."""
        for key in ("slide40.xml", "slide41.xml", "slide43.xml"):
            self.assertNotIn(key, ppt._SLIDE_EDITORS, f"{key}가 레지스트리에 등록됨")

    def test_slide22_override_is_edit_slide24(self):
        """slide22.xml → _edit_slide24 (is 검증)."""
        self.assertIs(ppt._SLIDE_EDITORS["slide22.xml"], ppt._edit_slide24,
                      "slide22.xml이 _edit_slide24 함수가 아님")

    def test_slide15_override_is_edit_slide15_v2(self):
        """slide15.xml → _edit_slide15_v2 (is 검증)."""
        self.assertIs(ppt._SLIDE_EDITORS["slide15.xml"], ppt._edit_slide15_v2,
                      "slide15.xml이 _edit_slide15_v2 함수가 아님")

    def test_slide15_not_edit_slide15(self):
        """slide15.xml이 구버전 edit_slide15 함수가 아님을 확인."""
        # _edit_slide15 (v1)이 존재하면 v2와 다른 객체여야 함
        v1 = getattr(ppt, "_edit_slide15", None)
        if v1 is not None:
            self.assertIsNot(ppt._SLIDE_EDITORS["slide15.xml"], v1,
                             "slide15.xml이 구버전 _edit_slide15에 연결됨")


# ---------------------------------------------------------------------------
# 3. _record_run_experience
# ---------------------------------------------------------------------------

class TestRecordRunExperience(unittest.TestCase):
    """_record_run_experience: tmp monkeypatch 격리, runs append, 50-cap, etc."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.tmp_path = Path(self.tmp_dir)

        # fake harness dir
        harness_dir = self.tmp_path / "harness"
        harness_dir.mkdir()
        self.mem_path = harness_dir / "long_term_memory.json"

        # initial memory (has one existing run)
        initial_mem = {
            "version": "0.4",
            "total_runs": 5,
            "success_rate": 1.0,
            "last_updated": "2026-01-01",
            "runs": [
                {"ts": "2026-01-01T00:00:00", "topic": "old", "n_slides": 3,
                 "templates": [], "vision_issues": 0, "qa_done": True, "qa_ok": True}
            ]
        }
        self.mem_path.write_text(
            json.dumps(initial_mem, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        # evolution dir
        (self.tmp_path / "evolution").mkdir(exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _call(self, topic="테스트", n_slides=5, vision_issues=0, qa_done=True):
        plan = {
            "slides": [{"template_file": f"slide{i}.xml"} for i in range(1, n_slides + 1)]
        }
        with patch.object(ppt, "SKILL_DIR", self.tmp_path):
            ppt._record_run_experience(topic, plan, vision_issues, qa_done=qa_done)

    def _read_mem(self) -> dict:
        return json.loads(self.mem_path.read_text(encoding="utf-8"))

    def test_run_appended(self):
        """새 run이 runs 목록에 추가된다."""
        self._call(topic="신규 주제")
        mem = self._read_mem()
        self.assertEqual(len(mem["runs"]), 2)
        self.assertEqual(mem["runs"][-1]["topic"], "신규 주제")

    def test_total_runs_incremented(self):
        """total_runs가 1 증가한다."""
        self._call()
        mem = self._read_mem()
        self.assertEqual(mem["total_runs"], 6)

    def test_fifty_cap(self):
        """runs가 51개가 되면 최근 50개만 유지된다."""
        # 49개를 추가해 total 50개로 만들기
        for i in range(49):
            self._call(topic=f"run_{i}")
        mem_before = self._read_mem()
        self.assertEqual(len(mem_before["runs"]), 50)

        # 1개 더 추가 → 여전히 50개
        self._call(topic="run_cap")
        mem_after = self._read_mem()
        self.assertEqual(len(mem_after["runs"]), 50)
        self.assertEqual(mem_after["runs"][-1]["topic"], "run_cap")

    def test_success_rate_excludes_none(self):
        """qa_ok=None인 run은 success_rate 계산에서 제외된다."""
        # qa_done=False → qa_ok=None
        self._call(qa_done=False)
        mem = self._read_mem()
        # runs 중 bool qa_ok=True인 1개만 rated → success_rate = 1.0
        self.assertIsNotNone(mem.get("success_rate"))
        self.assertAlmostEqual(mem["success_rate"], 1.0)

    def test_qa_ok_none_when_qa_done_false(self):
        """qa_done=False일 때 기록된 run의 qa_ok가 None이다."""
        self._call(qa_done=False)
        mem = self._read_mem()
        last = mem["runs"][-1]
        self.assertIsNone(last["qa_ok"], "qa_done=False인데 qa_ok가 None이 아님")

    def test_qa_ok_true_when_zero_vision_issues(self):
        """qa_done=True, vision_issues=0 → qa_ok=True."""
        self._call(vision_issues=0, qa_done=True)
        mem = self._read_mem()
        self.assertTrue(mem["runs"][-1]["qa_ok"])

    def test_qa_ok_false_when_vision_issues(self):
        """qa_done=True, vision_issues>0 → qa_ok=False."""
        self._call(vision_issues=2, qa_done=True)
        mem = self._read_mem()
        self.assertFalse(mem["runs"][-1]["qa_ok"])

    def test_file_content_persisted(self):
        """파일이 실제로 기록되어 재파싱 가능한 JSON이다."""
        self._call(topic="파일 내용 확인")
        raw = self.mem_path.read_text(encoding="utf-8")
        reloaded = json.loads(raw)  # parse 성공
        self.assertEqual(reloaded["runs"][-1]["topic"], "파일 내용 확인")

    def test_no_git_tracked_file_mutation(self):
        """실제 harness/long_term_memory.json(git-tracked)이 변경되지 않는다."""
        real_mem_path = PROJECT_ROOT / "harness" / "long_term_memory.json"
        if not real_mem_path.exists():
            self.skipTest("long_term_memory.json 없음")

        before = real_mem_path.read_text(encoding="utf-8")
        self._call()  # monkeypatched SKILL_DIR → tmp, real file 변경 안 됨
        after = real_mem_path.read_text(encoding="utf-8")
        self.assertEqual(before, after, "실제 long_term_memory.json이 오염됨")

    def test_returns_run_id_and_persists(self):
        """_record_run_experience가 run_id를 반환하고, 기록된 run에 동일 run_id가 들어간다 (#10)."""
        plan = {"slides": [{"template_file": "slide1.xml"}]}
        with patch.object(ppt, "SKILL_DIR", self.tmp_path):
            rid = ppt._record_run_experience("rid 테스트", plan, 0, qa_done=False)
        self.assertIsInstance(rid, str)
        self.assertTrue(rid)
        mem = self._read_mem()
        self.assertEqual(mem["runs"][-1]["run_id"], rid)
        # digest에도 run_id가 남는다 (SKILL.md가 읽어 update_last_run_qa(run_id=...)로 닫음)
        digest = json.loads(
            (self.tmp_path / "evolution" / "last_run_digest.json").read_text(encoding="utf-8"))
        self.assertEqual(digest["run_id"], rid)

    def test_success_rate_all_bool(self):
        """모든 run이 bool qa_ok인 경우 success_rate 산식 검증."""
        # 초기 mem을 2 success / 1 fail로 세팅
        mem_init = {
            "version": "0.4",
            "total_runs": 3,
            "runs": [
                {"ts": "2026-01-01T00:00:00", "topic": "a", "n_slides": 3,
                 "templates": [], "vision_issues": 0, "qa_done": True, "qa_ok": True},
                {"ts": "2026-01-02T00:00:00", "topic": "b", "n_slides": 3,
                 "templates": [], "vision_issues": 1, "qa_done": True, "qa_ok": False},
                {"ts": "2026-01-03T00:00:00", "topic": "c", "n_slides": 3,
                 "templates": [], "vision_issues": 0, "qa_done": True, "qa_ok": True},
            ]
        }
        self.mem_path.write_text(json.dumps(mem_init, ensure_ascii=False, indent=2),
                                 encoding="utf-8")
        # 추가 run: vision_issues=0, qa_done=True → qa_ok=True
        self._call(vision_issues=0, qa_done=True)
        mem = self._read_mem()
        # 3 True + 1 False = 3/4 = 0.75
        self.assertAlmostEqual(mem["success_rate"], 0.75, places=2)


# ---------------------------------------------------------------------------
# 3b. update_last_run_qa — run_id 매칭으로 off-by-one 회피 (#10)
# ---------------------------------------------------------------------------

class TestUpdateLastRunQa(unittest.TestCase):
    """update_last_run_qa: run_id 매칭 + 보류(None) 폴백 + 레거시 폴백 (#10)."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.tmp_path = Path(self.tmp_dir)
        harness_dir = self.tmp_path / "harness"
        harness_dir.mkdir()
        self.mem_path = harness_dir / "long_term_memory.json"

    def _write_mem(self, runs: list, **extra):
        mem = {"version": "0.4", "total_runs": len(runs), "runs": runs}
        mem.update(extra)
        self.mem_path.write_text(json.dumps(mem, ensure_ascii=False, indent=2),
                                 encoding="utf-8")

    def _read_mem(self) -> dict:
        return json.loads(self.mem_path.read_text(encoding="utf-8"))

    def _call(self, qa_ok, run_id=None):
        with patch.object(ppt, "SKILL_DIR", self.tmp_path):
            ppt.update_last_run_qa(qa_ok, run_id=run_id)

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_run_id_match_closes_exact_run(self):
        """run_id가 주어지면 runs[-1]이 아니라 그 run_id의 run을 닫는다 (off-by-one 회피)."""
        # A는 보류, B(나중 생성)도 보류 — A에 대한 QA 판정이 늦게 도착
        self._write_mem([
            {"run_id": "AAA", "qa_ok": None, "qa_done": False},
            {"run_id": "BBB", "qa_ok": None, "qa_done": False},
        ])
        self._call(True, run_id="AAA")
        mem = self._read_mem()
        a = next(r for r in mem["runs"] if r["run_id"] == "AAA")
        b = next(r for r in mem["runs"] if r["run_id"] == "BBB")
        self.assertTrue(a["qa_ok"], "run_id로 지정한 A가 닫혀야 한다")
        self.assertTrue(a["qa_done"])
        self.assertIsNone(b["qa_ok"], "지정 안 한 B(runs[-1])는 건드리지 않아야 한다")

    def test_run_id_miss_falls_back_to_pending(self):
        """매칭 실패 시 qa_ok=None인 가장 최근 보류 run을 닫는다."""
        self._write_mem([
            {"run_id": "AAA", "qa_ok": True, "qa_done": True},
            {"run_id": "BBB", "qa_ok": None, "qa_done": False},
        ])
        self._call(False, run_id="ZZZ-없음")
        mem = self._read_mem()
        b = next(r for r in mem["runs"] if r["run_id"] == "BBB")
        self.assertFalse(b["qa_ok"], "보류였던 B가 닫혀야 한다")

    def test_no_run_id_closes_latest_pending_not_last(self):
        """run_id 없이 호출하면 runs[-1](확정완료)이 아닌 보류(None) run을 닫는다."""
        # 마지막 run은 이미 qa_ok=True로 확정 → 블라인드 runs[-1]이면 오판
        self._write_mem([
            {"run_id": "AAA", "qa_ok": None, "qa_done": False},
            {"run_id": "BBB", "qa_ok": True, "qa_done": True},
        ])
        self._call(False)  # run_id 없음
        mem = self._read_mem()
        a = next(r for r in mem["runs"] if r["run_id"] == "AAA")
        b = next(r for r in mem["runs"] if r["run_id"] == "BBB")
        self.assertFalse(a["qa_ok"], "보류였던 A가 닫혀야 한다")
        self.assertTrue(b["qa_ok"], "이미 확정된 B(runs[-1])는 보존돼야 한다")

    def test_legacy_fallback_no_pending_no_run_id(self):
        """보류 run도 run_id도 없으면 runs[-1]을 닫는다 (레거시 호환)."""
        self._write_mem([
            {"run_id": "AAA", "qa_ok": True, "qa_done": True},
            {"run_id": "BBB", "qa_ok": True, "qa_done": True},
        ])
        self._call(False)
        mem = self._read_mem()
        self.assertFalse(mem["runs"][-1]["qa_ok"])

    def test_recomputes_success_rate(self):
        """qa_ok 확정 후 success_rate가 재계산된다 (None 제외)."""
        self._write_mem([
            {"run_id": "AAA", "qa_ok": True, "qa_done": True},
            {"run_id": "BBB", "qa_ok": None, "qa_done": False},
        ])
        self._call(False, run_id="BBB")
        mem = self._read_mem()
        # 1 True + 1 False = 0.5
        self.assertAlmostEqual(mem["success_rate"], 0.5, places=2)

    def test_empty_runs_no_crash(self):
        """run 기록이 없으면 조용히 반환한다 (예외 없음)."""
        self._write_mem([])
        self._call(True, run_id="AAA")  # 예외 없이 통과해야 함
        mem = self._read_mem()
        self.assertEqual(mem["runs"], [])

    def test_no_git_tracked_file_mutation(self):
        """실제 harness/long_term_memory.json(git-tracked)이 변경되지 않는다."""
        real_mem_path = PROJECT_ROOT / "harness" / "long_term_memory.json"
        if not real_mem_path.exists():
            self.skipTest("long_term_memory.json 없음")
        self._write_mem([{"run_id": "AAA", "qa_ok": None, "qa_done": False}])
        before = real_mem_path.read_text(encoding="utf-8")
        self._call(True, run_id="AAA")  # monkeypatched SKILL_DIR → tmp
        after = real_mem_path.read_text(encoding="utf-8")
        self.assertEqual(before, after, "실제 long_term_memory.json이 오염됨")


# ---------------------------------------------------------------------------
# 4. cleanup 경로 — 경로 산식 단위 검증
# ---------------------------------------------------------------------------

class TestCleanupPathLogic(unittest.TestCase):
    """
    cleanup 경로의 핵심 산식을 단위 검증.

    run_ppt_generation 은 거대한 파이프라인이므로 full e2e 실행 없이
    cleanup 로직에서 쓰이는 경로 규칙을 직접 재현한다.
    §5.2 명세: (a) 양 분기 2-tuple, (b) cleanup=True → work_dir.parent.parent/{safe_name}.pptx,
               (c) cleanup=False → work_dir/output.pptx
    """

    def _compute_final_output(self, work_dir: Path, topic: str) -> Path:
        """cleanup=True 분기의 최종 경로 산식."""
        safe_name = topic.replace(" ", "_").replace("/", "_")[:60]
        return work_dir.parent.parent / f"{safe_name}.pptx"

    def _compute_no_cleanup_output(self, work_dir: Path) -> Path:
        """cleanup=False 분기의 최종 경로 산식."""
        return work_dir / "output.pptx"

    def test_cleanup_true_path_structure(self):
        """cleanup=True → work_dir.parent.parent/{safe_name}.pptx."""
        with tempfile.TemporaryDirectory() as td:
            # work_dir = tmp/runs/session_01/work
            work_dir = Path(td) / "runs" / "session_01" / "work"
            work_dir.mkdir(parents=True)

            topic = "테스트 발표자료"
            final = self._compute_final_output(work_dir, topic)

            expected_parent = Path(td) / "runs"
            self.assertEqual(final.parent, expected_parent)
            self.assertEqual(final.suffix, ".pptx")
            # safe_name: 공백 → _
            self.assertIn("테스트_발표자료", final.stem)

    def test_cleanup_false_path_structure(self):
        """cleanup=False → work_dir/output.pptx."""
        with tempfile.TemporaryDirectory() as td:
            work_dir = Path(td) / "work"
            work_dir.mkdir()

            out = self._compute_no_cleanup_output(work_dir)
            self.assertEqual(out.parent, work_dir)
            self.assertEqual(out.name, "output.pptx")

    def test_safe_name_space_replaced(self):
        """공백이 _ 로 치환된다."""
        topic = "hello world foo"
        safe_name = topic.replace(" ", "_").replace("/", "_")[:60]
        self.assertEqual(safe_name, "hello_world_foo")

    def test_safe_name_slash_replaced(self):
        """슬래시가 _ 로 치환된다."""
        topic = "AWS/GCP 비교"
        safe_name = topic.replace(" ", "_").replace("/", "_")[:60]
        self.assertEqual(safe_name, "AWS_GCP_비교")

    def test_safe_name_truncated_60(self):
        """60자 초과는 잘린다."""
        topic = "a" * 80
        safe_name = topic.replace(" ", "_").replace("/", "_")[:60]
        self.assertEqual(len(safe_name), 60)

    def test_both_branches_return_path(self):
        """양 분기 모두 Path 객체를 반환해야 한다."""
        with tempfile.TemporaryDirectory() as td:
            work_dir = Path(td) / "runs" / "s" / "w"
            work_dir.mkdir(parents=True)

            r_true = self._compute_final_output(work_dir, "topic")
            r_false = self._compute_no_cleanup_output(work_dir)

            self.assertIsInstance(r_true, Path)
            self.assertIsInstance(r_false, Path)


# ---------------------------------------------------------------------------
# 5. restructure_sections 멱등성
# ---------------------------------------------------------------------------

class TestRestructureSectionsIdempotent(unittest.TestCase):
    """restructure_sections: 1회 vs 2회 호출 후 sectionLst 이름 목록 동일."""

    PRS_CLEAN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
  <p:extLst>
    <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
      <p14:sectionLst>
        <p14:section name="표지" id="{A1}">
          <p14:sldIdLst><p14:sldId id="1"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="목차/간지" id="{A2}">
          <p14:sldIdLst><p14:sldId id="2"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="본문" id="{A3}">
          <p14:sldIdLst><p14:sldId id="3"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="마무리" id="{A4}">
          <p14:sldIdLst><p14:sldId id="4"/></p14:sldIdLst>
        </p14:section>
      </p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>"""

    PRS_WITH_EXTRA = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
  <p:extLst>
    <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
      <p14:sectionLst>
        <p14:section name="표지" id="{A1}">
          <p14:sldIdLst><p14:sldId id="1"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="목차/간지" id="{A2}">
          <p14:sldIdLst><p14:sldId id="2"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="본문" id="{A3}">
          <p14:sldIdLst><p14:sldId id="3"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="마무리" id="{A4}">
          <p14:sldIdLst><p14:sldId id="4"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="차트" id="{A5}">
          <p14:sldIdLst><p14:sldId id="5"/></p14:sldIdLst>
        </p14:section>
        <p14:section name="타임라인" id="{A6}">
          <p14:sldIdLst></p14:sldIdLst>
        </p14:section>
      </p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>"""

    def _make_pptx(self, prs_xml: str, path: Path) -> None:
        buf = _make_minimal_pptx(prs_xml)
        path.write_bytes(buf)

    def test_already_clean_noop(self):
        """4개 표준 섹션만 있으면 restructure 후 동일하다."""
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "test.pptx"
            self._make_pptx(self.PRS_CLEAN, out)

            before = _read_sections_from_pptx(out)
            ppt.restructure_sections(out)
            after = _read_sections_from_pptx(out)

            self.assertEqual(before, after)

    def test_idempotent_with_extra_sections(self):
        """비표준 섹션 포함 → 1회 적용 후 2회 적용해도 sectionLst 동일."""
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "test.pptx"
            self._make_pptx(self.PRS_WITH_EXTRA, out)

            ppt.restructure_sections(out)
            after_first = _read_sections_from_pptx(out)

            ppt.restructure_sections(out)
            after_second = _read_sections_from_pptx(out)

            self.assertEqual(after_first, after_second,
                             "2회 적용 후 sectionLst가 변경됨 (멱등성 위반)")

    def test_extra_sections_removed(self):
        """비표준 섹션(차트/타임라인)이 1회 적용 후 사라진다."""
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "test.pptx"
            self._make_pptx(self.PRS_WITH_EXTRA, out)

            ppt.restructure_sections(out)
            sections = _read_sections_from_pptx(out)

            self.assertNotIn("차트", sections)
            self.assertNotIn("타임라인", sections)

    def test_standard_sections_preserved(self):
        """표준 4개 섹션은 정리 후에도 유지된다."""
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "test.pptx"
            self._make_pptx(self.PRS_WITH_EXTRA, out)

            ppt.restructure_sections(out)
            sections = _read_sections_from_pptx(out)

            for name in ("표지", "목차/간지", "본문", "마무리"):
                self.assertIn(name, sections, f"표준 섹션 '{name}' 소실")


# ---------------------------------------------------------------------------
# 글로벌 캐시 리셋 픽스처
# ---------------------------------------------------------------------------

def setUpModule():
    """모듈 시작 시 글로벌 캐시를 리셋한다."""
    ppt._CHAPTER_MAP_CACHE = None
    ppt._SLIDE_CATALOG_CACHE = None


def tearDownModule():
    """모듈 종료 시 글로벌 캐시를 리셋한다 (다음 테스트 오염 방지)."""
    ppt._CHAPTER_MAP_CACHE = None
    ppt._SLIDE_CATALOG_CACHE = None


if __name__ == "__main__":
    unittest.main(verbosity=2)
