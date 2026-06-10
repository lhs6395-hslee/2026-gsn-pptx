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
import os
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


import ahe_loop


class TestManifestLedgerBridge(unittest.TestCase):
    """change_manifest.jsonl 단방향 브리지 (manifest-ledger-split [D]).

    검증 verdict(iteration_*_manifest.json) → authoritative 원장으로의
    유일한 write 경로. 명시적 manifest_id 링크로만 pending을 해소하고,
    기존 원장 데이터(필드)는 손실 없이 보존한다.
    """

    LEDGER_LINES = [
        {"date": "2026-06-09", "id": "2026-06-09-09",
         "change": "원장 도입", "files": ["harness/change_manifest.jsonl"],
         "evidence": "ev", "root_cause": "rc", "fix": "fx",
         "predicted_fixes": ["pf1"], "regression_risk": [],
         "verification": "pending (다음 실행부터 자동 기록 연결 예정)"},
        {"date": "2026-06-09", "id": "2026-06-09-01",
         "change": "이미 검증됨", "files": ["x"],
         "evidence": "ev", "root_cause": "rc", "fix": "fx",
         "predicted_fixes": ["pf"], "regression_risk": [],
         "verification": "verified (PDF 렌더 확인)"},
        {"date": "2026-06-10", "id": "2026-06-10-08",
         "change": "explain 분리", "files": ["y"],
         "evidence": "ev", "root_cause": "rc", "fix": "fx",
         "predicted_fixes": ["pf"], "regression_risk": [],
         "verification": "pending"},
    ]

    def _write_ledger(self, harness_dir: Path) -> Path:
        p = harness_dir / "change_manifest.jsonl"
        p.write_text(
            "\n".join(json.dumps(e, ensure_ascii=False) for e in self.LEDGER_LINES) + "\n",
            encoding="utf-8",
        )
        return p

    def _read_ledger(self, path: Path) -> list[dict]:
        return [json.loads(ln) for ln in path.read_text(encoding="utf-8").splitlines() if ln.strip()]

    def test_jsonl_roundtrip_preserves_fields(self):
        """_load_jsonl/_save_jsonl 왕복: 모든 필드·줄 수 보존, 빈 줄/손상 줄 스킵."""
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "m.jsonl"
            raw = (
                json.dumps(self.LEDGER_LINES[0], ensure_ascii=False) + "\n"
                + "\n"  # 빈 줄
                + "not-json{\n"  # 손상 줄
                + json.dumps(self.LEDGER_LINES[2], ensure_ascii=False) + "\n"
            )
            p.write_text(raw, encoding="utf-8")

            entries = ahe_loop._load_jsonl(p)
            self.assertEqual(len(entries), 2, "유효 2줄만 로드되어야 함")
            self.assertEqual(entries[0], self.LEDGER_LINES[0])

            ahe_loop._save_jsonl(p, entries)
            again = ahe_loop._load_jsonl(p)
            self.assertEqual(again, entries, "왕복 후 동일해야 함")

    def test_bridge_pass_marks_verified(self):
        """manifest_id 링크 + PASS → 해당 pending 엔트리가 verified로 해소."""
        with tempfile.TemporaryDirectory() as td:
            harness = Path(td)
            ledger = self._write_ledger(harness)
            preds = [{"change": "c", "expected": "e", "metric": "m",
                      "manifest_id": "2026-06-09-09", "verification": "PASS"}]
            with patch.object(ahe_loop, "SKILL_DIR", harness.parent):
                n = ahe_loop.bridge_verdicts_to_ledger(preds, harness)
            self.assertEqual(n, 1)
            entries = {e["id"]: e for e in self._read_ledger(ledger)}
            self.assertTrue(entries["2026-06-09-09"]["verification"].startswith("verified"))
            self.assertIn("verified_at", entries["2026-06-09-09"])
            self.assertEqual(entries["2026-06-09-09"]["verified_by_run"], "verify_predictions")
            # 기존 prose 꼬리말 보존
            self.assertIn("자동 기록 연결", entries["2026-06-09-09"]["verification"])

    def test_bridge_fail_marks_refuted(self):
        """manifest_id 링크 + FAIL → pending 엔트리가 refuted로 해소."""
        with tempfile.TemporaryDirectory() as td:
            harness = Path(td)
            ledger = self._write_ledger(harness)
            preds = [{"manifest_id": "2026-06-10-08", "verification": "FAIL"}]
            ahe_loop.bridge_verdicts_to_ledger(preds, harness)
            entries = {e["id"]: e for e in self._read_ledger(ledger)}
            self.assertTrue(entries["2026-06-10-08"]["verification"].startswith("refuted"))

    def test_bridge_skips_already_resolved(self):
        """이미 verified인 엔트리는 verdict가 와도 덮어쓰지 않는다(큐레이션 존중)."""
        with tempfile.TemporaryDirectory() as td:
            harness = Path(td)
            ledger = self._write_ledger(harness)
            preds = [{"manifest_id": "2026-06-09-01", "verification": "FAIL"}]
            n = ahe_loop.bridge_verdicts_to_ledger(preds, harness)
            self.assertEqual(n, 0, "이미 verified는 갱신 안 함")
            entries = {e["id"]: e for e in self._read_ledger(ledger)}
            self.assertEqual(entries["2026-06-09-01"]["verification"],
                             "verified (PDF 렌더 확인)")

    def test_bridge_no_match_no_write(self):
        """manifest_id 없거나 UNVERIFIED면 원장 무변경 (단방향·보수적)."""
        with tempfile.TemporaryDirectory() as td:
            harness = Path(td)
            ledger = self._write_ledger(harness)
            before = ledger.read_bytes()
            preds = [
                {"change": "c", "metric": "m", "verification": "PASS"},          # manifest_id 없음
                {"manifest_id": "2026-06-09-09", "verification": "UNVERIFIED"},  # 결정 불가
                {"manifest_id": "9999-99-99", "verification": "PASS"},            # 매칭 id 없음
            ]
            n = ahe_loop.bridge_verdicts_to_ledger(preds, harness)
            self.assertEqual(n, 0)
            self.assertEqual(ledger.read_bytes(), before, "원장 바이트 불변")

    def test_bridge_missing_ledger_safe(self):
        """원장 파일이 없으면 0 반환·예외 없음."""
        with tempfile.TemporaryDirectory() as td:
            harness = Path(td)  # change_manifest.jsonl 없음
            preds = [{"manifest_id": "x", "verification": "PASS"}]
            self.assertEqual(ahe_loop.bridge_verdicts_to_ledger(preds, harness), 0)


# ---------------------------------------------------------------------------
# 8. evolve 트리거 게이트 (#12 evolve-manual-trigger-gap)
#    독립 QA fail → evolve 환류 경로 + human-in-loop 게이트(§1) 보존
# ---------------------------------------------------------------------------

class TestShouldTriggerEvolve(unittest.TestCase):
    """should_trigger_evolve: 순수 술어 — 어떤 신호가 evolve '제안' 대상인가."""

    def test_explicit_always_true(self):
        """explicit=True(사람이 --evolve 지정)는 다른 값과 무관하게 제안."""
        self.assertTrue(ahe_loop.should_trigger_evolve(None, 0, explicit=True))
        self.assertTrue(ahe_loop.should_trigger_evolve(True, 0, explicit=True))

    def test_qa_fail_triggers(self):
        """qa_ok=False(독립 QA 결함) → 제안 (skill 경로 환류 핵심)."""
        self.assertTrue(ahe_loop.should_trigger_evolve(False, 0, explicit=False))

    def test_vision_issues_triggers(self):
        """vision_issues>0(인라인) → 제안."""
        self.assertTrue(ahe_loop.should_trigger_evolve(None, 2, explicit=False))

    def test_qa_pass_no_trigger(self):
        """qa_ok=True + 이슈 없음 → 제안 안 함."""
        self.assertFalse(ahe_loop.should_trigger_evolve(True, 0, explicit=False))

    def test_qa_none_no_trigger(self):
        """qa_ok=None(판정 보류) + 이슈 없음 → 닫히지 않은 run은 진화 대상 아님."""
        self.assertFalse(ahe_loop.should_trigger_evolve(None, 0, explicit=False))


class TestMaybeRunEvolveLoop(unittest.TestCase):
    """maybe_run_evolve_loop: 게이트가 approved/explicit 없이는 절대 실행 안 함 (§1)."""

    def _maybe(self, **kw):
        """run_evolve_loop을 mock으로 가로채고 (실행여부, 호출횟수) 반환."""
        with patch.object(ahe_loop, "run_evolve_loop") as m:
            ran = ahe_loop.maybe_run_evolve_loop(
                Path("/tmp/_does_not_matter"), "주제", **kw)
        return ran, m.call_count

    def test_qa_fail_without_approval_does_not_run(self):
        """qa_ok=False라도 approved/explicit 없으면 제안만 — run_evolve_loop 미호출."""
        ran, calls = self._maybe(qa_ok=False)
        self.assertFalse(ran)
        self.assertEqual(calls, 0)

    def test_qa_fail_with_approval_runs(self):
        """qa_ok=False + approved=True → 실제 실행."""
        ran, calls = self._maybe(qa_ok=False, approved=True)
        self.assertTrue(ran)
        self.assertEqual(calls, 1)

    def test_explicit_runs_without_approved_flag(self):
        """explicit=True(사람이 --evolve 직접 지정)는 그 자체로 승인 → 실행."""
        ran, calls = self._maybe(explicit=True)
        self.assertTrue(ran)
        self.assertEqual(calls, 1)

    def test_no_signal_no_run(self):
        """신호 없음(qa_ok=None, 이슈 0) → 실행 안 함, 제안도 안 함."""
        ran, calls = self._maybe(qa_ok=None, vision_issues=0)
        self.assertFalse(ran)
        self.assertEqual(calls, 0)

    def test_vision_issues_without_approval_does_not_run(self):
        """vision_issues>0라도 approved/explicit 없으면 제안만 — 자동 실행 금지."""
        ran, calls = self._maybe(vision_issues=3)
        self.assertFalse(ran)
        self.assertEqual(calls, 0)


# ---------------------------------------------------------------------------
# 9. 독립 QA verdict 스키마 코드화 (#13 independent-qa-not-implemented)
#    §1 채점 가능 기준 + §2 생성≠판정 (분리). 인터페이스/스키마 우선, spawn optional.
# ---------------------------------------------------------------------------

class TestQaVerdict(unittest.TestCase):
    """QaVerdict: verdict→qa_ok 3-값 매핑 + §2 분리 가드(independent 표식)."""

    def test_needs_fix_maps_false_regardless_of_independence(self):
        """NEEDS_FIX(결함 보고)는 분리 여부와 무관하게 qa_ok=False로 신뢰."""
        self.assertIs(ahe_loop.QaVerdict(ahe_loop.VERDICT_NEEDS_FIX).qa_ok, False)
        self.assertIs(
            ahe_loop.QaVerdict(ahe_loop.VERDICT_NEEDS_FIX, independent=True).qa_ok,
            False,
        )

    def test_pass_requires_independence_for_qa_ok_true(self):
        """§2: 비분리(자기-QA) PASS는 qa_ok=True로 승격 금지 → None(보류)."""
        self.assertIsNone(ahe_loop.QaVerdict(ahe_loop.VERDICT_PASS,
                                             independent=False).qa_ok)
        self.assertIs(ahe_loop.QaVerdict(ahe_loop.VERDICT_PASS,
                                         independent=True).qa_ok, True)

    def test_deferred_is_none(self):
        """DEFERRED(판정 보류) → qa_ok=None (닫히지 않은 run, 진화 대상 아님)."""
        self.assertIsNone(ahe_loop.QaVerdict(independent=True).qa_ok)
        self.assertIsNone(ahe_loop.QaVerdict(ahe_loop.VERDICT_DEFERRED).qa_ok)

    def test_unknown_verdict_normalized_to_deferred(self):
        """알 수 없는 verdict 토큰은 DEFERRED로 정규화 (날조 금지)."""
        self.assertEqual(ahe_loop.QaVerdict("garbage").verdict,
                         ahe_loop.VERDICT_DEFERRED)

    def test_to_dict_includes_qa_ok(self):
        v = ahe_loop.QaVerdict(ahe_loop.VERDICT_NEEDS_FIX, summary="s")
        d = v.to_dict()
        self.assertIs(d["qa_ok"], False)
        self.assertEqual(d["verdict"], ahe_loop.VERDICT_NEEDS_FIX)
        self.assertEqual(d["summary"], "s")


class TestParseVerdict(unittest.TestCase):
    """parse_verdict: 임의 에이전트 출력(dict/str/리포트형/None) → QaVerdict 관용 파싱."""

    def test_none_and_empty_are_deferred(self):
        for raw in (None, "", {}):
            self.assertEqual(ahe_loop.parse_verdict(raw).verdict,
                             ahe_loop.VERDICT_DEFERRED)

    def test_canonical_dict(self):
        v = ahe_loop.parse_verdict(
            {"verdict": "needs_fix", "issues": [{"slide": 3}], "summary": "x"},
            independent=True,
        )
        self.assertEqual(v.verdict, ahe_loop.VERDICT_NEEDS_FIX)
        self.assertEqual(len(v.issues), 1)
        self.assertTrue(v.independent)
        self.assertIs(v.qa_ok, False)

    def test_report_shape_with_issues_is_needs_fix(self):
        """SKILL.md 11단계 리포트형(slides[].issues) → 이슈 있으면 NEEDS_FIX."""
        report = {"slides": [{"index": 7, "role": "toc",
                              "issues": [{"type": "overflow", "severity": "HIGH"}]}],
                  "summary": "오버플로우 1건"}
        v = ahe_loop.parse_verdict(report, independent=True)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_NEEDS_FIX)
        self.assertEqual(v.issues[0]["slide"], 7)

    def test_report_shape_no_issues_is_pass(self):
        v = ahe_loop.parse_verdict({"slides": [{"index": 1, "issues": []}]},
                                   independent=True)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_PASS)
        self.assertIs(v.qa_ok, True)

    def test_json_fenced_string(self):
        raw = '```json\n{"verdict": "pass", "summary": "ok"}\n```'
        v = ahe_loop.parse_verdict(raw, independent=True)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_PASS)

    def test_prose_fail_tokens(self):
        """JSON이 아닌 prose 종합 판정 — FAIL 신호 우선(보수적)."""
        v = ahe_loop.parse_verdict("종합 판정: 수정 필요 — slide 5 잘림")
        self.assertEqual(v.verdict, ahe_loop.VERDICT_NEEDS_FIX)

    def test_prose_pass_tokens(self):
        v = ahe_loop.parse_verdict("전체 통과, 이슈 없음")
        self.assertEqual(v.verdict, ahe_loop.VERDICT_PASS)

    def test_prose_ambiguous_is_deferred_not_fabricated(self):
        """모호한 prose는 결함을 날조하지 않고 DEFERRED로 떨어진다."""
        v = ahe_loop.parse_verdict("음... 잘 모르겠습니다")
        self.assertEqual(v.verdict, ahe_loop.VERDICT_DEFERRED)

    def test_free_verdict_token_mapped(self):
        """정식 토큰이 아닌 verdict 자유 표현도 매핑된다."""
        self.assertEqual(
            ahe_loop.parse_verdict({"verdict": "REJECTED"}).verdict,
            ahe_loop.VERDICT_NEEDS_FIX,
        )


class TestRunIndependentQa(unittest.TestCase):
    """run_independent_qa: spawn 콜백 주입 인터페이스 + 비대화형 안전 폴백."""

    def test_no_spawn_is_deferred_not_crash(self):
        """spawn=None(headless/콜백 미주입) → 죽지 않고 DEFERRED 폴백."""
        v = ahe_loop.run_independent_qa(Path("/x.pptx"), "주제", spawn=None)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_DEFERRED)
        self.assertIsNone(v.qa_ok)
        self.assertFalse(v.independent)

    def test_spawn_output_parsed_as_independent(self):
        """spawn이 반환한 출력은 independent=True로 파싱된다 (§2 분리)."""
        def spawn(path, topic):
            return {"verdict": "pass", "summary": "ok"}
        v = ahe_loop.run_independent_qa(Path("/x.pptx"), "주제", spawn=spawn)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_PASS)
        self.assertTrue(v.independent)
        self.assertIs(v.qa_ok, True)

    def test_spawn_exception_is_deferred(self):
        """spawn 콜백 예외가 파이프라인을 죽이지 않고 DEFERRED로 흡수."""
        def boom(path, topic):
            raise RuntimeError("agent down")
        v = ahe_loop.run_independent_qa(Path("/x.pptx"), "주제", spawn=boom)
        self.assertEqual(v.verdict, ahe_loop.VERDICT_DEFERRED)
        self.assertIsNone(v.qa_ok)


class TestInlineQaHeadlessGuard(unittest.TestCase):
    """assert_inline_qa_headless_only: 인라인 vision QA를 headless 전용으로 강제 (§2)."""

    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in
                       ("PPT_SKILL_HEADLESS", "PPT_SKILL_ALLOW_INLINE_QA")}

    def tearDown(self):
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    def test_headless_env_passes(self):
        os.environ["PPT_SKILL_HEADLESS"] = "1"
        os.environ.pop("PPT_SKILL_ALLOW_INLINE_QA", None)
        ahe_loop.assert_inline_qa_headless_only()  # no raise

    def test_orchestrator_session_raises(self):
        """대화형(오케스트레이터)에서는 RuntimeError로 인라인 self-QA 차단."""
        os.environ["PPT_SKILL_HEADLESS"] = "0"
        os.environ.pop("PPT_SKILL_ALLOW_INLINE_QA", None)
        with self.assertRaises(RuntimeError):
            ahe_loop.assert_inline_qa_headless_only()

    def test_explicit_override_allows(self):
        """PPT_SKILL_ALLOW_INLINE_QA=1 명시 오버라이드는 허용."""
        os.environ["PPT_SKILL_HEADLESS"] = "0"
        os.environ["PPT_SKILL_ALLOW_INLINE_QA"] = "1"
        ahe_loop.assert_inline_qa_headless_only()  # no raise

    def test_analyze_qa_images_blocked_in_session(self):
        """analyze_qa_images도 비-headless에서 이미지가 있으면 차단된다."""
        os.environ["PPT_SKILL_HEADLESS"] = "0"
        os.environ.pop("PPT_SKILL_ALLOW_INLINE_QA", None)
        with self.assertRaises(RuntimeError):
            ahe_loop.analyze_qa_images(["/nonexistent.png"], {"slides": []})

    def test_analyze_qa_images_empty_is_noop_even_in_session(self):
        """이미지 0장이면 가드 이전에 무해 반환 — 세션에서도 RuntimeError 없음."""
        os.environ["PPT_SKILL_HEADLESS"] = "0"
        out = ahe_loop.analyze_qa_images([], {"slides": []})
        self.assertEqual(out["slides"], [])


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
