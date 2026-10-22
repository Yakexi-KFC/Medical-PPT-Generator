"""Session and Streamlit regressions using synthetic images and mocked AI only."""

import copy
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from zipfile import ZipFile

from PIL import Image

try:
    import app_workflow as workflow
    from streamlit.testing.v1 import AppTest
except ImportError:
    workflow = None
    AppTest = None


class State(dict):
    def __getattr__(self, key):
        try:
            return self[key]
        except KeyError:
            raise AttributeError(key) from None

    def __setattr__(self, key, value):
        self[key] = value


def uploaded_image(name="report.png", color="white"):
    file = io.BytesIO()
    Image.new("RGB", (100, 100), color).save(file, format="PNG")
    file.name = name
    return file


def evidence(aid):
    return {"asset_id": aid, "title": "Report", "date": "", "category": "pathology",
            "findings": [], "transcription": "", "warnings": [],
            "regions": [{"id": f"{aid}-r1", "label": "First", "box": [0, 0, 500, 1000]},
                        {"id": f"{aid}-r2", "label": "Second", "box": [500, 0, 1000, 1000]}]}


REVIEW_APP = '''
import streamlit as st
import io
from PIL import Image
from app_workflow import review_assets
if "assets" not in st.session_state:
    raw = io.BytesIO()
    Image.new("RGB", (100,100), "white").save(raw, format="PNG")
    st.session_state.assets=[{"id":"img-001","name":"test.png","image_bytes":raw.getvalue(),"width":100,"height":100}]
    st.session_state.analyses={"img-001":{"asset_id":"img-001","title":"Report","date":"","category":"pathology","findings":[],"transcription":"","warnings":[],"regions":[{"id":"img-001-r1","label":"First","box":[0,0,500,1000]},{"id":"img-001-r2","label":"Second","box":[500,0,1000,1000]}]}}
    st.session_state.upload_signature="test_signature"
    st.session_state.refine_errors={}
    st.session_state.asset_errors={}
    st.session_state.analysis_versions={}
    st.session_state.crop_edits={}
assets, analyses, pending=review_assets()
st.session_state.review_result={"assets":assets,"analyses":analyses,"pending":pending}
'''


TEXT_APP = '''
import streamlit as st
import app_workflow as workflow
from case_pipeline import normalize_case
class FakeClient:
    def close(self):
        pass
def fake_extract(client, patient_text, analyses, model):
    st.session_state.call_count=st.session_state.get("call_count",0)+1
    return normalize_case({"baseline":{"chief_complaint":patient_text}})
workflow.make_client=lambda api_key: FakeClient()
workflow.extract_case=fake_extract
workflow.setting=lambda name, default="": "mock-only-no-API" if name == "DEEPSEEK_API_KEY" else default
workflow.render_workflow()
'''


RESULT_APP = '''
import streamlit as st
import io
from PIL import Image
from app_workflow import render_result, case_input_signature
from case_pipeline import normalize_case
from vision_pipeline import prepare_asset
if "fixture_assets" not in st.session_state:
    raw=io.BytesIO()
    image=Image.new("RGB", (100,80), "white")
    image.paste("blue", (25,0,75,40))
    image.save(raw,format="PNG")
    st.session_state.fixture_assets=[prepare_asset(raw.getvalue(),"synthetic-report.png","img-fixture")]
    analysis={"asset_id":"img-fixture","title":"Synthetic report","date":"2025-02-01","category":"pathology","findings":["Recorded finding"],"transcription":"Source report","warnings":[],"regions":[{"id":"img-fixture-r1","label":"Report panel","box":[0,0,1000,1000]}]}
    st.session_state.fixture_analyses=[analysis]
    case=normalize_case({"baseline":{},"treatments":[{"phase":"Recorded treatment","duration":"2025-03"}],"evidence_links":[{"asset_id":"img-fixture","section":"baseline","region_ids":["img-fixture-r1"]}]},[analysis])
    st.session_state.fixture_result={"signature":case_input_signature("Synthetic history",[analysis]),"case":case}
render_result(st.session_state.fixture_result,st.session_state.fixture_assets,st.session_state.fixture_analyses)
'''


@unittest.skipIf(workflow is None, "Install Streamlit and OpenAI to run workflow tests")
class SessionStateTests(unittest.TestCase):
    def setUp(self):
        self.state = State()
        self.stub_st = Mock(session_state=self.state)
        self.stub_st.progress.return_value = Mock()
        self.patch = patch.object(workflow, "st", self.stub_st)
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def test_same_filename_changed_bytes_invalidates_analyses_case_and_export(self):
        workflow.sync_uploads([uploaded_image()])
        old_id = self.state.assets[0]["id"]
        self.state.analyses[old_id] = evidence(old_id)
        self.state.case_result = {"old": True}
        self.state.export_result = {"old": True}
        self.state.crop_edits["old-region"] = {"box": [0, 0, 100, 100]}
        self.state.review_old = "edited value"
        workflow.sync_uploads([uploaded_image(color="black")])
        self.assertNotEqual(self.state.assets[0]["id"], old_id)
        self.assertEqual(self.state.analyses, {})
        self.assertEqual(self.state.crop_edits, {})
        self.assertNotIn("case_result", self.state)
        self.assertNotIn("export_result", self.state)
        self.assertNotIn("review_old", self.state)

    def test_identical_upload_rerun_preserves_completed_work(self):
        files = [uploaded_image()]
        workflow.sync_uploads(files)
        aid = self.state.assets[0]["id"]
        self.state.analyses[aid] = evidence(aid)
        self.state.case_result = {"reviewed": True}
        self.state.export_result = {"ppt": b"cached"}
        workflow.sync_uploads(files)
        self.assertEqual(self.state.case_result, {"reviewed": True})
        self.assertEqual(self.state.export_result["ppt"], b"cached")
        self.assertEqual(list(self.state.analyses), [aid])

    def test_removing_uploads_clears_old_case_and_images(self):
        workflow.sync_uploads([uploaded_image()])
        self.state.case_result = {"old": True}
        workflow.sync_uploads([])
        self.assertEqual(self.state.assets, [])
        self.assertEqual(self.state.analyses, {})
        self.assertNotIn("case_result", self.state)

    def test_reread_failure_preserves_previous_review_and_export(self):
        workflow.sync_uploads([uploaded_image()])
        asset = self.state.assets[0]
        previous = evidence(asset["id"])
        self.state.analyses[asset["id"]] = previous
        self.state.case_result = {"reviewed": True}
        client = Mock()
        with patch.object(workflow, "make_client", return_value=client), patch.object(
            workflow, "analyze_asset", return_value=evidence(asset["id"])
        ), patch.object(workflow, "refine_analysis", side_effect=TimeoutError("simulated")):
            with self.assertRaises(TimeoutError):
                workflow.reread_asset(asset, "mock-key")
        self.assertEqual(self.state.analyses[asset["id"]], previous)
        self.assertEqual(self.state.case_result, {"reviewed": True})
        client.close.assert_called_once()

    def test_reread_success_updates_only_one_image_and_invalidates_case(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("second.png", "black")])
        first, second = self.state.assets
        self.state.analyses = {a["id"]: evidence(a["id"]) for a in self.state.assets}
        self.state.case_result = {"old": True}
        self.state.export_result = {"old": True}
        revised = evidence(first["id"])
        revised["title"] = "Rechecked report"
        with patch.object(workflow, "make_client", return_value=Mock()), patch.object(
            workflow, "analyze_asset", return_value=revised
        ):
            workflow.reread_asset(first, "mock-key", refine=False)
        self.assertEqual(self.state.analyses[first["id"]]["title"], "Rechecked report")
        self.assertEqual(self.state.analyses[second["id"]], evidence(second["id"]))
        self.assertNotIn("case_result", self.state)
        self.assertNotIn("export_result", self.state)
        self.assertEqual(self.state.analysis_versions[first["id"]], 1)

    def test_one_image_failure_keeps_success_and_retry_only_calls_failed_item(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("second.png", "black")])
        first_id, second_id = [asset["id"] for asset in self.state.assets]
        client = Mock()
        with patch.object(workflow, "make_client", return_value=client), patch.object(
            workflow, "analyze_asset", side_effect=[evidence(first_id), TimeoutError("simulated")]
        ) as analyze:
            workflow.analyze_pending("mock-key", "deepseek-flash", False)
        self.assertEqual(analyze.call_count, 2)
        self.assertIn(first_id, self.state.analyses)
        self.assertIn(second_id, self.state.asset_errors)
        first_version = self.state.analysis_versions[first_id]
        with patch.object(workflow, "make_client", return_value=client), patch.object(
            workflow, "analyze_asset", return_value=evidence(second_id)
        ) as retry:
            workflow.analyze_pending("mock-key", "deepseek-flash", False)
        retry.assert_called_once()
        self.assertEqual(retry.call_args.args[1]["id"], second_id)
        self.assertEqual(self.state.analysis_versions[first_id], first_version)
        self.assertEqual(self.state.asset_errors, {})
        self.assertEqual(client.close.call_count, 2)

    def test_refinement_failure_retains_first_pass_and_does_not_rescan(self):
        workflow.sync_uploads([uploaded_image()])
        aid = self.state.assets[0]["id"]
        first_pass = evidence(aid)
        with patch.object(workflow, "make_client", return_value=Mock()), patch.object(
            workflow, "analyze_asset", return_value=first_pass
        ), patch.object(workflow, "refine_analysis", side_effect=TimeoutError("simulated")):
            workflow.analyze_pending("mock-key", "deepseek-flash", True)
        self.assertEqual(self.state.analyses[aid], first_pass)
        self.assertIn(aid, self.state.refine_errors)
        with patch.object(workflow, "make_client", return_value=Mock()), patch.object(
            workflow, "analyze_asset"
        ) as analyze, patch.object(workflow, "refine_analysis", return_value=first_pass):
            workflow.analyze_pending("mock-key", "deepseek-flash", True)
        analyze.assert_not_called()
        self.assertEqual(self.state.refine_errors, {})

    def test_excluded_image_is_not_sent_to_recognition(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("excluded.png", "black")])
        first_id, excluded_id = [asset["id"] for asset in self.state.assets]
        self.state[workflow.include_key(excluded_id)] = False
        with patch.object(workflow, "make_client", return_value=Mock()), patch.object(
            workflow, "analyze_asset", return_value=evidence(first_id)
        ) as analyze:
            workflow.analyze_pending("mock-key", "deepseek-flash", False)
        analyze.assert_called_once()
        self.assertEqual(analyze.call_args.args[1]["id"], first_id)
        self.assertNotIn(excluded_id, self.state.analyses)
        self.assertFalse(self.state[workflow.include_key(excluded_id)])

    def test_clinical_edits_invalidate_case_signature_but_crops_do_not(self):
        original = evidence("img-001")
        signature = workflow.case_input_signature("history", [original])
        crop_edit = copy.deepcopy(original)
        crop_edit["regions"][0]["box"] = [50, 50, 500, 500]
        self.assertEqual(signature, workflow.case_input_signature("history", [crop_edit]))
        clinical_edit = copy.deepcopy(original)
        clinical_edit["findings"] = ["Corrected original finding"]
        self.assertNotEqual(signature, workflow.case_input_signature("history", [clinical_edit]))
        self.assertNotEqual(signature, workflow.case_input_signature("new history", [original]))
        self.assertNotEqual(signature, workflow.case_input_signature("history", []))

    def test_review_import_applies_to_exact_source_and_invalidates_export(self):
        workflow.sync_uploads([uploaded_image()])
        aid = self.state.assets[0]["id"]
        revised = evidence(aid)
        revised["findings"] = ["Manually reviewed source finding"]
        revised["regions"][0]["box"] = [100, 50, 450, 800]
        self.state.case_result = {"old": True}
        self.state.export_result = {"old": True}
        self.assertEqual(workflow.import_reviewed_analyses(json.dumps({"analyses": [revised]}).encode()), 1)
        self.assertEqual(self.state.analyses[aid], revised)
        self.assertEqual(self.state.analysis_versions[aid], 1)
        self.assertNotIn("case_result", self.state)
        self.assertNotIn("export_result", self.state)

    def test_review_import_rejects_cross_case_or_invalid_crop_atomically(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("second.png", "black")])
        first, second = [a["id"] for a in self.state.assets]
        original = evidence(first)
        self.state.analyses[first] = original
        revised = evidence(first)
        revised["title"] = "Changed"
        invalid = evidence(second)
        invalid["regions"][0]["box"] = [500, 0, 100, 1000]
        for bad in [evidence("img-unrelated"), invalid, revised]:
            with self.assertRaises(ValueError):
                workflow.import_reviewed_analyses(json.dumps({"analyses": [revised, bad]}).encode())
            self.assertEqual(self.state.analyses[first], original)
            self.assertEqual(self.state.analysis_versions, {})

    def test_review_import_restores_case_in_uploaded_source_order(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("second.png", "black")])
        first, second = [a["id"] for a in self.state.assets]
        rows = [evidence(second), evidence(first)]
        payload = {"analyses": rows, "case": {"baseline": {"diagnosis": "Reviewed diagnosis"}}}
        workflow.import_reviewed_analyses(json.dumps(payload).encode(), "Reviewed history")
        self.assertEqual(self.state.case_result["case"]["baseline"]["diagnosis"], "Reviewed diagnosis")
        self.assertEqual(self.state.case_result["signature"],
                         workflow.case_input_signature("Reviewed history", list(reversed(rows))))

    def test_review_import_rejects_incomplete_case_before_mutating_evidence(self):
        workflow.sync_uploads([uploaded_image(), uploaded_image("second.png", "black")])
        aid = self.state.assets[0]["id"]
        self.state.case_result = {"old": True}
        with self.assertRaises(ValueError):
            workflow.import_reviewed_analyses(json.dumps({"analyses": [evidence(aid)],
                                                          "case": {"baseline": {}}}).encode())
        self.assertEqual(self.state.case_result, {"old": True})
        self.assertEqual(self.state.analyses, {})


@unittest.skipIf(AppTest is None, "Install Streamlit and OpenAI to run UI tests")
class StreamlitWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.old_tempdir = tempfile.tempdir
        test_temp_root = Path(__file__).resolve().parents[1] / "test-output" / "app-tests"
        test_temp_root.mkdir(parents=True, exist_ok=True)
        tempfile.tempdir = str(test_temp_root)

    @classmethod
    def tearDownClass(cls):
        tempfile.tempdir = cls.old_tempdir

    def setUp(self):
        self.original_functions = {name: getattr(workflow, name) for name in ("make_client", "extract_case", "setting")}

    def tearDown(self):
        for name, function in self.original_functions.items():
            setattr(workflow, name, function)

    def make_app(self, code):
        app = AppTest.from_string(code, default_timeout=15).run()
        self.assertEqual([error.message for error in app.exception], [])
        return app

    def assert_healthy(self, app):
        self.assertEqual([error.message for error in app.exception], [])

    def test_empty_region_fallback_download_can_be_reimported_without_changing_existing_ids(self):
        existing = [{"id": f"img-001-r{i}", "label": f"Region {i}", "box": [0, 0, 1000, 1000]}
                    for i in range(1, 7)]
        for regions in ([], existing):
            with self.subTest(region_count=len(regions)):
                code = REVIEW_APP.replace(
                    "    st.session_state.crop_edits={}",
                    "    st.session_state.crop_edits={}\n"
                    f"    st.session_state.analyses['img-001']['regions'] = {regions!r}",
                ) + '''
import json
from app_workflow import import_reviewed_analyses
download = json.dumps({"analyses": analyses}, ensure_ascii=False, indent=2)
st.session_state.review_download = download
st.download_button("Download review", download, "review.json", "application/json")
if st.button("Restore review"):
    st.session_state.restored_count = import_reviewed_analyses(download.encode("utf-8"))
'''
                app = self.make_app(code)
                expected_ids = [row["id"] for row in regions] or ["img-001-r1"]
                initial = app.session_state.review_result["analyses"][0]["regions"]
                self.assertEqual([row["id"] for row in initial], expected_ids)
                self.assertEqual(initial[0]["box"], [0, 0, 1000, 1000])
                app.slider[0].set_range(100, 900).run()
                app.slider[1].set_range(200, 800).run()
                downloaded = json.loads(app.session_state.review_download)["analyses"][0]
                self.assertEqual(len(app.get("download_button")), 1)
                self.assertEqual(downloaded["regions"][0]["box"], [100, 200, 900, 800])
                next(button for button in app.button if button.label == "Restore review").click().run()
                self.assertEqual(app.session_state.restored_count, 1)
                self.assertEqual(app.session_state.analyses["img-001"], downloaded)
                app.run()
                self.assertEqual(app.session_state.review_result["analyses"][0], downloaded)
                self.assertEqual(app.session_state.review_result["pending"], [])
                self.assert_healthy(app)

    def test_crop_coordinates_survive_switch_away_rerun_and_return(self):
        app = self.make_app(REVIEW_APP)
        app.slider[0].set_range(100, 400).run()
        app.selectbox[1].select_index(1).run()
        app.run()
        self.assertEqual(app.session_state.review_result["analyses"][0]["regions"][0]["box"], [100, 0, 400, 1000])
        app.selectbox[1].select_index(0).run()
        self.assertEqual(app.slider[0].value, (100, 400))
        self.assert_healthy(app)

    def test_region_exclusion_survives_switch_and_rerun(self):
        app = self.make_app(REVIEW_APP)
        app.checkbox[1].uncheck().run()
        app.selectbox[1].select_index(1).run()
        app.run()
        self.assertEqual([r["id"] for r in app.session_state.review_result["analyses"][0]["regions"]], ["img-001-r2"])
        app.selectbox[1].select_index(0).run()
        self.assertFalse(app.checkbox[1].value)
        self.assert_healthy(app)

    def test_exact_crop_inputs_update_both_sliders_and_exported_region(self):
        app = self.make_app(REVIEW_APP)
        for widget, value in zip(app.number_input, (100, 450, 200, 800)):
            widget.set_value(value)
        next(button for button in app.button if button.label == "应用裁剪范围").click().run()
        self.assertEqual(app.slider[0].value, (100, 450))
        self.assertEqual(app.slider[1].value, (200, 800))
        self.assertEqual(app.session_state.review_result["analyses"][0]["regions"][0]["box"], [100, 200, 450, 800])
        app.selectbox[1].select_index(1).run()
        app.selectbox[1].select_index(0).run()
        self.assertEqual(app.slider[0].value, (100, 450))
        self.assert_healthy(app)

    def test_asset_exclusion_preserves_edits_for_reinclusion(self):
        app = self.make_app(REVIEW_APP)
        app.text_input[0].set_value("Reviewed title").run()
        app.checkbox[0].uncheck().run()
        self.assertEqual(app.session_state.review_result["analyses"], [])
        self.assertEqual(app.session_state.review_result["pending"], [])
        app.run()
        app.checkbox[0].check().run()
        self.assertEqual(app.session_state.review_result["analyses"][0]["title"], "Reviewed title")
        self.assert_healthy(app)

    def test_excluding_unrecognized_image_removes_generation_blocker(self):
        app = self.make_app(REVIEW_APP)
        app.session_state.analyses = {}
        app.run()
        self.assertEqual(app.session_state.review_result["pending"], ["img-001"])
        app.checkbox[0].uncheck().run()
        self.assertEqual(app.session_state.review_result["pending"], [])
        self.assertEqual(app.session_state.review_result["assets"], [])
        self.assert_healthy(app)

    def test_text_only_download_persists_and_stale_result_is_hidden(self):
        app = self.make_app(TEXT_APP)
        app.text_area(key="patient_text").set_value("Synthetic complete case history without any uploaded images.").run()
        generate = next(button for button in app.button if "图文 PPT" in button.label)
        self.assertFalse(generate.disabled)
        generate.click().run()
        self.assertEqual(app.session_state.call_count, 1)
        self.assertEqual(len(app.get("download_button")), 2)
        exported = app.session_state.export_result["ppt"]
        app.run()
        self.assertEqual(app.session_state.call_count, 1)
        self.assertEqual(app.session_state.export_result["ppt"], exported)
        self.assertEqual(len(app.get("download_button")), 2)
        app.text_area(key="patient_text").set_value("Changed synthetic complete case history needs a new case.").run()
        self.assertEqual(len(app.get("download_button")), 0)
        self.assertEqual(app.session_state.call_count, 1)
        self.assert_healthy(app)

    def test_image_placement_and_crop_rebuild_actual_ppt_media(self):
        def embedded_sizes(ppt_bytes):
            sizes = []
            with ZipFile(io.BytesIO(ppt_bytes)) as package:
                for name in package.namelist():
                    if name.startswith("ppt/media/"):
                        with Image.open(io.BytesIO(package.read(name))) as image:
                            sizes.append(image.size)
            return sizes

        app = self.make_app(RESULT_APP)
        initial = copy.deepcopy(app.session_state.export_result)
        self.assertEqual(len(app.get("download_button")), 2)
        self.assertEqual(initial["manifest"][0]["asset_id"], "img-fixture")
        self.assertEqual(initial["manifest"][0]["section"], "baseline")
        self.assertEqual(embedded_sizes(initial["ppt"]), [(100, 80)])

        app.selectbox[0].select("treatment:0").run()
        moved = copy.deepcopy(app.session_state.export_result)
        self.assertNotEqual(moved["key"], initial["key"])
        self.assertEqual(moved["manifest"][0]["section"], "treatment")
        self.assertEqual(moved["manifest"][0]["treatment_index"], 0)
        self.assertGreater(moved["manifest"][0]["slide_number"], initial["manifest"][0]["slide_number"])
        self.assertEqual(len(app.get("download_button")), 2)

        changed = copy.deepcopy(app.session_state.fixture_analyses)
        changed[0]["regions"][0]["box"] = [250, 0, 750, 500]
        app.session_state.fixture_analyses = changed
        app.run()
        cropped = app.session_state.export_result
        self.assertNotEqual(cropped["key"], moved["key"])
        self.assertNotEqual(cropped["ppt"], moved["ppt"])
        self.assertEqual(embedded_sizes(cropped["ppt"]), [(50, 40)])
        self.assertEqual(cropped["manifest"][0]["section"], "treatment")
        self.assertEqual(cropped["manifest"][0]["region_ids"], ["img-fixture-r1"])
        self.assertEqual(len(app.get("download_button")), 2)
        self.assert_healthy(app)

    def test_real_entrypoint_opens_without_mandatory_secrets(self):
        app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"), default_timeout=15).run()
        self.assert_healthy(app)
        self.assertTrue(any(area.key == "patient_text" for area in app.text_area))

    def test_layout_version_change_invalidates_cached_download(self):
        app = self.make_app(RESULT_APP)
        previous_key = app.session_state.export_result["key"]
        with patch.object(workflow, "PPT_LAYOUT_VERSION", "test-layout-update"):
            app.run()
        self.assertNotEqual(app.session_state.export_result["key"], previous_key)
        self.assert_healthy(app)


if __name__ == "__main__":
    unittest.main()
