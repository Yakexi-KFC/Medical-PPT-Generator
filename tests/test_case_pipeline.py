"""Case schema and source mapping regression tests; no network or credentials."""

import copy
import json
import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from case_pipeline import CASE_PROMPT, extract_case, normalize_case


def analysis(asset_id="img-001", title="病理检查"):
    return {
        "asset_id": asset_id,
        "title": title,
        "date": "2025-03-01",
        "category": "pathology",
        "findings": ["Ki-67 15%"],
        "transcription": "Ki-67：15%；CK（+）",
        "warnings": ["仅供界面显示的识别问题"],
        "regions": [{"id": f"{asset_id}-r1", "label": "检查结果", "box": [0, 0, 1000, 1000]}],
    }


def response(payload, finish_reason="stop"):
    content = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)
    return SimpleNamespace(choices=[SimpleNamespace(
        finish_reason=finish_reason,
        message=SimpleNamespace(content=content),
    )])


class CaseNormalizationTests(unittest.TestCase):
    def test_rejects_non_object_or_missing_main_sections(self):
        for payload in (None, [], "{}", 1, {}, {"cover": {"title": "病例"}}):
            with self.subTest(payload=payload), self.assertRaises(ValueError):
                normalize_case(payload)

    def test_null_and_malformed_fields_produce_consistent_schema(self):
        normalized = normalize_case({
            "cover": None,
            "baseline": {"patient_info": ["不得转成字符串"], "chief_complaint": 7, "diagnosis": "  原文诊断  ", "key_exams": None},
            "treatments": None,
            "current_admission": [],
            "timeline_events": "not-an-array",
            "summary": None,
            "evidence_links": False,
        })
        self.assertEqual(normalized["cover"], {"title": "病例汇报"})
        self.assertEqual(normalized["baseline"], {
            "patient_info": "", "chief_complaint": "", "diagnosis": "原文诊断", "key_exams": "",
        })
        self.assertEqual(normalized["treatments"], [])
        self.assertEqual(normalized["current_admission"], {"exams": [], "imaging": "", "plan": []})
        self.assertEqual(normalized["summary"], {"highlights": [], "discussion": []})
        self.assertEqual(normalized["timeline_events"], [])
        self.assertEqual(normalized["evidence_links"], [])

    def test_scalar_and_list_text_do_not_invent_clinical_content(self):
        result = normalize_case({
            "baseline": {},
            "current_admission": {"exams": "  Hb 90 g/L  ", "plan": [None, "", "  已记录的方案  ", {"bad": "value"}, 12]},
            "summary": {"highlights": ["  原文结论  ", False], "discussion": 42},
        })
        self.assertEqual(result["current_admission"]["exams"], ["Hb 90 g/L"])
        self.assertEqual(result["current_admission"]["plan"], ["已记录的方案"])
        self.assertEqual(result["summary"], {"highlights": ["原文结论"], "discussion": []})

    def test_treatment_and_event_arrays_ignore_wrong_item_types(self):
        result = normalize_case({
            "treatments": [None, "一线治疗", {"phase": "  辅助治疗  ", "regimen": "原文方案", "duration": None, "imaging": 5}],
            "timeline_events": [None, [], {"date": "  2025-03  ", "phase": False, "event_type": "Treatment", "event": "原文事件"}],
        })
        self.assertEqual(result["treatments"], [{"phase": "辅助治疗", "regimen": "原文方案", "duration": "", "imaging": "", "markers": ""}])
        self.assertEqual(result["timeline_events"], [{"date": "2025-03", "phase": "", "event_type": "Treatment", "event": "原文事件"}])

    def test_does_not_mutate_case_or_evidence(self):
        evidence = [analysis()]
        payload = {"baseline": {}, "treatments": [{"phase": "辅助治疗"}], "evidence_links": [{"asset_id": "img-001", "section": "baseline"}]}
        before_payload, before_evidence = copy.deepcopy(payload), copy.deepcopy(evidence)
        result = normalize_case(payload, evidence)
        result["treatments"][0]["phase"] = "修改展示副本"
        self.assertEqual(payload, before_payload)
        self.assertEqual(evidence, before_evidence)

    def test_invalid_and_duplicate_asset_refs_cannot_create_extra_images(self):
        result = normalize_case({"baseline": {}, "evidence_links": [
            None, {"asset_id": []}, {"asset_id": {"id": "img-001"}},
            {"asset_id": "not-uploaded", "section": "baseline"},
            {"asset_id": "img-001", "section": "baseline", "caption": "第一次引用"},
            {"asset_id": "img-001", "section": "current_admission", "caption": "重复引用"},
        ]}, [analysis()])
        self.assertEqual(len(result["evidence_links"]), 1)
        self.assertEqual(result["evidence_links"][0]["caption"], "第一次引用")
        self.assertEqual(result["evidence_links"][0]["section"], "baseline")

    def test_region_refs_are_existing_local_ids_and_unique(self):
        result = normalize_case({"baseline": {}, "evidence_links": [{
            "asset_id": "img-001", "section": "baseline",
            "region_ids": ["invented-region", "img-001-r1", None, "img-001-r1", 9],
        }]}, [analysis()])
        self.assertEqual(result["evidence_links"][0]["region_ids"], ["img-001-r1"])

    def test_missing_links_preserve_every_included_asset_in_appendix(self):
        result = normalize_case({"baseline": {}, "evidence_links": []}, [analysis(), analysis("img-002", "基因检查")])
        self.assertEqual([link["asset_id"] for link in result["evidence_links"]], ["img-001", "img-002"])
        self.assertTrue(all(link["section"] == "appendix" and link["treatment_index"] is None for link in result["evidence_links"]))
        self.assertEqual(result["evidence_links"][1]["caption"], "基因检查")

    def test_unknown_or_unhashable_section_falls_back_to_appendix(self):
        for section in (None, [], {}, ["baseline"], 9, True, "unknown"):
            with self.subTest(section=section):
                result = normalize_case({"baseline": {}, "evidence_links": [{
                    "asset_id": "img-001", "section": section, "treatment_index": 0,
                }]}, [analysis()])
                self.assertEqual(result["evidence_links"][0]["section"], "appendix")
                self.assertIsNone(result["evidence_links"][0]["treatment_index"])

    def test_treatment_reference_remaps_indices_after_bad_items_removed(self):
        result = normalize_case({
            "treatments": [None, {"phase": "已记录的一线"}, "bad", {"phase": "已记录的二线"}],
            "evidence_links": [{"asset_id": "img-001", "section": "treatment", "treatment_index": 3}],
        }, [analysis()])
        self.assertEqual(result["evidence_links"][0]["treatment_index"], 1)
        self.assertEqual(result["treatments"][1]["phase"], "已记录的二线")

    def test_invalid_treatment_indices_do_not_silently_choose_first_phase(self):
        for index in (None, True, False, -1, 2, 0.0, "0", [], {}):
            with self.subTest(index=index):
                result = normalize_case({
                    "treatments": [{"phase": "第一阶段"}],
                    "evidence_links": [{"asset_id": "img-001", "section": "treatment", "treatment_index": index}],
                }, [analysis()])
                self.assertEqual(result["evidence_links"][0]["section"], "appendix")
                self.assertIsNone(result["evidence_links"][0]["treatment_index"])

    def test_reference_to_removed_treatment_goes_to_appendix(self):
        result = normalize_case({
            "treatments": [None, {"phase": "唯一有效阶段"}],
            "evidence_links": [{"asset_id": "img-001", "section": "treatment", "treatment_index": 0}],
        }, [analysis()])
        self.assertEqual(result["evidence_links"][0]["section"], "appendix")

    def test_non_treatment_links_clear_irrelevant_treatment_indices(self):
        result = normalize_case({"baseline": {}, "evidence_links": [{
            "asset_id": "img-001", "section": "current_admission", "treatment_index": {"invalid": "unused"},
        }]}, [analysis()])
        self.assertIsNone(result["evidence_links"][0]["treatment_index"])


class ExtractCaseTests(unittest.TestCase):
    def test_mock_client_receives_json_case_and_evidence_without_ui_warnings(self):
        client = Mock()
        client.chat.completions.create.return_value = response({"baseline": {"diagnosis": "原文诊断"}})
        evidence = [analysis()]
        before = copy.deepcopy(evidence)
        result = extract_case(client, "给定病史", evidence, model="deepseek-flash")
        client.chat.completions.create.assert_called_once()
        params = client.chat.completions.create.call_args.kwargs
        self.assertEqual(params["model"], "deepseek-flash")
        self.assertEqual(params["response_format"], {"type": "json_object"})
        self.assertEqual(params["messages"][0], {"role": "system", "content": CASE_PROMPT})
        supplied = json.loads(params["messages"][1]["content"])
        self.assertEqual(supplied["病史"], "给定病史")
        self.assertNotIn("warnings", supplied["附件检查"][0])
        self.assertEqual(supplied["附件检查"][0]["asset_id"], "img-001")
        self.assertEqual(supplied["附件检查"][0]["regions"], evidence[0]["regions"])
        self.assertEqual(evidence, before)
        self.assertEqual(result["baseline"]["diagnosis"], "原文诊断")
        self.assertEqual(result["evidence_links"][0]["section"], "appendix")

    def test_mock_client_handles_text_only_case(self):
        client = Mock()
        client.chat.completions.create.return_value = response({"treatments": []})
        result = extract_case(client, "没有附件的完整病史")
        self.assertEqual(result["evidence_links"], [])
        self.assertEqual(json.loads(client.chat.completions.create.call_args.kwargs["messages"][1]["content"])["附件检查"], [])

    def test_empty_partial_or_wrong_schema_response_never_becomes_case(self):
        for output in (response(""), response('{"baseline": {}', "length"), response("[]"), response({"wrong": "schema"})):
            with self.subTest(output=output):
                client = Mock()
                client.chat.completions.create.return_value = output
                with self.assertRaises(ValueError):
                    extract_case(client, "给定病史", [analysis()])

    def test_network_failure_propagates_without_substituting_a_fake_case(self):
        client = Mock()
        client.chat.completions.create.side_effect = TimeoutError("mock timeout")
        with self.assertRaises(TimeoutError):
            extract_case(client, "给定病史")

    def test_api_plan_gets_grounded_timeline_before_first_treatment_but_import_does_not(self):
        payload = {"baseline": {}, "treatments": [{"phase": "化疗"}], "presentation_plan": {"slides": [
            {"type": "cover", "title": "病例汇报"},
            {"type": "content", "title": "患者资料", "section": "baseline"},
            {"type": "content", "title": "治疗经过", "section": "treatment", "treatment_index": 0}]}}
        source = "2025-01-10予方案A治疗。2025-03-10建议方案B治疗。"
        client = Mock()
        client.chat.completions.create.return_value = response(payload)
        generated = extract_case(client, source)
        slides = generated["presentation_plan"]["slides"]
        self.assertEqual([slide["type"] for slide in slides], ["cover", "content", "timeline", "content"])
        self.assertEqual(slides[2]["events"], generated["timeline_events"])
        self.assertEqual(slides[2]["events"][-1]["event_type"], "Treatment Plan")
        self.assertFalse(any(slide["type"] == "timeline" for slide in
                             normalize_case(payload, source_text=source)["presentation_plan"]["slides"]))
        client.chat.completions.create.return_value = response(generated)
        self.assertEqual(sum(slide["type"] == "timeline" for slide in
                             extract_case(client, source)["presentation_plan"]["slides"]), 1)

    def test_api_does_not_drop_a_full_plan_or_add_an_unsupported_timeline(self):
        full = {"baseline": {}, "presentation_plan": {"slides": [{"type": "content", "title": f"正文{i}"} for i in range(30)]}}
        client = Mock()
        client.chat.completions.create.return_value = response(full)
        with self.assertRaisesRegex(ValueError, "已有30页"):
            extract_case(client, "2025-01-10予方案A治疗。")
        undated = extract_case(client, "未记录诊疗日期。")
        self.assertEqual(len(undated["presentation_plan"]["slides"]), 30)


class DatedPlacementTests(unittest.TestCase):
    @staticmethod
    def case(durations, section="appendix", index=None):
        return {"baseline": {}, "treatments": [{"duration": value} for value in durations],
                "evidence_links": [{"asset_id": "img-001", "section": section, "treatment_index": index}]}

    def locate(self, dates, exam_date, section="appendix", index=None, category="lab", transcript=""):
        item = analysis()
        item.update(date=exam_date, category=category, transcription=transcript)
        return normalize_case(self.case(dates, section, index), [item])["evidence_links"][0]

    def test_post_treatment_genetic_report_is_not_baseline(self):
        link = self.locate(["2022-01-10", "2022-02-20及2022-03-11"], "2022-02-28", "baseline", category="genetics")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 1))

    def test_lab_in_enumerated_phase_understands_inherited_years(self):
        link = self.locate(["2022-01-06、01-27、02-17、03-11及04-02", "2022-04-23"], "2022-03-31")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 0))

    def test_short_dates_handle_december_to_january_rollover(self):
        link = self.locate(["2022-12-16、01-06及01-27", "2023-02-10"], "2023-01-12")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 0))

    def test_dated_followup_between_phases_goes_after_preceding_phase(self):
        link = self.locate(["2022-01-10", "2022-02-20"], "2022-02-01")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 0))

    def test_report_before_first_phase_is_baseline(self):
        link = self.locate(["2022-02-20"], "2022-02-01")
        self.assertEqual((link["section"], link["treatment_index"]), ("baseline", None))

    def test_year_only_unknown_and_invalid_dates_do_not_guess(self):
        for exam_date in ("", "2022", "日期不详", "2022-02-30"):
            with self.subTest(exam_date=exam_date):
                self.assertEqual(self.locate(["2022-01-10", "2022-04-23"], exam_date)["section"], "appendix")

    def test_month_precision_spanning_phase_boundary_remains_unassigned(self):
        link = self.locate(["2022-01-10", "2022-03-15", "2022-04-20"], "2022-03")
        self.assertEqual(link["section"], "appendix")

    def test_history_screenshot_date_does_not_relocate_a_multiphase_summary(self):
        link = self.locate(["2022-01-10", "2022-04-23"], "2022-02-01", category="record")
        self.assertEqual(link["section"], "appendix")

    def test_complete_final_phase_is_not_extended_without_a_known_next_phase(self):
        self.assertEqual(self.locate(["2022-01-10至2022-02-20"], "2022-12-01")["section"], "appendix")

    def test_open_month_start_is_bounded_by_next_phase(self):
        link = self.locate(["2022年6月起", "2023年3月开始"], "2022-11-01")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 0))

    def test_overlapping_phases_do_not_arbitrarily_choose_a_new_phase(self):
        durations = ["2022-01-10至2022-03-20", "2022-02-10至2022-04-20"]
        ambiguous = self.locate(durations, "2022-03-01", "baseline")
        self.assertEqual(ambiguous["section"], "appendix")
        existing = self.locate(durations, "2022-03-01", "treatment", 1)
        self.assertEqual((existing["section"], existing["treatment_index"]), ("treatment", 1))

    def test_explicit_latest_current_assessment_is_preserved(self):
        link = self.locate(["2022年6月起"], "2022-11-01", "current_admission")
        self.assertEqual(link["section"], "current_admission")

    def test_valid_concurrent_local_treatment_association_is_not_overridden(self):
        link = self.locate(["2022-02-20及2022-03-11", "2022-03-09", "2022-04-01"], "2022-03-10", "treatment", 1)
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 1))

    def test_labelled_report_date_fills_missing_metadata_for_placement(self):
        link = self.locate(["2022-01-10", "2022-04-23"], "", transcript="检验日期：2022-03-31\n检查结果：原文")
        self.assertEqual((link["section"], link["treatment_index"]), ("treatment", 0))

    def test_conflicting_labelled_dates_are_not_silently_resolved(self):
        link = self.locate(["2022-01-10", "2022-04-23"], "", transcript="采样日期：2022-03-01\n报告日期：2022-03-31")
        self.assertEqual(link["section"], "appendix")

    def test_generator_evidence_retains_links_for_reconciliation(self):
        result = normalize_case({"baseline": {}}, (analysis(f"img-{i}") for i in range(2)))
        self.assertEqual(len(result["evidence_links"]), 2)


class ReportGroupingTests(unittest.TestCase):
    REPORT = ("病理诊断：(肺组织)上皮源性肿瘤，考虑为伴有神经内分泌分化的非小细胞肺癌。"
              "免疫组化结果：CK(+)、P40(-)、P63(-)、NapsinA(-)、TTF1(-)、CK7(+)、"
              "CK5/6(-)、CD56(-)、CgA(+)、Ki67(约10%+)、SYN(+)、PD-L1(E1L3N)(-)、"
              "PD-L1(Neg)(-)、CD34(-)、ALK-P(D5F3)(-)、ALK-P(Neg)(-)、INSM1(+)。")

    def items(self):
        first, second = analysis("whole"), analysis("detail")
        first.update(date="2022-02-01", transcription=self.REPORT)
        second.update(date="", transcription=self.REPORT.replace("(", "（").replace(")", "）"))
        return [first, second]

    def grouped(self, items):
        return normalize_case({"baseline": {}, "treatments": [{"duration": "2022-01-10"}, {"duration": "2022-03-20"}]}, items)

    def test_whole_and_detail_are_grouped_without_removing_either_source(self):
        items = self.items()
        before = copy.deepcopy(items)
        result = self.grouped(items)
        self.assertEqual(len(result["evidence_links"]), 2)
        self.assertEqual(len(result["evidence_groups"]), 1)
        self.assertEqual(result["evidence_groups"][0]["asset_ids"], ["whole", "detail"])
        self.assertEqual(result["evidence_groups"][0]["reason"], "matching_report_content")
        self.assertEqual(result["evidence_groups"][0]["section"], "treatment")
        self.assertEqual(result["evidence_groups"][0]["treatment_index"], 0)
        self.assertEqual(result["evidence_links"][1]["placement_basis"], "matching_dated_report")
        self.assertNotIn("excluded_asset_ids", result)
        self.assertEqual(items, before)

    def test_different_dates_never_merge_repeated_similar_reports(self):
        items = self.items()
        items[1]["date"] = "2022-02-02"
        self.assertEqual(self.grouped(items)["evidence_groups"], [])

    def test_different_report_identifiers_are_independent(self):
        items = self.items()
        items[0]["transcription"] += "报告编号：SYNTH-A123"
        items[1]["transcription"] += "报告编号：SYNTH-B456"
        self.assertEqual(self.grouped(items)["evidence_groups"], [])

    def test_opposite_sign_or_different_value_prevents_grouping(self):
        for replacement in ("CK(-)", "CK(弱+)"):
            items = self.items()
            items[1]["transcription"] = self.REPORT.replace("CK(+)", replacement)
            self.assertEqual(self.grouped(items)["evidence_groups"], [])
        items = self.items()
        items[1]["transcription"] = self.REPORT.replace("10%", "20%")
        self.assertEqual(self.grouped(items)["evidence_groups"], [])

    def test_short_template_or_title_similarity_is_insufficient(self):
        items = self.items()
        for item in items:
            item["transcription"] = "病理诊断：考虑恶性肿瘤，CK(+)。"
        self.assertEqual(self.grouped(items)["evidence_groups"], [])

    def test_repeated_history_is_not_a_duplicate_report(self):
        items = self.items()
        for item in items:
            item["category"] = "record"
        self.assertEqual(self.grouped(items)["evidence_groups"], [])

    def test_pairwise_check_prevents_undated_detail_bridging_two_dates(self):
        items = self.items()
        third = copy.deepcopy(items[0])
        third["asset_id"] = "independent"
        third["date"] = "2022-02-02"
        result = self.grouped([items[0], items[1], third])
        self.assertEqual(len(result["evidence_links"]), 3)
        self.assertEqual(result["evidence_groups"][0]["asset_ids"], ["whole", "detail"])


class GroundedTreatmentLabelTests(unittest.TestCase):
    def test_unsupported_lines_are_removed_without_changing_regimen_or_adjuvant_phase(self):
        payload = {"baseline": {}, "treatments": [{"phase": "一线化疗", "regimen": "记录中的方案A"},
                    {"phase": "第二线治疗", "regimen": "记录中的方案B"}, {"phase": "术后辅助"}],
                   "timeline_events": [{"phase": "一线", "event": "开始二线化疗"}],
                   "summary": {"highlights": ["二线化疗后复查"]}}
        result = normalize_case(payload, source_text="先予方案A，后改用方案B。术后辅助治疗。")
        self.assertEqual([item["phase"] for item in result["treatments"]], ["化疗", "阶段治疗", "术后辅助"])
        self.assertEqual(result["treatments"][0]["regimen"], "记录中的方案A")
        self.assertEqual(result["timeline_events"], [])  # No dated source statement supports this model event.
        self.assertEqual(result["summary"]["highlights"], ["化疗后复查"])

    def test_explicit_source_lines_are_retained_in_arabic_or_chinese_form(self):
        item = analysis()
        item["transcription"] = "原文：第2线治疗。"
        result = normalize_case({"treatments": [{"phase": "一线化疗"}, {"phase": "二线治疗"}]}, [item], source_text="原文明确记载第一线治疗。")
        self.assertEqual([tx["phase"] for tx in result["treatments"]], ["一线化疗", "二线治疗"])

    def test_model_findings_do_not_self_validate_invented_lines(self):
        item = analysis()
        item["findings"] = ["二线治疗"]
        item["transcription"] = "原文只记录治疗方案"
        result = normalize_case({"treatments": [{"phase": "二线化疗"}]}, [item], source_text="明确给出的病史")
        self.assertEqual(result["treatments"][0]["phase"], "化疗")

    def test_normalizing_a_previously_edited_case_without_sources_does_not_strip_labels(self):
        result = normalize_case({"treatments": [{"phase": "三线治疗"}]})
        self.assertEqual(result["treatments"][0]["phase"], "三线治疗")

    def test_extract_case_checks_model_phase_against_actual_patient_text(self):
        client = Mock()
        client.chat.completions.create.return_value = response({"treatments": [{"phase": "二线化疗"}]})
        self.assertEqual(extract_case(client, "先后使用两种化疗方案")["treatments"][0]["phase"], "化疗")


class GroundedTimelineTests(unittest.TestCase):
    def test_same_day_same_combination_plan_in_different_wording_is_not_repeated(self):
        item = analysis("record")
        item.update(category="record", transcription="2022-09-30诊疗意见：建议药物B、药物A联合用药。")
        events = normalize_case({"baseline": {}}, [item], source_text="2022-09-30门诊建议药物A联合药物B治疗。")["timeline_events"]
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["source_asset_id"], "patient_text")

    def test_different_plan_drug_dose_negation_or_qualifier_is_preserved(self):
        for second in ("药物A联合药物C治疗", "药物A 20mg联合药物B治疗", "不建议药物A联合药物B治疗", "考虑药物A联合药物B治疗"):
            with self.subTest(second=second):
                source = f"2022-09-30门诊建议药物A联合药物B治疗。\n2022-09-30建议{second}。"
                self.assertEqual(len(normalize_case({"baseline": {}}, source_text=source)["timeline_events"]), 2)

    def test_reviewed_pathology_date_does_not_become_a_biopsy_date(self):
        item = analysis("pathology")
        item.update(title="肺穿刺活检病理结果", date="2022-09-06")
        item["warnings"] = ["原图为送检日期，未明确穿刺操作日期"]
        events = normalize_case({"baseline": {}}, [item], source_text="")["timeline_events"]
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["event_type"], "Pathology")
        self.assertEqual(events[0]["event"], "肺穿刺活检病理结果")

    def test_admission_date_cannot_date_an_undated_biopsy(self):
        item = analysis("record")
        item.update(category="record", date="", transcription="2022-09-04胸部CT示病灶增大，行CT引导下肺穿刺活检。\n2022-09-05至2022-09-08 肿瘤内科 住院 肺恶性肿瘤。")
        payload = {"baseline": {}, "timeline_events": [{"date": "2022-09-05", "event_type": "Biopsy", "event": "CT引导下肺穿刺活检"}]}
        result = normalize_case(payload, [item], source_text="")
        self.assertEqual(len(result["timeline_events"]), 1)
        event = result["timeline_events"][0]
        self.assertEqual(event["date"], "2022-09-04")
        self.assertEqual(event["event_type"], "Imaging")
        self.assertNotIn("穿刺", event["event"])
        self.assertEqual(event["source_quote"], "2022-09-04胸部CT示病灶增大")

    def test_explicitly_dated_biopsy_is_retained(self):
        result = normalize_case({"baseline": {}}, source_text="2022-09-05行CT引导下肺穿刺活检。")
        self.assertEqual(result["timeline_events"][0]["date"], "2022-09-05")
        self.assertEqual(result["timeline_events"][0]["event_type"], "Biopsy")

    def test_all_recorded_treatment_dates_and_month_precision_are_preserved(self):
        payload = {"treatments": [{"phase": "治疗A", "duration": "2022-11-01、11-21、12-16"},
                                  {"phase": "治疗B", "duration": "2023年6月起"}]}
        source = "2022-11-01、11-21及12-16予方案A治疗。\n2023年6月起采用方案B化疗，共8次。"
        events = normalize_case(payload, source_text=source)["timeline_events"]
        self.assertEqual([event["date"] for event in events], ["2022-11-01", "2022-11-21", "2022-12-16", "2023-06"])
        self.assertEqual([event["phase"] for event in events], ["治疗A", "治疗A", "治疗A", "治疗B"])
        self.assertTrue(all(event["source_quote"] in source for event in events))

    def test_unmatched_model_events_are_not_supported_by_an_arbitrary_date(self):
        payload = {"baseline": {}, "timeline_events": [{"date": "2022-09-05", "event_type": "Surgery", "event": "肺切除术"}]}
        self.assertEqual(normalize_case(payload, source_text="2022-09-05入院。随后完成相关检查。")["timeline_events"], [])

    def test_same_day_independent_ct_and_mri_are_not_deduplicated(self):
        source = "2022-09-05胸部CT示肺部结节。\n2022-09-05头部MRI未见明显异常。"
        events = normalize_case({"baseline": {}}, source_text=source)["timeline_events"]
        self.assertEqual(len(events), 2)
        self.assertIn("CT", events[0]["event"])
        self.assertIn("MRI", events[1]["event"])

    def test_repeated_same_procedure_in_visit_record_is_not_repeated(self):
        item = analysis("record")
        item.update(category="record", transcription="2022-09-05开始胸部局部放疗，DT 60 Gy/20 F。")
        events = normalize_case({"baseline": {}}, [item], source_text="2022-09-05行胸部局部放疗，DT 60 Gy/20 F。")["timeline_events"]
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["source_asset_id"], "patient_text")

    def test_reviewed_lab_date_is_used_as_check_date_not_a_procedure(self):
        item = analysis("lab")
        item.update(category="lab", date="2022-03-31", title="肿瘤六项检验", transcription="NSE 11.20 ng/mL")
        events = normalize_case({"baseline": {}}, [item], source_text="")["timeline_events"]
        self.assertEqual(events[0]["date"], "2022-03-31")
        self.assertEqual(events[0]["event_type"], "Laboratory")
        self.assertEqual(events[0]["source_asset_id"], "lab")
        self.assertEqual(events[0]["date_basis"], "reviewed_examination_date")

    def test_offline_normalization_is_idempotent(self):
        payload = {"treatments": [{"phase": "化疗", "duration": "2022-01-10及01-31"}]}
        source = "2022-01-10及01-31予方案A治疗。"
        first = normalize_case(payload, source_text=source)
        self.assertEqual(normalize_case(first, source_text=source), first)


class PresentationPlanTests(unittest.TestCase):
    @staticmethod
    def payload(slide=None):
        return {"baseline": {}, "treatments": [{"phase": "同一治疗阶段"}],
                "presentation_plan": {"mode": "compact", "slides": [slide or {
                    "type": "content", "title": "临床评估", "section": "baseline",
                    "sections": [{"heading": "病理", "text": "Ki-67 5%\n会诊热点约20%"}],
                    "images": [{"asset_id": "img-001", "region_id": "img-001-r1", "label": "原始结果"}],
                    "source_ids": ["img-001"],
                }]}}

    def test_default_compact_and_explicit_detailed_modes(self):
        self.assertEqual(normalize_case({"baseline": {}})["presentation_mode"], "compact")
        self.assertEqual(normalize_case({"baseline": {}})["presentation_theme"], "standard_case")
        self.assertEqual(normalize_case({"baseline": {}, "presentation_theme": "legacy"})["presentation_theme"], "legacy")
        self.assertEqual(normalize_case({"baseline": {}, "presentation_mode": "detailed"})["presentation_mode"], "detailed")
        self.assertNotIn("presentation_plan", normalize_case({"baseline": {}}))
        for mode in (None, {}, [], False, "unknown"):
            with self.subTest(mode=mode), self.assertRaises(ValueError):
                normalize_case({"baseline": {}, "presentation_mode": mode})

    def test_valid_plan_survives_normalization_and_repeated_editing(self):
        payload = self.payload()
        payload["presentation_plan"]["slides"][0].update(subtitle="  已核对的用户编辑文本  ", layout="bottom")
        original = copy.deepcopy(payload)
        result = normalize_case(payload, [analysis()])
        slide = result["presentation_plan"]["slides"][0]
        self.assertEqual(slide["sections"], original["presentation_plan"]["slides"][0]["sections"])
        self.assertEqual(slide["subtitle"], "  已核对的用户编辑文本  ")
        self.assertEqual(slide["layout"], "bottom")
        self.assertEqual(payload, original)
        self.assertEqual(normalize_case(result, [analysis()]), result)

    def test_table_fields_preserved_as_editable_values(self):
        table = {"columns": ["日期", "结果"], "rows": [["2026-01", "CK(+)、CgA(+)"], ["2026-02", "原发灶64×39×47 mm"]],
                 "column_widths": [1, 3.5], "font_size": 18.5}
        slide = {"type": "table", "title": "影像对照", "table": table, "section": "treatment", "treatment_index": 0}
        result = normalize_case(self.payload(slide))
        self.assertEqual(result["presentation_plan"]["slides"][0]["table"], table)
        self.assertEqual(result["presentation_plan"]["slides"][0]["treatment_index"], 0)

    def test_invalid_plan_and_slide_count_are_rejected_not_truncated(self):
        for plan in (None, [], "plan", {"mode": "detailed", "slides": []}, {"slides": []}, {"slides": [None]},
                     {"slides": [{"type": "content", "title": "病例"}] * 31}):
            with self.subTest(plan=plan), self.assertRaises(ValueError):
                normalize_case({"baseline": {}, "presentation_plan": plan})
        result = normalize_case({"baseline": {}, "presentation_plan": {"slides": [{"type": "content", "title": "病例"}] * 30}})
        self.assertEqual(len(result["presentation_plan"]["slides"]), 30)

    def test_invalid_slide_type_title_section_and_layout_are_rejected(self):
        for key, value in (("type", "unknown"), ("type", []), ("title", ""), ("title", 7),
                           ("section", "unknown"), ("section", {}), ("layout", "left")):
            with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                payload = self.payload()
                payload["presentation_plan"]["slides"][0][key] = value
                normalize_case(payload, [analysis()])

    def test_invalid_region_cannot_become_full_original_image(self):
        for region_id in (None, "", [], "not-a-region", "img-002-r1"):
            with self.subTest(region_id=region_id), self.assertRaises(ValueError):
                payload = self.payload()
                payload["presentation_plan"]["slides"][0]["images"][0]["region_id"] = region_id
                normalize_case(payload, [analysis(), analysis("img-002")])

    def test_missing_optional_region_selects_existing_crop(self):
        payload = self.payload()
        picture = payload["presentation_plan"]["slides"][0]["images"][0]
        del picture["region_id"]
        result = normalize_case(payload, [analysis()])
        self.assertEqual(result["presentation_plan"]["slides"][0]["images"][0]["region_id"], "img-001-r1")
        no_crops = analysis()
        no_crops["regions"] = []
        with self.assertRaises(ValueError):
            normalize_case(payload, [no_crops])

    def test_full_picture_is_explicit_and_needs_no_crop_but_still_validates_source(self):
        payload = self.payload()
        payload["presentation_theme"] = "standard_case"
        slide = payload["presentation_plan"]["slides"][0]
        slide["layout"] = "image_left"
        slide["images"] = [{"asset_id": "img-001", "display_mode": "full", "label": "完整病理图像"}]
        record = analysis()
        record["regions"] = []
        result = normalize_case(payload, [record])
        self.assertEqual(result["presentation_plan"]["slides"][0]["images"], slide["images"])
        self.assertEqual(result["presentation_theme"], "standard_case")
        self.assertEqual(normalize_case(result, [record]), result)
        slide["images"][0]["asset_id"] = "other-patient"
        with self.assertRaises(ValueError):
            normalize_case(payload, [record])

    def test_ambiguous_full_and_invalid_picture_modes_fail_closed(self):
        for mode in ("unknown", None, [], True, "full"):
            payload = self.payload()
            payload["presentation_plan"]["slides"][0]["images"][0]["display_mode"] = mode
            with self.subTest(mode=mode), self.assertRaises(ValueError):
                normalize_case(payload, [analysis()])
        for theme in ("custom", None, [], True):
            with self.subTest(theme=theme), self.assertRaises(ValueError):
                normalize_case({"baseline": {}, "presentation_theme": theme})

    def test_explicit_region_mode_and_image_sidebar_preserve_selected_crop(self):
        for layout in ("image_left", "image_right"):
            payload = self.payload()
            slide = payload["presentation_plan"]["slides"][0]
            slide["layout"] = layout
            slide["images"][0]["display_mode"] = "region"
            result = normalize_case(payload, [analysis()])["presentation_plan"]["slides"][0]
            self.assertEqual(result["images"], slide["images"])
            self.assertEqual(result["layout"], layout)

    def test_report_pair_requires_two_ordered_valid_images(self):
        payload = self.payload()
        slide = payload["presentation_plan"]["slides"][0]
        slide["layout"] = "report_pair"
        slide["images"] = [{"asset_id": "img-001", "region_id": "img-001-r1", "label": "完整报告主体"},
                           {"asset_id": "img-002", "display_mode": "full", "label": "完整显微图组"}]
        result = normalize_case(payload, [analysis(), analysis("img-002")])["presentation_plan"]["slides"][0]
        self.assertEqual(result["layout"], "report_pair")
        self.assertEqual(result["images"], slide["images"])
        for count in (0, 1, 3):
            invalid = copy.deepcopy(payload)
            invalid["presentation_plan"]["slides"][0]["images"] = [slide["images"][0]] * count
            with self.subTest(count=count), self.assertRaises(ValueError):
                normalize_case(invalid, [analysis(), analysis("img-002")])
        slide["type"] = "cover"
        with self.assertRaises(ValueError):
            normalize_case(payload, [analysis(), analysis("img-002")])

    def test_cross_case_asset_and_source_ids_are_rejected(self):
        for value in ("other-case", None, [], 5):
            for field in ("images", "source_ids"):
                with self.subTest(field=field, value=value), self.assertRaises(ValueError):
                    payload = self.payload()
                    slide = payload["presentation_plan"]["slides"][0]
                    if field == "images":
                        slide["images"][0]["asset_id"] = value
                    else:
                        slide["source_ids"] = [value]
                    normalize_case(payload, [analysis()])

    def test_treatment_indices_are_checked_and_remapped(self):
        payload = self.payload({"type": "content", "title": "阶段治疗", "section": "treatment", "treatment_index": 2})
        payload["treatments"] = [None, {}, {"phase": "后续治疗"}]
        result = normalize_case(payload)
        self.assertEqual(result["presentation_plan"]["slides"][0]["treatment_index"], 1)
        for index in (True, "2", -1, 0, 8, None):
            with self.subTest(index=index), self.assertRaises(ValueError):
                payload["presentation_plan"]["slides"][0]["treatment_index"] = index
                normalize_case(payload)
        with self.assertRaises(ValueError):
            normalize_case(self.payload({"type": "content", "title": "基线", "section": "baseline", "treatment_index": 0}))

    def test_malformed_table_cannot_silently_drop_clinical_cells(self):
        for table in (None, {}, {"columns": [], "rows": []}, {"columns": ["日期", "结果"], "rows": [["日期"]]},
                      {"columns": ["日期"], "rows": [[4]]}, {"columns": ["日期"], "rows": "结果"}):
            with self.subTest(table=table), self.assertRaises(ValueError):
                normalize_case(self.payload({"type": "table", "title": "对照", "table": table}))
        with self.assertRaises(ValueError):
            normalize_case(self.payload({"type": "table", "title": "对照"}))

    def test_table_sizes_reject_nonfinite_boolean_and_invalid_dimensions(self):
        for key, value in (("column_widths", [0]), ("column_widths", [float("nan")]), ("column_widths", [True]),
                           ("column_widths", [1, 2]), ("font_size", 9), ("font_size", float("inf")), ("font_size", True)):
            with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                table = {"columns": ["结果"], "rows": [["既定结果"]], key: value}
                normalize_case(self.payload({"type": "table", "title": "对照", "table": table}))

    def test_oversized_text_and_arrays_fail_without_truncation(self):
        for field, value in (("title", "字" * 201), ("sections", [{"text": "字" * 6001}]),
                             ("sections", [{}] * 13), ("images", [{}] * 9), ("source_ids", ["img-001"] * 31)):
            with self.subTest(field=field), self.assertRaises(ValueError):
                payload = self.payload()
                payload["presentation_plan"]["slides"][0][field] = value
                normalize_case(payload, [analysis()])

    def test_sources_are_unique_and_optional_empty_arrays_remain_valid(self):
        payload = self.payload()
        payload["presentation_plan"]["slides"][0]["source_ids"] *= 2
        slide = normalize_case(payload, [analysis()])["presentation_plan"]["slides"][0]
        self.assertEqual(slide["source_ids"], ["img-001"])
        result = normalize_case(self.payload({"type": "cover", "title": "病例汇报"}))
        self.assertEqual(result["presentation_plan"]["slides"][0]["images"], [])

    def test_generated_plan_cannot_bypass_source_line_constraints(self):
        payload = self.payload({"type": "table", "title": "二线治疗", "section": "treatment", "treatment_index": 0,
                                "sections": [{"heading": "二线评估", "text": "二线方案A，继续同药减量"}],
                                "table": {"columns": ["二线治疗", "结果"], "rows": [["二线方案A", "已记录结果"]]}})
        slide = normalize_case(payload, source_text="2026-01-01使用方案A治疗，随后减量。")["presentation_plan"]["slides"][0]
        self.assertNotIn("二线", json.dumps(slide, ensure_ascii=False))
        preserved = normalize_case(payload)["presentation_plan"]["slides"][0]
        self.assertEqual(preserved["title"], "二线治疗")

    def test_prompt_groups_reports_and_does_not_count_local_treatment_as_new_line(self):
        self.assertIn("10–16页", CASE_PROMPT)
        self.assertIn("不单独增加系统治疗线数", CASE_PROMPT)
        self.assertIn("TAE", CASE_PROMPT)
        self.assertIn("不以方案序号推断线次", CASE_PROMPT)
        self.assertIn("可选返回presentation_plan", CASE_PROMPT)
        self.assertIn("完整句子和分点叙述为主", CASE_PROMPT)
        self.assertIn("完整的显微图组", CASE_PROMPT)
        self.assertIn("没有提供CT/MRI原始图像时不得合成", CASE_PROMPT)

    def test_line_only_title_remains_valid_after_grounding_and_renormalizing(self):
        payload = self.payload({"type": "table", "title": "二线", "table": {"columns": ["二线"], "rows": [["方案A"]]}})
        result = normalize_case(payload, source_text="方案A减量继续治疗。")
        self.assertEqual(result["presentation_plan"]["slides"][0]["title"], "治疗阶段")
        self.assertEqual(normalize_case(result, source_text="方案A减量继续治疗。"), result)


class TimelinePresentationTests(unittest.TestCase):
    @staticmethod
    def payload(events=None, **fields):
        slide = {"type": "timeline", "title": "诊疗时间轴", "section": "treatment", **fields}
        if events is not None:
            slide["events"] = events
        return {"baseline": {}, "treatments": [{"phase": "已记录阶段"}],
                "presentation_plan": {"slides": [slide]}}

    @staticmethod
    def event(**fields):
        return {"date": "2025-01至2025-04", "phase": "化疗", "event_type": "Treatment",
                "event": "记录中的方案A", **fields}

    def test_reviewed_ranges_and_all_nodes_survive_without_inferred_lines(self):
        nodes = [self.event(date=f"2025-{month:02d}") for month in range(1, 12)]
        nodes[0]["date"] = "2025.01—04"
        result = normalize_case(self.payload(nodes))
        slide = result["presentation_plan"]["slides"][0]
        self.assertEqual(slide["events"], nodes)
        self.assertEqual(len(slide["events"]), 11)
        self.assertIsNone(slide["treatment_index"])
        self.assertNotIn("一线", json.dumps(slide, ensure_ascii=False))
        self.assertEqual(normalize_case(result), result)

    def test_optional_events_use_grounded_timeline_and_preserve_provenance_on_restore(self):
        record = analysis()
        payload = self.payload()
        result = normalize_case(payload, [record], source_text="2025-01-10予方案A治疗。")
        self.assertEqual(result["presentation_plan"]["slides"][0]["events"], result["timeline_events"])
        self.assertIn("img-001", result["presentation_plan"]["slides"][0]["source_ids"])
        self.assertEqual(result["timeline_events"][0]["source_asset_id"], "patient_text")
        restored = normalize_case(result, [record])
        self.assertEqual(restored["timeline_events"], result["timeline_events"])
        self.assertEqual(restored["presentation_plan"], result["presentation_plan"])
        with self.assertRaises(ValueError):
            normalize_case(payload, source_text="没有注明诊疗日期。")

    def test_sources_are_validated_and_added_to_slide_references(self):
        node = self.event(source_asset_id="img-001", source_quote="2025年原始报告内容")
        slide = normalize_case(self.payload([node]), [analysis()])["presentation_plan"]["slides"][0]
        self.assertEqual(slide["source_ids"], ["img-001"])
        for invalid in ("another-case", [], None):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                normalize_case(self.payload([self.event(source_asset_id=invalid)]), [analysis()])

    def test_invalid_nodes_empty_arrays_and_mixed_layouts_are_rejected(self):
        for nodes in ([], "not-an-array", [None], [self.event(event="")],
                      [self.event(date=None)], [self.event(event_type="unknown")],
                      [self.event()] * 121, [self.event(event="字" * 6001)]):
            with self.subTest(nodes=type(nodes)), self.assertRaises(ValueError):
                normalize_case(self.payload(nodes))
        for field, value in (("sections", [{"text": "不得静默遗漏的内容"}]), ("images", [{}]),
                             ("table", {}), ("layout", "bottom")):
            with self.subTest(field=field), self.assertRaises(ValueError):
                normalize_case(self.payload([self.event()], **{field: value}))
        with self.assertRaises(ValueError):
            normalize_case({"baseline": {}, "presentation_plan": {"slides": [{
                "type": "content", "title": "正文", "events": [self.event()]}]}})

    def test_cross_phase_timeline_allows_no_index_but_validates_explicit_index(self):
        for value in (None, 0):
            result = normalize_case(self.payload([self.event()], treatment_index=value))
            self.assertEqual(result["presentation_plan"]["slides"][0]["treatment_index"], value)
        for value in (True, "0", -1, 1, []):
            with self.subTest(value=value), self.assertRaises(ValueError):
                normalize_case(self.payload([self.event()], treatment_index=value))

    def test_plans_remain_distinct_from_administered_treatment(self):
        for node in (self.event(event_type="Plan"), self.event(phase="门诊建议"),
                     self.event(event="拟予方案A治疗")):
            result = normalize_case(self.payload([node]))
            self.assertEqual(result["presentation_plan"]["slides"][0]["events"][0]["event_type"], "Treatment Plan")
        result = normalize_case(self.payload([self.event(date="2025-09-30", event="已经开始方案A治疗")]),
                                source_text="2025-09-30建议方案A治疗。")
        node = result["presentation_plan"]["slides"][0]["events"][0]
        self.assertEqual(node["event_type"], "Treatment Plan")
        self.assertIn("建议", node["event"])
        self.assertNotIn("已经开始", node["event"])
        with self.assertRaises(ValueError):
            normalize_case(self.payload([self.event(date="2025-09-30")]),
                           source_text="2025-09-30建议方案A治疗。2025-09-30建议方案B治疗。")

    def test_procedure_dates_cannot_inherit_imaging_or_admission_dates(self):
        node = self.event(date="2025-09-01", event_type="Biopsy", event="肺穿刺活检")
        for source in ("2025-09-01行CT检查，行肺穿刺活检。", "2025-09-01入院。"):
            with self.subTest(source=source), self.assertRaises(ValueError):
                normalize_case(self.payload([node]), source_text=source)
        result = normalize_case(self.payload([node]), source_text="2025-09-01行肺穿刺活检。")
        self.assertEqual(result["presentation_plan"]["slides"][0]["events"][0], node)

    def test_sourced_surgery_day_can_be_shown_as_month_but_month_cannot_gain_a_day(self):
        node = self.event(date="2024.08", event_type="Surgery", event="左肺上叶楔形切除术")
        result = normalize_case(self.payload([node]), source_text="2024-08-09行左肺上叶楔形切除术。")
        self.assertEqual(result["presentation_plan"]["slides"][0]["events"][0]["date"], "2024.08")
        node["date"] = "2024-08-09"
        with self.assertRaises(ValueError):
            normalize_case(self.payload([node]), source_text="2024年8月行左肺上叶楔形切除术。")

    def test_explicit_timeline_cannot_bypass_line_grounding(self):
        result = normalize_case(self.payload([self.event(date="2025-01-10", phase="二线化疗", event="二线方案A")]),
                                source_text="2025-01-10予方案A治疗。")
        node = result["presentation_plan"]["slides"][0]["events"][0]
        self.assertEqual((node["phase"], node["event"]), ("化疗", "方案A"))
        self.assertIn("重复给药合并", CASE_PROMPT)
        self.assertIn("type为timeline", CASE_PROMPT)
        self.assertIn("最多10个节点", CASE_PROMPT)


if __name__ == "__main__":
    unittest.main()
