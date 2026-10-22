import base64
import copy
import io
import json
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PIL import Image

import vision_pipeline as vision


def image_bytes(width=100, height=80, image_format="PNG"):
    image = Image.new("RGB", (width, height), "white")
    for x in range(width):
        for y in range(height):
            image.putpixel((x, y), (x % 256, y % 256, (x + y) % 256))
    output = io.BytesIO()
    image.save(output, format=image_format)
    return output.getvalue()


def payload(asset_id="img_test"):
    return {
        "asset_id": asset_id,
        "category": "pathology",
        "date": "2024-03-01",
        "title": "病理及免疫组化",
        "findings": ["Ki-67 约 15%"],
        "transcription": "Ki-67：约 15%；CK（+）",
        "regions": [{"label": "免疫组化结果", "box": [100, 200, 900, 800]}],
        "warnings": [],
    }


def response(data, finish_reason="stop"):
    content = data if isinstance(data, str) else json.dumps(data, ensure_ascii=False)
    return SimpleNamespace(choices=[SimpleNamespace(
        finish_reason=finish_reason, message=SimpleNamespace(content=content),
    )])


class PrepareAndCropTests(unittest.TestCase):
    def test_crop_uses_actual_original_pixels_and_stable_id(self):
        raw = image_bytes()
        asset = vision.prepare_asset(raw, r"C:\private\report.png", "img_test")
        self.assertEqual(asset["name"], "report.png")
        self.assertEqual((asset["width"], asset["height"]), (100, 80))
        region = {"box": [100, 200, 900, 800]}
        with Image.open(io.BytesIO(vision.crop_region(asset, region))) as crop:
            self.assertEqual(crop.size, (80, 48))
            self.assertEqual(crop.getpixel((0, 0)), (10, 16, 26))
            self.assertEqual(crop.getpixel((79, 47)), (89, 63, 152))
        self.assertEqual(
            vision.prepare_asset(raw, "same.png")["id"],
            vision.prepare_asset(raw, "different.png")["id"],
        )

    def test_crop_bytes_signature_and_non_divisible_bounds(self):
        with Image.open(io.BytesIO(vision.crop_region(image_bytes(7, 7), [100, 100, 900, 900]))) as crop:
            self.assertEqual(crop.size, (7, 7))
            self.assertEqual(crop.getpixel((0, 0)), (0, 0, 0))

    def test_exif_orientation_is_corrected_before_coordinates(self):
        original = Image.new("RGB", (40, 20), "red")
        original.paste("blue", (20, 0, 40, 20))
        exif = Image.Exif()
        exif[274] = 6  # Rotate 90 degrees clockwise for display.
        out = io.BytesIO()
        original.save(out, format="JPEG", exif=exif, quality=100, subsampling=0)
        asset = vision.prepare_asset(out.getvalue(), "oriented.jpg")
        self.assertEqual((asset["width"], asset["height"]), (20, 40))
        with Image.open(io.BytesIO(asset["image_bytes"])) as corrected:
            self.assertGreater(corrected.getpixel((10, 5))[0], 240)
            self.assertGreater(corrected.getpixel((10, 35))[2], 240)
            self.assertFalse(corrected.getexif())
        with Image.open(io.BytesIO(vision.crop_region(asset, [0, 500, 1000, 1000]))) as bottom:
            self.assertEqual(bottom.size, (20, 20))
            self.assertGreater(bottom.getpixel((10, 10))[2], 240)

    def test_valid_jpeg_above_previous_upload_limit_keeps_all_decoded_pixels(self):
        original = image_bytes(37, 29, "JPEG")
        # Some camera/application exports carry large unused trailing payloads.
        # The file limit concerns the upload bytes, not the decoded pixel count.
        large = original + bytes(21 * 1024 * 1024 - len(original))
        asset = vision.prepare_asset(large, "large-export.jpg")
        with Image.open(io.BytesIO(original)) as source, Image.open(io.BytesIO(asset["image_bytes"])) as normalized:
            self.assertEqual(normalized.format, "PNG")
            self.assertEqual(normalized.size, source.size)
            self.assertEqual(normalized.convert("RGB").tobytes(), source.convert("RGB").tobytes())

    def test_rejects_empty_corrupt_overlarge_animated_and_bad_id(self):
        for raw in (b"", b"not-an-image"):
            with self.subTest(raw=raw), self.assertRaises(vision.VisionPipelineError):
                vision.prepare_asset(raw, "invalid.png")
        with patch.object(vision, "MAX_UPLOAD_BYTES", 1), self.assertRaises(vision.VisionPipelineError):
            vision.prepare_asset(image_bytes(2, 2), "large.png")
        with patch.object(vision, "MAX_IMAGE_PIXELS", 3), self.assertRaises(vision.VisionPipelineError):
            vision.prepare_asset(image_bytes(2, 2), "large.png")
        with self.assertRaises(vision.VisionPipelineError):
            vision.prepare_asset(image_bytes(), "a.png", "../../escape")
        frames = io.BytesIO()
        Image.new("RGB", (2, 2), "red").save(
            frames, format="GIF", save_all=True, append_images=[Image.new("RGB", (2, 2), "blue")],
        )
        with self.assertRaises(vision.VisionPipelineError):
            vision.prepare_asset(frames.getvalue(), "animation.gif")

    def test_invalid_boxes_never_clamp_or_convert_units(self):
        for box in ([0., 0., 1., 1.], [0, 0, 1001, 1000], [-1, 0, 500, 500],
                    [400, 0, 300, 500], [0, 0, 0, 500], [0, 0, True, 500], [0, 1, 2]):
            with self.subTest(box=box), self.assertRaises(vision.VisionPipelineError):
                vision.crop_region(image_bytes(), box)


class ResponseValidationTests(unittest.TestCase):
    def test_accepts_completed_json_and_fenced_json(self):
        data = payload()
        self.assertEqual(vision.parse_json_response(response(data)), data)
        self.assertEqual(vision.parse_json_response("```json\n" + json.dumps(data) + "\n```"), data)
        raw_response = {"choices": [{"finish_reason": "stop", "message": {"content": "{}"}}]}
        self.assertEqual(vision.parse_json_response(raw_response), {})

    def test_rejects_truncation_even_when_partial_json_looks_valid(self):
        for reason in ("length", "content_filter", "tool_calls", None):
            with self.subTest(reason=reason), self.assertRaises(vision.VisionPipelineError):
                vision.parse_json_response(response({"findings": []}, reason))

    def test_rejects_empty_list_duplicate_keys_nan_and_trailing_text(self):
        for content in ("", "[]", '{"a":1,"a":2}', '{"a":NaN}', '{"a":1} trailing', '{"a":'):
            with self.subTest(content=content), self.assertRaises(vision.VisionPipelineError):
                vision.parse_json_response(response(content))

    def test_rejects_mismatched_source_schema_and_coordinate_units(self):
        for field, value in (
            ("asset_id", "other_source"), ("category", "invalid"), ("category", {}),
            ("findings", "not a list"), ("transcription", None),
            ("regions", [{"label": "bad", "box": [0.1, 0.1, 0.9, 0.9]}]),
            ("regions", [{"label": "bad", "box": [0, 0, 100, 100]}] * 7),
        ):
            data = payload()
            data[field] = value
            with self.subTest(field=field, value=value), self.assertRaises(vision.VisionPipelineError):
                vision.validate_analysis(data, "img_test")

    def test_assigns_local_region_ids_and_summary_has_source_ids(self):
        data = payload()
        data["regions"][0]["id"] = "untrusted-model-id"
        validated = vision.validate_analysis(data, "img_test")
        self.assertEqual(validated["regions"][0]["id"], "img_test-r1")
        summary = json.loads(vision.build_evidence_summary([validated]))
        self.assertEqual(summary[0]["asset_id"], "img_test")
        self.assertEqual(summary[0]["regions"][0]["id"], "img_test-r1")
        with self.assertRaises(vision.VisionPipelineError):
            vision.evidence_by_id([validated, validated])


class EvidenceSanitizationTests(unittest.TestCase):
    def test_removes_labelled_identifiers_and_preserves_medical_fields(self):
        data = vision.validate_analysis(payload(), "img_test")
        data["title"] = "姓名：示例对象甲；基因检测报告"
        data["findings"] = ["住院号：DEMO-001", "样本编号：SAMPLE-001；TMB：4.8 mut/Mb"]
        data["transcription"] = (
            "姓名：示例对象甲；性别：女；年龄：58岁；报告日期：2024-03-01\n"
            "病理号：DEMO-P1 | 样本类型：石蜡组织 | 样本编号：DEMO-S1\n"
            "送检医生：示例医生甲；电话：00000000000；地址：示例地址\n"
            "EGFR：c.2573T>G (p.L858R)；变异频率：19.3%；TMB：4.8 mut/Mb；MSI：MSS\n"
            "Ki-67：约 15%；CK（+）；检测基因数：525；检测日期：2024-03-02"
        )
        before = copy.deepcopy(data)
        result = vision.sanitize_analysis(data)
        self.assertEqual(data, before)
        self.assertEqual(result["title"], "基因检测报告")
        self.assertEqual(result["findings"], ["TMB：4.8 mut/Mb"])
        self.assertEqual(result["regions"], data["regions"])
        for expected in ("性别：女", "年龄：58岁", "报告日期：2024-03-01", "样本类型：石蜡组织",
                         "EGFR：c.2573T>G (p.L858R)", "变异频率：19.3%", "TMB：4.8 mut/Mb",
                         "MSI：MSS", "Ki-67：约 15%", "CK（+）", "检测基因数：525", "检测日期：2024-03-02"):
            self.assertIn(expected, result["transcription"])
        for removed in ("姓名", "住院号", "病理号", "样本编号", "送检医生", "电话", "地址", "DEMO", "示例对象"):
            self.assertNotIn(removed, result["transcription"])

    def test_handles_adjacent_clinical_fields_and_markdown_table_cells(self):
        data = payload()
        data["transcription"] = (
            "**姓名**：示例对象甲 年龄：58岁 性别：女\n"
            "住院号：DEMO-001报告日期：2024-03-01\n"
            "| 患者姓名 | 示例对象乙 | 样本类型 | 血液 |\n"
            "| 门诊号：DEMO-002 | 血小板：200 × 10^9/L |"
        )
        result = vision.sanitize_analysis(data)
        for removed in ("示例对象", "姓名", "住院号", "门诊号", "DEMO"):
            self.assertNotIn(removed, result["transcription"])
        for expected in ("年龄：58岁", "性别：女", "报告日期：2024-03-01", "样本类型 | 血液", "血小板：200 × 10^9/L"):
            self.assertIn(expected, result["transcription"])

    def test_drops_identity_header_regions_without_reindexing_or_changing_pixels(self):
        data = payload()
        data["regions"] = [
            {"label": "报告标题及送样信息", "box": [0, 0, 1000, 200]},
            {"label": "基因检测结果及表头", "box": [100, 200, 900, 800]},
        ]
        validated = vision.validate_analysis(data, "img_test")
        original_crop = vision.crop_region(image_bytes(), validated["regions"][1])
        clean = vision.sanitize_analysis(validated)
        self.assertEqual(clean["regions"], [validated["regions"][1]])
        self.assertEqual(clean["regions"][0]["id"], "img_test-r2")
        self.assertEqual(vision.crop_region(image_bytes(), clean["regions"][0]), original_crop)
        exported = json.loads(vision.build_evidence_summary([clean]))[0]
        self.assertEqual(exported["regions"], clean["regions"])
        self.assertEqual(vision.sanitize_analysis(clean), clean)

    def test_no_whole_image_fallback_when_all_regions_are_identity_headers(self):
        for label in ("报告标题及送样信息", "患者信息", "基本信息", "个人信息", "身份信息", "患者信息及病理诊断"):
            data = payload()
            data["regions"][0]["label"] = label
            before = copy.deepcopy(data)
            with self.subTest(label=label), self.assertRaisesRegex(vision.VisionPipelineError, "诊断或检测结果"):
                vision.sanitize_analysis(data)
            self.assertEqual(data, before)

    def test_omits_identifier_warnings_but_preserves_medical_uncertainties(self):
        data = payload()
        data["warnings"] = ["姓名模糊", "样本编号无法确认；EGFR 变异频率不清晰", "Ki-67 百分比模糊", "患者信息无法辨认"]
        self.assertEqual(vision.sanitize_analysis(data)["warnings"], ["EGFR 变异频率不清晰", "Ki-67 百分比模糊"])

    def test_export_helper_sanitizes_legacy_evidence(self):
        data = payload()
        data["transcription"] = "姓名：示例对象甲；检测基因数：525"
        exported = vision.build_evidence_summary([data])
        self.assertNotIn("示例对象甲", exported)
        self.assertIn("检测基因数：525", exported)

    def test_unrelated_clinical_table_rows_keep_spacing_and_empty_cells(self):
        data = payload()
        data["transcription"] = "| 基因 | 变异 | 频率 |\n| EGFR |  | 19.3% |\nKi-67  约 15%"
        self.assertEqual(vision.sanitize_analysis(data)["transcription"], data["transcription"])

    def test_unseparated_identifier_numbers_preserve_neighboring_clinical_text(self):
        data = payload()
        data["findings"] = ["病理号DEMO24001；CgA（+）；标本大小0.3 cm", "样本号SAMPLE-002；EGFR：c.2573T>G"]
        data["transcription"] = (
            "病理号DEMO24001报告时间：2024-08-09 11:26\n"
            "样本编号DEMO_S02；样本类型：石蜡组织；Ki-67：15%\n"
            "住院号DEMO-003；年龄：58岁；性别：女\n"
            "CDK6：4.2；TMB：5.0 mut/Mb"
        )
        result = vision.sanitize_analysis(data)
        self.assertEqual(result["findings"], ["CgA（+）；标本大小0.3 cm", "EGFR：c.2573T>G"])
        self.assertNotIn("DEMO", result["transcription"])
        for expected in ("报告时间：2024-08-09 11:26", "样本类型：石蜡组织", "Ki-67：15%",
                         "年龄：58岁", "性别：女", "CDK6：4.2", "TMB：5.0 mut/Mb"):
            self.assertIn(expected, result["transcription"])

    def test_explicit_doctor_signature_fields_only_are_removed(self):
        data = payload()
        data["transcription"] = (
            "初诊医生：示例医师甲；病理诊断：神经内分泌肿瘤\n"
            "审核医生：示例医师乙 | Ki-67：15%\n"
            "医师签名：示例医师丙；报告时间：2024-08-09 11:26\n"
            "病理诊断经审核后签发；TRPS1（−）"
        )
        result = vision.sanitize_analysis(data)
        for removed in ("初诊医生", "审核医生", "医师签名", "示例医师"):
            self.assertNotIn(removed, result["transcription"])
        for expected in ("病理诊断：神经内分泌肿瘤", "Ki-67：15%", "报告时间：2024-08-09 11:26",
                         "病理诊断经审核后签发；TRPS1（−）"):
            self.assertIn(expected, result["transcription"])


class RefinementTileTests(unittest.TestCase):
    def make_asset(self, width, height):
        raw = io.BytesIO()
        image = Image.new("RGB", (width, height), "white")
        image.putpixel((0, 0), (19, 23, 41))
        image.putpixel((width - 1, height - 1), (43, 47, 53))
        image.save(raw, format="PNG")
        return vision.prepare_asset(raw.getvalue(), "report.png", "img_test")

    def test_dense_tiles_cover_every_pixel_with_overlap_and_stable_source(self):
        asset = self.make_asset(1600, 1600)
        tiles = vision._source_refinement_tiles(asset)
        self.assertEqual(len(tiles), 4)
        coverage = Image.new("1", (1600, 1600), 0)
        rebuilt = Image.new("RGB", (1600, 1600))
        for tile in tiles:
            self.assertEqual(tile["asset_id"], "img_test")
            self.assertEqual(tile["source_size_px"], [1600, 1600])
            self.assertEqual(tile["tile_count"], 4)
            with Image.open(io.BytesIO(tile["image_bytes"])) as image:
                self.assertLessEqual(max(image.size), 1300)
                self.assertLessEqual(image.width * image.height, 1_300_000)
                coverage.paste(1, tuple(tile["bounds_px"]))
                rebuilt.paste(image, tuple(tile["bounds_px"][:2]))
        self.assertEqual(coverage.getextrema(), (1, 1))
        with Image.open(io.BytesIO(asset["image_bytes"])) as source:
            self.assertEqual(rebuilt.tobytes(), source.tobytes())
        x_starts = sorted({tile["bounds_px"][0] for tile in tiles})
        y_starts = sorted({tile["bounds_px"][1] for tile in tiles})
        self.assertGreaterEqual(1300 - (x_starts[1] - x_starts[0]), 120)
        self.assertGreaterEqual(1000 - (y_starts[1] - y_starts[0]), 120)

    def test_pixel_threshold_splits_even_when_both_sides_are_under_1300(self):
        asset = self.make_asset(1200, 1200)
        self.assertEqual(len(vision._source_refinement_tiles(asset)), 2)

    def test_tile_cap_errors_instead_of_omitting_report_remainder(self):
        asset = self.make_asset(1600, 1600)
        with patch.object(vision, "MAX_REFINEMENT_TILES", 4):
            with self.assertRaisesRegex(vision.VisionPipelineError, "全图概览超过"):
                vision._source_refinement_tiles(asset)


class ApiRequestTests(unittest.TestCase):
    def setUp(self):
        self.asset = vision.prepare_asset(image_bytes(), "report.png", "img_test")
        self.client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=Mock())))

    def test_image_request_uses_correct_model_json_and_original_detail(self):
        self.client.chat.completions.create.return_value = response(payload())
        result = vision.analyze_asset(self.client, self.asset)
        args = self.client.chat.completions.create.call_args.kwargs
        self.assertEqual(args["model"], "deepseek-flash")
        self.assertEqual(args["response_format"], {"type": "json_object"})
        image_part = args["messages"][1]["content"][1]["image_url"]
        self.assertEqual(image_part["detail"], "original")
        self.assertTrue(image_part["url"].startswith("data:image/png;base64,"))
        self.assertEqual(result["findings"], ["Ki-67 约 15%"])
        self.assertIn("0–1000", args["messages"][0]["content"])

    def test_transport_resize_does_not_resize_source_crop(self):
        raw = io.BytesIO()
        Image.new("RGB", (9000, 2), "white").save(raw, format="PNG")
        asset = vision.prepare_asset(raw.getvalue(), "long.png", "img_test")
        data_url = vision._api_image_url(asset["image_bytes"])
        with Image.open(io.BytesIO(base64.b64decode(data_url.split(",", 1)[1]))) as sent:
            self.assertLessEqual(sent.width, 8192)
        with Image.open(io.BytesIO(vision.crop_region(asset, [0, 0, 1000, 1000]))) as crop:
            self.assertEqual(crop.size, (9000, 2))

    def test_transport_byte_budget(self):
        raw = io.BytesIO()
        Image.effect_noise((256, 256), 100).convert("RGB").save(raw, format="PNG")
        url = vision._api_image_url(raw.getvalue(), max_bytes=2048)
        self.assertLessEqual(len(base64.b64decode(url.split(",", 1)[1])), 2048)

    def test_refinement_keeps_region_ids_boxes_and_first_pass_immutable(self):
        first = vision.validate_analysis(payload(), "img_test")
        first["warnings"] = ["局部日期模糊"]
        before = copy.deepcopy(first)
        self.client.chat.completions.create.return_value = response({
            "findings": ["Ki-67 约 15%；CK（+）"],
            "transcription": "Ki-67：约 15%；CK（+）", "warnings": [],
        })
        result = vision.refine_analysis(self.client, self.asset, first)
        self.assertEqual(first, before)
        self.assertEqual(result["regions"], before["regions"])
        self.assertEqual(result["asset_id"], before["asset_id"])
        self.assertEqual(result["warnings"], ["局部日期模糊"])
        self.assertEqual(result["findings"], ["Ki-67 约 15%；CK（+）"])
        parts = self.client.chat.completions.create.call_args.kwargs["messages"][1]["content"]
        self.assertEqual(sum(part["type"] == "image_url" for part in parts), 1)

    def test_both_api_passes_sanitize_text_and_keep_filtered_region_id(self):
        data = payload()
        data["transcription"] = "姓名：示例对象甲；Ki-67：约 15%；CK（+）"
        data["regions"].insert(0, {"label": "报告标题及送样信息", "box": [0, 0, 1000, 100]})
        self.client.chat.completions.create.return_value = response(data)
        first = vision.analyze_asset(self.client, self.asset)
        self.assertEqual(first["transcription"], "Ki-67：约 15%；CK（+）")
        self.assertEqual(first["regions"][0]["id"], "img_test-r2")
        self.client.chat.completions.create.return_value = response({
            "findings": ["样本编号：DEMO-001；Ki-67 约 15%"],
            "transcription": "住院号：DEMO-001；Ki-67：约 15%；CK（+）", "warnings": ["病理号模糊"],
        })
        second = vision.refine_analysis(self.client, self.asset, first)
        self.assertEqual(second["transcription"], first["transcription"])
        self.assertEqual(second["findings"], ["Ki-67 约 15%"])
        self.assertEqual(second["warnings"], [])
        self.assertEqual(second["regions"], first["regions"])

    def test_refinement_independently_reads_tiles_without_first_pass_results(self):
        raw = io.BytesIO()
        Image.new("RGB", (1600, 1600), "white").save(raw, format="PNG")
        asset = vision.prepare_asset(raw.getvalue(), "dense.png", "img_test")
        first = vision.validate_analysis(payload(), "img_test")
        first["regions"][0]["box"] = [0, 0, 1000, 1000]
        first["findings"] = ["INCORRECT_FIRST_PASS_MARKER"]
        first["transcription"] = "INCORRECT_TRANSCRIPTION_MARKER"
        self.client.chat.completions.create.return_value = response({
            "findings": ["CgA（+）；TRPS1（−）"],
            "transcription": "CgA（+）；TRPS1（−）；标本大小：0.3 cm", "warnings": [],
        })
        with patch.object(vision, "_api_image_url", wraps=vision._api_image_url) as encode:
            result = vision.refine_analysis(self.client, asset, first)
        parts = self.client.chat.completions.create.call_args.kwargs["messages"][1]["content"]
        text_parts = "\n".join(part["text"] for part in parts if part["type"] == "text")
        self.assertNotIn("INCORRECT_FIRST_PASS_MARKER", text_parts)
        self.assertNotIn("INCORRECT_TRANSCRIPTION_MARKER", text_parts)
        self.assertIn(first["date"], text_parts)
        self.assertIn(first["title"], text_parts)
        image_parts = [part["image_url"] for part in parts if part["type"] == "image_url"]
        self.assertEqual(len(image_parts), 5)
        binary_total = 0
        for part in image_parts:
            binary = base64.b64decode(part["url"].split(",", 1)[1])
            binary_total += len(binary)
            with Image.open(io.BytesIO(binary)) as image:
                self.assertLessEqual(max(image.size), 1300)
        self.assertLessEqual(binary_total, 24 * 1024 * 1024)
        self.assertLessEqual(sum(call.kwargs["max_bytes"] for call in encode.call_args_list), 24 * 1024 * 1024)
        self.assertEqual(result["regions"], first["regions"])
        self.assertIn("0.3 cm", result["transcription"])

    def test_narrow_slide_roi_never_removes_left_columns_or_source_pixels_from_reading(self):
        raw = io.BytesIO()
        source = Image.new("RGB", (1600, 1600), "white")
        source.paste((19, 23, 41), (0, 0, 200, 1600))
        source.paste((43, 47, 53), (1500, 1500, 1600, 1600))
        source.save(raw, format="PNG")
        asset = vision.prepare_asset(raw.getvalue(), "table.png", "img_test")
        first = vision.validate_analysis(payload(), "img_test")
        first["category"] = "genetics"
        first["regions"][0]["box"] = [900, 900, 1000, 1000]
        before = copy.deepcopy(first)
        self.client.chat.completions.create.return_value = response({
            "findings": ["EGFR：检测结果见表"], "transcription": "EGFR：检测结果见表", "warnings": [],
        })
        result = vision.refine_analysis(self.client, asset, first)
        parts = self.client.chat.completions.create.call_args.kwargs["messages"][1]["content"]
        images = [part["image_url"] for part in parts if part["type"] == "image_url"]
        self.assertEqual(len(images), 5)
        with Image.open(io.BytesIO(base64.b64decode(images[0]["url"].split(",", 1)[1]))) as overview:
            self.assertEqual(overview.size, (1000, 1000))
        rebuilt = Image.new("RGB", (1600, 1600))
        coverage = Image.new("1", (1600, 1600), 0)
        for expected, part in zip(vision._source_refinement_tiles(asset), images[1:]):
            with Image.open(io.BytesIO(base64.b64decode(part["url"].split(",", 1)[1]))) as tile:
                rebuilt.paste(tile, tuple(expected["bounds_px"][:2]))
                coverage.paste(1, tuple(expected["bounds_px"]))
        self.assertEqual(coverage.getextrema(), (1, 1))
        self.assertEqual(rebuilt.tobytes(), source.tobytes())
        self.assertEqual(result["regions"], before["regions"])
        self.assertEqual(first, before)
        prompt = "\n".join(part["text"] for part in parts if part["type"] == "text")
        self.assertIn("全图像素范围", prompt)
        self.assertIn("左侧基因名", prompt)
        self.assertIn("完整表头", prompt)

    def test_failed_refinement_does_not_wipe_first_pass(self):
        first = vision.validate_analysis(payload(), "img_test")
        before = copy.deepcopy(first)
        self.client.chat.completions.create.return_value = response({}, "length")
        with self.assertRaises(vision.VisionPipelineError):
            vision.refine_analysis(self.client, self.asset, first)
        self.assertEqual(first, before)

    def test_refinement_skips_imaging(self):
        first = payload()
        first["category"] = "imaging"
        self.assertEqual(vision.refine_analysis(self.client, self.asset, first), first)
        self.client.chat.completions.create.assert_not_called()

    def test_report_refinement_reads_complete_source_even_without_slide_regions(self):
        first = payload()
        first["regions"] = []
        self.client.chat.completions.create.return_value = response({
            "findings": ["Ki-67 约 15%"], "transcription": "Ki-67：约 15%", "warnings": [],
        })
        result = vision.refine_analysis(self.client, self.asset, first)
        parts = self.client.chat.completions.create.call_args.kwargs["messages"][1]["content"]
        images = [part["image_url"] for part in parts if part["type"] == "image_url"]
        self.assertEqual(len(images), 1)
        with Image.open(io.BytesIO(base64.b64decode(images[0]["url"].split(",", 1)[1]))) as full_source:
            self.assertEqual(full_source.size, (100, 80))
        self.assertEqual(result["regions"], [])

    def test_refinement_only_replaces_date_when_response_provides_nonempty_date(self):
        first = vision.validate_analysis(payload(), "img_test")
        before = copy.deepcopy(first)
        for date_fields, expected in (({}, "2024-03-01"), ({"date": ""}, "2024-03-01"),
                                      ({"date": "   "}, "2024-03-01"), ({"date": None}, "2024-03-01"),
                                      ({"date": "2024-03-05"}, "2024-03-05")):
            with self.subTest(date_fields=date_fields):
                self.client.chat.completions.create.return_value = response({
                    "findings": ["Ki-67 约 15%"], "transcription": "Ki-67：约 15%", "warnings": [],
                    **date_fields,
                })
                result = vision.refine_analysis(self.client, self.asset, first)
                self.assertEqual(result["date"], expected)
                self.assertEqual(result["title"], first["title"])
                self.assertEqual(first, before)

    def test_refinement_rejects_malformed_optional_date_without_mutating_original(self):
        first = vision.validate_analysis(payload(), "img_test")
        before = copy.deepcopy(first)
        self.client.chat.completions.create.return_value = response({
            "findings": ["Ki-67 约 15%"], "transcription": "Ki-67：约 15%", "warnings": [], "date": 20240305,
        })
        with self.assertRaises(vision.VisionPipelineError):
            vision.refine_analysis(self.client, self.asset, first)
        self.assertEqual(first, before)

    def test_safe_error_never_includes_secret_or_raw_payload(self):
        err = RuntimeError("sk-secret private patient data")
        err.status_code = 401
        result = vision.safe_error_message(err)
        self.assertIn("API Key", result)
        self.assertNotIn("sk-secret", result)
        self.assertNotIn("patient", result)
        self.assertNotIn("sk-secret", vision.safe_error_message(RuntimeError("sk-secret")))


if __name__ == "__main__":
    unittest.main()
