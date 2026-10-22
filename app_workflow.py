"""Session-scoped upload, review, and editable slide-placement workflow."""
import copy
import hashlib
import io
import json
import os

import streamlit as st
from openai import OpenAI

from case_pipeline import extract_case, normalize_case
from ppt_builder import AdvancedPPTMaker, PPT_LAYOUT_VERSION
from vision_pipeline import prepare_asset, analyze_asset, refine_analysis, crop_region, safe_error_message, sanitize_analysis, validate_analysis, VisionPipelineError

CATEGORIES = {"imaging": "影像学", "pathology": "病理 / 免疫组化", "genetics": "基因检测", "lab": "实验室检查", "record": "病历记录", "other": "其他检查"}


def setting(name, default=""):
    value = os.environ.get(name)
    if value:
        return value
    try:
        return st.secrets.get(name, default)
    except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
        return default


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def make_client(api_key):
    return OpenAI(api_key=api_key, base_url="https://api.deepseek.com", timeout=180.0, max_retries=1)


def sync_uploads(files):
    signature = [(file.name, hashlib.sha256(file.getvalue()).hexdigest()) for file in files]
    upload_key = digest(signature)
    if upload_key == st.session_state.get("upload_signature"):
        return
    st.session_state.upload_signature = upload_key
    st.session_state.assets = []
    st.session_state.analyses = {}
    st.session_state.asset_errors = {}
    st.session_state.upload_errors = []
    st.session_state.refine_errors = {}
    st.session_state.crop_edits = {}
    st.session_state.analysis_versions = {}
    st.session_state.pop("case_result", None)
    for key in list(st.session_state):
        if key.startswith(("review_", "placement_", "export_")):
            del st.session_state[key]
    if len(files) > 30 or sum(len(file.getvalue()) for file in files) > 240 * 1024 * 1024:
        st.session_state.upload_errors.append("每批最多 30 张、原始文件合计不超过 240 MB，请分批上传。")
        return
    pixels = 0
    for index, file in enumerate(files):
        try:
            asset = prepare_asset(file.getvalue(), file.name, f"img-{index + 1:03d}-{signature[index][1][:8]}")
            asset["source_index"] = index + 1
            pixels += asset["width"] * asset["height"]
            if pixels > 120_000_000:
                raise ValueError("本批图片合计超过 1.2 亿像素，请减少图片后重试。")
            st.session_state.assets.append(asset)
        except ValueError as error:
            st.session_state.upload_errors.append(f"{file.name}：{error}")


def analyze_pending(api_key, model, refine):
    pending = [a for a in st.session_state.assets if st.session_state.get(include_key(a["id"]), True) and (a["id"] not in st.session_state.analyses or a["id"] in st.session_state.refine_errors)]
    if not pending:
        st.info("所有图片已识别，可在下方调整裁剪和文字。")
        return
    progress = st.progress(0, text="正在识别检查图片…")
    client = make_client(api_key)
    try:
        for index, asset in enumerate(pending):
            aid = asset["id"]
            progress.progress(index / len(pending), text=f"{index + 1}/{len(pending)} · {asset['name']}")
            try:
                previous = st.session_state.analyses.get(aid)
                analysis = previous or analyze_asset(client, asset, model=model)
                st.session_state.analyses[aid] = analysis
                st.session_state.asset_errors.pop(aid, None)
                if refine and analysis["category"] in ("pathology", "genetics", "lab", "record"):
                    try:
                        st.session_state.analyses[aid] = refine_analysis(client, asset, analysis, model=model)
                        st.session_state.refine_errors.pop(aid, None)
                    except Exception as error:
                        st.session_state.refine_errors[aid] = safe_error_message(error)
                elif not refine:
                    st.session_state.refine_errors.pop(aid, None)
                if previous != st.session_state.analyses[aid]:
                    st.session_state.analysis_versions[aid] = st.session_state.analysis_versions.get(aid, 0) + 1
            except Exception as error:
                st.session_state.asset_errors[aid] = safe_error_message(error)
        st.session_state.pop("case_result", None)
    finally:
        client.close()
    progress.progress(1.0, text=f"已识别 {len(st.session_state.analyses)}/{len(st.session_state.assets)} 张图片")


def include_key(asset_id):
    return f"review_{st.session_state.upload_signature[:10]}_{asset_id}_include"


def reread_asset(asset, api_key, refine=True):
    """Replace only this image after a complete successful reread."""
    client = make_client(api_key)
    try:
        analysis = analyze_asset(client, asset, model="deepseek-flash")
        if refine and analysis["category"] in ("pathology", "genetics", "lab", "record"):
            analysis = refine_analysis(client, asset, analysis, model="deepseek-flash")
        aid = asset["id"]
        st.session_state.analyses[aid] = analysis
        st.session_state.analysis_versions[aid] = st.session_state.analysis_versions.get(aid, 0) + 1
        st.session_state.asset_errors.pop(aid, None)
        st.session_state.refine_errors.pop(aid, None)
        st.session_state.pop("case_result", None)
        st.session_state.pop("export_result", None)
    finally:
        client.close()


def review_assets(api_key="", refine=True):
    selected_assets, selected_analyses, pending = [], [], []
    for index, asset in enumerate(st.session_state.assets):
        aid = asset["id"]
        original = st.session_state.analyses.get(aid)
        sanitation_error = None
        if original:
            try:
                cleaned = sanitize_analysis(original)
                if cleaned != original:
                    st.session_state.analyses[aid] = cleaned
                    st.session_state.analysis_versions[aid] = st.session_state.analysis_versions.get(aid, 0) + 1
                original = cleaned
            except ValueError as error:
                sanitation_error = str(error)
        revision = st.session_state.analysis_versions.get(aid, 0)
        prefix = f"review_{st.session_state.upload_signature[:10]}_{revision}_{aid}"
        with st.expander(f"资料 {asset.get('source_index', index + 1):02d} · {asset['name']}" + (" · 已识别" if original else " · 待识别")):
            included = st.checkbox("纳入病例与幻灯", value=True, key=include_key(aid))
            if api_key and st.button("重新识别本图", key=f"reread_{aid}"):
                try:
                    with st.spinner("正在重新识别此图并精读结果…"):
                        reread_asset(asset, api_key, refine)
                    st.rerun()
                except Exception as error:
                    st.error(safe_error_message(error))
            if sanitation_error:
                st.error(sanitation_error)
                if included:
                    pending.append(aid)
                continue
            if not original:
                st.image(asset["image_bytes"], width=420)
                if aid in st.session_state.asset_errors:
                    st.error(st.session_state.asset_errors[aid])
                else:
                    st.info("点击上方“识别检查图片”，或取消纳入此图。")
                if included:
                    pending.append(aid)
                continue
            analysis = copy.deepcopy(original)
            for warning in analysis.get("warnings", []):
                st.warning(warning)
            if aid in st.session_state.refine_errors:
                st.warning("完整报告精读未完成，已保留首次识别结果。可再次点击识别重试。")
            col1, col2 = st.columns(2)
            analysis["title"] = col1.text_input("检查标题", value=analysis["title"], key=f"{prefix}_title")
            analysis["date"] = col2.text_input("检查日期", value=analysis["date"], key=f"{prefix}_date")
            analysis["category"] = st.selectbox("检查类型", list(CATEGORIES), index=list(CATEGORIES).index(analysis["category"]), format_func=CATEGORIES.get, key=f"{prefix}_category")
            findings = st.text_area("检查要点（每行一条）", value="\n".join(analysis["findings"]), key=f"{prefix}_findings", height=100)
            analysis["findings"] = [line.strip() for line in findings.splitlines() if line.strip()]
            analysis["transcription"] = st.text_area("识别原文", value=analysis["transcription"], key=f"{prefix}_transcription", height=140)
            regions = analysis["regions"]
            if not regions:
                # No existing regions can collide; keep exported IDs valid for re-import.
                regions = [{"id": f"{aid}-r1", "label": "完整检查图", "box": [0, 0, 1000, 1000]}]
            chosen = st.selectbox("预览与调整裁剪区域", list(range(len(regions))), format_func=lambda i: regions[i]["label"], key=f"{prefix}_active_region")
            region = regions[chosen]
            rkey = f"{prefix}_{region['id']}"
            saved = st.session_state.crop_edits.get(rkey, {"box": region["box"], "include": True})
            left, top, right, bottom = saved["box"]
            with st.popover("精确输入裁剪范围"):
                manual_key = f"{rkey}_manual_{left}_{top}_{right}_{bottom}"
                x1, x2 = st.columns(2)
                new_left = x1.number_input("左边界", 0, 1000, left, key=f"{manual_key}_left")
                new_right = x2.number_input("右边界", 0, 1000, right, key=f"{manual_key}_right")
                y1, y2 = st.columns(2)
                new_top = y1.number_input("上边界", 0, 1000, top, key=f"{manual_key}_top")
                new_bottom = y2.number_input("下边界", 0, 1000, bottom, key=f"{manual_key}_bottom")
                if st.button("应用裁剪范围", key=f"{manual_key}_apply"):
                    if new_left >= new_right or new_top >= new_bottom:
                        st.error("右边界和下边界必须大于左边界和上边界。")
                    else:
                        st.session_state[f"{rkey}_x"] = (new_left, new_right)
                        st.session_state[f"{rkey}_y"] = (new_top, new_bottom)
                        st.session_state.crop_edits[rkey] = {"box": [new_left, new_top, new_right, new_bottom], "include": saved["include"]}
                        st.rerun()
            left, right = st.slider("左右范围（0—1000）", 0, 1000, (left, right), key=f"{rkey}_x")
            top, bottom = st.slider("上下范围（0—1000）", 0, 1000, (top, bottom), key=f"{rkey}_y")
            include_region = st.checkbox("将此区域放入幻灯", value=saved["include"], key=f"{rkey}_include")
            # Widgets for inactive regions are removed by Streamlit; keep durable session data separately.
            st.session_state.crop_edits[rkey] = {"box": [left, top, right, bottom], "include": include_region}
            reviewed = []
            for region in regions:
                rkey = f"{prefix}_{region['id']}"
                saved = st.session_state.crop_edits.get(rkey, {"box": region["box"], "include": True})
                if not saved["include"]:
                    continue
                region["box"] = list(saved["box"])
                try:
                    cropped = crop_region(asset, region)
                    if region["id"] == regions[chosen]["id"]:
                        left_col, right_col = st.columns(2)
                        left_col.image(asset["image_bytes"], caption="原图", use_container_width=True)
                        right_col.image(cropped, caption="实际导出区域", use_container_width=True)
                    reviewed.append(region)
                except ValueError:
                    st.error("裁剪区域宽高必须大于零，请调整范围。")
                    if included:
                        pending.append(aid)
            if not reviewed:
                st.warning("至少保留一个裁剪区域，或取消纳入此图。")
                if included:
                    pending.append(aid)
            analysis["regions"] = reviewed
            if included:
                selected_assets.append(asset)
                selected_analyses.append(analysis)
    return selected_assets, selected_analyses, pending


def case_input_signature(patient_text, analyses):
    # Cropping changes only affect layout; corrected clinical content requires restructuring.
    clinical = [{key: value for key, value in a.items() if key not in ("regions", "warnings")} for a in analyses]
    return digest({"text": patient_text, "evidence": clinical})


def import_reviewed_analyses(raw, patient_text=None):
    """Restore reviewed evidence only for the exact images in this session."""
    if len(raw) > 2 * 1024 * 1024:
        raise VisionPipelineError("核对记录不能超过 2 MB。")
    try:
        payload = json.loads(raw)
    except (ValueError, UnicodeError):
        raise VisionPipelineError("核对记录不是有效的 JSON 文件。") from None
    rows = payload.get("analyses") if isinstance(payload, dict) else None
    if not isinstance(rows, list) or not 1 <= len(rows) <= 30:
        raise VisionPipelineError("核对记录需包含 1—30 张图片的 analyses 列表。")
    allowed = {asset["id"] for asset in st.session_state.assets}
    checked = {}
    for row in rows:
        aid = row.get("asset_id") if isinstance(row, dict) else None
        if not isinstance(aid, str) or aid not in allowed or aid in checked:
            raise VisionPipelineError("核对记录包含重复或不属于本批图片的来源编号，请核对上传图片。")
        analysis = sanitize_analysis(validate_analysis(row, aid, preserve_region_ids=True))
        if not analysis["regions"]:
            raise VisionPipelineError("每张核对图片至少保留一个裁剪区域。")
        checked[aid] = analysis
    restored_case = None
    narrative = patient_text if patient_text is not None else st.session_state.get("patient_text", "")
    ordered = [checked[asset["id"]] for asset in st.session_state.assets if asset["id"] in checked]
    if "case" in payload:
        if set(checked) != allowed:
            raise VisionPipelineError("恢复完整病例时，识别记录必须覆盖本批所有图片。")
        restored_case = normalize_case(payload["case"], ordered, source_text=narrative)
    # Commit only after every source and crop has passed validation.
    for aid, analysis in checked.items():
        st.session_state.analyses[aid] = analysis
        st.session_state.analysis_versions[aid] = st.session_state.analysis_versions.get(aid, 0) + 1
        st.session_state.asset_errors.pop(aid, None)
        st.session_state.refine_errors.pop(aid, None)
    st.session_state.pop("case_result", None)
    st.session_state.pop("export_result", None)
    if restored_case is not None:
        st.session_state.case_result = {
            "signature": case_input_signature(narrative, ordered), "case": restored_case,
        }
    return len(checked)


def render_result(result, assets, analyses, patient_text=None):
    case = normalize_case(copy.deepcopy(result["case"]), analyses, source_text=patient_text)
    has_plan = bool(case.get("presentation_plan", {}).get("slides"))
    if has_plan:
        st.caption(f"按已核对的病例结构整合为 {len(case['presentation_plan']['slides'])} 页，相近检查集中呈现。")
    if assets and not has_plan:
        st.subheader("幻灯放置位置")
        positions = {"baseline": "基线资料之后"}
        for i, tx in enumerate(case["treatments"]):
            positions[f"treatment:{i}"] = f"治疗阶段 {i + 1}：{tx['phase']} {tx['duration']}"
        positions.update({"current_admission": "本次评估之后", "appendix": "补充检查资料"})
        for link in case["evidence_links"]:
            current = f"treatment:{link['treatment_index']}" if link["section"] == "treatment" else link["section"]
            key = f"placement_{result['signature']}_{link['asset_id']}"
            chosen = st.selectbox(link["caption"] or "检查资料", list(positions), index=list(positions).index(current if current in positions else "appendix"), format_func=positions.get, key=key)
            link["section"] = chosen.split(":")[0]
            link["treatment_index"] = int(chosen.split(":")[1]) if chosen.startswith("treatment:") else None
            # Each currently reviewed crop is included; stale model region selections cannot drop one.
            link["region_ids"] = []
    export_key = digest({"case": case, "analyses": analyses, "assets": [a["id"] for a in assets], "layout_version": PPT_LAYOUT_VERSION})
    cached = st.session_state.get("export_result")
    if not cached or cached["key"] != export_key:
        try:
            maker = AdvancedPPTMaker(case, assets=assets, analyses=analyses)
            ppt = maker.build().getvalue()
            cached = {"key": export_key, "ppt": ppt, "manifest": maker.evidence_manifest, "warnings": maker.warnings}
            st.session_state.export_result = cached
        except Exception:
            st.error("幻灯排版失败，请检查病例结构和裁剪区域后重试。")
            return
    st.success("病例幻灯已生成。修改裁剪或放置位置后，下载文件会自动更新。")
    for warning in cached["warnings"]:
        st.warning(warning)
    left, right = st.columns(2)
    left.download_button("下载完整病例 PPT", data=cached["ppt"], file_name="病例汇报_图文版.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    audit = {"case": case, "analyses": analyses, "slide_images": cached["manifest"]}
    right.download_button("下载识别与配图记录", json.dumps(audit, ensure_ascii=False, indent=2), "病例配图记录.json", "application/json")
    with st.expander("查看病例结构与配图页码"):
        st.json(audit)
    with st.expander("病例全病程逻辑线"):
        baseline = case["baseline"]
        st.write(baseline["patient_info"])
        st.write(baseline["diagnosis"])
        st.write(baseline["key_exams"])
        for tx in case["treatments"]:
            st.write(f"{tx['phase']} · {tx['duration']}")
            st.write(f"方案：{tx['regimen']}")
            st.write(f"评估：{tx['imaging']} {tx['markers']}")
        for plan in case["current_admission"]["plan"]:
            st.write(plan)


def render_workflow():
    with st.sidebar:
        st.subheader("DeepSeek 配置")
        configured_key = setting("DEEPSEEK_API_KEY")
        session_key = st.text_input("API key（仅本次会话）", type="password", help="也可由部署环境的 DEEPSEEK_API_KEY 提供。不会写入代码或导出文件。")
        api_key = session_key.strip() or configured_key
        st.caption("已配置 API key" if api_key else "填入 API key 后即可识别和生成。")
        st.caption("识图模型：deepseek-flash")
        text_model = st.selectbox("病例整理模型", ["deepseek-v4-pro", "deepseek-flash"])
        refine = st.checkbox("精读完整报告", value=True, help="再次读取完整报告的小字与表格，避免配图裁剪遗漏检查内容。每张报告会增加一次调用。")
    st.subheader("1. 上传检查图片与病史")
    st.caption("支持影像、病理、免疫组化、基因检测和检验报告。每张最多 40 MB / 2000 万像素；每批最多 30 张、240 MB。")
    files = st.file_uploader("添加检查图片（可多选）", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="case_uploads") or []
    sync_uploads(files)
    for error in st.session_state.upload_errors:
        st.error(error)
    patient_text = st.text_area("补充病史与治疗经过（也可仅粘贴病历生成）", height=220, key="patient_text", placeholder="可补充主诉、诊断、各阶段治疗方案与时间、疗效记录及后续计划。")
    if st.session_state.assets:
        st.caption(f"已载入 {len(st.session_state.assets)} 张图片。点击识别后，原图和裁剪结果可逐一对照。")
        if st.button("识别检查图片 / 重试失败项", type="primary", disabled=not api_key):
            analyze_pending(api_key, "deepseek-flash", refine)
        st.subheader("2. 核对检查内容与裁剪")
    assets, analyses, pending = review_assets(api_key, refine)
    if analyses:
        st.download_button(
            "下载图片识别记录",
            json.dumps({"analyses": analyses}, ensure_ascii=False, indent=2),
            "图片识别记录.json", "application/json",
        )
        with st.expander("导入已核对的识别记录"):
            st.caption("可恢复识别文字与裁剪；含病例结构的配图记录会同时恢复整合幻灯。")
            imported = st.file_uploader("选择此前导出的识别记录 JSON", type=["json"], key="review_import")
            if st.button("应用核对记录", disabled=imported is None):
                try:
                    import_reviewed_analyses(imported.getvalue(), patient_text)
                    st.rerun()
                except ValueError as error:
                    st.error(str(error))
    signature = case_input_signature(patient_text, analyses)
    ready = not pending and not st.session_state.upload_errors and (bool(analyses) or len(patient_text.strip()) >= 20)
    if pending:
        st.info("纳入的图片尚有未识别或无效裁剪项。可重试识别，或取消纳入相应图片。")
    if st.button("整理病例并生成图文 PPT", type="primary", disabled=not api_key or not ready):
        client = make_client(api_key)
        try:
            with st.spinner("正在整理检查资料、病程和配图位置…"):
                case = extract_case(client, patient_text, analyses, model=text_model)
            st.session_state.case_result = {"signature": signature, "case": case}
            st.session_state.pop("export_result", None)
        except Exception as error:
            st.error(safe_error_message(error))
        finally:
            client.close()
    result = st.session_state.get("case_result")
    if result and result["signature"] == signature and ready:
        render_result(result, assets, analyses, patient_text)
    elif result:
        st.info("病例文字或纳入资料已修改，请重新整理病例后下载。")
