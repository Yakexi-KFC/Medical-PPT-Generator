import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import io
import json
import base64
import requests
from openai import OpenAI

# ==========================================
# 🔑 密钥配置区 (使用 Streamlit Secrets 保护)
# ==========================================
BAIDU_API_KEY = st.secrets["BAIDU_API_KEY"]
BAIDU_SECRET_KEY = st.secrets["BAIDU_SECRET_KEY"]
DEEPSEEK_API_KEY = st.secrets["DEEPSEEK_API_KEY"]

# ==========================================
# 1. 百度 OCR 图片识别模块
# ==========================================
def get_baidu_access_token():
    url = f"https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={BAIDU_API_KEY}&client_secret={BAIDU_SECRET_KEY}"
    headers = {'Content-Type': 'application/json', 'Accept': 'application/json'}
    response = requests.request("POST", url, headers=headers, data="")
    return response.json().get("access_token")

def perform_ocr(image_bytes, access_token):
    try:
        url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" + access_token
        img_base64 = base64.b64encode(image_bytes).decode('utf-8')
        payload = {'image': img_base64}
        headers = {'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json'}
        response = requests.request("POST", url, headers=headers, data=payload)
        result_json = response.json()
        if "words_result" in result_json:
            text_list = [item["words"] for item in result_json["words_result"]]
            return "\n".join(text_list)
        else:
            return f"[识别错误: {result_json.get('error_msg', '未知错误')}]"
    except Exception as e:
        return f"[请求异常: {str(e)}]"

# ==========================================
# 2. AI 结构化提取模块 (严防漏药、严卡线数标准)
# ==========================================
def extract_complex_case(patient_text):
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY, 
        base_url="https://api.deepseek.com"
    )
    system_prompt = """
    你是一位极其严谨的肿瘤内科主任医师。请阅读用户提供的真实长篇病历，将其拆解为标准的病例汇报结构。
    
    【核心指令 1：严谨的线数判定 (重点！绝不可错)】
    - **非手术患者**：如果患者未进行根治性手术，其初始治疗绝对属于【一线治疗】，严禁标记为新辅助或辅助。
    - **同线维持与调整（严禁误判为新一线）**：只要影像学未明确提示疾病进展（PD），任何因毒副反应的停药、减量、转换为维持治疗（如二线改为百泽安+索凡替尼维持），或者仅因肿瘤标志物升高而增加靶向药（如四线期间加用安罗替尼），都**绝对不能**算作新的一线！必须将其作为“方案调整”或“维持治疗”合并在当前的治疗线数阶段内！
    - 只有在明确记录影像学【PD】或【复发】后更改方案，才算开启下一线。
    
    【核心指令 2：治疗细节的完整保留 (严防遗漏)】
    - `regimen` 字段必须包含该线治疗的**全过程**。如果该阶段包含前期方案和后期的维持/加药，必须按时间顺序列在同一个 `regimen` 中（如：前期行AG+百泽安，后期行百泽安+索凡替尼维持；前期行四药，后期加用安罗替尼），严禁遗漏维持或加药部分！
    - 完整保留放疗、介入等局部治疗。
    
    必须严格输出为以下 JSON 格式：
    {
        "cover": {"title": "晚期XXX癌综合治疗病例汇报"},
        "baseline": {
            "patient_info": "患者姓名(只保留姓氏)、性别、年龄",
            "chief_complaint": "主诉",
            "diagnosis": "完整的临床及病理诊断（含分期）",
            "key_exams": "关键的基线检查结果"
        },
        "treatments": [
            {
                "phase": "阶段（如：一线治疗 / 四线治疗）", 
                "duration": "具体时间段（必须涵盖维持期或加药期）", 
                "regimen": "【严禁遗漏】完整保留该阶段所有的全身用药经过（必须包含同一线内的维持用药和加药调整原文）及局部治疗", 
                "imaging": "关键影像学评估结果",
                "markers": "肿瘤标志物变化情况"
            }
        ],
        "current_admission": {
            "exams": ["检验指标1", "检验指标2"],
            "imaging": "本次影像学结论",
            "plan": ["计划1", "计划2"]
        },
        "timeline_events": [
            {
                "date": "年月", 
                "phase": "线数（如'一线'、'四线(加药)'、'二线维持'、'评估'）",
                "event_type": "Treatment 或 Evaluation",
                "event": "事件简述(限15字，必须提取出中途维持或加药的事件)"
            }
        ],
        "summary": ["总结点1", "总结点2"]
    }
    注意：timeline_events 需提取全病程中所有的换线节点、同线加药/维持节点和评估节点，按先后排序，最多提取12个！
    """
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": patient_text}
        ],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

# ==========================================
# 3. PPT 生成模块
# ==========================================
class AdvancedPPTMaker:
    def __init__(self, data):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333) 
        self.prs.slide_height = Inches(7.5)
        
        self.data = self.clean_data(data)
        
        self.C_PRI = RGBColor(115, 21, 40)   
        self.C_ACC = RGBColor(0, 51, 102)  

    def clean_data(self, data):
        has_surgery = False
        full_text = json.dumps(data, ensure_ascii=False)
        if "根治术" in full_text or "切除术" in full_text or "手术切除" in full_text:
            has_surgery = True
            
        if not has_surgery:
            for tx in data.get("treatments", []):
                p = tx.get("phase", "")
                if "辅助" in p:
                    tx["phase"] = "一线治疗" 
            for evt in data.get("timeline_events", []):
                p = evt.get("phase", "")
                if "辅助" in p:
                    evt["phase"] = "一线"
        return data

    def add_header(self, slide, text):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.9))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.C_PRI
        shape.line.fill.background()
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.05), Inches(10), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    def make_cover(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(7.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.C_PRI
        tb = slide.shapes.add_textbox(Inches(1.5), Inches(3), Inches(10), Inches(2))
        p = tb.text_frame.paragraphs[0]
        p.text = self.data.get("cover", {}).get("title", "病例汇报")
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    def make_baseline(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "病例介绍 (基线资料)")
        base_data = self.data.get("baseline", {})
        content = f"【患者信息】 {base_data.get('patient_info', '')}\n\n" \
                  f"【主诉】 {base_data.get('chief_complaint', '')}\n\n" \
                  f"【临床诊断】\n{base_data.get('diagnosis', '')}\n\n" \
                  f"【关键检查/病理】\n{base_data.get('key_exams', '')}"
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = content
        p.font.size = Pt(20) 
        
    def make_treatments(self):
        for tx in self.data.get("treatments", []):
            phase_name = tx.get('phase', '阶段治疗')
            if "辅助" in phase_name and len(tx.get('regimen', '')) < 5:
                continue

            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide, f"治疗经过：{phase_name}")
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(6))
            tf = tb.text_frame
            tf.word_wrap = True 
            
            p1 = tf.paragraphs[0]
            p1.text = f"【治疗时间】 {tx.get('duration', '')}"
            p1.font.size = Pt(20) 
            p1.font.bold = True
            p1.font.color.rgb = self.C_PRI
            
            p2 = tf.add_paragraph()
            p2.text = f"\n【用药方案及局部治疗】\n{tx.get('regimen', '')}"
            p2.font.size = Pt(16) 
            
            p3 = tf.add_paragraph()
            p3.text = f"\n【影像学评估】\n{tx.get('imaging', '')}"
            p3.font.size = Pt(16) 
            p3.font.color.rgb = RGBColor(50, 50, 50)
            
            p4 = tf.add_paragraph()
            p4.text = f"\n【肿瘤标志物】\n{tx.get('markers', '')}"
            p4.font.size = Pt(16) 
            p4.font.color.rgb = self.C_ACC

    def make_current_admission(self):
        adm_data = self.data.get("current_admission")
        if not adm_data: return
        
        exams_list = adm_data.get("exams", [])
        exams_str = "\n".join([f"• {item}" for item in exams_list]) if isinstance(exams_list, list) else str(exams_list)
        imaging_str = adm_data.get("imaging", "")
        plan_list = adm_data.get("plan", [])
        plan_str = "\n".join([f"• {item}" for item in plan_list]) if isinstance(plan_list, list) else str(plan_list)
        
        total_len = len(exams_str) + len(imaging_str) + len(plan_str)
        is_split = len(plan_str) > 200 or total_len > 500
        
        if is_split:
            slide1 = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide1, "本次入院评估 (1/2)")
            tb1 = slide1.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(5.5))
            tf1 = tb1.text_frame
            tf1.word_wrap = True
            p_ex_title = tf1.paragraphs[0]
            p_ex_title.text = "【入院检验指标】"
            p_ex_title.font.bold = True
            p_ex_title.font.size = Pt(20)
            p_ex_title.font.color.rgb = self.C_PRI
            p_ex_body = tf1.add_paragraph()
            p_ex_body.text = exams_str + "\n"
            p_ex_body.font.size = Pt(18)
            p_im_title = tf1.add_paragraph()
            p_im_title.text = "【影像学评估】"
            p_im_title.font.bold = True
            p_im_title.font.size = Pt(20)
            p_im_title.font.color.rgb = self.C_PRI
            p_im_body = tf1.add_paragraph()
            p_im_body.text = imaging_str
            p_im_body.font.size = Pt(18)
            
            slide2 = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide2, "后续治疗与随访计划 (2/2)")
            tb2 = slide2.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(5.5))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p_pl_title = tf2.paragraphs[0]
            p_pl_title.text = "【治疗与随访计划】"
            p_pl_title.font.bold = True
            p_pl_title.font.size = Pt(20)
            p_pl_title.font.color.rgb = self.C_PRI
            p_pl_body = tf2.add_paragraph()
            p_pl_body.text = plan_str
            p_pl_body.font.size = Pt(18)
        else:
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide, "本次入院评估及计划 (转归)")
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(6))
            tf = tb.text_frame
            tf.word_wrap = True
            content = f"【入院检验指标】\n{exams_str}\n\n【影像学评估】\n{imaging_str}\n\n【后续计划】\n{plan_str}"
            p = tf.paragraphs[0]
            p.text = content
            p.font.size = Pt(16)

    def make_timeline(self):
        """支持高达 12 个节点的动态微缩时间轴"""
        events = self.data.get("timeline_events", [])
        if not events: return
        
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "全病程时间轴概览 (Timeline)")
        
        line_y = Inches(4.2)
        start_x = Inches(0.6) # 起点更靠左，留出更多空间
        total_width = 12.1 
        count = min(len(events), 12) # 放宽到最多 12 个节点
        
        # 极高密度（10-12个节点）下的极小卡片
        if count > 9:
            card_width = Inches(0.95)
            card_height = Inches(1.4)
            font_size_date = Pt(9)
            font_size_body = Pt(8)
        # 中高密度（7-9个节点）下的中卡片
        elif count > 6:
            card_width = Inches(1.3)
            card_height = Inches(1.2)
            font_size_date = Pt(10)
            font_size_body = Pt(9)
        # 正常密度
        else:
            card_width = Inches(1.6)
            card_height = Inches(1.2)
            font_size_date = Pt(12)
            font_size_body = Pt(11)

        main_line = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, start_x - Inches(0.2), line_y - Inches(0.05), Inches(total_width + 0.4), Inches(0.1))
        main_line.fill.solid()
        main_line.fill.fore_color.rgb = RGBColor(220, 220, 220) 
        main_line.line.fill.background()
        
        for i, evt in enumerate(events[:12]): 
            if count > 1:
                x = start_x + Inches(total_width * (i / (count - 1)))
            else:
                x = start_x + Inches(total_width / 2)

            event_text = evt.get("event", "")
            phase_text = evt.get("phase", "") 
            event_type = evt.get("event_type", "Treatment")
            
            is_pd = "进展" in event_text or "PD" in event_text.upper() or "复发" in event_text
            is_control = "PR" in event_text.upper() or "SD" in event_text.upper() or "缩小" in event_text
            
            if is_pd:
                node_color = RGBColor(220, 50, 50) 
            elif is_control and event_type == "Evaluation":
                node_color = RGBColor(46, 139, 87) 
            else:
                node_color = self.C_PRI 
            
            # 动态计算连接线和卡片位置，防重叠绝对安全算法
            stem_height = Inches(1.0)
            stem_top = line_y - stem_height if i % 2 == 0 else line_y
            stem = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, stem_top, Inches(0.03), stem_height) 
            stem.fill.solid()
            stem.fill.fore_color.rgb = node_color
            stem.line.fill.background()
            
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x - Inches(0.15), line_y - Inches(0.15), Inches(0.3), Inches(0.3))
            circle.fill.solid()
            circle.fill.fore_color.rgb = node_color
            circle.line.color.rgb = RGBColor(255, 255, 255) 
            circle.line.width = Pt(2)
            
            # 卡片紧贴连接线末端
            card_top = line_y - stem_height - card_height if i % 2 == 0 else line_y + stem_height
            card_x = x - (card_width / 2)
            
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_top, card_width, card_height)
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(250, 250, 250) 
            card.line.color.rgb = node_color 
            card.line.width = Pt(1.5)
            
            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.05)
            tf.margin_right = Inches(0.05)
            tf.margin_top = Inches(0.05)
            
            p0 = tf.paragraphs[0]
            p0.text = evt.get("date", "")
            p0.font.bold = True
            p0.font.size = font_size_date
            p0.font.color.rgb = node_color
            p0.alignment = PP_ALIGN.CENTER
            
            if phase_text and phase_text != "评估":
                p_phase = tf.add_paragraph()
                p_phase.text = f"【{phase_text}】"
                p_phase.font.size = font_size_body
                p_phase.font.bold = True
                p_phase.font.color.rgb = node_color
                p_phase.alignment = PP_ALIGN.CENTER
            elif event_type == "Evaluation":
                p_phase = tf.add_paragraph()
                p_phase.text = "【疗效评估】"
                p_phase.font.size = font_size_body
                p_phase.font.bold = True
                p_phase.font.color.rgb = node_color 
                p_phase.alignment = PP_ALIGN.CENTER

            p1 = tf.add_paragraph()
            p1.text = event_text
            p1.font.size = font_size_body
            p1.font.color.rgb = RGBColor(30, 30, 30)
            p1.alignment = PP_ALIGN.CENTER

    def make_summary(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "病例小结与思考")
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(5))
        tf = tb.text_frame
        tf.word_wrap = True
        for item in self.data.get("summary", []):
            p = tf.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(20)
            p.space_after = Pt(14)

    def build(self):
        self.make_cover()
        self.make_baseline()
        self.make_treatments()
        self.make_current_admission()
        self.make_timeline()
        self.make_summary()
        
        ppt_stream = io.BytesIO()
        self.prs.save(ppt_stream)
        ppt_stream.seek(0)
        return ppt_stream

# ==========================================
# 4. Streamlit 网页前端
# ==========================================
st.set_page_config(page_title="Pro级肿瘤病例PPT生成", layout="wide")
st.title("🩺 医疗级病史 PPT 自动生成排版系统")

tab1, tab2 = st.tabs(["📸 多图连拍识别 (OCR)", "📝 电子病历粘贴"])

if "ocr_result_text" not in st.session_state:
    st.session_state.ocr_result_text = ""

with tab1:
    st.markdown("### 第一步：批量上传病历图片")
    uploaded_files = st.file_uploader(
        "支持拍照上传多张化验单、出院小结等（按顺序多选即可）", 
        type=["png", "jpg", "jpeg"], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.info(f"📁 已选择 {len(uploaded_files)} 张图片。")
        if st.button("🔍 开始批量提取文字"):
            with st.spinner("正在呼叫百度高精度 OCR 引擎扫描所有图片..."):
                token = get_baidu_access_token()
                if not token:
                    st.error("获取百度 API 授权失败，请检查密钥。")
                else:
                    all_extracted_text = []
                    for i, file in enumerate(uploaded_files):
                        image_bytes = file.getvalue()
                        text = perform_ocr(image_bytes, token)
                        all_extracted_text.append(f"【第 {i+1} 页提取结果】\n{text}\n")
                    st.session_state.ocr_result_text = "\n".join(all_extracted_text)
            st.success("✅ 文字提取成功！请在下方核对。")

    st.markdown("### 第二步：人工校对与修改")
    final_text_to_process = st.text_area(
        "校对并补全病史（支持手动补充没拍全的信息）：", 
        value=st.session_state.ocr_result_text, 
        height=350
    )
    
    if st.button("🚀 校对无误，自动推断分线并生成 PPT", type="primary"):
        if len(final_text_to_process) < 20:
            st.warning("⚠️ 病史太短，请补充详细记录。")
        else:
            try:
                with st.spinner('🤖 AI 正在化身老总，按时间轴拆解并自动推断您的治疗线数...'):
                    case_json = extract_complex_case(final_text_to_process)
                with st.spinner('📊 正在为您自动绘制时间轴并排版幻灯片...'):
                    maker = AdvancedPPTMaker(case_json)
                    ppt_file = maker.build()
                st.success("✅ 专业版病例幻灯片已生成就绪！")
                st.download_button(
                    label="📥 立即下载 PPT (含Timeline优化版)",
                    data=ppt_file,
                    file_name="病例汇报_Pro版.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"❌ 运行出错，请核对：{str(e)}")

with tab2:
    st.markdown("如果你已经有电子版的长病历（如从医院系统拷贝），可以直接粘贴在这里。")
    patient_input = st.text_area("请贴入详细病史：", height=250)
    if st.button("🚀 开始深度解析并生成 PPT", key="btn_text"):
        if len(patient_input) < 20:
            st.warning("⚠️ 病史太短，请提供详细病历。")
        else:
            try:
                with st.spinner('🤖 AI 正在按时间轴拆解并自动推断治疗线数...'):
                    case_json = extract_complex_case(patient_input)
                with st.spinner('📊 正在为您自动排版幻灯片...'):
                    maker = AdvancedPPTMaker(case_json)
                    ppt_file = maker.build()
                st.success("✅ 专业版病例幻灯片已生成就绪！")
                col1, col2 = st.columns([2, 1])
                with col1:
                    with st.expander("点击查看 AI 解析出的结构化病历树"):
                        st.json(case_json)
                with col2:
                    st.download_button(
                        label="📥 立即下载 PPT",
                        data=ppt_file,
                        file_name="病例汇报_文本版.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"❌ 运行出错，请核对：{str(e)}")
