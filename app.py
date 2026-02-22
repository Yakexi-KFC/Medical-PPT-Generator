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
# 2. AI 结构化提取模块 (终极抗遗漏版：强保局部治疗与原文细节)
# ==========================================
def extract_complex_case(patient_text):
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY, 
        base_url="https://api.deepseek.com"
    )
    system_prompt = """
    你是一位极其严谨的肿瘤内科主任医师。请阅读用户提供的真实长篇病历，将其拆解为标准的病例汇报结构。
    
    【核心指令与肿瘤内科铁律 - 极其重要，严禁漏字】：
    1. 零删减原则：绝不要过度精简或自行概括！必须完整保留原病历中的详细客观描述。特别是【放疗】、【手术】、【介入微创】等局部治疗手段，绝对不允许遗漏！必须将它们一字不差地归入对应时间段的治疗过程中。
    2. 严格的线数划分铁律：
       - 只有在明确记录【疾病进展（PD）】或【复发】后彻底更改方案，才算开启下一线治疗。
       - 若未进展而更改/停用部分药物，必须判定为【同一线的维持治疗】。
    
    必须严格输出为以下 JSON 格式：
    {
        "cover": {"title": "晚期XXX癌综合治疗病例汇报"},
        "baseline": {
            "patient_info": "患者姓名(只保留姓氏加某某)、性别、年龄",
            "chief_complaint": "主诉（如无明确主诉，根据病史总结）",
            "diagnosis": "完整的临床及病理诊断（含分期）",
            "key_exams": "关键的病理、基因检测、免疫组化或其他重要基线检查结果"
        },
        "treatments": [
            {
                "phase": "遵守铁律推断的阶段（如：一线治疗 / 一线维持治疗 / 二线治疗）", 
                "duration": "具体时间段", 
                "regimen": "【严禁遗漏】完整保留该阶段所有的治疗措施原文（不仅包含化疗/靶向/免疫等全身用药，必须包含该阶段发生的放疗、手术、消融等局部治疗原文）", 
                "imaging": "关键影像学评估结果原文保留（必须注明是PR, SD还是PD，以及具体的病灶变化描述）",
                "markers": "肿瘤标志物变化情况原文保留（如CA19-9, CEA等的起伏，若原文未提及则写'未提及'）"
            }
        ],
        "timeline_events": [
            {
                "date": "年月", 
                "event_type": "必须填 'Treatment' 或 'Evaluation'",
                "event": "若是Treatment，填具体方案(如'一线:AG+百泽安'或'局部放疗')；若是Evaluation，填疗效(如'肺部PD'或'维持SD')"
            }
        ],
        "summary": ["基于原文提炼的治疗亮点总结1", "基于原文提炼的治疗亮点总结2"]
    }
    注意：timeline_events 需提取全病程中最重要的【治疗换线节点】、【局部重大治疗节点（如放疗/手术）】和【影像学评估节点】，按时间先后排序，最多不超过8个。
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
# 3. PPT 生成模块 (适配中大系主题色与学术排版)
# ==========================================
class AdvancedPPTMaker:
    def __init__(self, data):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333) 
        self.prs.slide_height = Inches(7.5)
        self.data = data
        
        # 换成了类似中山一院院徽的紫红色 (Burgundy/Maroon) 作为主色调
        self.C_PRI = RGBColor(115, 21, 40)   
        # 辅助色用沉稳的深蓝色
        self.C_ACC = RGBColor(0, 51, 102)  

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
        
        # 按照模板结构拼接
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
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide, f"治疗经过：{tx.get('phase', '阶段治疗')}")
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(6))
            tf = tb.text_frame
            tf.word_wrap = True 
            
            p1 = tf.paragraphs[0]
            p1.text = f"【治疗时间】 {tx.get('duration', '')}"
            p1.font.size = Pt(20) 
            p1.font.bold = True
            p1.font.color.rgb = self.C_PRI
            
            p2 = tf.add_paragraph()
            p2.text = f"\n【用药方案】\n{tx.get('regimen', '')}"
            p2.font.size = Pt(16) 
            
            p3 = tf.add_paragraph()
            p3.text = f"\n【影像学评估】\n{tx.get('imaging', '')}"
            p3.font.size = Pt(16) 
            p3.font.color.rgb = RGBColor(50, 50, 50)
            
            p4 = tf.add_paragraph()
            p4.text = f"\n【肿瘤标志物】\n{tx.get('markers', '')}"
            p4.font.size = Pt(16) 
            p4.font.color.rgb = self.C_ACC

    def make_timeline(self):
        """专业版时间轴：分离治疗节点与评估节点"""
        events = self.data.get("timeline_events", [])
        if not events: return
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "全病程时间轴概览 (Timeline)")
        
        line_y = Inches(4.2)
        main_line = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(0.5), line_y - Inches(0.05), Inches(12.3), Inches(0.1))
        main_line.fill.solid()
        main_line.fill.fore_color.rgb = RGBColor(220, 220, 220) 
        main_line.line.fill.background()
        
        start_x = Inches(1.0)
        interval = Inches(11.0 / max(len(events), 1)) 
        
        for i, evt in enumerate(events[:8]): 
            x = start_x + (i * interval)
            event_text = evt.get("event", "")
            event_type = evt.get("event_type", "Treatment")
            
            # 智能判断颜色：如果是PD/进展标红；如果是PR/SD标绿；如果是治疗方案则用主色调
            is_pd = "进展" in event_text or "PD" in event_text.upper() or "复发" in event_text
            is_control = "PR" in event_text.upper() or "SD" in event_text.upper() or "缩小" in event_text
            
            if is_pd:
                node_color = RGBColor(220, 50, 50) # 警示红
            elif is_control and event_type == "Evaluation":
                node_color = RGBColor(46, 139, 87) # 稳定绿
            else:
                node_color = self.C_PRI # 治疗紫红
            
            # 上下交错防止重叠
            stem_top = line_y - Inches(1.2) if i % 2 == 0 else line_y
            stem_height = Inches(1.2)
            stem = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.13), stem_top, Inches(0.04), stem_height)
            stem.fill.solid()
            stem.fill.fore_color.rgb = node_color
            stem.line.fill.background()
            
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, line_y - Inches(0.15), Inches(0.3), Inches(0.3))
            circle.fill.solid()
            circle.fill.fore_color.rgb = node_color
            circle.line.color.rgb = RGBColor(255, 255, 255) 
            circle.line.width = Pt(2)
            
            card_top = line_y - Inches(2.4) if i % 2 == 0 else line_y + Inches(1.2)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x - Inches(0.7), card_top, Inches(1.6), Inches(1.2))
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(250, 250, 250) 
            card.line.color.rgb = node_color 
            card.line.width = Pt(1.5)
            
            tf = card.text_frame
            tf.word_wrap = True
            
            # 日期
            p0 = tf.paragraphs[0]
            p0.text = evt.get("date", "")
            p0.font.bold = True
            p0.font.size = Pt(11)
            p0.font.color.rgb = node_color
            p0.alignment = PP_ALIGN.CENTER
            
            # 标签：区分是【评估】还是【方案】
            p_tag = tf.add_paragraph()
            p_tag.text = "【评估】" if event_type == "Evaluation" else "【方案】"
            p_tag.font.size = Pt(9)
            p_tag.font.bold = True
            p_tag.font.color.rgb = node_color
            p_tag.alignment = PP_ALIGN.CENTER
            
            # 事件内容
            p1 = tf.add_paragraph()
            p1.text = event_text
            p1.font.size = Pt(10)
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
                    label="📥 立即下载 PPT (含完整细节保留)",
                    data=ppt_file,
                    file_name="病例汇报_多图连拍版.pptx",
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

