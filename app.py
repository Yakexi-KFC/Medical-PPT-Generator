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
BAIDU_API_KEY = st.secrets["QfpFe95LYcIY5o1crrROWCi3"]
BAIDU_SECRET_KEY = st.secrets["aSvE1enC3zrL7IKCAKABlszyvP7RXYTZ"]
DEEPSEEK_API_KEY = st.secrets["sk-61e2d5846bd34ca5aa14f4fe92482f91"]

# ==========================================
# 1. 百度 OCR 图片识别模块 (支持批量识别)
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
# 2. AI 结构化提取模块 (强化：不删减原意 + 自动推断治疗线数)
# ==========================================
def extract_complex_case(patient_text):
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY, 
        base_url="https://api.deepseek.com"
    )
    system_prompt = """
    你是一位严谨的肿瘤内科主任医师。请阅读用户提供的真实长篇病历，将其拆解为标准的病例汇报结构。
    
    【核心指令与肿瘤内科铁律 - 极其重要】：
    1. 原汁原味：绝不要过度精简，必须尽可能保留原病历中的详细客观描述（如肿瘤大小数值、生化指标、用药剂量）。
    2. 严格的线数划分铁律（必须遵守）：
       - 只有在明确记录【疾病进展（PD）】或【复发】后彻底更改方案，才算开启下一线治疗（如二线、三线）。
       - 如果在未进展（如PR、CR、SD）的情况下，仅仅是停用部分毒副反应大的药物（如化疗），保留或替换免疫/靶向药物进行延续治疗，必须判定为【同一线的维持治疗】（例如：二线未进展时改为百泽安+索凡替尼，严禁称为三线，必须标为“二线维持治疗”）。
       - 手术前后的辅助/新辅助治疗，不计入晚期解救治疗的线数。
    
    必须严格输出为以下 JSON 格式：
    {
        "cover": {"title": "晚期XXX癌综合治疗病例汇报"},
        "baseline": {
            "info": "保留基本信息原文",
            "diagnosis": "保留诊断与分期原文",
            "molecular": "保留基因检测原文"
        },
        "treatments": [
            {
                "phase": "遵守铁律推断的阶段（如：一线治疗 / 一线维持治疗 / 进展后二线治疗）", 
                "duration": "具体时间段", 
                "regimen": "用药方案及调整经过原文", 
                "efficacy": "疗效评估原文"
            }
        ],
        "timeline_events": [
            {"date": "年月", "event": "核心事件摘要（如包含疾病进展，请写明'进展'或'PD'），限15个字内"}
        ],
        "summary": ["基于原文提炼的治疗亮点总结1", "基于原文提炼的治疗亮点总结2"]
    }
    注意：timeline_events 数组最多提取 6 个最重要的节点，按时间先后排序。
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
# 3. PPT 生成模块 (适配海量文字排版)
# ==========================================
class AdvancedPPTMaker:
    def __init__(self, data):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333) 
        self.prs.slide_height = Inches(7.5)
        self.data = data
        self.C_PRI = RGBColor(0, 51, 102)   
        self.C_ACC = RGBColor(0, 153, 153)  

    def add_header(self, slide, text):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(1.0))
        shape.fill.solid()
        shape.fill.fore_color.rgb = self.C_PRI
        shape.line.fill.background()
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(10), Inches(0.8))
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
        p.text = self.data["cover"]["title"]
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    def make_baseline(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "患者基线资料")
        base_data = self.data["baseline"]
        content = f"【基本信息】\n{base_data.get('info', '')}\n\n" \
                  f"【临床诊断】\n{base_data.get('diagnosis', '')}\n\n" \
                  f"【分子病理】\n{base_data.get('molecular', '')}"
        tb = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(11), Inches(6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = content
        p.font.size = Pt(18) # 调小字号容纳大量细节
        
    def make_treatments(self):
        for tx in self.data.get("treatments", []):
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.add_header(slide, f"治疗经过：{tx.get('phase', '阶段治疗')}")
            tb = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(11), Inches(6))
            tf = tb.text_frame
            tf.word_wrap = True 
            
            p1 = tf.paragraphs[0]
            p1.text = f"【治疗时间】 {tx.get('duration', '')}"
            p1.font.size = Pt(20) 
            p1.font.bold = True
            
            p2 = tf.add_paragraph()
            p2.text = f"\n【用药方案及调整经过】\n{tx.get('regimen', '')}"
            p2.font.size = Pt(16) # 调小字号，完美容纳大量保留的原始病历描述
            
            p3 = tf.add_paragraph()
            p3.text = f"\n【疗效评估与随访】\n{tx.get('efficacy', '')}"
            p3.font.size = Pt(16) 
            p3.font.color.rgb = self.C_ACC

    def make_timeline(self):
        """专业版时间轴：带引线、卡片、及语义色彩警示"""
        events = self.data.get("timeline_events", [])
        if not events: return
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "全病程时间轴概览 (Timeline)")
        
        # 1. 画一根带箭头的灰色主轴线
        line_y = Inches(4.2)
        main_line = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(1), line_y - Inches(0.05), Inches(11.3), Inches(0.1))
        main_line.fill.solid()
        main_line.fill.fore_color.rgb = RGBColor(220, 220, 220) # 浅灰主轴
        main_line.line.fill.background()
        
        start_x = Inches(1.5)
        interval = Inches(10 / max(len(events), 1)) 
        
        for i, evt in enumerate(events[:6]): 
            x = start_x + (i * interval)
            event_text = evt.get("event", "")
            
            # 【高级特效】语义识别颜色：如果事件包含“PD/进展/复发”，自动标红！否则用主色调蓝色。
            is_pd = "进展" in event_text or "PD" in event_text.upper() or "复发" in event_text
            node_color = RGBColor(220, 50, 50) if is_pd else self.C_PRI
            
            # 2. 画竖直连接线 (Stem)
            stem_top = line_y - Inches(1.2) if i % 2 == 0 else line_y
            stem_height = Inches(1.2)
            stem = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.13), stem_top, Inches(0.04), stem_height)
            stem.fill.solid()
            stem.fill.fore_color.rgb = node_color
            stem.line.fill.background()
            
            # 3. 画时间轴上的圆点
            circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, line_y - Inches(0.15), Inches(0.3), Inches(0.3))
            circle.fill.solid()
            circle.fill.fore_color.rgb = node_color
            circle.line.color.rgb = RGBColor(255, 255, 255) # 白色描边显得更精致
            circle.line.width = Pt(2)
            
            # 4. 画带有边框的圆角文本卡片
            card_top = line_y - Inches(2.2) if i % 2 == 0 else line_y + Inches(1.2)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x - Inches(0.8), card_top, Inches(1.8), Inches(1.0))
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(250, 250, 250) # 卡片白底
            card.line.color.rgb = node_color # 边框颜色跟随状态
            card.line.width = Pt(1.5)
            
            # 5. 往卡片里填字
            tf = card.text_frame
            tf.word_wrap = True
            
            p0 = tf.paragraphs[0]
            p0.text = evt.get("date", "")
            p0.font.bold = True
            p0.font.size = Pt(12)
            p0.font.color.rgb = node_color
            p0.alignment = PP_ALIGN.CENTER
            
            p1 = tf.add_paragraph()
            p1.text = event_text
            p1.font.size = Pt(11)
            p1.font.color.rgb = RGBColor(50, 50, 50)
            p1.alignment = PP_ALIGN.CENTER

    def make_summary(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.add_header(slide, "病例小结与思考")
        tb = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11), Inches(5))
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
# 4. Streamlit 网页前端 (支持多图批量上传)
# ==========================================
st.set_page_config(page_title="Pro级肿瘤病例PPT生成", layout="wide")
st.title("🩺 医疗级病史 PPT 自动生成排版系统")

tab1, tab2 = st.tabs(["📸 多图连拍识别 (OCR)", "📝 电子病历粘贴"])

if "ocr_result_text" not in st.session_state:
    st.session_state.ocr_result_text = ""

with tab1:
    st.markdown("### 第一步：批量上传病历图片")
    # 核心修改点：加入 accept_multiple_files=True 支持多选图片
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
                    # 循环处理每一张图片
                    for i, file in enumerate(uploaded_files):
                        image_bytes = file.getvalue()
                        text = perform_ocr(image_bytes, token)
                        all_extracted_text.append(f"【第 {i+1} 页提取结果】\n{text}\n")
                    
                    # 拼接所有文字
                    st.session_state.ocr_result_text = "\n".join(all_extracted_text)
            st.success("✅ 文字提取成功！请在下方核对。")

    st.markdown("### 第二步：人工校对与修改")
    st.info("💡 医疗数据容不得马虎，请核对 OCR 识别出的文字（尤其注意多页之间的拼接是否连贯），确认无误后再生成 PPT。")
    
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