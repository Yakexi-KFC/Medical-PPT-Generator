import streamlit as st

from app_workflow import render_workflow

st.set_page_config(page_title="肿瘤病例 PPT 生成工作台", layout="wide")

def inject_frontend_style():
    st.markdown(
        """
        <style>
        :root {
            --primary: #b00086;
            --primary-dark: #7e0067;
            --accent: #ec008c;
            --ink: #25232a;
            --muted: #6f6b76;
            --line: #ded9e1;
            --panel: #ffffff;
            --soft: #f7f5f8;
            --warning-bg: #fff4fa;
            --warning-line: #f3a4d1;
            --shadow: 0 18px 45px rgba(70, 58, 76, 0.10);
        }

        .stApp {
            background: linear-gradient(180deg, #faf8fb 0%, #f0edf2 44%, #fbfafc 100%);
            color: var(--ink);
        }

        .block-container {
            max-width: 1180px;
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        h1, h2, h3, p, label, span {
            letter-spacing: 0;
        }

        div[data-testid="stTabs"] button {
            min-height: 48px;
            border-radius: 8px 8px 0 0;
            font-weight: 800;
        }

        div[data-testid="stTabs"] button[aria-selected="true"] {
            color: var(--primary);
            border-color: var(--line);
            background: #ffffff;
        }

        .stButton > button,
        .stDownloadButton > button {
            width: 100%;
            min-height: 46px;
            border: 1px solid var(--primary);
            border-radius: 8px;
            background: var(--primary);
            color: #ffffff;
            font-weight: 850;
            box-shadow: 0 10px 20px rgba(176, 0, 134, 0.16);
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: var(--primary-dark);
            background: var(--primary-dark);
            color: #ffffff;
        }

        .stTextArea textarea,
        .stFileUploader section {
            border-radius: 8px;
        }

        .case-hero {
            position: relative;
            overflow: hidden;
            border: 1px solid rgba(122, 31, 54, 0.14);
            border-radius: 8px;
            padding: 34px 38px 30px;
            background:
                linear-gradient(135deg, rgba(176, 0, 134, 0.95), rgba(236, 0, 140, 0.88)),
                repeating-linear-gradient(90deg, rgba(255,255,255,0.10) 0, rgba(255,255,255,0.10) 1px, transparent 1px, transparent 22px);
            color: #fff;
            box-shadow: 0 22px 55px rgba(86, 0, 74, 0.15);
            margin-bottom: 18px;
        }

        .case-hero::after {
            content: "";
            position: absolute;
            right: -20px;
            bottom: -44px;
            width: 260px;
            height: 150px;
            border: 1px solid rgba(255,255,255,0.28);
            border-radius: 50%;
            transform: rotate(-15deg);
        }

        .case-kicker {
            position: relative;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 5px 10px;
            border: 1px solid rgba(255,255,255,0.22);
            border-radius: 6px;
            background: rgba(255,255,255,0.14);
            font-size: 13px;
            font-weight: 700;
        }

        .case-hero h1 {
            position: relative;
            max-width: 780px;
            margin: 16px 0 10px;
            color: #fff;
            font-size: clamp(30px, 4.2vw, 48px);
            line-height: 1.08;
            font-weight: 850;
        }

        .case-hero p {
            position: relative;
            max-width: 800px;
            margin: 0;
            color: rgba(255,255,255,0.86);
            font-size: 16px;
            line-height: 1.75;
        }

        .case-release {
            position: relative;
            z-index: 1;
            margin-top: 24px;
            padding-top: 18px;
            border-top: 1px solid rgba(255,255,255,0.28);
        }

        .case-release-label {
            display: flex;
            flex-wrap: wrap;
            align-items: baseline;
            gap: 10px;
            margin-bottom: 12px;
            font-size: 15px;
            font-weight: 700;
        }

        .case-release-label time {
            font-size: 13px;
            font-weight: 400;
            color: rgba(255,255,255,0.8);
        }

        .case-release-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 22px;
        }

        .case-release-item strong {
            display: block;
            margin-bottom: 5px;
            color: #ffffff;
            font-size: 15px;
        }

        .case-hero .case-release-item p {
            max-width: none;
            font-size: 14px;
            line-height: 1.7;
            color: rgba(255,255,255,0.9);
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 12px;
            margin: 18px 0 24px;
        }

        .metric-card {
            min-height: 86px;
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 14px 15px;
            background: rgba(255,255,255,0.88);
            box-shadow: 0 12px 28px rgba(70, 58, 76, 0.06);
        }

        .metric-label {
            margin: 0 0 5px;
            color: var(--muted);
            font-size: 12px;
            font-weight: 700;
        }

        .metric-value {
            margin: 0;
            color: var(--ink);
            font-size: 18px;
            font-weight: 800;
        }

        .section-card {
            border: 1px solid var(--line);
            border-left: 5px solid var(--primary);
            border-radius: 8px;
            padding: 18px 20px;
            margin: 16px 0 14px;
            background: var(--panel);
            box-shadow: var(--shadow);
        }

        .section-card.accent {
            border-left-color: var(--accent);
        }

        .section-kicker {
            margin: 0 0 6px;
            color: var(--primary);
            font-size: 12px;
            font-weight: 800;
        }

        .section-card.accent .section-kicker {
            color: var(--accent);
        }

        .section-title {
            margin: 0 0 6px;
            color: var(--ink);
            font-size: 22px;
            font-weight: 850;
        }

        .section-copy {
            margin: 0;
            color: var(--muted);
            font-size: 14px;
            line-height: 1.7;
        }

        .guideline-panel {
            border: 1px solid var(--warning-line);
            border-radius: 8px;
            padding: 15px 17px;
            margin: 4px 0 18px;
            background: var(--warning-bg);
            color: #6b3154;
            line-height: 1.75;
        }

        .footer-note {
            margin-top: 24px;
            padding: 13px 16px;
            border-top: 1px solid var(--line);
            color: var(--muted);
            font-size: 13px;
            text-align: center;
        }

        @media (max-width: 900px) {
            .metric-grid {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            .case-release-grid {
                grid-template-columns: 1fr;
                gap: 14px;
            }

            .case-hero {
                padding: 28px 22px;
            }

        }

        @media (max-width: 620px) {
            .metric-grid {
                grid-template-columns: 1fr;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_workbench_header():
    st.markdown(
        """
        <section class="case-hero">
          <div class="case-kicker">Case Slide Studio · Oncology Workflow</div>
          <h1>肿瘤病例 PPT 自动生成工作台</h1>
          <p>识别影像、病理与基因检测报告，裁剪关键区域，按病程整理为图文病例幻灯。</p>
          <section class="case-release" aria-label="本次更新说明">
            <div class="case-release-label">本次更新 <time datetime="2026-10-22">2026.10.22</time></div>
            <div class="case-release-grid">
              <div class="case-release-item">
                <strong>图片识别与裁剪</strong>
                <p>通过 DeepSeek API 识别影像、病理及基因检测资料，支持逐图核对和裁剪调整。</p>
              </div>
              <div class="case-release-item">
                <strong>标准病例排版</strong>
                <p>支持保留完整报告主体和病理图组，正文以叙述和分点为主，相近检查集中呈现。</p>
              </div>
              <div class="case-release-item">
                <strong>可编辑诊疗时间轴</strong>
                <p>按已记录日期梳理诊疗经过，区分已实施治疗与后续计划。</p>
              </div>
            </div>
          </section>
        </section>
        <section class="metric-grid" aria-label="功能概览">
          <article class="metric-card">
            <p class="metric-label">入口</p>
            <p class="metric-value">图片 / 病历</p>
          </article>
          <article class="metric-card">
            <p class="metric-label">结构化</p>
            <p class="metric-value">检查与病程</p>
          </article>
          <article class="metric-card">
            <p class="metric-label">输出</p>
            <p class="metric-value">PPTX 幻灯片</p>
          </article>
          <article class="metric-card">
            <p class="metric-label">附加</p>
            <p class="metric-value">裁剪与配图</p>
          </article>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_section_card(kicker, title, copy, accent=False):
    accent_class = " accent" if accent else ""
    st.markdown(
        f"""
        <article class="section-card{accent_class}">
          <p class="section-kicker">{kicker}</p>
          <h2 class="section-title">{title}</h2>
          <p class="section-copy">{copy}</p>
        </article>
        """,
        unsafe_allow_html=True,
    )



inject_frontend_style()
render_workbench_header()
render_workflow()
