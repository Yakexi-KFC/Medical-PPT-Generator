# 肿瘤病例 PPT 生成工作台

这是一个 Streamlit 应用，用于将病历图片 OCR 或电子病历文本转换为结构化病例资料，并自动生成适合病例讨论、科室汇报和学术交流的 PowerPoint。

## 功能

- 病历截图批量上传与百度 OCR 提取
- 电子病历文本粘贴输入
- DeepSeek 模型结构化解析病例、推断治疗线数与时间轴
- 自动生成 PPTX 病例汇报
- 页面端展示病例全病程逻辑线与底层 JSON

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署配置

部署到 Streamlit Cloud 或同类平台时，需要在 Secrets 中配置：

```toml
BAIDU_API_KEY = "..."
BAIDU_SECRET_KEY = "..."
DEEPSEEK_API_KEY = "..."
```
