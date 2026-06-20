---
name: pptx-zh2en
description: "PPT中英双向翻译。将中文PPT翻译为英文，或将英文PPT翻译为中文。保持排版、样式、图片不变，自动替换字体、段落级翻译、标点清洗、自动质检。"
version: "2.4.0"
author: "WorkBuddy"
created: "2026-06-19"
updated: "2026-06-20"
alwaysApply: false
keywords:
  - PPT翻译
  - pptx翻译
  - 中文PPT翻译
  - 英文PPT翻译
  - translate ppt
  - ppt translation
  - chinese to english ppt
  - english to chinese ppt
  - 中英文PPT转换
  - PPT英文版
  - PPT中文版
  - 学术PPT翻译
  - 字体替换
  - 段落级翻译
---

# PPTX 中英双向翻译 Skill

## 功能说明

将 PowerPoint (.pptx) 文件在中英文之间双向翻译：

| 功能 | 说明 |
|------|------|
| 🔄 **双向翻译** | 支持 `--direction zh2en`（中→英）和 `en2zh`（英→中）两个方向 |
| 📝 **段落级提取翻译** | 以段落为单位提取和翻译，避免 run 碎片化导致翻译断裂 |
| 🔤 **三层字体替换** | 精确映射 → 启发式判断 → 主题字体修改，适配所有字体 |
| 🧹 **标点自动清洗** | 翻译后自动将源语言标点替换为目标语言标点 |
| 🔍 **翻译后自动质检** | 自动检测残留源语言字符/标点/重复词/缺失空格 |
| ⚠️ **遗漏翻译检测** | apply 时检测 `translated == original` 的段落，防止漏译 |
| 📐 **文本框自动适配** | 保留原有 autofit 类型，仅对无设置的文本框开启换行+normAutofit |
| ⚡ **双语重复检测** | 检测原 PPT 中已有的翻译，避免重复 |
| 🎨 **保留全部格式** | 字号、颜色、加粗、布局、图片均不改动 |

---

## 触发场景

| 用户可能会说 | 执行方向 |
|-------------|---------|
| "帮我把这个PPT翻译成英文" / "PPT做个英文版" | zh2en（中→英） |
| "把这个英文PPT翻译成中文" / "PPT做个中文版" | en2zh（英→中） |
| "翻译这个PPT" / "PPT翻译一下" | 根据内容自动判断方向 |

---

## 工作流（标准流程 — 推荐）

> **依赖**：需安装 `pip install python-pptx`，可选安装 `pip install "markitdown[pptx]"` 用于预览 PPT 内容。

### 第一步：确认输入文件 + 领域识别

查找用户提供的 PPT 文件，用 markitdown 预览内容，**判断 PPT 所属领域**：

```bash
python3 -m markitdown input.pptx
```

**领域判断**：根据内容关键词自动识别，选择对应术语表：
- **农业/遥感**：棉花、玉米、冠层、表型、光合、无人机、叶面积指数 → 用农业术语表
- **AI/计算机科学**：智能体、大模型、记忆、检索、向量数据库、深度学习 → 用 AI 术语表
- **通用学术**：其他领域 → 通用翻译规范

### 第二步：提取文本（段落级）

> 脚本位于 skill 目录下的 `scripts/` 子目录，使用相对路径 `scripts/translate_pptx_inline.py` 调用。

```bash
# 中→英（默认）
python3 scripts/translate_pptx_inline.py \
  --mode extract --direction zh2en \
  --input input.pptx \
  --translations-file translations.json

# 英→中
python3 scripts/translate_pptx_inline.py \
  --mode extract --direction en2zh \
  --input input.pptx \
  --translations-file translations.json
```

输出 JSON 格式（每条 = 一个段落）：

```json
[
  {
    "id": 0,
    "slide": 1,
    "shape_name": "TextBox 15",
    "shape_id": 16,
    "sub_path": "body",
    "para_idx": 0,
    "run_count": 3,
    "location": "slide",
    "original": "搜索引擎、代码运行、API调用结果",
    "translated": "",
    "font_name": "Microsoft YaHei UI",
    "has_existing_english": false,
    "existing_english": ""
  }
]
```

> **⚡ 双语提示**：`has_existing_english=true` 的条目，`existing_english` 字段已提取出原文中的英文，可直接填入 `translated`。

### 第三步：AI 翻译所有文本

**你（Agent）直接翻译所有条目，填写 `translated` 字段。**

> ⚠️ **关键要求：必须用 JSON 中每条的 `original` 字段完整文本进行翻译，不要手动缩写或从预览中复制。**
> 用 `id` 字段作为索引逐条处理，确保 100% 覆盖，不遗漏任何一条。

翻译要求：
- 以**段落**为单位翻译，保证语义完整
- 标题使用 **Title Case**（zh2en 方向）
- 保留数字、公式、单位符号不变
- 学术引用（如 `Monteith et al., 1977`）保留原文不翻译
- 文件路径保留原文不翻译
- `has_existing_english=true` 的条目，优先使用 `existing_english`
- 参考下方术语表保证专业词汇准确
- 已是目标语言的段落（如备注页已是中文），直接复制到 `translated`

### 第四步：写入翻译 + 生成 PPT

```bash
# 中→英
python3 scripts/translate_pptx_inline.py \
  --mode apply --direction zh2en \
  --input input.pptx \
  --translations-file translations.json \
  --output output_EN.pptx

# 英→中
python3 scripts/translate_pptx_inline.py \
  --mode apply --direction en2zh \
  --input input.pptx \
  --translations-file translations.json \
  --output output_ZH.pptx
```

此步骤自动完成：
1. 段落级文本替换（翻译写入首 run，清空其余 run）
2. 中文标点清洗
3. 三层字体替换
4. **文本框自动适配**（保留原有 spAutoFit/normAutofit，仅对无设置的开启换行+normAutofit）
5. **自动质检**（残留中文/标点/重复词/缺失空格）

### 第五步：查看质检结果

apply 完成后会自动输出质检报告。如果发现问题，修正 JSON 后重新执行第四步。

---

## 备选方案（外部 API 全自动模式）

```bash
export OPENAI_API_KEY="sk-..."
# 中→英（默认）
python3 scripts/translate_pptx.py \
  --input input.pptx \
  --output output_EN.pptx \
  --model gpt-4o-mini

# 英→中
python3 scripts/translate_pptx.py \
  --direction en2zh \
  --input input.pptx \
  --output output_ZH.pptx \
  --model gpt-4o-mini
```

---

## 翻译质量标准

| 文本类型 | 规范 |
|----------|------|
| 幻灯片标题 | Title Case，简洁有力 |
| 正文要点 | 流畅学术英文，保持原意 |
| 表格内容 | 与正文术语一致 |
| 备注页 | 完整翻译，保留演讲口吻 |
| 标点符号 | 全部使用英文标点（, . ; : ? !） |
| 专有名词 | 使用下表标准译名 |

---

## 专业术语表

### 农业 / 遥感领域

| 中文 | 标准英文 |
|------|---------|
| 数字表型 / 数字表型组学 | digital phenotyping / digital phenomics |
| 冠层 | canopy |
| 无人机遥感 | UAV remote sensing |
| 光合速率 / 净光合速率 | photosynthetic rate / net photosynthetic rate |
| 叶面积指数 | leaf area index (LAI) |
| 株高 | plant height |
| 生物量 | biomass |
| 高产栽培 | high-yield cultivation |
| 数字孪生 | digital twin |
| 点云 | point cloud |
| 3D重建 / 三维重建 | 3D reconstruction |
| 高光谱 / 多光谱 | hyperspectral / multispectral |
| 归一化植被指数 | NDVI |
| 图像分割 | image segmentation |

### AI / 计算机科学领域

| 中文 | 标准英文 |
|------|---------|
| 智能体 | agent |
| 大模型 / 大语言模型 | large language model (LLM) |
| 工作记忆 | working memory |
| 情景记忆 | episodic memory |
| 语义记忆 | semantic memory |
| 程序记忆 | procedural memory |
| 上下文窗口 | context window |
| 检索增强生成 | retrieval-augmented generation (RAG) |
| 向量数据库 | vector database |
| 知识图谱 | knowledge graph |
| 用户画像 | user profile |
| 提示词 | prompt |
| 深度学习 | deep learning |
| 机器学习 | machine learning |
| 多智能体 | multi-agent |
| 自主进化 | self-evolving |
| 持久化 | persistence |
| 标准流程 | standard operating procedure (SOP) |

### 通用学术

| 中文 | 标准英文 |
|------|---------|
| 研究背景 | research background |
| 材料与方法 | materials and methods |
| 主要结果 | key results |
| 结论与展望 | conclusions and future perspectives |
| 国家自然科学基金 | National Natural Science Foundation of China |

---

## 字体替换策略（三层覆盖）

| 层级 | 场景 | 处理方式 |
|------|------|---------|
| ① Run 级已知字体 | `run.font.name` 命中映射表 | 精确替换 |
| ② Run 级未知字体 | 中文字体但不在映射表 | 启发式：含"宋"→Times New Roman，含"黑"→Arial |
| ③ 主题字体 | `run.font.name` 为 None，继承主题 | 修改 theme1.xml + 显式设为 Calibri |

常见映射：微软雅黑→Calibri，宋体→Times New Roman，黑体→Arial，思源黑体→Arial

---

## 已知限制

| 限制 | 说明 |
|------|------|
| 图片内嵌文字 | 无法处理，需手动修改 |
| 图表数据系列名 | 部分不在文本框里，需手动改 |
| SmartArt 文字 | 支持有限，可能遗漏 |
| .ppt 格式 | 只支持 .pptx |
| 文本溢出 | 英文比中文长 20-40%，建议检查 |

---

## 输出

最终文件保存为 `output_EN.pptx`（中→英）或 `output_ZH.pptx`（英→中），文件名可通过 `--output` 参数指定。
