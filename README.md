# PPTX 中英双向翻译 Skill

> 将中文 PowerPoint 翻译为英文，或将英文翻译为中文。保持原有排版、样式、图片不变，自动替换字体，段落级翻译避免碎片化。
>
> 支持 CodeBuddy、Claude Code、Cursor、Cline、Windsurf 等主流 AI 编程智能体。

## 📦 安装

### 通用前置依赖

无论用哪个智能体，都需要先安装 Python 依赖：

```bash
pip install python-pptx
pip install "markitdown[pptx]"  # 可选，用于预览 PPT 内容
```

### 最简单的方式：告诉你的 AI 助手

直接拷贝下面这句话发给你的 AI 助手：

> 帮我安装 PPTX 中英双向翻译 Skill，仓库地址：https://github.com/smiler488/pptx-zh2en

Agent 会自动克隆仓库并安装到对应的 Skill 目录。

### 手动安装

将本仓库克隆到你项目下的 Skill 目录，不同 IDE 对应的路径：

| IDE | Skill 目录 |
|-----|-----------|
| CodeBuddy | `.codebuddy/skills/pptx-zh2en/` |
| Claude Code | `.claude/skills/pptx-zh2en/` |
| Cursor | `.cursor/skills/pptx-zh2en/` |
| Windsurf | `.windsurf/skills/pptx-zh2en/` |
| Cline | `.cline/skills/pptx-zh2en/` |
| 通用 | `.agents/skills/pptx-zh2en/` |

```bash
# 示例：安装到 Cursor
git clone https://github.com/smiler488/pptx-zh2en.git \
  .cursor/skills/pptx-zh2en
```

只要目录下有 `SKILL.md`，Agent 下次启动就会自动加载这个 Skill。

### CodeBuddy 插件市场安装

```bash
/plugin marketplace add smiler488/pptx-zh2en
/plugin install pptx-zh2en@pptx-zh2en-marketplace
```

### 独立使用（脱离任何智能体）

脚本本身是纯 Python，不依赖任何智能体平台，可以独立运行：

```bash
pip install python-pptx

# 中→英
# Step 1: 提取文本
python scripts/translate_pptx_inline.py --mode extract --direction zh2en -i chinese.pptx -t trans.json
# Step 2: 编辑 trans.json，填写 translated 字段（人工或用任意翻译工具）
# Step 3: 写回翻译
python scripts/translate_pptx_inline.py --mode apply --direction zh2en -i chinese.pptx -t trans.json -o english.pptx

# 英→中
# Step 1: 提取文本
python scripts/translate_pptx_inline.py --mode extract --direction en2zh -i english.pptx -t trans.json
# Step 2: 编辑 trans.json，填写 translated 字段
# Step 3: 写回翻译
python scripts/translate_pptx_inline.py --mode apply --direction en2zh -i english.pptx -t trans.json -o chinese.pptx
```

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 🔄 **双向翻译** | 支持 `zh2en`（中→英）和 `en2zh`（英→中）两个方向 |
| 📝 **段落级提取翻译** | 以段落为单位提取和翻译，避免 run 碎片化导致翻译断裂 |
| 🔤 **三层字体替换** | 精确映射 → 启发式判断 → 主题字体修改，适配所有字体 |
| 🧹 **标点自动清洗** | 翻译后自动将源语言标点替换为目标语言标点 |
| 🔍 **翻译后自动质检** | 自动检测残留源语言字符/标点/重复词/缺失空格 |
| ⚠️ **遗漏翻译检测** | apply 时检测 `translated == original` 的段落，防止漏译 |
| 📐 **文本框自动适配** | 保留原有 autofit 类型，仅对无设置的文本框开启 normAutofit |
| ⚡ **双语重复检测** | 检测原 PPT 中已有的翻译，自动去重 |
| 🎨 **保留全部格式** | 字号、颜色、加粗、布局、图片均不改动 |

## 🚀 使用方法

### 标准流程（AI 内联翻译 — 推荐）

```bash
# Step 1: 提取文本（段落级）
python scripts/translate_pptx_inline.py \
  --mode extract --direction zh2en \
  --input chinese.pptx \
  --translations-file translations.json

# Step 2: 编辑 translations.json，填写 translated 字段
# （由 AI 或人工完成翻译）

# Step 3: 写入翻译 + 字体替换 + 自动适配
python scripts/translate_pptx_inline.py \
  --mode apply --direction zh2en \
  --input chinese.pptx \
  --translations-file translations.json \
  --output english.pptx
```

### 外部 API 全自动模式

如果有 OpenAI 兼容 API Key，可一步完成提取+翻译+写回：

```bash
export OPENAI_API_KEY="sk-..."
export OPENAI_BASE_URL="https://api.openai.com/v1"

# 中→英（默认）
python scripts/translate_pptx.py \
  --input chinese.pptx \
  --output english.pptx \
  --model gpt-4o-mini

# 英→中
python scripts/translate_pptx.py \
  --direction en2zh \
  --input english.pptx \
  --output chinese.pptx \
  --model gpt-4o-mini
```

### 英→中方向

```bash
# 提取
python scripts/translate_pptx_inline.py \
  --mode extract --direction en2zh \
  --input english.pptx \
  --translations-file translations.json

# 写入
python scripts/translate_pptx_inline.py \
  --mode apply --direction en2zh \
  --input english.pptx \
  --translations-file translations.json \
  --output chinese.pptx
```

## 📁 项目结构

```
pptx-zh2en/
├── SKILL.md                          # Skill 说明文件（跨平台通用，Agent 自动识别）
├── skill.json                        # 机器可读配置（名称、版本、关键词）
├── scripts/
│   ├── translate_pptx_inline.py      # 主脚本（内联模式）
│   └── translate_pptx.py             # 备选脚本（API 模式）
├── .codebuddy-plugin/
│   └── marketplace.json              # CodeBuddy 插件市场清单
├── README.md
├── LICENSE
└── .gitignore
```

> **SKILL.md 是跨智能体通用标准**：CodeBuddy、Claude Code、Cursor、Cline、Windsurf 等都通过识别 SKILL.md 的 YAML frontmatter（`name`、`description`、`keywords`）来发现和触发技能。

## 🌾 专业术语表

内置三个领域的术语表：

- **农业/遥感**：数字表型、冠层、无人机遥感、光合速率、叶面积指数、株高、生物量等
- **AI/计算机科学**：智能体、大模型、工作记忆、情景记忆、向量数据库、知识图谱等
- **通用学术**：研究背景、材料与方法、主要结果、结论与展望等

## 🔧 字体替换规则

### 中→英方向

| 原中文字体 | → 替换为 |
|-----------|----------|
| 微软雅黑 / Microsoft YaHei | Calibri |
| 宋体 / SimSun | Times New Roman |
| 黑体 / SimHei | Arial |
| 思源黑体 / Noto Sans CJK | Arial |
| 思源宋体 / Noto Serif CJK | Times New Roman |
| 其他中文字体 | 启发式判断（含"宋"→Times，含"黑"→Arial） |

### 英→中方向

| 原英文字体 | → 替换为 |
|-----------|----------|
| Calibri | 微软雅黑 |
| Times New Roman | 宋体 |
| Arial | 黑体 |

## ⚠️ 已知限制

| 限制 | 说明 |
|------|------|
| 图片内嵌文字 | 无法处理，需手动修改 |
| 图表数据系列名 | 部分不在文本框里，需手动改 |
| SmartArt 文字 | 支持有限，可能遗漏 |
| .ppt 格式 | 只支持 .pptx |
| 文本溢出 | 英文比中文长 20-40%，建议检查文本框 |

## 📄 许可证

MIT License — 详见 [LICENSE](LICENSE)
