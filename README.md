# PPTX 中英双向翻译 Skill

> 将中文 PowerPoint 翻译为英文，或将英文翻译为中文。保持原有排版、样式、图片不变，自动替换字体，段落级翻译避免碎片化。

## 📦 安装

### 方式一：CodeBuddy 插件市场安装（推荐）

在 CodeBuddy 中执行以下命令：

```bash
# 1. 添加插件市场
/plugin marketplace add smiler488/pptx-zh2en

# 2. 安装插件
/plugin install pptx-zh2en@pptx-zh2en-marketplace
```

安装 Python 依赖：

```bash
pip install python-pptx
```

### 方式二：手动安装（git clone）

```bash
# 克隆仓库到 CodeBuddy skills 目录
git clone https://github.com/smiler488/pptx-zh2en.git ~/.codebuddy/skills/pptx-zh2en

# 安装依赖
pip install python-pptx
```

> 手动安装时，请将 `skills/pptx-zh2en/` 目录下的内容复制到 `~/.codebuddy/skills/pptx-zh2en/`。

### 方式三：独立使用（脱离 CodeBuddy）

脚本可以独立运行，只需安装 `python-pptx`：

```bash
pip install python-pptx
python scripts/translate_pptx_inline.py --mode extract --direction zh2en -i input.pptx -t trans.json
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

### 标准流程（CodeBuddy 内联翻译 — 推荐）

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

如果有多 OpenAI 兼容 API Key，可一步完成提取+翻译+写回：

```bash
export OPENAI_API_KEY="sk-..."
export OPENAI_BASE_URL="https://api.openai.com/v1"

python scripts/translate_pptx.py \
  --input chinese.pptx \
  --output english.pptx \
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
├── .codebuddy-plugin/
│   └── marketplace.json              # 插件市场清单
├── plugins/
│   └── pptx-zh2en/
│       ├── .codebuddy-plugin/
│       │   └── plugin.json           # 插件清单
│       └── skills/
│           └── pptx-zh2en/
│               ├── SKILL.md          # Skill 说明文件
│               └── scripts/
│                   ├── translate_pptx_inline.py  # 主脚本（内联模式）
│                   └── translate_pptx.py         # 备选脚本（API 模式）
├── README.md
├── LICENSE
└── .gitignore
```

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
