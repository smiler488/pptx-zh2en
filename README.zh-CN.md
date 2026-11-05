# pptx-zh2en

[English](README.md) | 中文

一个 Agent Skill，外加两个 Python 脚本，用来把 PowerPoint 在中英文之间互译，字体、排版、图片和表格都保持原样。

我的报告大多要准备中英两个版本。常见的翻译做法是把文字复制出去、再把译文贴回来，这样会把段落拆得零零碎碎，英文版里还会残留中文字体和全角标点。这里改成分工完成：脚本把每个段落连同位置信息提取成 JSON，智能体翻译 JSON，脚本再写回 PPT，同时替换字体、清洗标点并做质检。

## 安装

对 AI 助手说：

```text
请根据 https://raw.githubusercontent.com/smiler488/pptx-zh2en/main/install.md，安装 pptx-zh2en。
```

也可以手动安装：

```bash
git clone --depth 1 https://github.com/smiler488/pptx-zh2en.git ~/.claude/skills/pptx-zh2en
pip install -r ~/.claude/skills/pptx-zh2en/requirements.txt
```

通过 CodeBuddy 插件市场安装：先执行 `/plugin marketplace add smiler488/pptx-zh2en`，再执行 `/plugin install pptx-zh2en@pptx-zh2en-marketplace`。

## 使用

装好 skill 后，附上 `.pptx` 文件，对智能体说"把这个 PPT 翻成英文版"（或中文版）就行。结果保存为 `<文件名>_EN.pptx` 或 `<文件名>_ZH.pptx`。

不用智能体，也可以直接运行脚本：

```bash
# 1. 把段落提取到 JSON
python scripts/translate_pptx_inline.py --mode extract --direction zh2en \
  --input talk.pptx --translations-file talk.json
# 2. 填写每条记录的 "translated" 字段（手动翻译、用翻译工具或大模型都可以）
# 3. 写回 PPT，替换字体、清洗标点并质检
python scripts/translate_pptx_inline.py --mode apply --direction zh2en \
  --input talk.pptx --translations-file talk.json --output talk_EN.pptx
```

通过任意 OpenAI 兼容接口全自动翻译：

```bash
export OPENAI_API_KEY=...          # 用其他服务商时同时设置 OPENAI_BASE_URL
python scripts/translate_pptx.py --direction zh2en --input talk.pptx --output talk_EN.pptx --model gpt-4o-mini
```

## 能处理什么

- 文本框、占位符、表格、组合形状和备注页。
- 每个段落对应 JSON 里的一条记录，带幻灯片编号和形状路径；PPT 里已有的双语内容会被识别出来，避免重复翻译。
- 字体分三层替换：常见中文字体按映射表替换（微软雅黑 → Calibri，宋体 → Times New Roman，黑体 → Arial）；不在表里的中文字体按名称推断；主题字体直接改主题 XML。
- 全角标点转换成目标语言的标点。漏译的段落、残留的源语言字符、重复单词和缺失的空格都会报告出来。
- 列出可能溢出的文本框（英文通常比中文长 20–40%）。
- `SKILL.md` 里附有农业/遥感、AI 和通用学术三类术语表。

局限：

- 图片里的文字、部分图表标签和 SmartArt 无法翻译。
- 只支持 `.pptx`，不支持 `.ppt`。
- 同一段落里的混合格式（比如只有一个词加粗）会统一成该段第一个 run 的格式。

## 本次修订

- API 模式做"英译中"时会直接报错（这个脚本里没有定义 `contains_english`），已修复。
- SKILL.md 原来用相对路径调用脚本，智能体在用户的工作目录里运行时会找不到脚本；现在改用 skill 的安装路径。
- frontmatter 精简为标准字段，作者从工具名改为我本人，新增 `install.md` 和 `requirements.txt`，删掉没用到的 import。

已在一份包含标题、要点、表格和备注的测试 PPT 上测过中译英和英译中两个方向。

## 许可证

MIT，见 [LICENSE](LICENSE)。邓良超（Liangchao Deng）。
