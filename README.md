# pptx-zh2en

English | [中文](README.zh-CN.md)

An Agent Skill and two Python scripts that translate PowerPoint files between Chinese and English. Fonts, layout, images and tables stay in place.

I give most of my talks in both languages. The common way of translating slides (copy text out, paste translation back) breaks paragraphs into fragments and leaves Chinese fonts and full-width punctuation in the English version. Here the work is split instead: the script extracts every paragraph to JSON with its location, the assistant translates the JSON, and the script writes it back, swaps fonts, fixes punctuation and checks the result.

## Install

Tell your assistant:

```text
Install pptx-zh2en following https://raw.githubusercontent.com/smiler488/pptx-zh2en/main/install.md
```

Or by hand:

```bash
git clone --depth 1 https://github.com/smiler488/pptx-zh2en.git ~/.claude/skills/pptx-zh2en
pip install -r ~/.claude/skills/pptx-zh2en/requirements.txt
```

CodeBuddy marketplace: `/plugin marketplace add smiler488/pptx-zh2en`, then `/plugin install pptx-zh2en@pptx-zh2en-marketplace`.

## Use

With the skill installed, say "translate this deck into English" (or into Chinese) and attach the `.pptx`. The output is saved as `<name>_EN.pptx` or `<name>_ZH.pptx`.

Without an assistant, run the scripts directly:

```bash
# 1. extract paragraphs to JSON
python scripts/translate_pptx_inline.py --mode extract --direction zh2en \
  --input talk.pptx --translations-file talk.json
# 2. fill the "translated" field of every entry (by hand, a translation tool or an LLM)
# 3. write back, replace fonts, clean punctuation, run the checks
python scripts/translate_pptx_inline.py --mode apply --direction zh2en \
  --input talk.pptx --translations-file talk.json --output talk_EN.pptx
```

Fully automatic mode through any OpenAI-compatible API:

```bash
export OPENAI_API_KEY=...          # and OPENAI_BASE_URL for other providers
python scripts/translate_pptx.py --direction zh2en --input talk.pptx --output talk_EN.pptx --model gpt-4o-mini
```

## What it handles

- Text boxes, placeholders, tables, grouped shapes and speaker notes.
- One JSON entry per paragraph, with slide number and shape path. Existing bilingual lines are detected so they are not translated twice.
- Fonts in three layers: known Chinese fonts are mapped (Microsoft YaHei → Calibri, SimSun → Times New Roman, SimHei → Arial); unknown CJK fonts are mapped by name; theme fonts are changed in the theme XML.
- Full-width punctuation converted for the target language. Untranslated paragraphs, leftover source characters, doubled words and missing spaces are reported.
- Text boxes that are likely to overflow are listed (English is usually 20–40 % longer).
- Glossaries for agriculture / remote sensing, AI and general academic terms are in `SKILL.md`.

Limits: text inside images, some chart labels and SmartArt are not translated. Only `.pptx` is supported, not `.ppt`. Mixed formatting inside one paragraph (one bold word, for example) takes the format of the paragraph's first run.

## Changes in this revision

- API mode crashed for English → Chinese (`contains_english` was not defined in that script).
- SKILL.md called the scripts with a relative path, which fails when the assistant runs in the user's folder. It now uses the skill's install path.
- Frontmatter reduced to standard keys. Author changed from a tool name to me. Added `install.md` and `requirements.txt`. Removed unused imports.

Tested on a generated deck with title, bullet points, a table and notes in both directions.

## License

MIT, see [LICENSE](LICENSE). Liangchao Deng (邓良超).
