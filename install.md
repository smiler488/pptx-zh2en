# pptx-zh2en 安装说明（给 AI 助手）

> 用户只需对 AI 助手说：
> `请根据 https://raw.githubusercontent.com/smiler488/pptx-zh2en/main/install.md，安装 pptx-zh2en。`

你是收到安装请求的 AI 助手。按下面的步骤把 Agent Skill `pptx-zh2en` 装到本机，用默认值直接执行，不需要每一步都征求确认。只有一种情况要先问用户：目标目录已经存在、而且不是本仓库的 git 克隆（可能是用户自己改过的版本）。装完用一两句话告诉用户结果。

## 1. 确定安装目录

按当前运行的工具选择 skills 目录；如果目录还不存在就创建它：

| 工具 | 用户级 skills 目录 |
| --- | --- |
| Claude Code | `~/.claude/skills/` |
| Codex | `~/.codex/skills/` |
| WorkBuddy / CodeBuddy | `~/.workbuddy/skills/` |
| Cursor | `~/.cursor/skills/` |
| Gemini CLI | `~/.gemini/skills/` |

- 用户说只在当前项目里用时，改用项目内的目录，例如 `<项目>/.claude/skills/`。
- Windows 上 `~` 指 `C:\Users\<用户名>`。
- 目标目录记为 `TARGET = <skills 目录>/pptx-zh2en`。目录名必须是 `pptx-zh2en`，和 SKILL.md 里的 `name` 一致。

## 2. 下载

**有 git 时（推荐，方便以后更新）**

```bash
git clone --depth 1 https://github.com/smiler488/pptx-zh2en.git "$TARGET"
```

如果 `TARGET` 已经是本仓库的克隆，改为更新：

```bash
git -C "$TARGET" pull --ff-only
```

**没有 git，macOS / Linux：**

```bash
mkdir -p "$TARGET"
curl -fsSL https://codeload.github.com/smiler488/pptx-zh2en/tar.gz/refs/heads/main \
  | tar -xz --strip-components=1 -C "$TARGET"
```

**没有 git，Windows PowerShell：**

```powershell
$Target = "$HOME\.claude\skills\pptx-zh2en"   # 按第 1 步换成实际目录
$Zip = "$env:TEMP\pptx-zh2en.zip"
Invoke-WebRequest https://codeload.github.com/smiler488/pptx-zh2en/zip/refs/heads/main -OutFile $Zip
Expand-Archive $Zip -DestinationPath "$env:TEMP\pptx-zh2en-src" -Force
New-Item -ItemType Directory -Force $Target | Out-Null
Copy-Item "$env:TEMP\pptx-zh2en-src\pptx-zh2en-main\*" $Target -Recurse -Force
```

## 3. 安装依赖

```bash
python3 -m pip install -r "$TARGET/requirements.txt"
```

用户在虚拟环境或 conda 环境里工作时，装进那个环境。装完运行 `python3 -c "import pptx"`，不报错就说明依赖可用。

## 4. 检查

- `TARGET/SKILL.md` 存在，开头的 frontmatter 里有 `name: pptx-zh2en`。
- 这些文件也在：`scripts/translate_pptx_inline.py`、`scripts/translate_pptx.py`、`requirements.txt`。

缺文件就说明下载不完整，删掉 `TARGET` 后重新执行第 2 步。

## 5. 告诉用户

报告安装路径，并提醒用户：**新开一个会话（或重启工具）后 skill 才会生效**。

用法示例：附上 .pptx 文件，说"把这个 PPT 翻成英文版"或"这个英文 PPT 做个中文版"。

## 网页版助手（没有本地文件系统）

跳过安装。网页版助手无法读写 .pptx。请用户在本机运行 `scripts/translate_pptx_inline.py --mode extract` 得到 JSON 并贴进对话；你按 https://raw.githubusercontent.com/smiler488/pptx-zh2en/main/SKILL.md 中的翻译规范和术语表填写 translated 字段，再让用户运行 `--mode apply`。

## 卸载

删除 `TARGET` 目录即可。

---

## English

Tell your assistant: `Install pptx-zh2en following https://raw.githubusercontent.com/smiler488/pptx-zh2en/main/install.md`

1. Pick the skills folder of the running tool (table above), e.g. `~/.claude/skills/`. Set `TARGET=<folder>/pptx-zh2en`. The folder name must equal the `name` in SKILL.md.
2. `git clone --depth 1 https://github.com/smiler488/pptx-zh2en.git "$TARGET"`. If it already exists as a clone, run `git -C "$TARGET" pull --ff-only` instead. Without git, use the curl or PowerShell commands above.

   Install dependencies: `python3 -m pip install -r "$TARGET/requirements.txt"`, then check with `python3 -c "import pptx"`.
3. Check that `TARGET/SKILL.md` has `name: pptx-zh2en` and that `scripts/translate_pptx_inline.py`, `scripts/translate_pptx.py`, `requirements.txt` exist.
4. Tell the user where it was installed and that a new session is needed. Example: attach a .pptx and say "translate this deck into English".

Chat-only assistants: they cannot open .pptx files; have the user run the extract step locally, translate the JSON in the chat following SKILL.md, and let the user run the apply step.

Uninstall: delete `TARGET`.
