#!/usr/bin/env python3
"""Validate this skill repository: SKILL.md frontmatter, install.md and README links.

Run from the repository root:  python .github/scripts/validate_skill.py
"""
import re
import subprocess
import sys
from pathlib import Path

ALLOWED = {"name", "description", "license", "allowed-tools", "metadata", "compatibility"}
errors = []
root = Path(".")
repo = Path.cwd().name

text = (root / "SKILL.md").read_text(encoding="utf-8")
m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
if not m:
    sys.exit("SKILL.md: missing frontmatter")
fm = m.group(1)
keys = {line.split(":")[0] for line in fm.splitlines() if re.match(r"^[A-Za-z_-]+:", line)}
if keys - ALLOWED:
    errors.append(f"SKILL.md: non-standard top-level keys {sorted(keys - ALLOWED)}")
name = re.search(r"^name:\s*(.+)$", fm, re.M)
desc = re.search(r"^description:\s*(.+)$", fm, re.M)
name = name.group(1).strip().strip('"') if name else ""
if not re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", name) or len(name) > 64:
    errors.append(f"SKILL.md: invalid name {name!r}")
if not desc or len(desc.group(1).strip().strip('"')) > 1024:
    errors.append("SKILL.md: description missing or longer than 1024 characters")

install = (root / "install.md").read_text(encoding="utf-8")
for url in re.findall(r"https://raw\.githubusercontent\.com/smiler488/([^/]+)/([^/]+)/install\.md", install):
    if url[1] != "main":
        errors.append(f"install.md: raw link uses branch {url[1]!r}, expected main")
if f"name: {name}" not in install:
    errors.append(f"install.md: does not check for 'name: {name}'")
listed = re.search(r"这些文件也在：(.*)", install)
for f in re.findall(r"`([^`]+)`", listed.group(1) if listed else ""):
    if not (root / f).exists():
        errors.append(f"install.md: listed file {f} is missing")

tracked = subprocess.run(["git", "ls-files", "*.md"], capture_output=True, text=True).stdout.split()
for md in tracked:
    for link in re.findall(r"\]\(([^)#\s]+)", Path(md).read_text(encoding="utf-8")):
        if not link.startswith(("http", "mailto:")) and not (Path(md).parent / link).exists():
            errors.append(f"{md}: broken link {link}")

for e in errors:
    print("ERROR", e)
print(f"{name}: {'OK' if not errors else f'{len(errors)} problem(s)'}")
sys.exit(1 if errors else 0)
