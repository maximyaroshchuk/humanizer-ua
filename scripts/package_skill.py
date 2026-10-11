#!/usr/bin/env python3
"""Збирає dist/humanizer-ua.zip для завантаження скіла в claude.ai.

Використання:
    python3 scripts/package_skill.py

Архів містить теку humanizer-ua/ із SKILL.md, references/, scripts/ і knowledge/.
У claude.ai: Settings → Capabilities → Skills → Upload skill → вибрати цей zip.
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INCLUDE = ["SKILL.md", "LICENSE", "references", "scripts", "knowledge"]
SKIP = {"__pycache__", "package_skill.py", "validate-package.py"}


def main():
    out_dir = os.path.join(ROOT, "dist")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "humanizer-ua.zip")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in INCLUDE:
            path = os.path.join(ROOT, item)
            if os.path.isfile(path):
                z.write(path, f"humanizer-ua/{item}")
                continue
            for folder, dirs, files in os.walk(path):
                dirs[:] = sorted(d for d in dirs if d not in SKIP)
                for name in sorted(files):
                    if name in SKIP or name.endswith(".pyc"):
                        continue
                    full = os.path.join(folder, name)
                    z.write(full, "humanizer-ua/" + os.path.relpath(full, ROOT))
    print(out)


if __name__ == "__main__":
    main()
