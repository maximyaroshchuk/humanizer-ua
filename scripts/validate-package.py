#!/usr/bin/env python3
"""Перевіряє файли пакета humanizer-ua без сторонніх залежностей.

Використання:
    python3 scripts/validate-package.py

Перевіряє: метадані SKILL.md, одну версію в усіх файлах, однаковий опис у маніфестах,
нумерацію патернів у references/patterns.md і її збіг із картою в SKILL.md та
таблицею в README.md, посилання §N, формат словника references/surzhyk.md,
обсяг SKILL.md і коректність evals/evals.json.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL_WORD_LIMIT = 4500


def read(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError as error:
        raise SystemExit(f"Не вдалося прочитати {path.relative_to(ROOT)}: {error}")


def load_json(path: Path):
    try:
        return json.loads(read(path))
    except json.JSONDecodeError as error:
        raise SystemExit(f"Виправ JSON у {path.relative_to(ROOT)}: {error}")


def need(match, message):
    if match is None:
        raise SystemExit(message)
    return match


SKILL = read(ROOT / "SKILL.md")
README = read(ROOT / "README.md")
CHANGELOG = read(ROOT / "CHANGELOG.md")
PATTERNS = read(ROOT / "references" / "patterns.md")
SURZHYK = read(ROOT / "references" / "surzhyk.md")
LINT = read(ROOT / "scripts" / "lint.py")
PLUGIN = load_json(ROOT / ".claude-plugin" / "plugin.json")
MARKETPLACE = load_json(ROOT / ".claude-plugin" / "marketplace.json")
EVALS = load_json(ROOT / "evals" / "evals.json")

# --- метадані SKILL.md ---
meta = need(re.match(r"\A---\n(.*?)\n---\n", SKILL, re.S), "SKILL.md має починатися з YAML-метаданих").group(1)
for field in ("version:", "compatibility:", "allowed-tools:"):
    if re.search(rf"(?m)^{re.escape(field)}", meta):
        raise SystemExit(f"Прибери непідтримуване поле YAML верхнього рівня: {field[:-1]}")
if not re.search(r"(?m)^name: humanizer-ua$", meta):
    raise SystemExit("У SKILL.md має бути name: humanizer-ua")
description = " ".join(need(re.search(r"(?m)^description: \|\n((?:  .*\n)+)", meta + "\n"),
                            "Запиши description у SKILL.md блоком із відступом").group(1).split())

if len(description) > 1024:
    raise SystemExit(f"description у SKILL.md довший за 1024 символи ({len(description)}): Claude.ai його не прийме")

# --- версія ---
versions = {
    "SKILL.md": need(re.search(r'(?m)^\s+version:\s*["\']?(\d+\.\d+\.\d+)["\']?\s*$', meta),
                     "Додай metadata.version у SKILL.md").group(1),
    "CHANGELOG.md": need(re.search(r"(?m)^## (\d+\.\d+\.\d+)\b", CHANGELOG),
                         "Додай заголовок версії «## X.Y.Z» у CHANGELOG.md").group(1),
    "plugin.json": str(PLUGIN.get("version", "")),
}
if len(set(versions.values())) != 1:
    raise SystemExit(f"Одна версія в усіх файлах: {versions}")
version = versions["SKILL.md"]

# --- маніфести ---
if {PLUGIN.get("name"), *(p.get("name") for p in MARKETPLACE.get("plugins", []))} != {"humanizer-ua"}:
    raise SystemExit("Назва плагіна в маніфестах має бути humanizer-ua")
if PLUGIN.get("skills") != ["./"]:
    raise SystemExit("Плагін Claude має вказувати skills: [\"./\"]")
manifest_descriptions = {PLUGIN.get("description", ""),
                         *(p.get("description", "") for p in MARKETPLACE.get("plugins", []))}
if len(manifest_descriptions) != 1 or not description.startswith(next(iter(manifest_descriptions))):
    raise SystemExit("Опис у маніфестах має дорівнювати першому реченню description зі SKILL.md: "
                     f"{sorted(manifest_descriptions)}")
skill_files = {p.relative_to(ROOT) for p in ROOT.rglob("SKILL.md") if "dist" not in p.parts}
if skill_files != {Path("SKILL.md")}:
    raise SystemExit("Має бути один SKILL.md у корені репо")

# --- патерни ---
patterns = {int(n): name.strip() for n, name in re.findall(r"(?m)^### (\d+)\. (.+)$", PATTERNS)}
numbers = list(patterns)
count = len(numbers)
if not count or numbers != list(range(1, count + 1)):
    raise SystemExit(f"Пронумеруй патерни в references/patterns.md з 1 без пропусків: {numbers}")
if f"({count} патерн" not in PATTERNS.splitlines()[0]:
    raise SystemExit(f"Заголовок references/patterns.md має назвати кількість: «({count} патерни)»")


def table(text, where):
    rows = {int(n): name.strip() for n, name in re.findall(r"(?m)^\| (\d+) \| ([^|]+?) \|", text)}
    if sorted(rows) != numbers:
        raise SystemExit(f"Таблиця патернів у {where} має містити 1–{count} по одному разу: {sorted(rows)}")
    wrong = [f"{n}: «{rows[n]}» ≠ «{patterns[n]}»" for n in numbers if rows[n] != patterns[n]]
    if wrong:
        raise SystemExit(f"Назви патернів у {where} мають збігатися з references/patterns.md: " + "; ".join(wrong))


table(SKILL, "SKILL.md")
table(README, "README.md")
for name, text in (("SKILL.md", SKILL), ("README.md", README)):
    if f"{count} патерн" not in text:
        raise SystemExit(f"{name} має називати актуальну кількість патернів: {count}")

# --- посилання §N ---
for name, text in (("SKILL.md", SKILL), ("references/patterns.md", PATTERNS),
                   ("references/surzhyk.md", SURZHYK), ("scripts/lint.py", LINT)):
    bad = sorted({int(n) for n in re.findall(r"§(\d+)", text)} - set(numbers))
    if bad:
        raise SystemExit(f"{name}: посилання § мають вказувати на патерни 1–{count}: {bad}")

# --- словник ---
entry_rx = re.compile(r"^- \*\*(.+?)\*\*.*?→\s*(.+?)\s·\s(ERROR|WARN)\s·\s§(\d+)\s·\s`(.+)`\s*$")
entries = 0
for i, line in enumerate(SURZHYK.splitlines(), 1):
    if not line.startswith("- **"):
        continue
    m = entry_rx.match(line)
    if not m:
        raise SystemExit(f"references/surzhyk.md:{i}: рядок не відповідає формату «- **…** → … · РІВЕНЬ · §N · `regex`»")
    try:
        re.compile(m.group(5), re.I)
    except re.error as error:
        raise SystemExit(f"references/surzhyk.md:{i}: некоректний регулярний вираз: {error}")
    entries += 1
if entries < 50:
    raise SystemExit(f"У словнику лише {entries} записів — схоже, формат зламано")

# --- обсяг ---
words = len(SKILL.split())
if words > SKILL_WORD_LIMIT:
    raise SystemExit(f"SKILL.md читається щоразу повністю: тримай до {SKILL_WORD_LIMIT} слів, зараз {words}")

# --- evals ---
evals = EVALS.get("evals", [])
ids = [e.get("id") for e in evals]
if not evals or len(ids) != len(set(ids)):
    raise SystemExit("evals/evals.json: потрібні сценарії з унікальними id")
for e in evals:
    if not e.get("prompt") or not e.get("expectations"):
        raise SystemExit(f"evals/evals.json: сценарій {e.get('id')} має мати prompt і expectations")

print(f"humanizer-ua v{version}: {count} патернів, {entries} записів словника, "
      f"{len(evals)} сценаріїв, SKILL.md {words} слів — пакет коректний")
