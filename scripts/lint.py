#!/usr/bin/env python3
"""Лінтер ШІ-ознак і мовної норми для humanizer-ua. Тільки стандартна бібліотека.

Використання:
    python3 scripts/lint.py текст.docx     # Word: читає абзаци, заголовки, жирний
    python3 scripts/lint.py текст.md       # або .txt
    python3 scripts/lint.py < текст.md     # stdin
    python3 scripts/lint.py --json текст.docx
    python3 scripts/lint.py --self-test

ERROR  — жорстка заборона SKILL.md, порушення норми зі словника або артефакт
         копіпасту з чат-бота. Текст не готовий: код виходу 1.
WARN   — маркер патерна. Оцінюй кластерами: одне «ключовий» нічого не доводить.
Вердикт: (errors*3 + warnings) на 300 слів → чисто (0–3) / переглянути (до 10) / переписати (більше).

Лінт ганяють лише по чистовику: без цитат «було», резюме і журналу змін.
Номери § відповідають references/patterns.md.
"""
import json
import os
import re
import statistics
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SURZHYK = os.path.join(HERE, "..", "references", "surzhyk.md")
APOS = "’'ʼ"
UPPER = "А-ЯІЇЄҐ"
LOWER = "а-яіїєґ"

# --- клас A: сліди копіпасту з чат-ботів (один слід = вердикт) ---
ARTIFACTS = [
    ("A сноска ChatGPT", re.compile(r":contentReference\[oaicite:\d+\]|oai_citation|\boaicite:\d+")),
    ("A пошукова мітка", re.compile(r"\bturn\d+(?:search|file|fetch|image|news|view|ref)\d+|citeturn")),
    ("A utm чат-бота", re.compile(r"utm_source=(?:chatgpt|copilot|openai|perplexity)\.(?:com|ai)|referrer=grok\.com")),
    ("A позначка Gemini", re.compile(r"\[cite_start\]|\[cite:\s*\d+|grounding-api-redirect|\[span_\d+\]")),
    ("A внутрішня сноска", re.compile(r"【\d+[†:][^】]*】|\]\(sandbox:/mnt/data/")),
    ("A залишок міркувань", re.compile(r"</?think>")),
    ("A файл Perplexity", re.compile(r"ppl-ai-file-upload")),
    ("A позначка Grok", re.compile(r"grok_card|grok_render_citation")),
    ("A плейсхолдер", re.compile(r"\[(?:вставте|вкажіть|додайте|ім[’'ʼ]я|назва компанії|посилання)[^\]]{0,40}\]"
                                 r"|INSERT_\w+|PASTE_\w+|\b20\d\d-XX-XX\b|\bXX\.XX\.20\d\d\b", re.I)),
]
EMOJI = "\U0001F300-\U0001FAFF☀-➿⬀-⯿"
ZERO_WIDTH = re.compile(r"[​‌⁠﻿]|(?<![%s])‍|‍(?![%s])" % (EMOJI, EMOJI))
CODE = re.compile(r"```.*?```|`[^`\n]+`", re.S)
URL = re.compile(r"https?://\S+|www\.\S+")

# --- жорсткі заборони ---
ERRORS = [
    # §12: потрібна повна фігура контрасту. «не тільки штатні працівники» без
    # другої частини — уточнення, не патерн. «не просто так» — ідіома.
    ("§12 не просто X, а Y", re.compile(
        r"\bне\s+просто\s+(?!так\b)[^.!?\n]{1,80}?,\s*(?:а|але|проте)\s+\S"
        r"|\bне\s+просто\s+(?!так\b)[^.!?\n]{1,40}[.!]\s+Це\s+\S"
        r"|\b[ЙйІі]деться\s+не\s+(?:лише|тільки|просто)\b"
        r"|\b(?:це|Це)\s+не\s+просто\b[^.!?\n]{1,60}[—–-]\s*це\b", re.I)),
    ("§21 рублені фрагменти", re.compile(
        r"(?:^|[.!?]\s+)(?:Без|Жодн\w+|Ніяк\w+|Нуль|Ані)\s[^.!?\n]{1,35}[.!]\s+(?:Без|Жодн\w+|Ніяк\w+|Нуль|Ані)\s")),
    ("§36 мат-знаки в прозі", re.compile(r"[→⇒←↔⟶≈≠≥≤]|(?<=\S)\s[=<>+&]\s(?=\S)|\bvs\.?(?=\s)")),
    ("§35 дефіс замість тире", re.compile(r"(?<=[^\s\-|])\s-\s(?=\S)")),
]
HR_LINE = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,})\s*$")

# --- маркери (WARN), кластери вирішують ---
PHRASES = [
    ("§6 канцелярит", ["здійснювати", "здійснює", "здійснюємо", "здійснюють", "проводити роботу",
                       "надавати допомогу", "надає допомогу", "забезпечувати виконання", "з метою",
                       "шляхом", "у рамках", "в рамках", "на предмет", "відповідно до вимог",
                       "з урахуванням", "в частині", "у частині"]),
    ("§7 являє собою", ["являє собою", "являють собою", "виступає в ролі", "виступає у ролі", "слугує",
                        "постає як", "є важливим", "є ключовим", "є необхідним", "є невід"]),
    ("§13 кікер", ["і це змінює все", "подумайте над цим", "у цьому вся суть", "в цьому вся суть",
                   "мораль проста", "це чудово ілюструє", "звучить просто? так і є"]),
    ("§14 псевдоглибина", ["по суті,", "якщо копнути глибше", "справжнє питання", "на глибинному рівні",
                           "у самому серці", "суть у тому", "це нова валюта", "валюта майбутнього"]),
    ("§15 анонс", ["давайте розберемося", "давайте розглянемо", "давайте поговоримо", "пориньмо",
                   "поринемо", "зануримося", "ось що вам потрібно знати", "а тепер найцікавіше",
                   "без зайвих слів"]),
    ("§16 фальшива довірливість", ["чесно?", "будьмо відверті", "будьмо чесні", "скажу прямо",
                                   "відкрию секрет", "ось у чому річ"]),
    ("§25 словник ШІ", ["ключов", "відіграє", "невід’ємн", "невід'ємн", "невідʼємн", "сприяє", "сприяють",
                        "демонструє", "демонструють", "підкреслює", "свідчить про", "унікальн",
                        "неабиям", "всебічн", "всеосяжн", "вирішальн", "трансформ", "інноваційн",
                        "екосистем", "розкрити потенціал", "розкриває потенціал", "на новий рівень",
                        "у сучасному світі", "в сучасному світі", "в епоху", "у епоху",
                        "в сучасних реаліях", "у сучасних реаліях", "безперечно", "безсумнівно",
                        "неймовірн", "захопливий світ", "багатогранн"]),
    ("§26 роздута значущість", ["знаменує", "поворотним моментом", "поворотний момент",
                                "увійшло в історію", "назавжди змінило", "важко переоцінити",
                                "світового рівня", "новий етап"]),
    ("§28 рекламна мова", ["неперевершен", "бездоганн", "мальовнич", "справжня перлина",
                           "для найвибагливіших", "незабутн", "у самому серці", "в самому серці"]),
    ("§29 розмите посилання", ["експерти вважають", "дослідження показують", "науковці стверджують",
                               "аналітики зазначають", "на думку фахівців", "на думку експертів",
                               "як відомо", "загальновизнано", "фахівці наголошують"]),
    ("§30 виклики й перспективи", ["попри низку викликів", "низку викликів", "стикається з викликами",
                                   "стикається з певними", "незважаючи на виклики", "попри ці виклики",
                                   "значні перспективи"]),
    ("§32 фінал-підсумок", ["підсумовуючи", "отже, можна зробити висновок", "на завершення",
                            "майбутнє виглядає", "попереду захопливі", "можна сказати, що",
                            "таким чином, можна"]),
    ("§37 залишки чат-бота", ["чудове запитання", "чудове питання", "гарне питання", "звісно! ось",
                              "сподіваюся, це допоможе", "сподіваюся, це було корисно",
                              "сподіваємося, це було корисно", "сподіваємося, це допоможе",
                              "якщо маєте запитання", "якщо у вас є запитання", "радий допомогти",
                              "рада допомогти", "як мовна модель", "як штучний інтелект",
                              "станом на мої", "я не маю доступу до", "ось ваш текст",
                              "ось оновлений текст", "дайте знати, якщо"]),
    ("§15 штамп-анонс", ["важливо зазначити", "варто зазначити", "слід зазначити",
                              "варто відзначити", "слід підкреслити", "необхідно враховувати",
                              "варто звернути увагу", "не можна не згадати", "варто пам’ятати",
                              "варто пам'ятати", "важливо пам’ятати", "важливо пам'ятати"]),
    ("§40 терапевтичний тон", ["і це нормально", "і це абсолютно нормально", "ви не самотні",
                               "ви не одні", "будьте до себе", "дозвольте собі", "кожен проходить свій"]),
]
WARN_RX = [
    ("§12 не лише X, а й Y (лиши, лише якщо обидві половини несуть факт)", re.compile(
        r"\bне\s+(?:лише|тільки)\b[^.!?\n]{1,80}?\b(?:а\s+й|але\s+й|а\s+ще\s+й|але\s+також|ба\s+й|а\s+і|але\s+і)\b", re.I)),
    ("§12 це не X, а Y", re.compile(r"\b[Цц]е\s+не\s+(?!просто\b|лише\b|тільки\b)[^.,!?\n]{1,40},\s+а\s+\S")),
    ("§4 пасив з орудним", re.compile(
        r"\b\w+(?:ється|ються|вся|лася|лося|лися|ний|на|не|ні|тий|та|те|ті)\s+"
        r"(?:керівництвом|адміністрацією|компанією|командою|користувачами|фахівцями|спеціалістами|"
        r"урядом|нами|ними|автором|авторами|системою|директором|менеджером|розробниками)\b", re.I)),
    ("§27 дієприслівниковий хвіст", re.compile(
        r",\s+(?:підкреслюючи|демонструючи|відображаючи|засвідчуючи|сприяючи|забезпечуючи|"
        r"роблячи|наголошуючи|ілюструючи|відкриваючи|створюючи)\b[^.!?\n]*[.!?]")),
    ("§20 двокрапка-підводка", re.compile(
        r"(?:^|[.!?]\s+)(?:Найцікавіше|Найважливіше|Головне|Результат|Висновок|Суть|Секрет|Відповідь|"
        r"Підсумок|Правда|Проблема|Факт|Ключ|Ось що \w+|Ось як|Ось чому|Мораль)[^.!?:\n]{0,25}:\s")),
    ("§17 риторичне запитання", re.compile(r"(?:^|[.!?]\s+)[^.!?\n]{3,60}\?\s+[%s]" % UPPER)),
    ("§19 тире перед «відкриттям»", re.compile(
        r"\s[—–]\s+(?:неймовірн|вражаюч|колосальн|величезн|очевидн|простий|проста|просте)\w*[.!]", re.I)),
    ("§9 звертання в називному", re.compile(
        r"(?:^|[.!?]\s+)(?:Шановн(?:ий|а)|Дорог(?:ий|а)|Любий|Люба|Привіт|Вітаю|Добрий день|Пане|Пані)"
        r",?\s+([%s][%s%s]+(?:\s+[%s][%s%s]+)?)\s*[,!]" % (UPPER, LOWER, APOS, UPPER, LOWER, APOS))),
    ("§35 дріб через крапку", re.compile(r"(?<![\d.])\d+\.\d+(?![\d.])\s?(?:%|млн|млрд|тис|разів|кг|км|грн|₴|\$|€)")),
    ("§35 розряди через кому", re.compile(r"(?<![\d,.])\d{1,3}(?:,\d{3})+(?!\d|,\d)")),
    ("§35 валюта перед числом", re.compile(r"[$€£]\s?\d")),
    ("§35 нарощення", re.compile(r"\d-(?:ий|ого|ому|ою)\b|\b[IVXLC]{2,}-(?:й|го|му|ий|ого|е|ті|те)\b")),
    ("§35 скорочення без пробілу", re.compile(r"\bт\.(?:д|ін|зв|п)\.|\b[%s]\.[%s]\.\s?[%s][%s]" % (UPPER, UPPER, UPPER, LOWER))),
    ("§35 № без пробілу", re.compile(r"[№§]\d")),
    ("§35 стара норма: проект", re.compile(r"\bпроект\w*", re.I)),
    ("§35 «пів» через дефіс", re.compile(r"\bпів-[%s]" % (LOWER + UPPER))),
]
VOCATIVE_OK = re.compile(r"(?:е|о|у|ю|є|ій|ю|і)$")  # Олександре, Олено, Андрію, Наталю
DOTTED = re.compile(r"(?:^|[.!?]\s+)([%s]{3,8})\.\s+([%s]{3,8})\.\s+([%s]{3,8})\.(?=\s|$)" % (LOWER, LOWER, LOWER))
ABBR = {"тис", "млн", "млрд", "див", "напр", "прил", "изд", "вид", "ред", "упор", "інш", "стор", "табл",
        "рис", "мал", "акад", "проф", "доц", "канд", "грн", "обл", "род", "див", "ім", "вул", "буд"}
FORMAL_OPEN = re.compile(r"^(?:У|В)\s+(?:сучасному світі|наш час|епоху|сучасних реаліях)|^(?:Сьогодні|Нині),\s+коли")
SENT_SPLIT = re.compile(r"(?<=[.!?…])\s+(?=[%s«\"0-9])" % UPPER)
WORD = re.compile(r"[%s%s%s\w-]+" % (LOWER, UPPER, APOS))
VERB_END = re.compile(r"^([%s]{3,}?)(?:ує|ують|ює|юють|ить|ять|ать|ає|ають|яє|яють|ується|юється)$" % LOWER)
COMMON_VERBS = {"має", "мають", "може", "можуть", "буде", "будуть", "стає", "стають", "знає", "хоче", "бачить"}


def load_dictionary(path=SURZHYK):
    """Читає рядки `- **…** → … · ERROR|WARN · §N · `regex`` зі словника."""
    entries = []
    line_rx = re.compile(r"^- \*\*(.+?)\*\*.*?→\s*(.+?)\s·\s(ERROR|WARN)\s·\s(§\d+)\s·\s`(.+)`\s*$")
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                m = line_rx.match(line.rstrip("\n"))
                if m:
                    shown, fix, level, num, rx = m.groups()
                    entries.append((level, f"{num} {shown} → {fix}", re.compile(rx, re.I)))
    except OSError:
        pass
    return entries


DICTIONARY = load_dictionary()


# --- розбір тексту ---

def strip_frontmatter(lines):
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return [""] * (i + 1) + lines[i + 1:]
    return lines


def prose_line(line):
    """Рядок без коду, посилань і розмітки на початку; порожній для коду й таблиць."""
    s = line.strip()
    if s.startswith("|"):
        return ""
    s = CODE.sub(" ", s)
    s = URL.sub(" ", s)
    s = re.sub(r"^(?:#{1,6}\s+|>\s*|[-*+]\s+|\d+[.)]\s+)", "", s)
    return s


def blocks(lines):
    """Абзаци (списки рядків) прози, без блоків коду."""
    out, cur, fence = [], [], False
    for line in lines:
        if line.strip().startswith("```"):
            fence = not fence
            continue
        if fence:
            continue
        if line.strip():
            cur.append(line)
        elif cur:
            out.append(cur)
            cur = []
    if cur:
        out.append(cur)
    return out


def sentences(text):
    text = re.sub(r"\s+", " ", text).strip()
    return [s for s in SENT_SPLIT.split(text) if s]


def ctx(text, start, end, width=35):
    a, b = max(0, start - width), min(len(text), end + width)
    return ("…" if a else "") + text[a:b].strip() + ("…" if b < len(text) else "")


# --- перевірки ---

QUOTE = re.compile(r"«[^«»\n]*»|„[^„“\n]*“")


def quote_spans(s):
    return [m.span() for m in QUOTE.finditer(s)]


def downgrade_in_quote(level, line, rule, context, start, spans):
    """Збіг усередині «лапок» — цитата чи назва: не правимо, тому лише WARN."""
    if level == "ERROR" and any(a < start < b for a, b in spans):
        return ("WARN", line, f"{rule} (у лапках — цитату й назву не правимо)", context)
    return (level, line, rule, context)


def lint(text):
    findings = []  # (рівень, рядок, правило, контекст)
    lines = strip_frontmatter(text.splitlines())
    fence = False
    code_free = CODE.sub(" ", text)

    for i, raw in enumerate(lines, 1):
        if raw.strip().startswith("```"):
            fence = not fence
            continue
        if fence:
            continue
        scan_raw = CODE.sub(" ", raw)
        for rule, rx in ARTIFACTS:
            m = rx.search(scan_raw)
            if m:
                findings.append(("ERROR", i, rule, ctx(scan_raw, m.start(), m.end())))
        if ZERO_WIDTH.search(scan_raw):
            findings.append(("WARN", i, "A символ нульової ширини", "невидимий символ у рядку"))
        if HR_LINE.match(raw) and i > 1:
            findings.append(("ERROR", i, "§33 розділювач --- у тексті", raw.strip()))
            continue

        heading = raw.lstrip().startswith("#")
        s = prose_line(raw)
        if not s:
            continue
        low = s.lower()
        quoted = quote_spans(s)
        for rule, rx in ERRORS:
            if heading and rule.startswith("§35"):
                continue
            for m in rx.finditer(s):
                findings.append(downgrade_in_quote("ERROR", i, rule, ctx(s, m.start(), m.end()), m.start(), quoted))
        for level, rule, rx in DICTIONARY:
            for m in rx.finditer(s):
                findings.append(downgrade_in_quote(level, i, rule, ctx(s, m.start(), m.end()), m.start(), quoted))
        for rule, words in PHRASES:
            for w in words:
                pos = low.find(w)
                if pos >= 0:
                    findings.append(("WARN", i, f"{rule}: «{w}»", ctx(s, pos, pos + len(w))))
        for rule, rx in WARN_RX:
            for m in rx.finditer(s):
                if rule.startswith("§9") and VOCATIVE_OK.search(m.group(1).split()[0]):
                    continue
                findings.append(("WARN", i, rule, ctx(s, m.start(), m.end())))
        if re.search("[%s]" % EMOJI, s) and (heading or re.match(r"^\s*(?:[-*+]\s+)?[%s]" % EMOJI, raw.strip())):
            findings.append(("WARN", i, "§33 емодзі в заголовку чи на початку пункту", s[:60]))
        if heading:
            words = [w for w in WORD.findall(s) if len(w) > 2]
            if len(words) >= 3 and all(w[0].isupper() for w in words[1:]):
                findings.append(("WARN", i, "§34 заголовок З Великих Літер", s[:70]))
        caps = [w for w in re.findall(r"\b[%s]{6,}\b" % UPPER, s)]
        if caps and not heading:
            findings.append(("WARN", i, "§13 слово ВЕЛИКИМИ", ", ".join(caps[:3])))
        for m in DOTTED.finditer(s):
            if not set(m.groups()) & ABBR:
                findings.append(("WARN", i, "§13 слова. через. крапку.", m.group().strip()))

    findings += structure(lines, code_free)
    return findings


def structure(lines, code_free):
    out = []
    paras = blocks(lines)
    prose = []  # (перший рядок абзацу як заголовок?, текст)
    for b in paras:
        first = b[0].lstrip()
        if first.startswith("|"):
            continue
        text = " ".join(prose_line(l) for l in b).strip()
        if text:
            prose.append((first.startswith("#"), first.startswith(("-", "*", "+")) or re.match(r"\d+[.)]\s", first), text, b))

    body = [t for h, lst, t, _ in prose if not h and not lst]
    all_sent = [s for t in body for s in sentences(t)]

    # §19 щільність тире
    dashes = sum(len(re.findall(r"\s[—–]\s", t)) for t in body)
    if len(all_sent) >= 5 and dashes / len(all_sent) >= 0.5:
        out.append(("WARN", 0, "§19 тире як універсальна зв’язка",
                    f"{dashes} тире на {len(all_sent)} речень — перевір, де правопис його не вимагає"))
    if re.search(r"\s—\s", code_free) and re.search(r"\s–\s", code_free):
        out.append(("WARN", 0, "§35 різні тире в одному тексті", "і — (довге), і – (коротке): тримай один знак"))

    # §35 прямі лапки й апостроф: один підсумковий рядок, бо це звичка набору, а не ШІ
    quotes = len(re.findall(r'(?<!\w)"[^"\n]{1,120}"(?!\w)', code_free))
    if quotes:
        out.append(("WARN", 0, "§35 прямі лапки", f"{quotes} раз(и) \"так\" — у чистовику «ялинки»"))
    apos = len(re.findall(r"[%s]'[%s]" % (LOWER + UPPER, LOWER), code_free))
    if apos:
        out.append(("WARN", 0, "§35 прямий апостроф", f"{apos} раз(и) ' — у чистовику ’ або ʼ"))

    # §31 каскад обережностей
    hedges = re.compile(r"\b(?:можливо|ймовірно|певною мірою|деякою мірою|у певному сенсі|потенційно|мабуть|напевно|дещо)\b", re.I)
    for s in all_sent:
        n = len(hedges.findall(s))
        if n >= 3:
            out.append(("WARN", 0, "§31 каскад обережностей", f"{n} застереження: {s[:70]}"))

    # §18 тріади
    triad = re.compile(r"\b[\w’'ʼ-]{3,},\s+[\w’'ʼ-]{3,}\s+(?:і|й|та)\s+[\w’'ʼ-]{3,}\b")
    triads = sum(len(triad.findall(t)) for t in body)
    if triads >= 3:
        out.append(("WARN", 0, "§18 тріади", f"{triads} переліки «X, Y і Z» — чи справді по три?"))

    # §22 повтор дієслова, §23 однакові початки
    def verb_stems(s):
        stems = set()
        for w in WORD.findall(s.lower()):
            if w in COMMON_VERBS:
                continue
            m = VERB_END.match(w)
            if m:
                stems.add(m.group(1))
        return stems

    for a, b in zip(all_sent, all_sent[1:]):
        common = verb_stems(a) & verb_stems(b)
        if common:
            out.append(("WARN", 0, "§22 повтор дієслова", f"«{sorted(common)[0]}…» у сусідніх реченнях: {b[:60]}"))
    firsts = [(WORD.findall(s) or [""])[0].lower() for s in all_sent]
    for k in range(len(firsts) - 2):
        if firsts[k] and len(firsts[k]) > 1 and firsts[k] == firsts[k + 1] == firsts[k + 2]:
            out.append(("WARN", 0, "§23 однакові початки", f"три речення поспіль починаються з «{firsts[k]}»"))
            break

    # ритм
    lengths = [len(WORD.findall(s)) for s in all_sent]
    if len(lengths) >= 6:
        mean = statistics.mean(lengths)
        if mean and statistics.pstdev(lengths) / mean < 0.3:
            out.append(("WARN", 0, "ритм монотонний", f"речення однакової довжини (~{mean:.0f} слів)"))
    if len(lengths) >= 8 and min(lengths) > 6:
        out.append(("WARN", 0, "ритм без коротких речень", "жодного речення коротшого за 7 слів"))

    # §24 абзаци й списки під копірку
    sizes = [len(WORD.findall(t)) for t in body]
    if len(sizes) >= 4:
        mean = statistics.mean(sizes)
        if mean >= 20 and all(abs(x - mean) <= 0.2 * mean for x in sizes):
            out.append(("WARN", 0, "§24 абзаци під копірку", f"{len(sizes)} абзаців по ~{mean:.0f} слів"))
    lists = [len(b) for h, lst, _, b in prose if lst]
    if len(lists) >= 2 and all(n == 3 for n in lists):
        out.append(("WARN", 0, "§24 списки по три", f"усі {len(lists)} списки рівно по три пункти"))
    labels = sum(1 for h, lst, t, b in prose if lst for l in b if re.match(r"^\s*[-*+]\s+\*\*[^*]+\*\*\s*[:—–-]", l))
    if labels >= 3:
        out.append(("WARN", 0, "§24 пункти «**мітка:** пояснення»", f"{labels} однакових пунктів"))

    # §33 жирний перебір
    words_total = len(WORD.findall(code_free))
    bold = len(re.findall(r"\*\*[^*\n]+\*\*", code_free))
    if bold >= 4 and words_total and bold * 40 > words_total:
        out.append(("WARN", 0, "§33 жирний перебір", f"{bold} виділень на {words_total} слів"))

    # §10 «Ви» з великої посеред речення
    mid_vy = re.findall(r"(?<=[\w,;:—–-] )(?:Ви|Вас|Вам|Вами|Ваш\w*)\b", code_free)
    if len(mid_vy) >= 2:
        out.append(("WARN", 0, "§10 «Ви» з великої", f"{len(mid_vy)} разів посеред речення — у масовому тексті пиши «ви»"))

    # §25 формальне відкриття, §38 заголовок-переказ
    if body and FORMAL_OPEN.search(body[0]):
        out.append(("WARN", 0, "§25 формальне відкриття", body[0][:60]))
    for (h, _, t, _), nxt in zip(prose, prose[1:]):
        if h and not nxt[0] and not nxt[1]:
            hw = {w.lower()[:5] for w in WORD.findall(t) if len(w) >= 4}
            first = (sentences(nxt[2]) or [""])[0]
            fw = {w.lower()[:5] for w in WORD.findall(first) if len(w) >= 4}
            if len(hw) >= 3 and len(hw & fw) / len(hw) >= 0.6:
                out.append(("WARN", 0, "§38 заголовок переказано першим реченням", first[:70]))
    return out


def verdict(errors, warnings, words=300):
    """Бал на 300 слів, щоб довгий текст не отримував «переписати» лише за довжину."""
    score = round((errors * 3 + warnings) * 300 / max(words, 300), 1)
    if score <= 3:
        return score, "чисто"
    if score <= 10:
        return score, "переглянути"
    return score, "переписати"


def read_input(path):
    if path is None or path == "-":
        return sys.stdin.read()
    if path.lower().endswith(".docx"):
        sys.path.insert(0, HERE)
        import docx_tool
        return docx_tool.render_markdown(path, all_parts=True)
    with open(path, encoding="utf-8") as f:
        return f.read()


def report(findings, text="", as_json=False):
    errors = [f for f in findings if f[0] == "ERROR"]
    warns = [f for f in findings if f[0] == "WARN"]
    words = len(WORD.findall(CODE.sub(" ", text)))
    score, label = verdict(len(errors), len(warns), words)
    if as_json:
        json.dump({"errors": len(errors), "warnings": len(warns), "words": words, "score": score, "verdict": label,
                   "findings": [dict(zip(("level", "line", "rule", "context"), f)) for f in findings]},
                  sys.stdout, ensure_ascii=False, indent=1)
        print()
    else:
        for level, line, rule, context in sorted(findings, key=lambda f: (f[0] != "ERROR", f[1])):
            where = f"L{line}" if line else "текст"
            print(f"{level:5}  {where:6} {rule}  ←  {context}")
        print(f"\nERROR: {len(errors)}  WARN: {len(warns)}  слів: {words}  бал на 300 слів: {score}  вердикт: {label}")
        if any(f[2].startswith("A ") for f in errors):
            print("Є артефакт копіпасту з чат-бота: текст скопійовано з ШІ.")
    return 1 if errors else 0


def self_test():
    def rules(text, level="ERROR"):
        return [f[2] for f in lint(text) if f[0] == level]

    assert DICTIONARY, "словник references/surzhyk.md не прочитано"

    bad = ("Це не просто сервіс, а справжній партнер. Даний проект являється ключовим, "
           "і ми приймаємо участь на протязі року. Швидкість > якості. Кит - ссавець. "
           "Без коду. Без налаштувань. Лише результат.")
    got = " | ".join(rules(bad))
    for need in ("§12", "§1 являється", "§1 приймати участь", "§1 на протязі", "§36", "§35 дефіс", "§21"):
        assert need in got, (need, got)

    ok = ("Кит — ссавець. У 2010—2018 роках місто з’являється на мапах. "
          "У звіт входять не тільки штатні працівники. Він прийшов не просто так. "
          "Рейс Київ — Львів відправляється о 10:30. Слово «пів» пишемо окремо: пів року.")
    assert not rules(ok), lint(ok)

    for p in ("Йдеться не лише про гроші.", "Це не просто інструмент. Це спосіб мислення.",
              "Це не просто сервіс, а партнер."):
        assert any(r.startswith("§12") for r in rules(p)), p
    for p in ("Це не лише швидко, а й дешево.", "Не тільки студенти, але й викладачі прийшли."):
        assert any(r.startswith("§12") for r in rules(p, "WARN")), p
        assert not any(r.startswith("§12") for r in rules(p)), p
    assert any("§5" in r for r in rules("По закінченню курсу видаємо сертифікат."))
    assert not rules("Порт засновано по завершенні будівництва.")

    calques = "Ця функція грає ключову роль. Існуючі правила застаріли. По закінченню курсу видаємо сертифікат."
    got = " | ".join(rules(calques))
    assert "§2" in got and "§3" in got and "§5" in got, got

    warn = ("Важливо зазначити, що в сучасному світі ключові рішення сприяють розвитку. "
            "Сподіваюся, це допоможе! Шановний Олександр, дякуємо.")
    wr = " | ".join(rules(warn, "WARN"))
    for need in ("§15", "§25", "§37", "§9"):
        assert need in wr, (need, wr)
    assert not any(r.startswith("§9") for r in rules("Шановний Олександре, дякуємо.", "WARN"))

    art = "Дані за 2024 рік :contentReference[oaicite:2]{index=2} і https://x.com/?utm_source=chatgpt.com"
    assert sum(r.startswith("A ") for r in rules(art)) == 2, rules(art)
    assert not any(r.startswith("A ") for r in rules("Мітка `turn0search3` у прикладі.")), "цитата в бектиках"

    quoted = "Директор сказав: «Ми являємося лідерами на протязі десяти років»."
    assert not rules(quoted), rules(quoted)
    assert any("у лапках" in r for r in rules(quoted, "WARN"))

    hr = "Перший абзац.\n\n---\n\nДругий абзац."
    assert any("---" in r for r in rules(hr))
    fm = "---\nname: x\n---\n\nТекст."
    assert not rules(fm), rules(fm)

    head = "# Як Почати Роботу З Платформою\n\nТекст."
    assert any(r.startswith("§34") for r in rules(head, "WARN"))

    typo = 'Ціна $1,500, знижка 7.5% до 5-ого числа в розділі "Умови", п\'ять днів.'
    tw = " | ".join(rules(typo, "WARN"))
    for need in ("валюта", "розряди", "дріб", "нарощення", "прямі лапки", "апостроф"):
        assert need in tw, (need, tw)
    assert not rules("Див. стор. табл. 5.", "WARN"), "скорочення — не «слова. через. крапку.»"

    # docx: лінтер читає Word-файл через docx_tool
    import tempfile
    sys.path.insert(0, HERE)
    import docx_tool
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "зразок.docx")
        with open(path, "wb") as f:
            f.write(docx_tool.build_docx(docx_tool.DEMO))
        got = " | ".join(f[2] for f in lint(read_input(path)))
        for need in ("§12", "§1 являється", "§3", "§34", "§21", "§37"):
            assert need in got, (need, got)
    print("lint self-test: OK")


def main(argv):
    as_json = "--json" in argv
    argv = [a for a in argv if a != "--json"]
    if argv and argv[0] == "--self-test":
        self_test()
        return 0
    if argv and argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    text = read_input(argv[0] if argv else None)
    return report(lint(text), text, as_json)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
