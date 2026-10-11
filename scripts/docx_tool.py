#!/usr/bin/env python3
"""Читання й правка .docx для humanizer-ua. Тільки стандартна бібліотека.

Використання:
    python3 scripts/docx_tool.py extract файл.docx [--all-parts]   > абзаци.json
    python3 scripts/docx_tool.py text файл.docx [--all-parts]      # текст для лінтера
    python3 scripts/docx_tool.py apply файл.docx правки.json [--out-dir ТЕКА] [--author ІМ'Я]
    python3 scripts/docx_tool.py demo вихід.docx                   # зразок із ШІ-ознаками
    python3 scripts/docx_tool.py --self-test

extract  друкує JSON зі списком непорожніх абзаців: id, стиль, текст, прапорці.
apply    читає правки [{"id": "12", "new": "новий текст"}] і пише поруч з оригіналом
         два файли: <ім'я>.рецензія.docx (режим рецензування Word, кожну правку
         можна прийняти чи відхилити) і <ім'я>.чистовик.docx (правки прийнято).
         Оригінал не змінюється. Форматування незмінених слів зберігається.

Захищені фрагменти абзацу (гіперпосилання, поля, зображення, виноски) не
змінюються. Їхній текст має лишитися в новому тексті абзацу без змін, інакше
абзац пропускається і потрапляє у звіт як «правте вручну».
"""
import datetime
import difflib
import io
import itertools
import json
import os
import re
import sys
import zipfile
from xml.dom import minidom

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# Діти w:r, з яких складається звичайний текст. Run з іншими дітьми захищений.
RUN_TEXT_CHILDREN = {"rPr", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen",
                     "lastRenderedPageBreak"}
# Діти w:p, які не несуть тексту і лишаються на місці.
TRANSPARENT = {"pPr", "bookmarkStart", "bookmarkEnd", "proofErr", "permStart",
               "permEnd", "commentRangeStart", "commentRangeEnd"}
# Наявні відстежені зміни: такий абзац не правимо, щоб не змішати рецензії.
TRACKED = {"ins", "del", "moveFrom", "moveTo"}
EXTRA_PARTS = re.compile(r"^word/(?:header\d*|footer\d*|footnotes|endnotes)\.xml$")
PUA = 0xE000  # атоми абзацу позначаємо символами з приватної зони Unicode
TOKEN = re.compile(r"\w+|\s+|[^\w\s]")


# --- пакет docx -------------------------------------------------------------

def read_package(path):
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        data = {info.filename: z.read(info.filename) for info in infos}
    return infos, data


def write_package(path, infos, data):
    with zipfile.ZipFile(path, "w") as z:
        for info in infos:
            z.writestr(info, data[info.filename], compress_type=info.compress_type)


def main_part(data):
    """Шлях до головного XML документа з _rels/.rels (зазвичай word/document.xml)."""
    rels = data.get("_rels/.rels")
    if rels:
        dom = minidom.parseString(rels)
        for rel in dom.getElementsByTagNameNS(REL_NS, "Relationship"):
            if rel.getAttribute("Type").endswith("/officeDocument"):
                return rel.getAttribute("Target").lstrip("/")
    return "word/document.xml"


def text_parts(data, all_parts):
    parts = [main_part(data)]
    if all_parts:
        parts += sorted(name for name in data if EXTRA_PARTS.match(name))
    return [p for p in parts if p in data]


def part_label(part, main):
    return "" if part == main else os.path.basename(part)[:-4] + ":"


# --- модель абзацу ----------------------------------------------------------

def local(node):
    return node.localName or node.nodeName.split(":")[-1]


def w_children(node):
    return [c for c in node.childNodes if c.nodeType == c.ELEMENT_NODE]


def in_fallback(node):
    parent = node.parentNode
    while parent is not None and parent.nodeType == parent.ELEMENT_NODE:
        if parent.namespaceURI == MC_NS and local(parent) == "Fallback":
            return True
        parent = parent.parentNode
    return False


def paragraphs(dom):
    return [p for p in dom.getElementsByTagNameNS(W_NS, "p") if not in_fallback(p)]


def run_text(run):
    out = []
    for child in w_children(run):
        name = local(child)
        if name == "t":
            out.append("".join(n.data for n in child.childNodes if n.nodeType == n.TEXT_NODE))
        elif name == "tab":
            out.append("\t")
        elif name in ("br", "cr"):
            out.append("\n")
        elif name == "noBreakHyphen":
            out.append("‑")
    return "".join(out)


def visible_text(node):
    """Текст захищеного фрагмента: усі w:t нащадки, без коду полів і видаленого."""
    out = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        out.append("".join(n.data for n in t.childNodes if n.nodeType == n.TEXT_NODE))
    return "".join(out)


def is_text_run(run):
    for child in w_children(run):
        name = local(child)
        if name not in RUN_TEXT_CHILDREN:
            return False
        if name == "br" and child.getAttributeNS(W_NS, "type") not in ("", "textWrapping"):
            return False  # розрив сторінки чи колонки зберігаємо як атом
    return True


def field_char(run):
    for child in w_children(run):
        if local(child) == "fldChar":
            return child.getAttributeNS(W_NS, "fldCharType")
    return None


def analyze(p):
    """Розкладає абзац на елементи: ('run', node, text) | ('atom', [nodes], text) | ('keep', node).

    Повертає (items, tracked). tracked=True, якщо в абзаці вже є відстежені зміни.
    """
    items, depth, field_nodes = [], 0, []
    tracked = any(p.getElementsByTagNameNS(W_NS, name).length for name in TRACKED)
    for child in w_children(p):
        name = local(child)
        if child.namespaceURI == W_NS and name == "r":
            kind = field_char(child)
            if kind == "begin" or depth:
                field_nodes.append(child)
                depth += 1 if kind == "begin" else 0
                depth -= 1 if kind == "end" else 0
                if depth == 0:
                    text = "".join(visible_text(n) for n in field_nodes)
                    items.append(("atom", field_nodes, text))
                    field_nodes = []
                continue
            if is_text_run(child):
                items.append(("run", child, run_text(child)))
            else:
                items.append(("atom", [child], visible_text(child)))
        elif child.namespaceURI == W_NS and name in TRANSPARENT:
            items.append(("keep", child))
        else:
            items.append(("atom", [child], visible_text(child)))
    if field_nodes:  # поле без кінця: усе, що лишилось, захищене
        items.append(("atom", field_nodes, "".join(visible_text(n) for n in field_nodes)))
    return items, tracked


def para_text(items):
    return "".join(item[2] for item in items if item[0] in ("run", "atom"))


def is_bold(run):
    for child in w_children(run):
        if local(child) == "rPr":
            for prop in w_children(child):
                if local(prop) == "b":
                    return prop.getAttributeNS(W_NS, "val") not in ("0", "false", "off")
    return False


def para_props(p, style_names):
    style, heading, listed = "", False, False
    for child in w_children(p):
        if local(child) != "pPr":
            continue
        for prop in w_children(child):
            name = local(prop)
            if name == "pStyle":
                style = prop.getAttributeNS(W_NS, "val")
            elif name == "numPr":
                listed = True
            elif name == "outlineLvl":
                heading = True
    full = style_names.get(style, style).lower()
    if full.startswith(("heading", "title", "заголовок", "назва")) or style.lower().startswith("heading"):
        heading = True
    if "list" in full or "спис" in full:
        listed = True
    return style_names.get(style, style), heading, listed


def style_names(data):
    raw = data.get("word/styles.xml")
    if not raw:
        return {}
    names = {}
    dom = minidom.parseString(raw)
    for st in dom.getElementsByTagNameNS(W_NS, "style"):
        sid = st.getAttributeNS(W_NS, "styleId")
        for name in st.getElementsByTagNameNS(W_NS, "name"):
            names[sid] = name.getAttributeNS(W_NS, "val")
    return names


def iter_paragraphs(data, all_parts=False):
    """(id, part, dom, p, items, tracked) для кожного абзацу всіх текстових частин."""
    main = main_part(data)
    for part in text_parts(data, all_parts):
        dom = minidom.parseString(data[part])
        for index, p in enumerate(paragraphs(dom)):
            items, tracked = analyze(p)
            yield f"{part_label(part, main)}{index}", part, dom, p, items, tracked


# --- extract / text ---------------------------------------------------------

def extract(path, all_parts=False):
    _, data = read_package(path)
    names = style_names(data)
    result = []
    for pid, _, _, p, items, tracked in iter_paragraphs(data, all_parts):
        text = para_text(items)
        if not text.strip():
            continue
        style, heading, listed = para_props(p, names)
        entry = {"id": pid, "text": text}
        if style:
            entry["style"] = style
        if heading:
            entry["heading"] = True
        if listed:
            entry["list"] = True
        locked = [item[2] for item in items if item[0] == "atom" and item[2]]
        if locked:
            entry["locked"] = locked
        if tracked:
            entry["tracked"] = True
        result.append(entry)
    return {"file": os.path.basename(path), "paragraphs": result}


def bold_markdown(text):
    """« слово » → « **слово** »: зірочки всередині пробілів, як вимагає Markdown."""
    core = text.strip()
    if not core:
        return text
    lead = text[:len(text) - len(text.lstrip())]
    tail = text[len(text.rstrip()):]
    return f"{lead}**{core}**{tail}"


def render_markdown(path, all_parts=False):
    """Текст документа у вигляді Markdown: заголовки з #, списки з -, жирний у **."""
    _, data = read_package(path)
    names = style_names(data)
    blocks = []
    for _, _, _, p, items, _ in iter_paragraphs(data, all_parts):
        pieces = []  # (текст, жирний)
        for item in items:
            if item[0] == "run":
                pieces.append((item[2], is_bold(item[1])))
            elif item[0] == "atom":
                pieces.append((item[2], False))
        merged = []
        for text, bold in pieces:
            if merged and merged[-1][1] == bold:
                merged[-1] = (merged[-1][0] + text, bold)
            else:
                merged.append((text, bold))
        line = "".join(bold_markdown(t) if b else t for t, b in merged).strip()
        if not line:
            continue
        _, heading, listed = para_props(p, names)
        if heading:
            line = "# " + line.replace("**", "")
        elif listed:
            line = "- " + line
        blocks.append(line)
    return "\n\n".join(blocks) + "\n"


# --- apply ------------------------------------------------------------------

def tokens(s):
    return [(m.start(), m.group()) for m in TOKEN.finditer(s)]


class Skip(Exception):
    pass


def merge_opcodes(opcodes, old_tok):
    """Зливає правки, між якими лише пробіл: «Варто Обрати» → одна заміна, а не дві.

    Так рецензія в Word читається фразами, а не окремими словами.
    """
    merged = []
    for index, (op, i1, i2, j1, j2) in enumerate(opcodes):
        is_gap = (op == "equal" and merged and merged[-1][0] != "equal"
                  and index + 1 < len(opcodes) and opcodes[index + 1][0] != "equal"
                  and all(not t.strip() for _, t in old_tok[i1:i2]))
        if merged and merged[-1][0] != "equal" and (op != "equal" or is_gap):
            _, a1, _, b1, _ = merged[-1]
            merged[-1] = ("replace", a1, i2, b1, j2)
        else:
            merged.append((op, i1, i2, j1, j2))
    return merged


def plan_edit(items, new_text):
    """Порівнює старий і новий текст абзацу. Повертає потік операцій.

    Операції: ('keep'|'del'|'ins', текст, run-власник) і ('atom', номер атома).
    """
    owners, old_chars, atoms = [], [], []
    for item in items:
        if item[0] == "run":
            old_chars.extend(item[2])
            owners.extend([("run", item[1])] * len(item[2]))
        elif item[0] == "atom":
            old_chars.append(chr(PUA + len(atoms)))
            owners.append(("atom", len(atoms)))
            atoms.append(item[2])
    old = "".join(old_chars)

    # У новому тексті шукаємо текст захищених фрагментів по черзі.
    new, pos = new_text, 0
    for index, text in enumerate(atoms):
        if not text:
            continue
        found = new.find(text, pos)
        if found < 0:
            raise Skip(f"захищений фрагмент «{text[:40]}» змінено або видалено — правте вручну")
        marker = chr(PUA + index)
        new = new[:found] + marker + new[found + len(text):]
        pos = found + 1

    old_tok, new_tok = tokens(old), tokens(new)
    matcher = difflib.SequenceMatcher(None, [t for _, t in old_tok], [t for _, t in new_tok],
                                      autojunk=False)
    stream, last_run = [], None
    first_run = next((o[1] for o in owners if o[0] == "run"), None)

    def old_range(i1, i2, kind):
        nonlocal last_run
        if i1 == i2:
            return
        start = old_tok[i1][0]
        end = old_tok[i2 - 1][0] + len(old_tok[i2 - 1][1])
        for offset in range(start, end):
            owner = owners[offset]
            if owner[0] == "atom":
                if kind == "del" and atoms[owner[1]]:
                    raise Skip(f"захищений фрагмент «{atoms[owner[1]][:40]}» зачеплено правкою — правте вручну")
                stream.append(("atom", owner[1]))
            else:
                stream.append((kind, old[offset], owner[1]))
                last_run = owner[1]

    def new_range(j1, j2):
        for _, text in new_tok[j1:j2]:
            for ch in text:
                if PUA <= ord(ch) < PUA + len(atoms):
                    raise Skip("захищений фрагмент переставлено — правте вручну")
                stream.append(("ins", ch, last_run or first_run))

    for op, i1, i2, j1, j2 in merge_opcodes(matcher.get_opcodes(), old_tok):
        if op == "equal":
            old_range(i1, i2, "keep")
        elif op == "delete":
            old_range(i1, i2, "del")
        elif op == "insert":
            new_range(j1, j2)
        else:
            old_range(i1, i2, "del")
            new_range(j1, j2)
    return stream


class Builder:
    def __init__(self, dom, author, date, ids):
        self.dom, self.author, self.date, self.ids = dom, author, date, ids
        self.prefix = dom.documentElement.prefix or "w"

    def el(self, name):
        return self.dom.createElementNS(W_NS, f"{self.prefix}:{name}")

    def attr(self, node, name, value):
        node.setAttributeNS(W_NS, f"{self.prefix}:{name}", value)

    def run(self, owner, text, deleted=False):
        r = self.el("r")
        if owner is not None:
            for child in w_children(owner):
                if local(child) == "rPr":
                    r.appendChild(child.cloneNode(True))
        buf = []

        def flush():
            if buf:
                t = self.el("delText" if deleted else "t")
                t.setAttributeNS(XML_NS, "xml:space", "preserve")
                t.appendChild(self.dom.createTextNode("".join(buf)))
                r.appendChild(t)
                buf.clear()

        for ch in text:
            if ch in "\t\n‑":
                flush()
                r.appendChild(self.el({"\t": "tab", "\n": "br", "‑": "noBreakHyphen"}[ch]))
            else:
                buf.append(ch)
        flush()
        return r

    def revision(self, kind, run):
        wrap = self.el(kind)
        self.attr(wrap, "id", str(next(self.ids)))
        self.attr(wrap, "author", self.author)
        self.attr(wrap, "date", self.date)
        wrap.appendChild(run)
        return wrap


def rebuild(p, items, stream, builder, review):
    """Замінює текстові runs абзацу на нові відповідно до потоку операцій."""
    atom_nodes = [item[1] for item in items if item[0] == "atom"]
    for item in items:
        if item[0] == "run":
            p.removeChild(item[1])
    groups = []  # (вид, власник, текст) або ('atom', номер)
    for op in stream:
        if op[0] == "atom":
            groups.append(op)
            continue
        kind, ch, owner = op
        if not review:
            if kind == "del":
                continue
            kind = "keep"
        if groups and groups[-1][0] == kind and groups[-1][1] is owner:
            groups[-1] = (kind, owner, groups[-1][2] + ch)
        else:
            groups.append((kind, owner, ch))

    next_atom = 0

    def place(node):
        if next_atom < len(atom_nodes):
            anchor = atom_nodes[next_atom][0]
            anchor.parentNode.insertBefore(node, anchor)
        else:
            p.appendChild(node)

    for group in groups:
        if group[0] == "atom":
            next_atom = group[1] + 1
            continue
        kind, owner, text = group
        if kind == "keep":
            place(builder.run(owner, text))
        elif kind == "del":
            place(builder.revision("del", builder.run(owner, text, deleted=True)))
        else:
            place(builder.revision("ins", builder.run(owner, text)))


def drop_paragraph(p, builder, review):
    """Видаляє абзац цілком: у рецензії — позначкою видалення, у чистовику — вузол.

    Абзац із розривом розділу (w:sectPr) і єдиний абзац комірки лишаємо порожнім:
    без них Word вважає файл пошкодженим.
    """
    ppr = next((c for c in w_children(p) if local(c) == "pPr"), None)
    if ppr is not None and any(local(c) == "sectPr" for c in w_children(ppr)):
        return
    if review:
        if ppr is None:
            ppr = builder.el("pPr")
            p.insertBefore(ppr, p.firstChild)
        rpr = next((c for c in w_children(ppr) if local(c) == "rPr"), None)
        if rpr is None:
            rpr = builder.el("rPr")
            tail = next((c for c in w_children(ppr) if local(c) in ("sectPr", "pPrChange")), None)
            ppr.insertBefore(rpr, tail) if tail is not None else ppr.appendChild(rpr)
        mark = builder.el("del")
        builder.attr(mark, "id", str(next(builder.ids)))
        builder.attr(mark, "author", builder.author)
        builder.attr(mark, "date", builder.date)
        rpr.appendChild(mark)
        return
    siblings = [c for c in w_children(p.parentNode) if c.namespaceURI == W_NS and local(c) == "p"]
    if len(siblings) > 1:
        p.parentNode.removeChild(p)


def max_revision_id(data, parts):
    top = 0
    for part in parts:
        for m in re.finditer(rb'w:id="(\d+)"', data[part]):
            top = max(top, int(m.group(1)))
    return top + 1


def serialize(dom):
    return (XML_HEADER + dom.documentElement.toxml()).encode("utf-8")


def load_edits(path):
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    if isinstance(raw, dict):
        raw = raw.get("edits", [])
    edits = {}
    for entry in raw:
        edits[str(entry["id"])] = entry
    return edits


def apply(path, edits, out_dir=None, author="humanizer-ua", all_parts=True):
    infos, data = read_package(path)
    parts = text_parts(data, all_parts)
    # Word записує в w:date місцевий час із суфіксом Z і так само його читає.
    date = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    start_id = max_revision_id(data, parts)
    outputs = {}
    report = {"applied": [], "skipped": []}
    for review in (True, False):
        out = dict(data)
        dom_by_part, builders = {}, {}
        ids = itertools.count(start_id)  # id правок унікальні в усьому документі
        for pid, part, dom, p, items, tracked in iter_paragraphs(data, all_parts):
            dom_by_part[part] = dom
            edit = edits.get(pid)
            if edit is None:
                continue
            builder = builders.setdefault(part, Builder(dom, author, date, ids))
            try:
                if tracked:
                    raise Skip("абзац уже має відстежені зміни — прийміть їх і повторіть")
                current = para_text(items)
                if "old" in edit and edit["old"] != current:
                    raise Skip("текст абзацу не збігається з extract")
                if edit["new"] == current:
                    raise Skip("без змін")
                stream = plan_edit(items, edit["new"])
                rebuild(p, items, stream, builder, review)
                if not edit["new"] and not any(item[0] == "atom" for item in items):
                    drop_paragraph(p, builder, review)
                if review:
                    report["applied"].append(pid)
            except Skip as reason:
                if review:
                    report["skipped"].append({"id": pid, "reason": str(reason)})
        for part, dom in dom_by_part.items():
            out[part] = serialize(dom)
        outputs[review] = out
    unknown = sorted(set(edits) - set(report["applied"]) - {s["id"] for s in report["skipped"]})
    for pid in unknown:
        report["skipped"].append({"id": pid, "reason": "абзацу з таким id немає"})

    stem = os.path.splitext(os.path.basename(path))[0]
    folder = out_dir or os.path.dirname(os.path.abspath(path))
    review_path = os.path.join(folder, f"{stem}.рецензія.docx")
    clean_path = os.path.join(folder, f"{stem}.чистовик.docx")
    write_package(review_path, infos, outputs[True])
    write_package(clean_path, infos, outputs[False])
    report.update({"review": review_path, "clean": clean_path})
    return report


# --- зразок і самоперевірка -------------------------------------------------

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '</Types>')
SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:settings xmlns:w="{W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/hyperlink" Target="https://example.com" TargetMode="External"/>'
    '</Relationships>')
STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
    '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    '</w:styles>')


def _r(text, bold=False):
    props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _p(inner, style=None):
    props = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{props}{inner}</w:p>"


def build_docx(paragraph_xml):
    """Мінімальний коректний docx із заданими абзацами (рядки XML w:p)."""
    body = "".join(paragraph_xml)
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:mc="{MC_NS}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
        f'<w:body>{body}<w:sectPr/></w:body></w:document>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/document.xml", document)
    return buf.getvalue()


DEMO = [
    _p(_r("Чому Варто Обрати Нашу Платформу"), "Heading1"),
    _p(_r("Наша платформа — це не просто сервіс, а справжній партнер для бізнесу. ")
       + _r("Давайте розберемося,", bold=True) + _r(" чому вже понад 300 компаній приймають участь у програмі.")),
    _p(_r("В сучасному світі автоматизація являється ключовим фактором успіху. "
          "Існуючі рішення не забезпечують необхідної швидкості, надійності та гнучкості.")),
    _p(_r("Результат: витрати скорочуються на 30 %. Детальніше — на ")
       + '<w:hyperlink r:id="rId9">' + _r("нашому сайті") + "</w:hyperlink>" + _r(".")),
    _p(_r("Без зайвих налаштувань. Без прихованих платежів. Лише результат.")),
    _p(_r("Підсумовуючи, можна сказати, що співпраця з нами — це інвестиція у Ваше майбутнє. "
          "Сподіваємося, це було корисно!")),
]


def accept_all(xml_bytes):
    """Імітує «Прийняти всі зміни»: прибирає w:del, розгортає w:ins."""
    dom = minidom.parseString(xml_bytes)
    for node in list(dom.getElementsByTagNameNS(W_NS, "del")):
        node.parentNode.removeChild(node)
    for node in list(dom.getElementsByTagNameNS(W_NS, "ins")):
        for child in list(node.childNodes):
            node.parentNode.insertBefore(child, node)
        node.parentNode.removeChild(node)
    return [para_text(analyze(p)[0]) for p in paragraphs(dom)]


def self_test():
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "зразок.docx")
        with open(src, "wb") as f:
            f.write(build_docx(DEMO))

        info = extract(src)
        texts = {p["id"]: p for p in info["paragraphs"]}
        assert texts["0"]["heading"], texts["0"]
        assert texts["3"]["locked"] == ["нашому сайті"], texts["3"]
        assert "**Давайте розберемося,**" in render_markdown(src)

        edits_path = os.path.join(tmp, "правки.json")
        edits = [
            {"id": "0", "new": "Чому варто обрати нашу платформу"},
            {"id": "1", "new": "Понад 300 компаній уже беруть участь у програмі."},
            {"id": "3", "new": "Витрати скорочуються на 30 %. Детальніше — на нашому сайті."},
            {"id": "4", "new": "Без налаштувань і прихованих платежів."},
            {"id": "5", "new": "Співпраця з нами окупається за пів року."},
            {"id": "99", "new": "немає такого"},
        ]
        with open(edits_path, "w", encoding="utf-8") as f:
            json.dump(edits, f, ensure_ascii=False)
        report = apply(src, load_edits(edits_path), out_dir=tmp)
        assert set(report["applied"]) == {"0", "1", "3", "4", "5"}, report
        assert [s["id"] for s in report["skipped"]] == ["99"], report

        _, review = read_package(report["review"])
        _, clean = read_package(report["clean"])
        _, orig = read_package(src)
        doc_review, doc_clean = review["word/document.xml"], clean["word/document.xml"]
        assert b"<w:ins " in doc_review and b"<w:del " in doc_review
        assert b"<w:ins " not in doc_clean and b"<w:del " not in doc_clean
        assert b'mc:Ignorable="w14"' in doc_review and b"xmlns:w14=" in doc_review
        assert review["word/styles.xml"] == orig["word/styles.xml"]

        clean_texts = [para_text(analyze(p)[0]) for p in paragraphs(minidom.parseString(doc_clean))]
        assert clean_texts[1] == "Понад 300 компаній уже беруть участь у програмі.", clean_texts[1]
        assert clean_texts[3] == "Витрати скорочуються на 30 %. Детальніше — на нашому сайті.", clean_texts[3]
        assert accept_all(doc_review) == clean_texts, (accept_all(doc_review), clean_texts)
        assert b"<w:hyperlink" in doc_clean, "гіперпосилання має лишитися"
        # Жирний «Давайте розберемося,» видалено, решта абзацу звичайна.
        clean_info = extract(report["clean"])
        assert clean_info["paragraphs"][0]["text"] == "Чому варто обрати нашу платформу"

        # Видалення абзацу цілком: у чистовику абзацу немає, у рецензії — позначка.
        report3 = apply(src, {"2": {"id": "2", "new": ""}}, out_dir=tmp)
        assert report3["applied"] == ["2"], report3
        _, rev3 = read_package(report3["review"])
        _, cln3 = read_package(report3["clean"])
        assert b"<w:rPr><w:del " in rev3["word/document.xml"], "позначка видалення абзацу"
        cln_paras = paragraphs(minidom.parseString(cln3["word/document.xml"]))
        assert len(cln_paras) == len(DEMO) - 1, len(cln_paras)

        # Правка, що ламає захищений фрагмент, пропускається.
        broken = {"3": {"id": "3", "new": "Витрати скорочуються. Детальніше на сайті."}}
        report2 = apply(src, broken, out_dir=tmp)
        assert report2["applied"] == [] and "захищений" in report2["skipped"][0]["reason"], report2

        # Оригінал не змінено.
        _, again = read_package(src)
        assert again == orig
    print("docx_tool self-test: OK")


# --- CLI --------------------------------------------------------------------

def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if argv[0] == "--self-test":
        self_test()
        return 0
    cmd, args = argv[0], argv[1:]
    all_parts = "--all-parts" in args
    args = [a for a in args if a != "--all-parts"]

    def option(name, default=None):
        if name in args:
            i = args.index(name)
            value = args[i + 1]
            del args[i:i + 2]
            return value
        return default

    if cmd == "extract" and len(args) == 1:
        json.dump(extract(args[0], all_parts), sys.stdout, ensure_ascii=False, indent=1)
        print()
    elif cmd == "text" and len(args) == 1:
        sys.stdout.write(render_markdown(args[0], all_parts))
    elif cmd == "apply":
        out_dir = option("--out-dir")
        author = option("--author", "humanizer-ua")
        if len(args) != 2:
            print(__doc__)
            return 2
        report = apply(args[0], load_edits(args[1]), out_dir, author)
        json.dump(report, sys.stdout, ensure_ascii=False, indent=1)
        print()
    elif cmd == "demo" and len(args) == 1:
        os.makedirs(os.path.dirname(os.path.abspath(args[0])), exist_ok=True)
        with open(args[0], "wb") as f:
            f.write(build_docx(DEMO))
        print(args[0])
    else:
        print(__doc__)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
