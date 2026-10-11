"use strict";

const $ = (id) => document.getElementById(id);
const FAMILIES = [
  "Мовна норма", "Інсценізація", "Ритм за правилом", "Роздування",
  "Форматування й типографіка", "Залишки чату", "Читач і структура", "Сліди копіпасту",
];
const FAMILY_RANGES = [[1, 11], [12, 17], [18, 24], [25, 32], [33, 36], [37, 39], [40, 42]];
const VERDICTS = {
  "чисто": ["clean", "Ознак ШІ майже немає"],
  "переглянути": ["review", "Є ознаки — варто переглянути"],
  "переписати": ["rewrite", "Багато ознак ШІ"],
};
const LEVELS = { 1: "слабкі ознаки ШІ", 2: "імовірно ШІ", 3: "майже напевно ШІ" };

const state = { health: null, job: null, view: "marks", poll: null };

// --- утиліти ---------------------------------------------------------------

function esc(s) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

function patternOf(rule) {
  const m = /^§(\d+)/.exec(rule);
  return m ? Number(m[1]) : null;
}

function familyOf(rule) {
  if (rule.startsWith("A")) return 7;
  const n = patternOf(rule);
  if (n == null) return null;
  return FAMILY_RANGES.findIndex(([a, b]) => n >= a && n <= b);
}

function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return `${n} ${one}`;
  if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return `${n} ${few}`;
  return `${n} ${many}`;
}

function toast(text, error = false) {
  const el = $("toast");
  el.textContent = text;
  el.className = "toast" + (error ? " error" : "");
  el.hidden = false;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.hidden = true; }, error ? 7000 : 3500);
}

async function api(path, options = {}) {
  const res = await fetch(path, options);
  const data = await res.json().catch(() => ({ error: "сервер повернув не JSON" }));
  if (!res.ok) throw new Error(data.error || `помилка ${res.status}`);
  return data;
}

// --- завантаження ----------------------------------------------------------

async function upload(file) {
  if (!file) return;
  if (!/\.docx$/i.test(file.name)) {
    toast("Потрібен файл Word у форматі .docx. Старий .doc спершу збережіть як .docx.", true);
    return;
  }
  const drop = $("drop");
  drop.classList.add("busy");
  drop.querySelector("strong").textContent = "Читаю документ…";
  try {
    const job = await api("/api/upload", {
      method: "POST",
      headers: { "X-Filename": encodeURIComponent(file.name) },
      body: file,
    });
    state.job = job;
    state.view = "marks";
    render();
  } catch (e) {
    toast(e.message, true);
  } finally {
    drop.classList.remove("busy");
    drop.querySelector("strong").textContent = "Покладіть сюди документ Word";
  }
}

function resetToDrop() {
  stopPoll();
  state.job = null;
  $("desk").classList.remove("loaded");
  $("side").hidden = true;
  $("tabs").hidden = true;
  $("sheet").replaceChildren($("drop"));
  $("file").value = "";
}

// --- розмітка абзацу -------------------------------------------------------

function aiSpans(text, item) {
  const out = [];
  for (const q of item.quotes || []) {
    const quote = q.trim();
    if (quote.length < 3) continue;
    const at = text.indexOf(quote);
    if (at >= 0) out.push({ level: "AI", rule: `Claude: ${item.reason}`, start: at, end: at + quote.length });
  }
  return out;
}

function markup(text, spans) {
  if (!spans.length) return esc(text);
  const cuts = new Set([0, text.length]);
  spans.forEach((s) => { cuts.add(s.start); cuts.add(s.end); });
  const points = [...cuts].sort((a, b) => a - b);
  let html = "";
  for (let i = 0; i < points.length - 1; i++) {
    const a = points[i], b = points[i + 1];
    const piece = esc(text.slice(a, b));
    const cover = spans.filter((s) => s.start <= a && s.end >= b);
    if (!cover.length) { html += piece; continue; }
    const cls = cover.some((s) => s.level === "AI") ? "ai" : cover.some((s) => s.level === "ERROR") ? "err" : "warn";
    const ids = cover.map((s) => s._i).join(" ");
    html += `<mark class="${cls}" data-s="${ids}">${piece}</mark>`;
  }
  return html;
}

function marginBadges(spans) {
  const seen = new Map();
  for (const s of spans) {
    const key = s.level === "AI" ? "ШІ" : (patternOf(s.rule) ?? (s.rule.startsWith("A") ? "A" : "•"));
    const rank = s.level === "AI" ? 3 : s.level === "ERROR" ? 2 : 1;
    const prev = seen.get(key);
    if (!prev || prev.rank < rank) seen.set(key, { rank, ids: [...(prev?.ids || []), s._i] });
    else prev.ids.push(s._i);
  }
  return [...seen.entries()]
    .sort((x, y) => y[1].rank - x[1].rank || String(x[0]).localeCompare(String(y[0]), "uk", { numeric: true }))
    .map(([key, v]) => {
      const cls = v.rank === 3 ? "ai" : v.rank === 2 ? "err" : "";
      const label = key === "ШІ" ? "ШІ" : key === "A" ? "A" : key === "•" ? "•" : `§${key}`;
      return `<span class="mark-n ${cls}" data-s="${v.ids.join(" ")}">${label}</span>`;
    }).join("");
}

function paraScore(spans) {
  let score = 0;
  for (const s of spans) {
    const fam = familyOf(s.rule);
    if (s.level === "ERROR") score += fam === 0 ? 2 : 3;
    else if (s.level === "WARN") score += 1;
  }
  return score;
}

// --- рендер ----------------------------------------------------------------

let SPANS = [];

function renderSheet() {
  const job = state.job;
  const view = state.view;
  const sheet = $("sheet");
  SPANS = [];
  const analysis = view === "clean" && job.rewrite ? job.rewrite.after : job.analysis;
  const detect = view === "marks" ? job.detect : null;
  const diffs = job.rewrite?.diffs || {};
  const skipped = new Map((job.rewrite?.skipped || []).map((s) => [s.id, s.reason]));

  const html = analysis.paragraphs.map((p) => {
    const cls = ["para"];
    if (p.heading) cls.push("h");
    if (p.list) cls.push("li");
    let body, margin = "", notes = "";

    if (view === "diff") {
      const d = diffs[p.id];
      if (d) {
        body = d.ops.map(([op, t]) => op === "eq" ? esc(t) : op === "del" ? `<del>${esc(t)}</del>` : `<ins>${esc(t)}</ins>`).join("");
        const pats = d.patterns.map((n) => `<span class="mark-n err" title="Патерн ${n}">§${n}</span>`).join("");
        margin = pats;
        if (d.note) notes += `<div class="edit-note">${esc(d.note)}</div>`;
        if (!d.applied) notes += `<div class="skip-note">Не застосовано у Word: ${esc(skipped.get(p.id) || "абзац пропущено")}</div>`;
      } else {
        cls.push("dim");
        body = esc(p.text);
      }
    } else {
      const spans = (analysis.spans[p.id] || []).map((s) => ({ ...s }));
      const item = detect?.paragraphs?.[p.id];
      if (detect && !detect.verdict && !detect.checked.includes(p.id)) cls.push("pending");
      if (item) spans.push(...aiSpans(p.text, item));
      spans.forEach((s) => { s._i = SPANS.length; SPANS.push(s); });
      body = markup(p.text, spans);
      margin = marginBadges(spans);
      if (item) {
        cls.push(`ai${item.level}`);
        notes += `<div class="claude-note">${esc(LEVELS[item.level] || "")}: ${esc(item.reason)}</div>`;
      } else {
        const score = paraScore(spans);
        if (score >= 4) cls.push("hot"); else if (score >= 2) cls.push("warm");
      }
    }
    return `<div class="${cls.join(" ")}" data-id="${esc(p.id)}"><div class="margin">${margin}</div><div class="text">${body}</div>${notes}</div>`;
  }).join("");
  sheet.innerHTML = html || `<p class="text">У документі немає тексту.</p>`;
  sheet.classList.toggle("no-warn", !$("showWarn").checked);
}

function renderSide() {
  const job = state.job;
  const st = job.analysis.stats;
  $("fileName").textContent = job.name;
  $("fileMeta").textContent = `${plural(st.words, "слово", "слова", "слів")}, ${plural(job.analysis.paragraphs.length, "абзац", "абзаци", "абзаців")}`;

  const [cls, label] = VERDICTS[st.verdict] || ["review", st.verdict];
  let claude = "";
  if (job.detect && !job.detect.verdict) {
    const n = Object.keys(job.detect.paragraphs).length;
    claude = `<div class="claude-verdict"><div class="verdict-label" style="font-size:20px">Claude перевіряє…</div>
      <p>Перевірено ${job.detect.checked.length} з ${job.detect.total} абзаців, позначено ${plural(n, "абзац", "абзаци", "абзаців")}. Результати з’являються на аркуші одразу.</p></div>`;
  } else if (job.detect) {
    const n = Object.keys(job.detect.paragraphs).length;
    const ai = job.detect.verdict !== "найімовірніше людина";
    claude = `<div class="claude-verdict"><div class="verdict-label" style="font-size:20px;color:var(${ai ? "--pencil" : "--green"})">Claude: ${esc(job.detect.verdict)}</div>
      <p>Позначено ${plural(n, "абзац", "абзаци", "абзаців")}. ${job.detect.summary.map(esc).join(" ")}</p></div>`;
  }
  $("verdict").className = `card verdict ${cls}`;
  $("verdict").innerHTML = `<h2>Лінтер</h2><div class="verdict-label">${label}</div>
    <p>${plural(st.errors, "помилка", "помилки", "помилок")}, ${plural(st.warnings, "попередження", "попередження", "попереджень")} — ${String(st.score).replace(".", ",")} бала на 300 слів.</p>
    <p>Лінтер шукає ознаки, а не доводить авторство: людина теж пише «ключовий» і тріади.</p>${claude}`;

  const max = Math.max(1, ...st.families);
  $("families").innerHTML = st.families.map((n, i) => n ? `<li><span>${FAMILIES[i]}</span><span class="n">${n}</span><div class="bar-row"><span style="width:${(n / max) * 100}%"></span></div></li>` : "").join("");
  $("docNotes").innerHTML = job.analysis.document.map((d) => `<p class="doc-note">${esc(d.rule)}: ${esc(d.context)}</p>`).join("");

  const running = job.state === "running";
  const noClaude = !state.health?.claude;
  $("detectBtn").disabled = running || noClaude;
  $("rewriteBtn").disabled = running || noClaude;
  $("detectBtn").title = $("rewriteBtn").title = noClaude ? "claude CLI не знайдено — встановіть Claude Code і увійдіть" : "";

  const prog = $("progress");
  prog.hidden = !(running || job.log.length);
  const [done, total] = job.progress;
  $("fill").classList.toggle("indeterminate", running && !total);
  $("fill").style.width = total ? `${(done / total) * 100}%` : running ? "" : "100%";
  const what = job.mode === "detect" ? "Перевірка" : "Правка";
  const detectRun = job.mode === "detect" && job.detect;
  $("progressText").textContent = running
    ? (detectRun ? `${what}: ${job.detect.checked.length} з ${job.detect.total} абзаців`
      : total ? `${what}: ${done} з ${total} частин тексту` : `${what}: готую…`)
    : job.state === "error" ? `Помилка: ${job.error}` : job.state === "cancelled" ? "Скасовано" : `${what} завершена`;
  $("cancelBtn").hidden = !running;
  $("log").innerHTML = job.log.map((l) => `<li>${esc(l)}</li>`).join("");

  renderResult();
}

function renderResult() {
  const job = state.job;
  const box = $("result");
  const r = job.rewrite;
  if (!r) { box.hidden = true; return; }
  const before = job.analysis.stats, after = r.after.stats;
  const changed = Object.values(r.diffs).filter((d) => d.applied).length;
  const base = `/api/jobs/${job.id}/download`;
  box.hidden = false;
  box.innerHTML = `<h2>Готово</h2>
    <div class="compare">
      <div><b>${before.errors} → ${after.errors}</b>помилок</div>
      <div><b>${before.warnings} → ${after.warnings}</b>попереджень</div>
      <div><b>${changed}</b>абзаців змінено</div>
    </div>
    <div class="downloads">
      <a class="button primary" href="${base}/clean" download>Завантажити чистовик</a>
      <a class="button secondary" href="${base}/review" download>Завантажити з рецензуванням</a>
    </div>
    <p>У файлі з рецензуванням кожну правку видно у Word: її можна прийняти чи відхилити.</p>
    ${r.skipped.length ? `<p><b>Не застосовано:</b></p><ul>${r.skipped.map((s) => `<li>абзац ${esc(s.id)}: ${esc(s.reason)}</li>`).join("")}</ul>` : ""}
    ${r.questions.length ? `<p><b>Питання до автора:</b></p><ul>${r.questions.map((q) => `<li>${esc(q)}</li>`).join("")}</ul>` : ""}
    ${r.summary.length ? `<p>${r.summary.map(esc).join(" ")}</p>` : ""}
    ${job.cost ? `<p>Вартість викликів за прайсом API: $${job.cost.toFixed(2)} (з підпискою Claude окремо не сплачується).</p>` : ""}`;
}

function renderTabs() {
  const tabs = $("tabs");
  tabs.hidden = false;
  for (const b of tabs.querySelectorAll("button")) {
    b.classList.toggle("on", b.dataset.view === state.view);
    if (b.dataset.view !== "marks") b.disabled = !state.job.rewrite;
  }
}

function render() {
  if (!state.job) return resetToDrop();
  $("desk").classList.add("loaded");
  $("side").hidden = false;
  renderTabs();
  renderSheet();
  renderSide();
}

// --- Claude ----------------------------------------------------------------

async function start(mode) {
  const job = state.job;
  const body = JSON.stringify({ genre: $("genre").value, model: $("model").value, only_flagged: $("onlyFlagged").checked });
  try {
    state.job = await api(`/api/jobs/${job.id}/${mode}`, { method: "POST", body, headers: { "Content-Type": "application/json" } });
    state._checked = 0;
    render();
    startPoll();
  } catch (e) {
    toast(e.message, true);
  }
}

function startPoll() {
  stopPoll();
  state.poll = setInterval(async () => {
    try {
      const job = await api(`/api/jobs/${state.job.id}`);
      const finished = job.state !== "running";
      const hadRewrite = !!state.job.rewrite;
      state.job = job;
      if (finished) {
        stopPoll();
        if (job.state === "done" && job.mode === "rewrite" && !hadRewrite) state.view = "diff";
        if (job.state === "done") toast(job.mode === "detect" ? "Глибоку перевірку завершено" : "Текст виправлено — можна завантажити");
        if (job.state === "error") toast(job.error, true);
        render();
      } else {
        const before = state._checked || 0;
        const now = job.detect?.checked?.length || 0;
        state._checked = now;
        if (job.mode === "detect" && now !== before && state.view === "marks") {
          const y = window.scrollY;
          renderSheet();
          window.scrollTo(0, y);
        }
        renderSide();
      }
    } catch (e) {
      stopPoll();
      toast(e.message, true);
    }
  }, 1000);
}

function stopPoll() {
  if (state.poll) clearInterval(state.poll);
  state.poll = null;
}

// --- підказки --------------------------------------------------------------

function showTip(target) {
  const ids = (target.dataset.s || "").split(" ").filter(Boolean).map(Number);
  const spans = ids.map((i) => SPANS[i]).filter(Boolean);
  if (!spans.length) return hideTip();
  const seen = new Set();
  const rows = spans.filter((s) => !seen.has(s.rule) && seen.add(s.rule)).map((s) => {
    const lv = s.level === "AI" ? "ШІ" : s.level === "ERROR" ? "Помилка" : "Ознака";
    return `<div><span class="lv">${lv}.</span> ${esc(s.rule)}</div>`;
  });
  const tip = $("tip");
  tip.innerHTML = rows.join("");
  tip.hidden = false;
  const r = target.getBoundingClientRect();
  const w = tip.offsetWidth, h = tip.offsetHeight;
  let x = Math.min(Math.max(8, r.left), window.innerWidth - w - 8);
  let y = r.bottom + 6;
  if (y + h > window.innerHeight - 8) y = r.top - h - 6;
  tip.style.left = `${x}px`;
  tip.style.top = `${y}px`;
  document.querySelectorAll(".focus").forEach((el) => el.classList.remove("focus"));
  for (const i of ids) document.querySelectorAll(`[data-s~="${i}"]`).forEach((el) => el.classList.add("focus"));
}

function hideTip() {
  $("tip").hidden = true;
  document.querySelectorAll(".focus").forEach((el) => el.classList.remove("focus"));
}

// --- події -----------------------------------------------------------------

function bind() {
  const drop = $("drop");
  $("file").addEventListener("change", (e) => upload(e.target.files[0]));
  for (const ev of ["dragenter", "dragover"]) {
    document.addEventListener(ev, (e) => { e.preventDefault(); if (!state.job) drop.classList.add("over"); });
  }
  document.addEventListener("dragleave", (e) => { if (!e.relatedTarget) drop.classList.remove("over"); });
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("over");
    const file = e.dataTransfer.files[0];
    if (file && state.job?.state === "running") return toast("Зачекайте, поки завершиться поточна обробка.", true);
    upload(file);
  });
  $("another").addEventListener("click", () => {
    if (state.job?.state === "running") return toast("Спершу скасуйте поточну обробку.", true);
    resetToDrop();
    $("file").click();
  });
  $("showWarn").addEventListener("change", () => $("sheet").classList.toggle("no-warn", !$("showWarn").checked));
  $("detectBtn").addEventListener("click", () => start("detect"));
  $("rewriteBtn").addEventListener("click", () => start("rewrite"));
  $("cancelBtn").addEventListener("click", () => api(`/api/jobs/${state.job.id}/cancel`, { method: "POST" }).catch((e) => toast(e.message, true)));
  $("tabs").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-view]");
    if (!b || b.disabled) return;
    state.view = b.dataset.view;
    renderTabs();
    renderSheet();
  });
  $("sheet").addEventListener("mouseover", (e) => {
    const t = e.target.closest("[data-s]");
    if (t) showTip(t); else hideTip();
  });
  $("sheet").addEventListener("mouseleave", hideTip);
  window.addEventListener("scroll", hideTip, { passive: true });
}

async function init() {
  bind();
  const pill = $("claude");
  try {
    state.health = await api("/api/health");
    if (state.health.claude) {
      pill.className = "claude ok";
      pill.textContent = `claude ${(state.health.version || "").split(" ")[0]}`;
    } else {
      pill.className = "claude bad";
      pill.textContent = "claude не знайдено: працює лише лінтер";
    }
  } catch {
    pill.className = "claude bad";
    pill.textContent = "сервер не відповідає";
  }
}

init();
