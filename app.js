/* הלו"ז של הילדים — נשאב חי מהטבלה בגוגל שיטס.
 * מבנה הטבלה: שורה 1 ימים (תאים ממוזגים), שורה 2 שמות ילדים, עמודה A שעות,
 * תאים ממוזגים לאורך = משך הזמן, ולרוחב = כמה ילדים. שורות בלי שעה בתחתית = הערות.
 * כתובת הטבלה לא נמצאת בקוד: היא נשמרת ב-Supabase מאחורי הקוד המשפחתי. */
"use strict";

const SB_URL = "https://mjeupkuecrbkswpvhjvl.supabase.co";
const SB_KEY = "sb_publishable_FDH64uyNtZ99o1pv_ZxpUg_AuBeG2B6";
const TRIP_ID = "london-2026-11";
const TZ = "Asia/Jerusalem";
const PX_PER_MIN = 2.4;
const REFRESH_MS = 60 * 1000;

const KID_ALIASES = [
  { id: "gal", names: ["גל", "גולי"], label: "גולי", col: "#2e7d4f", soft: "#e3f2e8" },
  { id: "or", names: ["אור", "אורי"], label: "אורי", col: "#7b4fb0", soft: "#efe6f8" },
  { id: "shai", names: ["שי"], label: "שי", col: "#d9480f", soft: "#fdeadf" },
];
const EXTRA_COLORS = [["#0f766e", "#dcf3f0"], ["#b45309", "#fdf0d8"], ["#be185d", "#fce4ef"]];
const SHARED = { col: "#1450b4", soft: "#e6eefb" };
const DAY_NAMES = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];

const $ = (id) => document.getElementById(id);
const LS = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
  del(k) { try { localStorage.removeItem(k); } catch (e) {} },
};
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pad = (n) => String(n).padStart(2, "0");
const hhmm = (m) => `${pad(Math.floor(m / 60))}:${pad(m % 60)}`;
const rng = (a, b) => `<span dir="ltr" class="rng">${hhmm(a)}–${hhmm(b)}</span>`;

const S = { code: null, sheetId: null, model: null, dayIndex: 0, dayChosen: false, fetchedAt: null, live: false };

/* ---------- זמן ---------- */
function israelNow() {
  const q = new URLSearchParams(location.search).get("now");
  if (q && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(q)) {
    return { date: q.slice(0, 10), minutes: +q.slice(11, 13) * 60 + +q.slice(14, 16), fake: true };
  }
  const f = new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
  const p = Object.fromEntries(f.formatToParts(new Date()).map((x) => [x.type, x.value]));
  return { date: `${p.year}-${p.month}-${p.day}`, minutes: +p.hour * 60 + +p.minute };
}
function londonLabel(n) {
  if (n.fake) return hhmm((n.minutes - 120 + 1440) % 1440); // בנובמבר לונדון שעתיים אחרי ישראל
  return new Intl.DateTimeFormat("he-IL", { timeZone: "Europe/London", hour: "2-digit", minute: "2-digit", hourCycle: "h23" }).format(new Date());
}
const dayName = (date) => DAY_NAMES[new Date(date + "T12:00:00Z").getUTCDay()];
const dayDM = (date) => { const d = new Date(date + "T12:00:00Z"); return `${d.getUTCDate()}.${d.getUTCMonth() + 1}`; };

/* ---------- xlsx בדפדפן ---------- */
async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function unzipXlsx(buf) {
  const bytes = new Uint8Array(buf), view = new DataView(buf);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= 0; i--) if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error("not a zip");
  const count = view.getUint16(eocd + 10, true);
  let off = view.getUint32(eocd + 16, true);
  const files = {}, utf8 = new TextDecoder("utf-8");
  for (let i = 0; i < count; i++) {
    const method = view.getUint16(off + 10, true), csize = view.getUint32(off + 20, true);
    const nameLen = view.getUint16(off + 28, true), extraLen = view.getUint16(off + 30, true), commentLen = view.getUint16(off + 32, true);
    const localOff = view.getUint32(off + 42, true);
    const name = utf8.decode(bytes.subarray(off + 46, off + 46 + nameLen));
    if (name.endsWith(".xml") || name.endsWith(".rels")) {
      const start = localOff + 30 + view.getUint16(localOff + 26, true) + view.getUint16(localOff + 28, true);
      const raw = bytes.subarray(start, start + csize);
      files[name] = utf8.decode(method === 8 ? await inflateRaw(raw) : raw);
    }
    off += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}
const decodeEntities = (s) => s
  .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
function textOfTags(xml) {
  let out = ""; const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g; let m;
  while ((m = re.exec(xml))) out += decodeEntities(m[1]);
  return out;
}
const attr = (tag, name) => { const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`)); return m ? decodeEntities(m[1]) : null; };
function colRow(ref) {
  const m = ref.match(/([A-Z]+)(\d+)/); let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { c: col - 1, r: parseInt(m[2], 10) - 1 };
}
function firstSheet(files) {
  const shared = [];
  const ss = files["xl/sharedStrings.xml"];
  if (ss) { const re = /<si>([\s\S]*?)<\/si>/g; let m; while ((m = re.exec(ss))) shared.push(textOfTags(m[1])); }
  const rels = {};
  for (const tag of (files["xl/_rels/workbook.xml.rels"] || "").match(/<Relationship\b[^>]*>/g) || []) rels[attr(tag, "Id")] = attr(tag, "Target");
  const sheetTag = ((files["xl/workbook.xml"] || "").match(/<sheet\b[^>]*>/g) || [])[0];
  let target = (rels[attr(sheetTag || "", "r:id")] || "worksheets/sheet1.xml").replace(/^\//, "");
  if (!target.startsWith("xl/")) target = "xl/" + target;
  const xml = files[target] || "";
  const cells = new Map();
  const cellRe = /<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let m;
  while ((m = cellRe.exec(xml))) {
    const head = "<c " + m[1] + ">", body = m[2] || "";
    const ref = attr(head, "r"); if (!ref) continue;
    const type = attr(head, "t"); const v = body.match(/<v>([\s\S]*?)<\/v>/);
    let text = null;
    if (type === "inlineStr") text = textOfTags(body);
    else if (v) text = type === "s" ? shared[parseInt(v[1], 10)] : decodeEntities(v[1]);
    if (text == null || !String(text).trim()) continue;
    const { r, c } = colRow(ref);
    cells.set(r + "," + c, String(text).trim());
  }
  const merges = []; const mr = /<mergeCell ref="([^"]+)"/g;
  while ((m = mr.exec(xml))) { const [a, b] = m[1].split(":"); const p = colRow(a), q = colRow(b || a); merges.push({ r1: p.r, c1: p.c, r2: q.r, c2: q.c }); }
  return { cells, merges };
}

/* ---------- פענוח הלו"ז ---------- */
const PHONE_RE = /0\d{1,2}[-\s]?\d{7}/g;
const WARN_RE = /((?:לא לשכוח|אסור לאחר)[^,.!\n]*?)(?=\s[-–]\s|[,.!\n]|$)/;

function parseTime(v) {
  if (v == null) return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return +m[1] * 60 + +m[2];
  const num = parseFloat(s);
  if (!isNaN(num) && num >= 0 && num < 1.01 && /^[\d.]+$/.test(s)) return Math.round(num * 1440);
  return null;
}
function kidFor(name, extras) {
  const n = String(name || "").trim();
  const known = KID_ALIASES.find((k) => k.names.includes(n));
  if (known) return known;
  if (!n) return null;
  let e = extras.find((x) => x.label === n);
  if (!e) { const c = EXTRA_COLORS[extras.length % EXTRA_COLORS.length]; e = { id: "k-" + n, label: n, col: c[0], soft: c[1] }; extras.push(e); }
  return e;
}

function buildModel({ cells, merges }) {
  const mergeAt = (r, c) => merges.find((m) => m.r1 === r && m.c1 === c) || { r1: r, c1: c, r2: r, c2: c };
  const year = +israelNow().date.slice(0, 4);
  // ימים
  const days = [];
  for (const [key, text] of cells) {
    const [r, c] = key.split(",").map(Number);
    if (r !== 0 || c === 0) continue;
    const dm = text.match(/(\d{1,2})[./](\d{1,2})/); if (!dm) continue;
    const m = mergeAt(r, c);
    days.push({ date: `${year}-${pad(+dm[2])}-${pad(+dm[1])}`, title: text, who: text.replace(/^.*?\d{1,2}[./]\d{1,2}(?:[./]\d{2,4})?/, "").replace(/^[\s\-–:,|]+/, "").trim(), c1: m.c1, c2: m.c2, items: [] });
  }
  days.sort((a, b) => a.c1 - b.c1);
  // ילדים לפי עמודה
  const extras = [], kidOfCol = {};
  for (const [key, text] of cells) {
    const [r, c] = key.split(",").map(Number);
    if (r === 1 && c > 0) kidOfCol[c] = kidFor(text, extras);
  }
  const kidsUsed = []; const seenKid = new Set();
  Object.keys(kidOfCol).map(Number).sort((a, b) => a - b).forEach((c) => { const k = kidOfCol[c]; if (k && !seenKid.has(k.id)) { seenKid.add(k.id); kidsUsed.push(k); } });
  // שעות בעמודה A
  const rowTime = {};
  let maxR = 0;
  for (const [key, text] of cells) { const [r, c] = key.split(",").map(Number); maxR = Math.max(maxR, r); if (c === 0 && r >= 2) { const t = parseTime(text); if (t != null) rowTime[r] = t; } }
  const timeRows = Object.keys(rowTime).map(Number).sort((a, b) => a - b);
  const lastTimeRow = timeRows[timeRows.length - 1] ?? 1;
  const endOf = (r2) => rowTime[r2 + 1] ?? ((rowTime[r2] ?? 0) + 30);
  // פריטים
  for (const [key, text] of cells) {
    const [r, c] = key.split(",").map(Number);
    if (c === 0 || r < 2 || r > lastTimeRow || rowTime[r] == null && !timeRows.some((t) => t > r)) continue;
    const day = days.find((d) => c >= d.c1 && c <= d.c2); if (!day) continue;
    const m = mergeAt(r, c);
    let start = rowTime[m.r1];
    if (start == null) { const prev = timeRows.filter((t) => t < m.r1).pop(); start = prev != null ? rowTime[prev] : 0; }
    const end = Math.max(endOf(m.r2), start + 15);
    const kids = [];
    for (let cc = m.c1; cc <= Math.min(m.c2, day.c2); cc++) { const k = kidOfCol[cc]; if (k && !kids.includes(k.id)) kids.push(k.id); }
    if (!kids.length) continue;
    day.items.push({ id: `${r},${c}`, start, end, kids, text, phones: (text.match(PHONE_RE) || []), warn: (text.match(WARN_RE) || [])[1] || "" });
  }
  // איחוד תאים זהים באותו זמן (למשל "רטלין" אצל שני ילדים)
  for (const d of days) {
    const merged = [];
    for (const it of d.items.sort((a, b) => a.start - b.start || a.id.localeCompare(b.id))) {
      const same = merged.find((x) => x.start === it.start && x.end === it.end && x.text === it.text);
      if (same) it.kids.forEach((k) => { if (!same.kids.includes(k)) same.kids.push(k); });
      else merged.push(it);
    }
    d.items = merged;
  }
  // הערות: שורות בלי שעה מתחת ללו"ז
  const notes = [];
  for (let r = lastTimeRow + 1; r <= maxR; r++) {
    const parts = [];
    for (const [key, text] of cells) { const [rr, cc] = key.split(",").map(Number); if (rr === r) parts.push([cc, text]); }
    parts.sort((a, b) => a[0] - b[0]);
    const t = parts.map((p) => p[1]).join(" ").trim();
    if (t) notes.push(t);
  }
  // טלפונים מכל הטבלה
  const contacts = [], seenPh = new Set();
  const collect = (text) => {
    for (const seg of text.split(/[|\n]/)) {
      const phs = seg.match(PHONE_RE); if (!phs) continue;
      for (const ph of phs) {
        const digits = ph.replace(/\D/g, ""); if (seenPh.has(digits)) continue; seenPh.add(digits);
        const clean = (x) => x.replace(/[:.,]+/g, " ").replace(/\s+/g, " ").trim();
        const parts = seg.slice(0, seg.indexOf(ph)).split(/\s[-–]\s|,|טלפון/).map(clean).filter(Boolean);
        const name = parts.pop() || "";
        let label = parts.length && parts[0] !== name ? `${name} (${parts[0]})` : name;
        contacts.push({ label: label || "טלפון", phone: digits.length === 10 ? `${digits.slice(0, 3)}-${digits.slice(3)}` : ph });
      }
    }
  };
  days.forEach((d) => d.items.forEach((i) => collect(i.text)));
  notes.forEach(collect);
  return { days: days.filter((d) => d.items.length || true), kids: kidsUsed, notes, contacts };
}

/* ---------- נתונים ---------- */
async function rpc(fn, body) {
  const r = await fetch(`${SB_URL}/rest/v1/rpc/${fn}`, { method: "POST", headers: { apikey: SB_KEY, "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error("http " + r.status);
  return r.json();
}
async function unlock(code) {
  const res = await rpc("trip_get", { p_id: TRIP_ID, p_code: code });
  if (!res.ok) return false;
  S.code = code; S.sheetId = res.state.sheetId;
  LS.set("lks_code", code); LS.set("lks_sheet", S.sheetId);
  return true;
}
async function fetchSheet() {
  const res = await fetch(`https://docs.google.com/spreadsheets/d/${S.sheetId}/export?format=xlsx&t=${Date.now()}`, { cache: "no-store" });
  if (!res.ok) throw new Error("sheet " + res.status);
  const model = buildModel(firstSheet(await unzipXlsx(await res.arrayBuffer())));
  if (!model.days.length) throw new Error("no days");
  return model;
}
async function refresh(scroll) {
  try {
    const model = await fetchSheet();
    S.model = model; S.fetchedAt = new Date(); S.live = true;
    LS.set("lks_model", JSON.stringify({ model, at: S.fetchedAt.toISOString() }));
  } catch (e) {
    S.live = false;
    if (!S.model) { const c = JSON.parse(LS.get("lks_model") || "null"); if (c) { S.model = c.model; S.fetchedAt = new Date(c.at); } }
  }
  if (S.model) { render(); if (scroll) setTimeout(scrollToNow, 300); }
}

/* ---------- תצוגה ---------- */
const kidById = (id) => S.model.kids.find((k) => k.id === id) || KID_ALIASES.find((k) => k.id === id);
const colorFor = (kids) => (kids.length === 1 ? kidById(kids[0]) : SHARED);
const kidNames = (kids) => {
  const names = S.model.kids.filter((k) => kids.includes(k.id)).map((k) => k.label);
  if (names.length === S.model.kids.length && names.length > 2) return "כל הילדים";
  return names.length > 1 ? names.slice(0, -1).join(", ") + " ו" + names[names.length - 1] : names[0] || "";
};
const cleanLabel = (t) => t.replace(PHONE_RE, "").replace(/טלפון[^:\n]*:\s*\S*\s*$/gm, "").replace(/[:\s]+$/gm, "").replace(/\n{2,}/g, "\n").trim();

function todayIndex() {
  const { date } = israelNow(); const days = S.model.days;
  const i = days.findIndex((d) => d.date === date);
  if (i >= 0) return i;
  return date < days[0].date ? 0 : days.length - 1;
}

function render() {
  const n = israelNow();
  if (!S.dayChosen) S.dayIndex = todayIndex();
  const day = S.model.days[S.dayIndex];
  $("top-sub").textContent = `יום ${dayName(day.date)} ${dayDM(day.date)}${n.fake ? " · תצוגת בדיקה" : ""}`;
  $("who-bar").hidden = !day.who;
  $("who-bar").innerHTML = day.who ? `<span class="who-ic" aria-hidden="true">🧡</span><span><small>מי עם הילדים ביום ${dayName(day.date)}</small>${esc(day.who)}</span>` : "";
  $("clocks").innerHTML = `<div class="il">${hhmm(n.minutes)}</div><div class="ld">לונדון ${londonLabel(n)}</div>`;
  renderTicker();
  renderDays(n);
  renderNow(n);
  renderTimeline(n);
  renderPanels();
  $("edit-link").href = `https://docs.google.com/spreadsheets/d/${S.sheetId}/edit`;
  const t = S.fetchedAt;
  $("sync-note").textContent = S.live
    ? `מתעדכן מהטבלה של מורן · ${pad(t.getHours())}:${pad(t.getMinutes())} · לרענון הקישו כאן`
    : "אין חיבור כרגע. מוצג הלו\"ז האחרון שנשמר בטלפון";
  $("sync-note").classList.toggle("stale", !S.live);
  const stack = document.querySelector(".sticky-stack");
  document.documentElement.style.setProperty("--stick", stack.offsetHeight + 6 + "px");
}

function renderTicker() {
  const items = S.model.notes.filter((t) => t.length <= 140);
  const el = $("ticker");
  if (!items.length) { el.hidden = true; return; }
  el.hidden = false;
  const one = items.map((t) => `<span class="ticker-item ${/לא לשכוח|חשוב|אסור/.test(t) ? "important" : ""}"><span class="dot"></span>${esc(t)}</span>`).join("");
  $("ticker-track").innerHTML = one + one;
  const len = items.join("").length;
  $("ticker-track").style.setProperty("--ticker-dur", Math.max(25, len * 0.22) + "s");
}

function renderDays(n) {
  $("days").innerHTML = S.model.days.map((d, i) => `
    <button class="day-chip ${i === S.dayIndex ? "active" : ""} ${d.date === n.date ? "today" : ""}" data-i="${i}" aria-pressed="${i === S.dayIndex}">
      <span class="d-name">${dayName(d.date)}</span><span class="d-date">${dayDM(d.date)}</span>
    </button>`).join("");
  $("days").querySelectorAll(".day-chip").forEach((b) => (b.onclick = () => {
    S.dayIndex = +b.dataset.i; S.dayChosen = true; render(); window.scrollTo({ top: 0 });
    if (S.model.days[S.dayIndex].date === israelNow().date) setTimeout(scrollToNow, 200);
  }));
}

function renderNow(n) {
  const banner = $("now-banner"); const days = S.model.days;
  const first = days[0].date, last = days[days.length - 1].date;
  if (n.date < first) {
    const diff = Math.round((new Date(first + "T12:00:00Z") - new Date(n.date + "T12:00:00Z")) / 864e5);
    banner.hidden = false;
    banner.innerHTML = `<div class="nb-head"><span class="pulse"></span>עוד מעט</div>
      <div class="nb-msg">אמא ואבא נוסעים ${diff === 1 ? "מחר" : `בעוד ${diff} ימים`}<small>הלו"ז מתחיל ביום ${dayName(first)} ${dayDM(first)}</small></div>`;
    return;
  }
  if (n.date > last) {
    banner.hidden = false;
    banner.innerHTML = `<div class="nb-msg">אמא ואבא חזרו הביתה 🎉<small>תודה ענקית על הכול!</small></div>`;
    return;
  }
  const day = days.find((d) => d.date === n.date);
  if (!day || S.model.days[S.dayIndex].date !== n.date) { banner.hidden = true; return; }
  const m = n.minutes;
  // מצב לכל ילד, ואיחוד ילדים שנמצאים באותו מקום
  const groups = [];
  for (const k of S.model.kids) {
    const mine = day.items.filter((i) => i.kids.includes(k.id));
    const cur = mine.filter((i) => i.start <= m && m < i.end).sort((a, b) => b.start - a.start)[0];
    const next = mine.filter((i) => i.start > m).sort((a, b) => a.start - b.start)[0];
    const key = cur ? "c" + cur.id : next ? "n" + next.id : "home";
    const g = groups.find((x) => x.key === key);
    if (g) g.kids.push(k); else groups.push({ key, kids: [k], cur, next });
  }
  const rows = groups.map(({ kids, cur, next }) => {
    const ids = kids.map((k) => k.id);
    const c = kids.length === 1 ? kids[0] : SHARED;
    const who = kidNames(ids);
    const style = `--kcol:${c.col};--ksoft:${c.soft}`;
    const av = kids.length === 1 ? esc(kids[0].label[0]) : `<span style="font-size:13px">${kids.map((k) => esc(k.label[0])).join("")}</span>`;
    let info;
    if (cur) {
      const pct = Math.round(((m - cur.start) / (cur.end - cur.start)) * 100);
      info = `<div class="nb-kid">${esc(who)} עכשיו</div><div class="nb-label">${esc(cleanLabel(cur.text))}</div>
        <div class="nb-time">${rng(cur.start, cur.end)}</div><div class="nb-progress"><i style="width:${pct}%"></i></div>`;
    } else if (next) {
      info = `<div class="nb-kid">${esc(who)} · בבית</div><div class="nb-label">בהמשך ב-${hhmm(next.start)}: ${esc(cleanLabel(next.text))}</div>`;
    } else {
      info = `<div class="nb-kid">${esc(who)}</div><div class="nb-label">${m >= 19 * 60 ? "בבית, ערב ושינה 😴" : "בבית"}</div>`;
    }
    const target = cur || next;
    return `<button class="nb-row" style="${style}" ${target ? `data-item="${target.id}"` : ""}><span class="nb-av">${av}</span><div class="nb-info">${info}</div></button>`;
  }).join("");
  banner.hidden = false;
  banner.innerHTML = `<div class="nb-head"><span class="pulse"></span>מה קורה עכשיו</div><div class="nb-rows">${rows}</div>`;
  banner.querySelectorAll("[data-item]").forEach((b) => (b.onclick = () => openDetail(day, b.dataset.item)));
}

function renderTimeline(n) {
  const day = S.model.days[S.dayIndex];
  const kids = S.model.kids, N = kids.length || 1;
  const items = day.items;
  const dayStart = Math.min(7 * 60, ...items.map((i) => i.start));
  const dayEnd = Math.max(20 * 60, ...items.map((i) => i.end));
  const y = (min) => (min - dayStart) * PX_PER_MIN;
  const height = y(dayEnd) + 10;
  let gutter = "", lines = "";
  for (let t = Math.ceil(dayStart / 30) * 30; t <= dayEnd; t += 30) {
    gutter += `<div class="tl-mark ${t % 60 === 0 ? "hour" : ""}" style="top:${y(t)}px">${hhmm(t)}</div>`;
    lines += `<div class="tl-line" style="top:${y(t)}px"></div>`;
  }
  const colW = 100 / N, GAP = 4; // px
  const colsBg = kids.map((k, i) => `<div class="tl-colbg" style="--kcol:${k.col};inset-inline-start:calc(${i * colW}% + ${GAP / 2}px);width:calc(${colW}% - ${GAP}px)"></div>`).join("");
  const isToday = day.date === n.date, m = n.minutes;
  const cards = items.map((it) => {
    const idx = kids.map((k, i) => (it.kids.includes(k.id) ? i : -1)).filter((i) => i >= 0);
    const a = Math.min(...idx), b = Math.max(...idx);
    const c = colorFor(it.kids);
    const h = (it.end - it.start) * PX_PER_MIN;
    const isNow = isToday && it.start <= m && m < it.end;
    const past = (isToday && m >= it.end) || day.date < n.date;
    return `<button class="act ${h < 60 ? "compact" : ""} ${isNow ? "now-active" : ""} ${past && !isNow ? "past" : ""}" data-item="${it.id}"
      style="top:${y(it.start) + 2}px;height:${Math.max(h - 4, 28)}px;inset-inline-start:calc(${a * colW}% + ${GAP / 2}px);width:calc(${(b - a + 1) * colW}% - ${GAP}px);--col:${c.col};--soft:${c.soft}"
      aria-label="${esc(hhmm(it.start) + " " + kidNames(it.kids) + ": " + it.text)}">
      ${isNow ? '<span class="now-tag">עכשיו</span>' : ""}
      <span class="a-time">${rng(it.start, it.end)}</span>
      ${it.warn ? `<span class="a-warn">⚠️ ${esc(it.warn)}</span>` : ""}
      <span class="a-label">${esc(cleanLabel(it.text))}</span>
    </button>`;
  }).join("");
  const showNow = isToday && m >= dayStart - 15 && m <= dayEnd + 15;
  $("timeline").innerHTML = `
    <div class="tl-gutter"><div class="spacer"></div><div class="tl-canvas" style="height:${height}px">${gutter}</div></div>
    <div class="tl-main">
      <div class="tl-heads" style="--n:${N}">${kids.map((k) => `<div class="tl-head" style="--kcol:${k.col};--ksoft:${k.soft}">${esc(k.label)}</div>`).join("")}</div>
      <div class="tl-canvas" style="height:${height}px">${colsBg}${lines}${cards}
        ${showNow ? `<div class="now-line" style="top:${y(m)}px"><span>${hhmm(m)}</span></div>` : ""}
      </div>
    </div>`;
  $("timeline").querySelectorAll(".act").forEach((el) => (el.onclick = () => openDetail(day, el.dataset.item)));
}

function renderPanels() {
  const { contacts, notes } = S.model;
  $("phones").hidden = !contacts.length;
  $("phones").innerHTML = `<h3>📞 טלפונים</h3>` + contacts.map((c) => `
    <div class="contact"><div class="c-info"><div class="c-name">${esc(c.label)}</div><div class="c-ph">${esc(c.phone)}</div></div>
    <a class="call" href="tel:${esc(c.phone.replace(/\D/g, ""))}">חיוג</a></div>`).join("");
  $("notes").hidden = !notes.length;
  $("notes").innerHTML = `<h3>💡 חשוב לזכור</h3>` + notes.map((t) => `<div class="note"><span>${esc(t).replace(PHONE_RE, (ph) => `<a class="ph" dir="ltr" href="tel:${ph.replace(/\D/g, "")}">${ph}</a>`)}</span></div>`).join("");
}

function scrollToNow() {
  const el = document.querySelector(".now-line") || document.querySelector(".act.now-active");
  if (!el) return;
  const top = el.getBoundingClientRect().top + window.scrollY - window.innerHeight * 0.4;
  window.scrollTo({ top: Math.max(top, 0), behavior: "smooth" });
}

/* ---------- פרטים ---------- */
function openDetail(day, id) {
  const it = day.items.find((i) => i.id === id); if (!it) return;
  const kids = S.model.kids.filter((k) => it.kids.includes(k.id));
  $("detail-sheet").innerHTML = `<div class="sheet-handle"></div>
    <div class="d-kids">${kids.map((k) => `<span class="d-kid" style="--kcol:${k.col};--ksoft:${k.soft}">${esc(k.label)}</span>`).join("")}</div>
    <div class="d-time">יום ${dayName(day.date)} · ${hhmm(it.start)} עד ${hhmm(it.end)}</div>
    <div class="d-text">${esc(it.text)}</div>
    ${it.warn ? `<div class="d-warn">⚠️ ${esc(it.warn)}</div>` : ""}
    <div class="d-actions">${it.phones.map((p) => `<a class="call" href="tel:${esc(p.replace(/\D/g, ""))}">📞 להתקשר ${esc(p)}</a>`).join("")}</div>
    <button class="d-close" id="d-close">סגירה</button>`;
  $("detail-sheet").classList.add("open"); $("sheet-backdrop").classList.add("open");
  $("d-close").onclick = closeDetail;
  $("d-close").focus({ preventScroll: true });
}
function closeDetail() { $("detail-sheet").classList.remove("open"); $("sheet-backdrop").classList.remove("open"); }
$("sheet-backdrop").onclick = closeDetail;
document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDetail(); });

/* ---------- הפעלה ---------- */
function showGate(bad) {
  $("app").hidden = true; $("gate").hidden = false; $("gate-err").hidden = !bad;
  $("gate-form").onsubmit = async (e) => {
    e.preventDefault();
    const v = $("gate-code").value.trim().toLowerCase(); if (!v) return;
    try { if (await unlock(v)) start(); else showGate(true); } catch (err) { $("gate-err").textContent = "אין חיבור לאינטרנט. נסו שוב."; $("gate-err").hidden = false; }
  };
}
function start() {
  $("gate").hidden = true; $("app").hidden = false;
  const c = JSON.parse(LS.get("lks_model") || "null");
  if (c) { S.model = c.model; S.fetchedAt = new Date(c.at); render(); }
  refresh(true);
  setInterval(() => { if (!document.hidden) refresh(false); }, REFRESH_MS);
  setInterval(() => { if (S.model) render(); }, 30 * 1000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) refresh(false); });
  $("sync-note").onclick = () => refresh(false);
}
(async function boot() {
  const m = location.hash.match(/k=([a-z0-9]+)/i);
  if (m) history.replaceState(null, "", location.pathname + location.search);
  const code = m ? m[1].toLowerCase() : LS.get("lks_code");
  const sheet = LS.get("lks_sheet");
  if (code && sheet && !m) { S.code = code; S.sheetId = sheet; start(); unlock(code).catch(() => {}); return; }
  if (!code) { showGate(false); return; }
  try { if (await unlock(code)) start(); else { LS.del("lks_code"); showGate(true); } }
  catch (e) { if (sheet) { S.code = code; S.sheetId = sheet; start(); } else showGate(false); }
})();
