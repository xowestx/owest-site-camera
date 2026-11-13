/* O-Plus Client Portal: shared helpers for index.html (client) and pm.html (team). */

// The portal's own Apps Script Web App (/exec URL). Separate from Site Camera's API.
const API_URL = "PASTE_PORTAL_EXEC_URL_HERE";

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MON_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
/** Page language: the client page can switch to Arabic; the team page stays English. */
let LANG = "en";
const AR = () => LANG === "ar";

/** "2026-11-13 09:15" -> "13 Nov 2026 · 9:15 AM"; "2026-11-13" -> "13 Nov 2026". */
function when(s, dateOnly) {
  const m = String(s || "").match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (!m) return esc(s || "");
  let out = `${+m[3]} ${(AR() ? MON_AR : MON)[+m[2] - 1]} ${m[1]}`;
  if (m[4] && !dateOnly) { const h = +m[4]; out += ` · ${h % 12 || 12}:${m[5]} ${AR() ? (h < 12 ? "ص" : "م") : (h < 12 ? "AM" : "PM")}`; }
  return out;
}

async function api(payload) {
  if (/PASTE_/.test(API_URL)) throw new Error("The portal is not connected yet (API_URL in portal.js).");
  const ctl = new AbortController(); const tm = setTimeout(() => ctl.abort(), 60000);
  let r;
  try {
    r = await fetch(API_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload), redirect: "follow", credentials: "omit", signal: ctl.signal });
  } catch (e) { throw new Error("No connection. Please check your internet and try again."); }
  finally { clearTimeout(tm); }
  const j = await r.json().catch(() => ({ ok: false, error: "Unexpected answer from the server." }));
  if (!j.ok) { const e = new Error(j.error || "Something went wrong."); e.code = j.code; throw e; }
  return j;
}

/** Status word -> brand tone (pill class). */
const TONE = {
  "Sent for approval": "waiting", "Approved": "done", "Rejected": "problem", "Replaced": "idle",
  "Under review": "active", "Replied": "waiting", "Resolved": "done",
  "Open": "active", "Fixed": "waiting", "Please check": "waiting", "Closed": "done"
};
const STATUS_AR = { "Sent for approval": "بانتظار موافقتك", "Approved": "تمت الموافقة", "Rejected": "مطلوب تعديل", "Replaced": "تم الاستبدال",
  "Under review": "قيد المراجعة", "Replied": "تم الرد", "Resolved": "تم الحل", "Open": "مفتوحة", "Fixed": "تم الإصلاح", "Please check": "برجاء المراجعة", "Closed": "مغلقة" };
const pill = s => `<span class="pill ${TONE[s] || "idle"}">${esc(AR() ? STATUS_AR[s] || s : s)}</span>`;
const CHEV = `<svg class="chev" viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>`;
const BACK = label => `<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg>${esc(label)}`;
const head = (eyebrow, title, sub) => `<div class="head"><div class="eyebrow">${esc(eyebrow)}</div><h1 class="title">${esc(title)}</h1><div class="bar"></div>${sub ? `<p class="sub">${esc(sub)}</p>` : ""}</div>`;
const STAGES_EN = ["Initiation", "Design", "Quotation", "Construction", "Quality Inspection", "Hand Over"];
const STAGES_AR = ["البداية", "التصميم", "عرض السعر", "التنفيذ", "فحص الجودة", "التسليم"];
function stepper(stage) {
  if (stage == null) return "";
  const STAGES = AR() ? STAGES_AR : STAGES_EN;
  const tick = `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>`;
  return `<div class="steps" aria-label="${AR() ? "مرحلة المشروع" : "Project stage"}: ${esc(STAGES[stage])}">${STAGES.map((s, i) =>
    `<div class="${i < stage ? "done" : i === stage ? "now" : ""}"><i>${i < stage ? tick : ""}</i>${esc(s)}</div>`).join("")}</div>`;
}

/** Item details: lines "Label: value" become a table; anything else stays a paragraph. */
function details(text) {
  const lines = String(text || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
  if (!lines.length) return "";
  const kv = lines.filter(l => /^[^:]{1,40}:\s*\S/.test(l));
  if (kv.length === lines.length) return `<div class="kv">${kv.map(l => { const i = l.indexOf(":"); return `<div><span>${esc(l.slice(0, i))}</span><span>${esc(l.slice(i + 1).trim())}</span></div>`; }).join("")}</div>`;
  return `<p class="para">${esc(lines.join("\n"))}</p>`;
}

/** Photos are fetched through the API (Drive files stay private); cached per page. */
const IMG = {};
function loadImages(root, auth) {
  root.querySelectorAll("img[data-id]").forEach(async el => {
    const id = el.dataset.id;
    if (IMG[id]) { el.src = IMG[id]; return; }
    try { const j = await api(Object.assign({ action: "portal.img", id }, auth)); IMG[id] = j.data; el.src = j.data; }
    catch (e) { el.alt = "Photo unavailable"; }
  });
}

let toastT;
function toast(t) {
  let el = $("toast");
  if (!el) { el = document.createElement("div"); el.id = "toast"; el.className = "toast"; el.setAttribute("role", "status"); document.body.appendChild(el); }
  el.textContent = t; el.hidden = false;
  clearTimeout(toastT); toastT = setTimeout(() => el.hidden = true, 3200);
}

/** Phone photo -> JPEG data URL, longest side 1280 px, about 450 KB. */
async function compress(file) {
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: "from-image" }); }
  catch (e) { const u = URL.createObjectURL(file); bmp = new Image(); bmp.src = u; await bmp.decode(); }
  const scale = Math.min(1, 1280 / Math.max(bmp.width, bmp.height));
  const c = document.createElement("canvas"); c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
  c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
  let q = .8, url = c.toDataURL("image/jpeg", q);
  while (url.length > 620000 && q > .4) { q -= .1; url = c.toDataURL("image/jpeg", q); }
  return url;
}

function sheet(html) {
  closeSheet();
  const d = document.createElement("div"); d.className = "sheet"; d.id = "sheet";
  d.innerHTML = `<div class="p" role="dialog" aria-modal="true">${html}</div>`;
  d.addEventListener("click", e => { if (e.target === d) closeSheet(); });
  document.body.appendChild(d);
  const f = d.querySelector("input,textarea"); if (f) setTimeout(() => f.focus(), 50);
  return d;
}
function closeSheet() { const s = $("sheet"); if (s) s.remove(); }

/** Snag rooms offered as suggestions (free text still allowed). */
const ROOMS = ["Entrance", "Living room", "Dining", "Kitchen", "Master bedroom", "Bedroom 2", "Bedroom 3", "Master bathroom", "Guest bathroom", "Corridor", "Balcony", "Garden", "Roof"];
const ROOMS_AR = ["المدخل", "الريسبشن", "السفرة", "المطبخ", "غرفة النوم الرئيسية", "غرفة نوم 2", "غرفة نوم 3", "حمام رئيسي", "حمام الضيوف", "الطرقة", "البلكونة", "الحديقة", "الروف"];
const snagTitle = s => `Snag #${s.no}${s.location ? " · " + s.location : ""}`;
