/**
 * O West Site Camera — backend (Google Apps Script bound to the "O West Site Camera" sheet).
 *
 * Website (GitHub Pages) -> this Web app:
 *   GET  ?action=units          units + their items, engineers, contractors, trades
 *   POST {action:"upload"}      one photo -> Drive + Photos tab
 *   POST {action:"delete"}      remove a photo of the open visit
 *   POST {action:"close"}       the "Done" button: work done + manpower + notes -> Visits + Manpower tabs
 *   POST {action:"report"}      a report PDF on demand. Daily per engineer: anyone.
 *                               Manpower (any date range) and weekly client reports: manager only.
 *   POST {action:"login"}       manager password -> token valid 30 days on that phone
 *
 * Manager password: sheet menu O West Camera > Set manager password.
 *
 * Scheduled (no Claude, no person needed):
 *   dailyJob   every day at "Daily report hour"   -> one PDF per engineer + manpower-by-contractor PDF
 *   weeklyJob  every "Weekly report day"          -> one PDF per unit (what we did this week)
 *
 * Deploy: Execute as = Me, Who has access = Anyone.
 * After pasting: run setup() once, then Deploy > Manage deployments > New version.
 */

/* =========================== config =========================== */

const TZ = 'Africa/Cairo';
const T = {                    // tab names
  units: 'Units', scope: 'Scope', engineers: 'Engineers', contractors: 'Contractors',
  trades: 'Trades', unitContractors: 'Unit Contractors', settings: 'Settings', photos: 'Photos', visits: 'Visits',
  manpower: 'Manpower', reports: 'Reports'
};
const PHOTO_HEADERS = ['Photo ID', 'Visit ID', 'Uploaded at', 'Date', 'Unit', 'Project', 'Item', 'Note', 'By', 'File name', 'Photo link', 'File ID', 'Status'];
const VISIT_HEADERS = ['Visit ID', 'Date', 'Unit', 'Project', 'Items', 'By', 'Started', 'Closed', 'Photos', 'Status', 'Folder link', 'Work done', 'Manpower total', 'Notes'];
const MP_HEADERS = ['Visit ID', 'Date', 'Unit', 'Project', 'Contractor', 'Trade', 'Count', 'By', 'Status'];
const REPORT_HEADERS = ['Made at', 'Report', 'Period', 'For', 'PDF link', 'Sent to'];
const PHOTOS_ROOT = 'O West Site Photos';
const REPORTS_ROOT = 'O West Reports';
const GENERAL = 'General';
const MAX_BASE64_CHARS = 8 * 1024 * 1024;

/* =========================== web API =========================== */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'units';
    if (action === 'units') return json_(Object.assign({ ok: true }, bootData_()));
    if (action === 'ping') return json_({ ok: true });
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) { return json_({ ok: false, error: msg_(err) }); }
}

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'upload') return json_(Object.assign({ ok: true }, upload_(b)));
    if (b.action === 'delete') return json_(Object.assign({ ok: true }, deletePhoto_(b)));
    if (b.action === 'close') return json_(Object.assign({ ok: true }, closeVisit_(b)));
    if (b.action === 'report') return json_(Object.assign({ ok: true }, report_(b)));
    if (b.action === 'login') return json_(Object.assign({ ok: true }, login_(b)));
    if (b.action === 'sync') { requireAdmin_(b.token); return json_(Object.assign({ ok: true }, syncClickUp_())); }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) { return json_({ ok: false, error: msg_(err), code: err && err.code }); }
}

/* =========================== manager access =========================== */

const TOKEN_DAYS = 30;

function login_(b) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_HASH');
  if (!hash) throw new Error('No manager password yet. Set it from the sheet menu: O West Camera > Set manager password.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('loginFails') || 0);
  if (fails >= 5) throw new Error('Too many wrong tries. Wait 10 minutes.');
  if (sha_(str_(b.password) + props.getProperty('ADMIN_SALT')) !== hash) {
    cache.put('loginFails', String(fails + 1), 600);
    throw new Error('Wrong password.');
  }
  cache.remove('loginFails');
  const exp = Date.now() + TOKEN_DAYS * 86400000;
  return { token: exp + '.' + sign_(String(exp)), expires: exp };
}

function requireAdmin_(token) {
  const t = str_(token), i = t.indexOf('.');
  const exp = Number(t.slice(0, i));
  if (i < 0 || !exp || exp < Date.now() || sign_(String(exp)) !== t.slice(i + 1)) {
    const e = new Error('Manager access needed. Unlock with the password.'); e.code = 'auth'; throw e;
  }
}

function sign_(v) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(v, key)).replace(/=+$/, '');
}
function sha_(v) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8)); }

/** Sheet menu. A new password also logs out every phone that was unlocked before. */
function menuSetPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Manager password', 'Type the new manager password (at least 6 characters).\nPhones unlocked with the old password will need the new one.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText().trim();
  if (pw.length < 6) return ui.alert('Too short. Use at least 6 characters.');
  const props = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid();
  props.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: sha_(pw + salt), ADMIN_SECRET: Utilities.getUuid() + Utilities.getUuid() });
  ui.alert('Manager password saved.');
}

/* =========================== reports on demand (website) =========================== */

function report_(b) {
  const type = str_(b.type);
  let file, label, period;
  if (type === 'daily') {
    const day = cleanDay_(b.day), eng = str_(b.engineer);
    if (!eng) throw new Error('Choose the engineer.');
    file = dailyEngineerPdf_(day, eng);
    if (!file) throw new Error(`No finished visits for ${eng} on ${day}.`);
    label = 'Daily engineer'; period = day;
  } else if (type === 'manpower') {
    requireAdmin_(b.token);
    const from = cleanDay_(b.from), to = cleanDay_(b.to || b.from);
    if (to < from) throw new Error('The end date is before the start date.');
    if ((parseDay_(to) - parseDay_(from)) / 86400000 > 92) throw new Error('Choose 3 months or less.');
    file = manpowerPdf_(from, to);
    if (!file) throw new Error(`No finished visits between ${from} and ${to}.`);
    label = 'Manpower'; period = from === to ? from : `${from} → ${to}`;
  } else if (type === 'weekly') {
    requireAdmin_(b.token);
    const end = cleanDay_(b.end), u = findUnit_(b.unit).unit;
    file = weeklyUnitPdf_(u, end);
    if (!file) throw new Error(`No finished visits on ${u} in the 7 days up to ${end}.`);
    label = 'Weekly client'; period = `${fmt_(new Date(parseDay_(end).getTime() - 6 * 86400000), 'yyyy-MM-dd')} → ${end}`;
  } else throw new Error('Unknown report.');
  logReport_(label, period, b.engineer || b.unit || 'All', file, '(opened on website)');
  return { name: file.getName(), pdf: Utilities.base64Encode(file.getBlob().getBytes()) };
}

function dailyEngineerPdf_(day, eng) {
  const list = closedVisits_(v => v.date === day && v.by.toLowerCase() === eng.toLowerCase()).sort((a, b) => a.closed.localeCompare(b.closed));
  if (!list.length) return null;
  const html = dailyHtml_(settings_(), day, list[0].by, list, photosByVisit_(list), crewByVisit_(list));
  return savePdf_(html, reportsFolder_('Daily', day), `Daily ${day} - ${list[0].by}.pdf`);
}

function manpowerPdf_(from, to) {
  const visits = closedVisits_(v => v.date >= from && v.date <= to);
  if (!visits.length) return null;
  const html = manpowerHtml_(settings_(), from, to, visits, crewByVisit_(visits));
  return savePdf_(html, reportsFolder_('Manpower', to), from === to ? `Manpower ${from}.pdf` : `Manpower ${from} to ${to}.pdf`);
}

function weeklyUnitPdf_(unit, endDay) {
  const startDay = fmt_(new Date(parseDay_(endDay).getTime() - 6 * 86400000), 'yyyy-MM-dd');
  const list = closedVisits_(v => v.unit.toLowerCase() === unit.toLowerCase() && v.date >= startDay && v.date <= endDay).sort((a, b) => a.closed.localeCompare(b.closed));
  if (!list.length) return null;
  const cl = rows_(T.units).find(r => low_(r['Unit']) === unit.toLowerCase()) || {};
  const html = weeklyHtml_(settings_(), unit, { name: str_(cl['Client name']) }, startDay, endDay, list, photosByVisit_(list), crewByVisit_(list));
  return savePdf_(html, reportsFolder_('Weekly', endDay), `Weekly ${endDay} - ${safeUnit_(unit)}.pdf`);
}

function cleanDay_(v) {
  const s = str_(v);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Choose a date.');
  return s;
}

/* =========================== lists for the website =========================== */

function bootData_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('boot');
  if (hit) return JSON.parse(hit);

  const scope = {};
  rows_(T.scope).forEach(r => {
    const u = low_(r['Unit']), item = str_(r['Item']);
    if (!u || !item) return;
    (scope[u] = scope[u] || []).indexOf(item) < 0 && scope[u].push(item);
  });
  // contractors on each unit: assigned in ClickUp first, then ones used before in the app
  const alias = aliases_();
  const used = {};
  const add = (u, c) => { c = alias[c.toLowerCase()] || c; (used[u] = used[u] || []).indexOf(c) < 0 && used[u].push(c); };
  rows_(T.unitContractors).forEach(r => { const u = low_(r['Unit']), c = str_(r['Contractor']); if (u && c) add(u, c); });
  rows_(T.manpower).forEach(r => {
    if (str_(r['Status']) === 'Deleted') return;
    const u = low_(r['Unit']), c = str_(r['Contractor']);
    if (u && c) add(u, c);
  });
  const units = rows_(T.units)
    .filter(r => str_(r['Unit']) && low_(r['Active']) !== 'no')
    .map(r => {
      const unit = str_(r['Unit']);
      return { unit: unit, project: str_(r['Project']), items: scope[unit.toLowerCase()] || [], contractors: used[unit.toLowerCase()] || [] };
    });
  const out = {
    units: units,
    engineers: rows_(T.engineers).filter(r => str_(r['Name']) && low_(r['Active']) !== 'no').map(r => str_(r['Name'])),
    contractors: rows_(T.contractors).filter(r => str_(r['Contractor']) && low_(r['Active']) !== 'no' && !str_(r['Same as'])).map(r => str_(r['Contractor'])).sort(),
    trades: rows_(T.trades).filter(r => str_(r['Trade'])).map(r => ({ en: str_(r['Trade']), ar: str_(r['Arabic']) }))
  };
  try { cache.put('boot', JSON.stringify(out), 600); } catch (e) {}
  return out;
}

function findUnit_(name) {
  const k = low_(name);
  const u = bootData_().units.find(x => x.unit.toLowerCase() === k);
  if (!u) throw new Error('This unit is not in the active list. Reload the page.');
  return u;
}

function clearBoot_() { CacheService.getScriptCache().remove('boot'); }

/** Editing any list tab refreshes what the app sees. */
function onEdit(e) {
  try {
    const n = e.range.getSheet().getName();
    if ([T.units, T.scope, T.engineers, T.contractors, T.trades].indexOf(n) >= 0) clearBoot_();
  } catch (err) {}
}

/* =========================== upload / delete =========================== */

function upload_(p) {
  const photoId = cleanId_(p.photoId, 'P');
  const visitId = cleanId_(p.visitId, 'V');
  const cache = CacheService.getScriptCache();
  const done = cache.get('p:' + photoId);            // same photo sent twice after a slow network
  if (done) return JSON.parse(done);

  if (typeof p.base64 !== 'string' || !p.base64) throw new Error('No photo received.');
  if (p.base64.length > MAX_BASE64_CHARS) throw new Error('Photo is too large.');

  const u = findUnit_(p.unit);
  const item = u.items.indexOf(p.item) >= 0 ? p.item : GENERAL;
  const tag = item.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
  const safeUnit = safeUnit_(u.unit);
  const by = str_(p.by).slice(0, 60);
  const note = str_(p.note).slice(0, 300);

  const now = new Date();
  const day = fmt_(now, 'yyyy-MM-dd');
  const folder = dayFolder_(safeUnit, day);
  const name = `${safeUnit}_${day}_${fmt_(now, 'HHmmss')}_${tag}_${photoId.slice(-5)}.jpg`;
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.base64), 'image/jpeg', name));
  file.setDescription([u.unit, item, note, by && ('By ' + by), 'Photo ' + photoId, 'Visit ' + visitId].filter(Boolean).join(' | '));

  sheet_(T.photos, PHOTO_HEADERS).appendRow([
    photoId, visitId, fmt_(now, 'yyyy-MM-dd HH:mm:ss'), day, u.unit, u.project, item, note, by,
    name, file.getUrl(), file.getId(), 'OK'
  ]);
  const out = { photoId: photoId, name: name };
  try { cache.put('p:' + photoId, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}

function deletePhoto_(p) {
  const photoId = cleanId_(p.photoId, 'P');
  const visitId = cleanId_(p.visitId, 'V');
  const sh = sheet_(T.photos, PHOTO_HEADERS);
  const cell = sh.getRange(1, 1, sh.getLastRow(), 1).createTextFinder(photoId).matchEntireCell(true).findNext();
  if (!cell) throw new Error('Photo not found.');
  const row = sh.getRange(cell.getRow(), 1, 1, PHOTO_HEADERS.length).getValues()[0];
  if (row[1] !== visitId) throw new Error('This photo belongs to another visit.');
  if (row[12] !== 'OK') return { photoId: photoId, already: true };
  trashFile_(row[11]);
  sh.getRange(cell.getRow(), 13).setValue('Deleted');
  CacheService.getScriptCache().remove('p:' + photoId);
  return { photoId: photoId };
}

/* =========================== close visit (Done) =========================== */

function closeVisit_(p) {
  const visitId = cleanId_(p.visitId, 'V');
  const u = findUnit_(p.unit);
  const by = str_(p.by).slice(0, 60);
  const workDone = str_(p.workDone).slice(0, 2000);
  const notes = str_(p.notes).slice(0, 2000);
  if (!by) throw new Error('Choose your name.');
  if (!workDone) throw new Error('Write what was done today.');

  const alias = aliases_();
  const crew = (Array.isArray(p.manpower) ? p.manpower : []).map(m => ({
    contractor: (alias[low_(m.contractor)] || str_(m.contractor)).slice(0, 80),
    trade: str_(m.trade).slice(0, 40),
    count: Math.max(0, Math.min(500, Math.round(Number(m.count) || 0)))
  })).filter(m => m.contractor && m.count > 0);
  if (!crew.length && !p.noWorkers) throw new Error('Add the manpower on site, or tick "No workers today".');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const vs = sheet_(T.visits, VISIT_HEADERS);
    if (vs.getLastRow() > 1 && vs.getRange(1, 1, vs.getLastRow(), 1).createTextFinder(visitId).matchEntireCell(true).findNext()) {
      return { visitId: visitId, already: true };
    }
    let count = 0; const items = [];
    rows_(T.photos).forEach(r => {
      if (r['Visit ID'] === visitId && r['Status'] === 'OK') {
        count++;
        if (items.indexOf(r['Item']) < 0) items.push(r['Item']);
      }
    });
    if (!count) throw new Error('No photos were saved for this visit yet.');

    const now = new Date();
    const day = fmt_(now, 'yyyy-MM-dd');
    const started = Number(p.startedAt) ? fmt_(new Date(Number(p.startedAt)), 'yyyy-MM-dd HH:mm:ss') : '';
    let link = '';
    try { link = folderPath_([PHOTOS_ROOT, safeUnit_(u.unit), day]).getUrl(); } catch (e) {}
    const total = crew.reduce((s, m) => s + m.count, 0);

    vs.appendRow([visitId, day, u.unit, u.project, items.join(', '), by, started, fmt_(now, 'yyyy-MM-dd HH:mm:ss'),
      count, 'Closed', link, workDone, total, notes]);

    if (crew.length) {
      const mp = sheet_(T.manpower, MP_HEADERS);
      const rows = crew.map(m => [visitId, day, u.unit, u.project, m.contractor, m.trade, m.count, by, 'OK']);
      mp.getRange(mp.getLastRow() + 1, 1, rows.length, MP_HEADERS.length).setValues(rows);

      // new contractor names typed in the app join the list
      const known = rows_(T.contractors).map(r => low_(r['Contractor']));
      const fresh = [];
      crew.forEach(m => { const k = m.contractor.toLowerCase(); if (known.indexOf(k) < 0 && fresh.indexOf(k) < 0) { fresh.push(k); sheet_(T.contractors, ['Contractor', 'Active', 'Same as']).appendRow([m.contractor, 'Yes', '']); } });
    }
    clearBoot_();
    return { visitId: visitId, count: count, items: items, manpower: total };
  } finally { lock.releaseLock(); }
}

/* =========================== DAILY REPORTS =========================== */

/** Scheduled. Also on the sheet menu. */
function dailyJob() { makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), true); }

function makeDaily_(day, send) {
  const S = settings_();
  const visits = closedVisits_(v => v.date === day);
  if (!visits.length) return { visits: 0 };
  const photos = photosByVisit_(visits);
  const crew = crewByVisit_(visits);
  const folder = reportsFolder_('Daily', day);
  const made = [];

  // 1) one report per engineer
  const byEng = groupBy_(visits, v => v.by || 'Unknown');
  Object.keys(byEng).sort().forEach(eng => {
    const list = byEng[eng].sort((a, b) => a.closed.localeCompare(b.closed));
    const html = dailyHtml_(S, day, eng, list, photos, crew);
    const pdf = savePdf_(html, folder, `Daily ${day} - ${eng}.pdf`);
    made.push({ kind: 'Daily engineer', forWho: eng, file: pdf, eng: eng });
  });

  // 2) manpower by contractor, everyone together
  const mpHtml = manpowerHtml_(S, day, day, visits, crew);
  const mpPdf = savePdf_(mpHtml, folder, `Manpower ${day}.pdf`);
  made.push({ kind: 'Daily manpower', forWho: 'All contractors', file: mpPdf });

  if (send) {
    const to = S['Daily report to'];
    const total = visits.reduce((s, v) => s + v.manpower, 0);
    if (to) {
      MailApp.sendEmail({
        to: to,
        subject: `${S['Company name'] || 'O-Plus'} · Daily site report · ${day} · ${visits.length} units · ${total} workers`,
        htmlBody: dailyEmailBody_(day, visits, crew, made),
        attachments: made.map(m => m.file.getBlob())
      });
    }
    if (yes_(S['Send each engineer their own report'])) {
      const emails = {};
      rows_(T.engineers).forEach(r => { if (str_(r['Email'])) emails[low_(r['Name'])] = str_(r['Email']); });
      made.filter(m => m.eng && emails[m.eng.toLowerCase()]).forEach(m => {
        MailApp.sendEmail({ to: emails[m.eng.toLowerCase()], subject: `Your site report · ${day}`, body: `Attached: your site report for ${day}.`, attachments: [m.file.getBlob()] });
      });
    }
    made.forEach(m => logReport_(m.kind, day, m.forWho, m.file, to || ''));
  } else {
    made.forEach(m => logReport_(m.kind, day, m.forWho, m.file, '(made from menu, not emailed)'));
  }
  return { visits: visits.length, files: made.length, folder: folder.getUrl() };
}

function dailyHtml_(S, day, eng, list, photos, crew) {
  const total = list.reduce((s, v) => s + v.manpower, 0);
  let h = head_(`تقرير الموقع اليومي`, `${dayAr_(day)} · ${esc_(eng)}`, `${S['Company name'] || 'O-Plus'} · متابعة تنفيذ الوحدات`);
  h += `<h2>ملخص اليوم (${list.length} وحدات)</h2>
  <table class="sum"><tr><th style="width:24px">#</th><th style="width:120px">الوحدة</th><th style="width:48px">الوقت</th><th style="width:60px">العمالة</th><th>الشغل</th></tr>`;
  list.forEach((v, i) => {
    h += `<tr><td>${i + 1}</td><td class="mono">${esc_(v.unit)}</td><td class="mono">${v.closed.slice(11, 16)}</td><td>${v.manpower || '—'}</td><td>${esc_(v.workDone).replace(/\n/g, '<br>')}</td></tr>`;
  });
  h += `</table><p class="foot">إجمالي العمالة: <b>${total}</b> فرد · ${list.length} وحدة · الأوقات بتوقيت القاهرة.</p>`;

  list.forEach((v, i) => {
    h += `<div class="page">`;
    h += `<table class="bar"><tr><td class="mono big">${i + 1} · ${esc_(v.unit)}</td><td style="text-align:left">المعاينة ${v.closed.slice(11, 16)} · ${esc_(v.project)}</td></tr></table>`;
    h += photoGrid_(photos[v.id] || [], 0);
    const c = crew[v.id] || [];
    h += `<table class="two"><tr>
      <td><div class="lbl">العمالة (Manpower) <span class="pill">${v.manpower || 0}</span></div>${crewList_(c)}</td>
      <td><div class="lbl">الحالة (Status)</div><div>${esc_(v.workDone).replace(/\n/g, '<br>')}</div></td></tr></table>`;
    if (v.notes) h += `<div class="note"><b>ملاحظات:</b> ${esc_(v.notes).replace(/\n/g, '<br>')}</div>`;
    h += `</div>`;
  });
  return h + '</body></html>';
}

function manpowerHtml_(S, from, to, visits, crew) {
  const byC = {};
  visits.forEach(v => (crew[v.id] || []).forEach(m => {
    const c = byC[m.contractor] = byC[m.contractor] || { total: 0, units: {}, trades: {} };
    c.total += m.count;
    c.units[v.unit] = (c.units[v.unit] || 0) + m.count;
    const t = m.trade || '—';
    c.trades[t] = (c.trades[t] || 0) + m.count;
  }));
  const names = Object.keys(byC).sort((a, b) => byC[b].total - byC[a].total);
  const grand = names.reduce((s, n) => s + byC[n].total, 0);
  const noCrew = visits.filter(v => !(crew[v.id] || []).length);

  const range = from !== to;
  const unitsCount = Object.keys(groupBy_(visits, v => v.unit)).length;
  let h = head_('العمالة حسب المقاول', range ? `${dayAr_(from)} ← ${dayAr_(to)}` : dayAr_(from), `${S['Company name'] || 'O-Plus'} · Manpower by contractor`);
  h += `<table class="kpi"><tr><td><div class="n">${grand}</div>${range ? 'إجمالي العمالة (عامل × يوم)' : 'إجمالي العمالة'}</td><td><div class="n">${names.length}</div>مقاولين</td><td><div class="n">${unitsCount}</div>وحدات</td>${range ? `<td><div class="n">${Object.keys(groupBy_(visits, v => v.date)).length}</div>أيام شغل</td>` : ''}</tr></table>`;
  h += `<table class="sum"><tr><th>المقاول</th><th style="width:60px">العدد</th><th>حسب الحرفة</th><th>حسب الوحدة</th></tr>`;
  names.forEach(n => {
    const c = byC[n];
    h += `<tr><td><b>${esc_(n)}</b></td><td class="mono big">${c.total}</td>
      <td>${Object.keys(c.trades).map(t => `${esc_(tradeAr_(t))} ${c.trades[t]}`).join('<br>')}</td>
      <td class="mono">${Object.keys(c.units).sort().map(u => `${esc_(u)}: ${c.units[u]}`).join('<br>')}</td></tr>`;
  });
  h += `<tr class="tot"><td>الإجمالي</td><td class="mono big">${grand}</td><td></td><td></td></tr></table>`;
  if (range) {
    const byDay = groupBy_(visits, v => v.date);
    h += `<h2>حسب اليوم</h2><table class="sum"><tr><th style="width:150px">اليوم</th><th style="width:60px">العمالة</th><th>المقاولين</th></tr>`;
    Object.keys(byDay).sort().forEach(d => {
      const c = {};
      byDay[d].forEach(v => (crew[v.id] || []).forEach(m => c[m.contractor] = (c[m.contractor] || 0) + m.count));
      const tot = Object.keys(c).reduce((s, k) => s + c[k], 0);
      h += `<tr><td>${dayAr_(d)}</td><td class="mono big">${tot}</td><td>${Object.keys(c).sort((a, b) => c[b] - c[a]).map(k => `${esc_(k)} ${c[k]}`).join(' · ')}</td></tr>`;
    });
    h += `</table>`;
  }
  if (noCrew.length && !range) h += `<p class="foot">وحدات من غير عمالة النهارده: ${noCrew.map(v => esc_(v.unit)).join('، ')}</p>`;
  return h + '</body></html>';
}

function dailyEmailBody_(day, visits, crew, made) {
  const byC = {};
  visits.forEach(v => (crew[v.id] || []).forEach(m => byC[m.contractor] = (byC[m.contractor] || 0) + m.count));
  const rows = Object.keys(byC).sort((a, b) => byC[b] - byC[a]).map(c => `<tr><td style="padding:4px 10px">${esc_(c)}</td><td style="padding:4px 10px;text-align:right"><b>${byC[c]}</b></td></tr>`).join('');
  const eng = groupBy_(visits, v => v.by);
  return `<div style="font-family:Arial,sans-serif;font-size:14px">
    <p><b>Daily site report · ${day}</b><br>${visits.length} unit visits · ${visits.reduce((s, v) => s + v.manpower, 0)} workers</p>
    <p><b>Manpower by contractor</b></p><table style="border-collapse:collapse;border:1px solid #ddd">${rows || '<tr><td style="padding:4px 10px">No manpower recorded</td></tr>'}</table>
    <p><b>Engineers</b><br>${Object.keys(eng).map(e => `${esc_(e)}: ${eng[e].length} units`).join('<br>')}</p>
    <p>PDFs attached (${made.length}). Copies are in Drive: ${REPORTS_ROOT} / Daily / ${day}.</p></div>`;
}

/* =========================== WEEKLY CLIENT REPORTS =========================== */

/** Scheduled on the weekly day. Also on the sheet menu. */
function weeklyJob() { makeWeekly_(fmt_(new Date(), 'yyyy-MM-dd'), true); }

function makeWeekly_(endDay, send) {
  const S = settings_();
  const end = parseDay_(endDay);
  const start = new Date(end.getTime() - 6 * 86400000);
  const startDay = fmt_(start, 'yyyy-MM-dd');
  const period = `${startDay} → ${endDay}`;
  const visits = closedVisits_(v => v.date >= startDay && v.date <= endDay);
  if (!visits.length) return { units: 0 };
  const photos = photosByVisit_(visits);
  const crew = crewByVisit_(visits);
  const clients = {};
  rows_(T.units).forEach(r => clients[low_(r['Unit'])] = { name: str_(r['Client name']), email: str_(r['Client email']) });
  const folder = reportsFolder_('Weekly', endDay);
  const made = [];

  const byUnit = groupBy_(visits, v => v.unit);
  Object.keys(byUnit).sort().forEach(unit => {
    const list = byUnit[unit].sort((a, b) => a.closed.localeCompare(b.closed));
    const cl = clients[unit.toLowerCase()] || {};
    const html = weeklyHtml_(S, unit, cl, startDay, endDay, list, photos, crew);
    const pdf = savePdf_(html, folder, `Weekly ${endDay} - ${safeUnit_(unit)}.pdf`);
    made.push({ unit: unit, client: cl, file: pdf });
  });

  const sentTo = {};
  if (send) {
    const copy = S['Weekly report copy to'];
    if (copy) {
      const links = made.map(m => `<li>${esc_(m.unit)} · ${esc_(m.client.name || '')} · <a href="${m.file.getUrl()}">PDF</a></li>`).join('');
      MailApp.sendEmail({
        to: copy,
        subject: `${S['Company name'] || 'O-Plus'} · Weekly client reports · ${period} · ${made.length} units`,
        htmlBody: `<div style="font-family:Arial,sans-serif;font-size:14px"><p><b>Weekly client reports · ${period}</b></p><ul>${links}</ul>
          <p>Send to clients is <b>${yes_(S['Send weekly to clients']) ? 'ON' : 'OFF'}</b> (Settings tab).</p></div>`
      });
      made.forEach(m => sentTo[m.unit] = copy);
    }
    if (yes_(S['Send weekly to clients'])) {
      made.forEach(m => {
        if (!m.client.email) return;
        MailApp.sendEmail({
          to: m.client.email,
          subject: `Weekly progress · ${m.unit} · ${period}`,
          htmlBody: `<div style="font-family:Arial,sans-serif;font-size:14px"><p>Dear ${esc_(m.client.name || 'Client')},</p>
            <p>Please find attached this week's progress report for your unit <b>${esc_(m.unit)}</b> (${period}).</p>
            <p>Best regards,<br>${esc_(S['Company name'] || 'O-Plus')}</p></div>`,
          attachments: [m.file.getBlob()]
        });
        sentTo[m.unit] = [sentTo[m.unit], m.client.email].filter(Boolean).join(', ');
      });
    }
  }
  made.forEach(m => logReport_('Weekly client', period, m.unit, m.file, send ? (sentTo[m.unit] || '') : '(made from menu, not emailed)'));
  return { units: made.length, folder: folder.getUrl() };
}

function weeklyHtml_(S, unit, cl, startDay, endDay, list, photos, crew) {
  const days = groupBy_(list, v => v.date);
  const dayKeys = Object.keys(days).sort();
  const items = [];
  list.forEach(v => v.items.split(',').map(s => s.trim()).filter(Boolean).forEach(i => items.indexOf(i) < 0 && items.push(i)));
  const workerDays = list.reduce((s, v) => s + v.manpower, 0);

  let h = head_('تقرير الأسبوع', `${esc_(unit)}${cl.name ? ' · ' + esc_(cl.name) : ''}`, `${S['Company name'] || 'O-Plus'} · ${startDay} → ${endDay}`);
  h += `<table class="kpi"><tr><td><div class="n">${dayKeys.length}</div>أيام شغل</td><td><div class="n">${workerDays}</div>إجمالي العمالة (عامل × يوم)</td><td><div class="n">${list.reduce((s, v) => s + v.photos, 0)}</div>صور</td></tr></table>`;
  if (items.length) h += `<p><b>البنود:</b> ${items.map(esc_).join(' · ')}</p>`;
  h += `<h2>ملخص الأسبوع</h2><table class="sum"><tr><th style="width:110px">اليوم</th><th style="width:60px">العمالة</th><th>الشغل اللي اتعمل</th></tr>`;
  dayKeys.forEach(d => {
    const vs = days[d];
    h += `<tr><td>${dayAr_(d)}</td><td>${vs.reduce((s, v) => s + v.manpower, 0) || '—'}</td><td>${vs.map(v => esc_(v.workDone).replace(/\n/g, '<br>')).join('<br>')}</td></tr>`;
  });
  h += `</table>`;
  dayKeys.forEach(d => {
    const vs = days[d];
    let ph = [];
    vs.forEach(v => ph = ph.concat(photos[v.id] || []));
    h += `<div class="page"><table class="bar"><tr><td class="big">${dayAr_(d)}</td><td style="text-align:left" class="mono">${esc_(unit)}</td></tr></table>`;
    h += photoGrid_(ph, 12);
    h += `<div class="box"><div class="lbl">الحالة (Status)</div>${vs.map(v => esc_(v.workDone).replace(/\n/g, '<br>')).join('<br>')}</div>`;
    const tot = vs.reduce((s, v) => s + v.manpower, 0);
    if (tot) h += `<div class="box"><div class="lbl">العمالة في الموقع</div>${tot} فرد</div>`;
    h += `</div>`;
  });
  return h + '</body></html>';
}

/* =========================== shared report pieces =========================== */

function head_(title, sub, line) {
  return `<html dir="rtl" lang="ar"><head><meta charset="utf-8"><style>
    @page { size: A4; margin: 14mm 12mm; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    body { font-family: Arial, sans-serif; font-size: 11pt; color: #0d1b3e; direction: rtl; }
    .hdr { color: #0d1b3e; padding: 10px 2px 12px; border-top: 8px solid #0d1b3e; border-bottom: 4px solid #e9b88f; }
    .hdr .t { font-size: 20pt; font-weight: bold; }
    .hdr .s { font-size: 12pt; margin-top: 4px; }
    .hdr .l { font-size: 9pt; color: #5d6577; margin-top: 4px; }
    h2 { font-size: 13pt; margin: 16px 0 8px; }
    table { border-collapse: collapse; width: 100%; }
    .sum th { background: #eceef4; text-align: right; padding: 6px; font-size: 10pt; border: 1px solid #dfe2ea; }
    .sum td { padding: 6px; border: 1px solid #dfe2ea; vertical-align: top; font-size: 10pt; }
    .sum .tot td { background: #f4f5f8; font-weight: bold; }
    .mono { font-family: "Courier New", monospace; direction: ltr; unicode-bidi: embed; }
    .big { font-size: 13pt; font-weight: bold; }
    .foot { font-size: 9pt; color: #5d6577; }
    .page { page-break-before: always; }
    .bar td { background: #eceef4; color: #0d1b3e; padding: 8px 10px; border-right: 6px solid #e9b88f; }
    .grid td { width: 33%; padding: 4px; vertical-align: top; text-align: center; }
    .grid img { width: 100%; max-height: 230px; border: 1px solid #dfe2ea; }
    .cap { font-size: 8.5pt; color: #5d6577; margin-top: 2px; }
    .grid tr, .sum tr, .two, .box, .note { page-break-inside: avoid; }
    .item-h { font-weight: bold; margin: 10px 0 2px; border-bottom: 2px solid #e9b88f; padding-bottom: 2px; }
    .item-h span { color: #5d6577; font-weight: normal; }
    .two td { width: 50%; vertical-align: top; border: 1px solid #dfe2ea; padding: 8px; }
    .lbl { font-weight: bold; margin-bottom: 6px; }
    .pill { background: #e9b88f; color: #0d1b3e; padding: 1px 8px; border-radius: 8px; font-size: 9pt; }
    .note, .box { border: 1px solid #dfe2ea; padding: 8px; margin-top: 8px; }
    .kpi td { text-align: center; border: 1px solid #dfe2ea; padding: 10px; font-size: 9pt; color: #5d6577; }
    .kpi .n { font-size: 20pt; font-weight: bold; color: #0d1b3e; }
    ul { margin: 0; padding-right: 18px; }
  </style></head><body>
  <div class="hdr"><div class="t">${title}</div><div class="s">${sub}</div><div class="l">${line}</div></div>`;
}

/** All photos of a visit (or a day), grouped by item, 3 per row. max = 0 means no limit. */
function photoGrid_(list, max) {
  if (!list.length) return '<p class="foot">مفيش صور.</p>';
  const ph = max ? list.slice(0, max) : list;
  const imgs = images_(ph.map(p => p.fileId));
  const groups = groupBy_(ph, p => p.item || GENERAL);
  const names = Object.keys(groups);
  let h = '';
  names.forEach(name => {
    const g = groups[name];
    if (names.length > 1) h += `<div class="item-h">${esc_(name)} <span>(${g.length})</span></div>`;
    h += '<table class="grid">';
    for (let i = 0; i < g.length; i += 3) {
      h += '<tr>';
      for (let j = i; j < i + 3; j++) {
        const p = g[j];
        h += p ? `<td>${imgs[p.fileId] ? `<img src="${imgs[p.fileId]}">` : ''}<div class="cap">${esc_(p.note || p.item)}</div></td>` : '<td></td>';
      }
      h += '</tr>';
    }
    h += '</table>';
  });
  if (max && list.length > max) h += `<p class="foot">+ ${list.length - max} صور تانية في فولدر الوحدة على الدرايف.</p>`;
  return h;
}

function crewList_(c) {
  if (!c.length) return '<div>مفيش عمالة.</div>';
  return '<ul>' + c.map(m => `<li>${esc_(m.contractor)}${m.trade ? ' · ' + esc_(tradeAr_(m.trade)) : ''}: <b>${m.count}</b></li>`).join('') + '</ul>';
}

/** Small JPEG data URIs for the PDF, fetched in parallel. */
function images_(ids) {
  const out = {};
  if (!ids.length) return out;
  const token = ScriptApp.getOAuthToken();
  const reqs = ids.map(id => ({ url: 'https://drive.google.com/thumbnail?sz=w600&id=' + id, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }));
  let res = [];
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) {}
  ids.forEach((id, i) => {
    const r = res[i];
    let blob = null;
    if (r && r.getResponseCode() === 200 && /^image\//.test(String(r.getHeaders()['Content-Type'] || ''))) blob = r.getBlob();
    if (!blob) { try { blob = DriveApp.getFileById(id).getBlob(); } catch (e) {} }
    if (blob) out[id] = 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  });
  return out;
}

function savePdf_(html, folder, name) {
  const pdf = Utilities.newBlob(html, 'text/html', 'r.html').getAs('application/pdf').setName(name);
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  return folder.createFile(pdf);
}

function logReport_(kind, period, forWho, file, sentTo) {
  sheet_(T.reports, REPORT_HEADERS).appendRow([fmt_(new Date(), 'yyyy-MM-dd HH:mm'), kind, period, forWho, file.getUrl(), sentTo]);
}

/* =========================== reading data =========================== */

function closedVisits_(filter) {
  return rows_(T.visits).map(r => ({
    id: str_(r['Visit ID']), date: day_(r['Date']), unit: str_(r['Unit']), project: str_(r['Project']),
    items: str_(r['Items']), by: str_(r['By']), closed: dt_(r['Closed']), photos: Number(r['Photos']) || 0,
    status: str_(r['Status']), workDone: str_(r['Work done']), manpower: Number(r['Manpower total']) || 0, notes: str_(r['Notes'])
  })).filter(v => v.id && v.status === 'Closed' && filter(v));
}

function photosByVisit_(visits) {
  const ids = {}; visits.forEach(v => ids[v.id] = 1);
  const out = {};
  rows_(T.photos).forEach(r => {
    if (!ids[r['Visit ID']] || r['Status'] !== 'OK') return;
    (out[r['Visit ID']] = out[r['Visit ID']] || []).push({ fileId: str_(r['File ID']), item: str_(r['Item']), note: str_(r['Note']) });
  });
  return out;
}

function crewByVisit_(visits) {
  const ids = {}; visits.forEach(v => ids[v.id] = 1);
  const out = {};
  rows_(T.manpower).forEach(r => {
    if (!ids[r['Visit ID']] || str_(r['Status']) === 'Deleted') return;
    (out[r['Visit ID']] = out[r['Visit ID']] || []).push({ contractor: str_(r['Contractor']), trade: str_(r['Trade']), count: Number(r['Count']) || 0 });
  });
  return out;
}

/** Rows of a tab as objects keyed by header name (column order doesn't matter). */
function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  return v.slice(1).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}

function settings_() {
  const out = {};
  rows_(T.settings).forEach(r => { if (str_(r['Setting'])) out[str_(r['Setting'])] = str_(r['Value']); });
  return out;
}

let TRADE_AR = null;
function tradeAr_(t) {
  if (!TRADE_AR) { TRADE_AR = {}; rows_(T.trades).forEach(r => TRADE_AR[low_(r['Trade'])] = str_(r['Arabic'])); }
  return TRADE_AR[low_(t)] || t;
}

/* =========================== sheet menu (for you) =========================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('O West Camera')
    .addItem("Make today's daily reports (no email)", 'menuDailyToday')
    .addItem("Send today's daily reports now", 'menuDailySend')
    .addItem('Make weekly client reports, week to date (no email)', 'menuWeekly')
    .addSeparator()
    .addItem('Delete selected photos', 'menuDeletePhotos')
    .addItem('Delete selected visits (photos + manpower)', 'menuDeleteVisits')
    .addSeparator()
    .addItem('Set manager password', 'menuSetPassword')
    .addItem('Connect ClickUp (API token)', 'menuSetClickUpToken')
    .addItem('Sync from ClickUp now', 'menuSync')
    .addItem('Apply schedule from Settings', 'installTriggers')
    .addItem('Refresh lists in the app', 'menuRefresh')
    .addToUi();
}

function menuDailyToday() { reportAlert_(makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), false), 'visits'); }
function menuDailySend() { reportAlert_(makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), true), 'visits'); }
function menuWeekly() { reportAlert_(makeWeekly_(fmt_(new Date(), 'yyyy-MM-dd'), false), 'units'); }
function reportAlert_(r, key) {
  const ui = SpreadsheetApp.getUi();
  if (!r[key]) return ui.alert('No closed visits in this period yet.');
  ui.alert(`Done. ${r[key]} ${key}. PDFs are in Drive > ${REPORTS_ROOT}, and listed in the Reports tab.`);
}
function menuRefresh() { clearBoot_(); SpreadsheetApp.getUi().alert('Done. The app loads the new lists on next open.'); }

function menuDeletePhotos() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== T.photos) return ui.alert('Open the Photos tab and select the rows to delete.');
  const rows = selectedRows_(sh);
  if (!rows.length) return ui.alert('Select at least one photo row.');
  if (ui.alert(`Delete ${rows.length} photo(s)? Files go to Drive trash (30 days to restore).`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let n = 0;
  rows.forEach(r => {
    const v = sh.getRange(r, 1, 1, PHOTO_HEADERS.length).getValues()[0];
    if (v[12] !== 'OK') return;
    trashFile_(v[11]); sh.getRange(r, 13).setValue('Deleted'); n++;
  });
  ui.alert(`${n} photo(s) deleted.`);
}

function menuDeleteVisits() {
  const ui = SpreadsheetApp.getUi();
  const vs = SpreadsheetApp.getActiveSheet();
  if (vs.getName() !== T.visits) return ui.alert('Open the Visits tab and select the visit rows to delete.');
  const rows = selectedRows_(vs);
  if (!rows.length) return ui.alert('Select at least one visit row.');
  if (ui.alert(`Delete ${rows.length} visit(s) with their photos and manpower? Files go to Drive trash.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ids = rows.map(r => vs.getRange(r, 1).getValue());
  let n = 0;
  const ps = sheet_(T.photos, PHOTO_HEADERS);
  if (ps.getLastRow() > 1) ps.getRange(2, 1, ps.getLastRow() - 1, PHOTO_HEADERS.length).getValues().forEach((v, i) => {
    if (ids.indexOf(v[1]) >= 0 && v[12] === 'OK') { trashFile_(v[11]); ps.getRange(i + 2, 13).setValue('Deleted'); n++; }
  });
  const mp = sheet_(T.manpower, MP_HEADERS);
  if (mp.getLastRow() > 1) mp.getRange(2, 1, mp.getLastRow() - 1, MP_HEADERS.length).getValues().forEach((v, i) => {
    if (ids.indexOf(v[0]) >= 0) mp.getRange(i + 2, 9).setValue('Deleted');
  });
  rows.forEach(r => vs.getRange(r, 10).setValue('Deleted'));
  clearBoot_();
  ui.alert(`${rows.length} visit(s) and ${n} photo(s) deleted. Reports made after this ignore them.`);
}

/* =========================== schedule & setup =========================== */

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['dailyJob', 'weeklyJob', 'syncJob'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  const S = settings_();
  const dh = clampHour_(S['Daily report hour'], 19), wh = clampHour_(S['Weekly report hour'], 17);
  const days = { sunday: 'SUNDAY', monday: 'MONDAY', tuesday: 'TUESDAY', wednesday: 'WEDNESDAY', thursday: 'THURSDAY', friday: 'FRIDAY', saturday: 'SATURDAY' };
  const wd = days[low_(S['Weekly report day'])] || 'THURSDAY';
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(dh).inTimezone(TZ).create();
  ScriptApp.newTrigger('weeklyJob').timeBased().onWeekDay(ScriptApp.WeekDay[wd]).atHour(wh).inTimezone(TZ).create();
  const sync = !!PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (sync) ScriptApp.newTrigger('syncJob').timeBased().everyMinutes(30).create();
  try { SpreadsheetApp.getUi().alert(`Schedule set: daily reports at ${dh}:00, weekly reports every ${wd.charAt(0) + wd.slice(1).toLowerCase()} at ${wh}:00 (Cairo).` + (sync ? ' ClickUp sync every 30 minutes.' : ' ClickUp sync is off until you connect ClickUp.')); } catch (e) {}
}

/** Run once after pasting the code. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  folderPath_([PHOTOS_ROOT]);
  folderPath_([REPORTS_ROOT]);
  sheet_(T.photos, PHOTO_HEADERS); sheet_(T.visits, VISIT_HEADERS); sheet_(T.manpower, MP_HEADERS); sheet_(T.reports, REPORT_HEADERS);
  // make sure Visits has the newer columns
  const vs = ss.getSheetByName(T.visits);
  vs.getRange(1, 1, 1, VISIT_HEADERS.length).setValues([VISIT_HEADERS]).setFontWeight('bold');
  clearBoot_();
  installTriggers();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });   // asks for the permission reports need
  MailApp.getRemainingDailyQuota();
}

/* =========================== helpers =========================== */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function msg_(e) { return String(e && e.message || e); }
function str_(v) { return v == null ? '' : String(v).trim(); }
function low_(v) { return str_(v).toLowerCase(); }
function yes_(v) { return ['yes', 'y', 'true', '1', 'نعم', 'اه', 'أيوه'].indexOf(low_(v)) >= 0; }
function esc_(s) { return str_(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function day_(v) { return v instanceof Date ? Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : str_(v).slice(0, 10); }
function dt_(v) { return v instanceof Date ? Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss') : str_(v); }
function parseDay_(s) { const p = s.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12)); }
function clampHour_(v, d) { const n = Math.round(Number(v)); return n >= 0 && n <= 23 ? n : d; }
function safeUnit_(u) { return u.replace(/[\/\\]+/g, '-').replace(/\s+/g, ' ').trim(); }
function groupBy_(list, f) { const o = {}; list.forEach(x => (o[f(x)] = o[f(x)] || []).push(x)); return o; }

function dayAr_(s) {
  const d = parseDay_(s);
  const wd = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'][d.getUTCDay()];
  const mo = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'][d.getUTCMonth()];
  return `${wd} ${d.getUTCDate()} ${mo} ${d.getUTCFullYear()}`;
}

function cleanId_(v, prefix) {
  const s = String(v || '');
  if (!new RegExp('^' + prefix + '-[A-Za-z0-9-]{6,40}$').test(s)) throw new Error('Bad request. Reload the page.');
  return s;
}

function trashFile_(id) { if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} } }

function selectedRows_(sh) {
  const out = [];
  const list = sh.getActiveRangeList();
  if (!list) return out;
  list.getRanges().forEach(rg => { for (let r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r > 1 && out.indexOf(r) < 0) out.push(r); });
  return out;
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Folder by path under My Drive, ids cached. */
function folderPath_(parts) {
  const cache = CacheService.getScriptCache();
  const key = 'fp:' + parts.join('/');
  const id = cache.get(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const props = PropertiesService.getScriptProperties();
  let f;
  const rootKey = 'ROOT:' + parts[0];
  const rid = props.getProperty(rootKey) || (parts[0] === PHOTOS_ROOT ? props.getProperty('ROOT_ID') : null);
  if (rid) { try { f = DriveApp.getFolderById(rid); } catch (e) {} }
  if (!f) { const it = DriveApp.getFoldersByName(parts[0]); f = it.hasNext() ? it.next() : DriveApp.createFolder(parts[0]); props.setProperty(rootKey, f.getId()); }
  for (let i = 1; i < parts.length; i++) { const it = f.getFoldersByName(parts[i]); f = it.hasNext() ? it.next() : f.createFolder(parts[i]); }
  try { cache.put(key, f.getId(), 21600); } catch (e) {}
  return f;
}

function dayFolder_(safeUnit, day) {
  const lock = LockService.getScriptLock();
  const cache = CacheService.getScriptCache();
  const key = 'fp:' + [PHOTOS_ROOT, safeUnit, day].join('/');
  const id = cache.get(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  lock.waitLock(30000);
  try { return folderPath_([PHOTOS_ROOT, safeUnit, day]); } finally { lock.releaseLock(); }
}

function reportsFolder_(kind, day) { return folderPath_([REPORTS_ROOT, kind, day]); }


/* =========================== CLICKUP SYNC =========================== */
/*
 * Every 30 minutes (and from the menu / the manager page) this reads the ClickUp
 * "Main Tracking Sheet" list and rebuilds, for every task in "site execution":
 *   Units (Source = ClickUp)  unit, project, client name/email
 *   Scope                     what the client bought (Request Type, or the description for "Others")
 *   Unit Contractors          the Contractor field (comma = several)
 * Rows with Source = Manual (e.g. the Authority Units) are never touched.
 * A unit you set Active = No stays No after syncing.
 */

const CU = {
  list: '1201250000001676',
  status: 'site execution',
  f: {
    requestType: '8f185b3c-db17-4dac-b577-d377a949db9f',
    description: 'af390993-91be-4db5-b8d9-c4a2dca60ce8',
    contractor: '8d4a9069-1ade-4b4e-9702-ef05ff3ecf97',
    client: '75747da4-250c-42e7-8d65-98d9d503060f',
    email: 'ecdc4bdb-8444-4074-b32c-cd4c36274f25'
  },
  skipTypes: ['design fees'],
  skipContractors: ['', 'pending', 'design team', '-', 'n/a', 'tbd'],
  typeNames: {
    'standard finishing package': 'Standard Finishing', 'customized finishing package': 'Customized Finishing',
    'ff&e': 'Furniture (FF&E)', 'ac`s': 'ACs', "ac's": 'ACs', 'bua enclosure': 'BUA Enclosure'
  }
};

function syncJob() { try { syncClickUp_(); } catch (e) { console.error(e); } }

function menuSync() {
  const r = syncClickUp_();
  SpreadsheetApp.getUi().alert(`ClickUp sync done: ${r.units} units, ${r.items} items, ${r.contractors} unit-contractor links. ${r.newContractors ? r.newContractors + ' new contractor name(s) added to Contractors.' : ''}`);
}

function menuSetClickUpToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Connect ClickUp', 'Paste your ClickUp personal API token.\nClickUp > avatar > Settings > Apps > API Token (starts with pk_).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (!/^pk_/.test(t)) return ui.alert('That does not look like a ClickUp token (it should start with pk_).');
  PropertiesService.getScriptProperties().setProperty('CLICKUP_TOKEN', t);
  try { syncClickUp_(); } catch (e) { return ui.alert('Token saved, but the first sync failed: ' + msg_(e)); }
  installTriggers();
}

function syncClickUp_() {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (!token) throw new Error('ClickUp is not connected. Sheet menu: O West Camera > Connect ClickUp.');
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    // 1) read every task in site execution
    const tasks = [];
    for (let page = 0; page < 30; page++) {
      const url = `https://api.clickup.com/api/v2/list/${CU.list}/task?statuses[]=${encodeURIComponent(CU.status)}&subtasks=true&include_closed=true&page=${page}`;
      const res = UrlFetchApp.fetch(url, { headers: { Authorization: token }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) throw new Error('ClickUp answered ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
      const j = JSON.parse(res.getContentText());
      (j.tasks || []).forEach(t => tasks.push(t));
      if (j.last_page || !(j.tasks || []).length) break;
    }

    // 2) turn tasks into units / items / contractors
    const alias = aliases_();
    const knownC = {};
    rows_(T.contractors).forEach(r => knownC[low_(r['Contractor'])] = 1);
    const units = {}, scope = {}, uc = {}, freshC = [];
    tasks.forEach(t => {
      const unit = normUnit_(t.name);
      if (!unit) return;
      const k = unit.toLowerCase();
      const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
      const u = units[k] = units[k] || { unit: unit, client: '', email: '' };
      u.client = u.client || str_(f[CU.f.client] && f[CU.f.client].value);
      u.email = u.email || str_(f[CU.f.email] && f[CU.f.email].value);

      const type = dropName_(f[CU.f.requestType]);
      if (type && CU.skipTypes.indexOf(type.toLowerCase()) < 0) {
        let item = CU.typeNames[type.toLowerCase().trim()] || type.trim();
        if (item.toLowerCase() === 'others') {
          item = str_(f[CU.f.description] && f[CU.f.description].value).split('\n')[0].trim().slice(0, 40) || 'Other works';
        }
        (scope[k] = scope[k] || []).indexOf(item) < 0 && scope[k].push(item);
      }
      str_(f[CU.f.contractor] && f[CU.f.contractor].value).split(',').forEach(c => {
        c = c.trim().replace(/\s+/g, ' ');
        if (CU.skipContractors.indexOf(c.toLowerCase()) >= 0) return;
        const name = alias[c.toLowerCase()] || c;
        if (!knownC[name.toLowerCase()]) { knownC[name.toLowerCase()] = 1; freshC.push(name); }
        (uc[k] = uc[k] || []).indexOf(name) < 0 && uc[k].push(name);
      });
    });

    // 3) write, keeping Manual rows and Active = No choices
    const ss = SpreadsheetApp.getActive();
    const uSh = ss.getSheetByName(T.units);
    const oldUnits = rows_(T.units);
    const manual = {}, active = {};
    oldUnits.forEach(r => {
      const k = low_(r['Unit']); if (!k) return;
      if (low_(r['Source']) === 'manual') manual[k] = r;
      active[k] = str_(r['Active']) || 'Yes';
    });
    const uRows = [];
    Object.keys(manual).forEach(k => { const r = manual[k]; uRows.push([str_(r['Unit']), str_(r['Active']) || 'Yes', str_(r['Project']), str_(r['Client name']), str_(r['Client email']), 'Manual']); });
    Object.keys(units).sort().forEach(k => {
      if (manual[k]) return;
      const u = units[k];
      uRows.push([u.unit, active[k] === 'No' ? 'No' : 'Yes', 'Extra Works', u.client, u.email, 'ClickUp']);
    });
    rewrite_(uSh, ['Unit', 'Active', 'Project', 'Client name', 'Client email', 'Source'], uRows);

    const keepRows = (tab, cols) => rows_(tab).filter(r => manual[low_(r['Unit'])]).map(r => cols.map(c => str_(r[c])));
    const sRows = keepRows(T.scope, ['Unit', 'Item']);
    Object.keys(scope).sort().forEach(k => { if (!manual[k]) scope[k].forEach(i => sRows.push([units[k].unit, i])); });
    rewrite_(sheet_(T.scope, ['Unit', 'Item']), ['Unit', 'Item'], sRows);

    const cRows = keepRows(T.unitContractors, ['Unit', 'Contractor']);
    Object.keys(uc).sort().forEach(k => { if (!manual[k]) uc[k].forEach(c => cRows.push([units[k].unit, c])); });
    rewrite_(sheet_(T.unitContractors, ['Unit', 'Contractor']), ['Unit', 'Contractor'], cRows);

    const cSh = sheet_(T.contractors, ['Contractor', 'Active', 'Same as']);
    freshC.forEach(n => cSh.appendRow([n, 'Yes', '']));

    setSetting_('Last ClickUp sync', fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ` · ${Object.keys(units).length} units`);
    clearBoot_();
    return { units: Object.keys(units).length, items: sRows.length, contractors: cRows.length, newContractors: freshC.length };
  } finally { lock.releaseLock(); }
}

function normUnit_(name) {
  let n = str_(name).replace(/\s+/g, '').toUpperCase();
  if (!/^OW\//.test(n)) return '';
  return n;
}

function dropName_(f) {
  if (!f || f.value == null || f.value === '') return '';
  const opts = (f.type_config && f.type_config.options) || [];
  const v = f.value;
  const o = opts.find(x => x.id === v || x.orderindex === v || String(x.orderindex) === String(v));
  return o ? str_(o.name) : (typeof v === 'object' && v.name ? str_(v.name) : '');
}

/** "Same as" column in Contractors: other spellings -> one name. */
function aliases_() {
  const a = {};
  rows_(T.contractors).forEach(r => { const s = str_(r['Same as']); if (s) a[low_(r['Contractor'])] = s; });
  return a;
}

function rewrite_(sh, headers, rows) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(headers.length, sh.getLastColumn())).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function setSetting_(key, value) {
  const sh = sheet_(T.settings, ['Setting', 'Value', 'What it does']);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (str_(v[i][0]) === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value, 'Filled automatically']);
}
