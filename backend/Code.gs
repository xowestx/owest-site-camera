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
const MP_HEADERS = ['Visit ID', 'Date', 'Unit', 'Project', 'Contractor', 'Trade', 'Count', 'By', 'Status', 'Item'];
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
  let info = null;
  try { info = cuReportInfo_(unit); } catch (e) { console.error('clickup info ' + unit + ': ' + msg_(e)); }
  const html = weeklyHtml_(settings_(), unit, { name: str_(cl['Client name']) }, startDay, endDay, list, photosByVisit_(list), crewByVisit_(list), info);
  return savePdf_(html, reportsFolder_('Weekly', endDay), `Weekly Progress Report - ${safeUnit_(unit)} - ${dateEn_(endDay)}.pdf`);
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
    item: u.items.indexOf(str_(m.item)) >= 0 ? str_(m.item) : GENERAL,
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
      const rows = crew.map(m => [visitId, day, u.unit, u.project, m.contractor, m.trade, m.count, by, 'OK', m.item]);
      mp.getRange(mp.getLastRow() + 1, 1, rows.length, MP_HEADERS.length).setValues(rows);

      // new contractor names typed in the app join the list
      const known = rows_(T.contractors).map(r => low_(r['Contractor']));
      const fresh = [];
      crew.forEach(m => { const k = m.contractor.toLowerCase(); if (known.indexOf(k) < 0 && fresh.indexOf(k) < 0) { fresh.push(k); sheet_(T.contractors, ['Contractor', 'Active', 'Same as']).appendRow([m.contractor, 'Yes', '']); } });
    }
    clearBoot_();
    try {
      unitLog_(u, { day: day, started: started, closed: fmt_(now, 'yyyy-MM-dd HH:mm:ss'), by: by, items: items, count: count,
        workDone: workDone, notes: notes, crew: crew, total: total, link: link });
    } catch (e) { console.error('unit log', e); }
    return { visitId: visitId, count: count, items: items, manpower: total };
  } finally { lock.releaseLock(); }
}

/* =========================== UNIT FILE (site log) =========================== */
/* Each unit folder in Drive has one Google Doc "<unit> - Site log" with every visit, newest at the top:
   date, time, engineer, items, work done, manpower by contractor, internal notes, link to that day's photos. */

function unitLog_(u, v) {
  const safe = safeUnit_(u.unit);
  const folder = folderPath_([PHOTOS_ROOT, safe]);
  const name = `${safe} - Site log`;
  const props = PropertiesService.getScriptProperties();
  const key = 'log:' + safe.toLowerCase();
  let doc = null;
  const id = props.getProperty(key);
  if (id) { try { doc = DocumentApp.openById(id); } catch (e) {} }
  if (!doc) {
    const it = folder.getFilesByName(name);
    if (it.hasNext()) doc = DocumentApp.openById(it.next().getId());
    else {
      doc = DocumentApp.create(name);
      const f = DriveApp.getFileById(doc.getId());
      folder.addFile(f); DriveApp.getRootFolder().removeFile(f);
      const b = doc.getBody();
      b.clear();
      b.appendParagraph(u.unit).setHeading(DocumentApp.ParagraphHeading.TITLE);
      b.appendParagraph(`${u.project} · Site log · newest visit first`).setItalic(true);
    }
    props.setProperty(key, doc.getId());
  }
  const b = doc.getBody();
  let at = 2;                                   // right under the title lines
  const put = (text, heading, bold) => {
    const p = b.insertParagraph(at++, text);
    if (heading) p.setHeading(heading); else p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setBold(!!bold);
    return p;
  };
  put(`${v.day} · ${hm_(v.started) || '?'}–${hm_(v.closed)} · ${v.by}`, DocumentApp.ParagraphHeading.HEADING2);
  if (v.items.length) put('Items: ' + v.items.join(', '));
  put('Work done: ' + v.workDone);
  put(`Manpower (${v.total}): ` + (v.crew.length ? v.crew.map(m => `${m.contractor} – ${m.item}${m.trade ? ' – ' + m.trade : ''}: ${m.count}`).join(' · ') : 'none'));
  if (v.notes) put('Notes (internal): ' + v.notes, null, true);
  const ph = put(`Photos (${v.count}): open folder`);
  if (v.link) ph.editAsText().setLinkUrl(v.link);
  put('');
  doc.saveAndClose();
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
  const nPh = list.reduce((s, v) => s + (photos[v.id] || []).length, 0);
  const first = list.map(v => v.started || v.closed).sort()[0] || '', last = list.map(v => v.closed).sort().slice(-1)[0] || '';
  let h = head_('تقرير الموقع اليومي', 'Daily Site Report', `${dayAr_(day)} · ${day}`, `${esc_(eng)} · ${esc_(S['Company name'] || 'O-Plus')}`);
  h += kpis_([[list.length, 'وحدات', 'Units'], [total, 'عمالة', 'Workers'], [nPh, 'صور', 'Photos'], [`<span class="mono">${hm_(first)}–${hm_(last)}</span>`, 'الوقت', 'Time on site']]);
  const clean = polishDay_(list);   // AI: one clean line + short bullets per visit (internal, nothing hidden)
  h += `<h2>ملخص اليوم <span class="en">Summary · ${list.length} units</span></h2><table class="sum">
    <tr><th class="c" style="width:22px">#</th><th style="width:118px">الوحدة <span class="en">Unit</span></th><th style="width:52px">الوقت <span class="en">Time</span></th>
    <th>الشغل <span class="en">Work</span></th><th class="c" style="width:52px">العمالة</th></tr>`;
  list.forEach((v, i) => {
    const c = clean[v.id];
    h += `<tr><td class="c">${i + 1}</td><td class="mono b">${esc_(v.unit)}</td><td class="mono">${hm_(v.started) || hm_(v.closed) || '—'}</td>
      <td>${c && c.line ? esc_(c.line) : br_(v.workDone)}</td><td class="c b">${v.manpower || '—'}</td></tr>`;
  });
  h += `<tr class="tot"><td></td><td colspan="3">إجمالي العمالة <span class="en">Total workers</span></td><td class="c">${total}</td></tr>`;
  h += `</table><p class="foot">الأوقات بتوقيت القاهرة · Cairo time</p>`;
  list.forEach((v, i) => {
    h += `<div class="page">${unitHead_(i + 1, v, eng)}`;
    h += photoGrid_(photos[v.id] || [], 0);
    h += workAndCrew_(v, crew[v.id] || [], clean[v.id]);
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
  const range = from !== to;
  const days = groupBy_(visits, v => v.date);
  const unitsCount = Object.keys(groupBy_(visits, v => v.unit)).length;
  let h = head_('العمالة حسب المقاول', 'Manpower by Contractor', range ? `${from} → ${to}` : `${dayAr_(from)} · ${from}`, esc_(S['Company name'] || 'O-Plus'));
  const k = [[grand, range ? 'عامل × يوم' : 'عمالة', range ? 'Worker-days' : 'Workers'], [names.length, 'مقاولين', 'Contractors'], [unitsCount, 'وحدات', 'Units']];
  if (range) k.push([Object.keys(days).length, 'أيام', 'Days']);
  h += kpis_(k);
  h += `<table class="sum"><tr><th>المقاول <span class="en">Contractor</span></th><th class="c" style="width:56px">العدد</th>
    <th style="width:150px">حسب الحرفة <span class="en">By trade</span></th><th style="width:190px">حسب الوحدة <span class="en">By unit</span></th></tr>`;
  names.forEach(n => {
    const c = byC[n];
    h += `<tr><td class="b">${esc_(n)}</td><td class="c big">${c.total}</td>
      <td class="sm">${Object.keys(c.trades).sort((a, b) => c.trades[b] - c.trades[a]).map(t => `${esc_(tradeAr_(t))}: <b>${c.trades[t]}</b>`).join('<br>')}</td>
      <td class="sm mono">${Object.keys(c.units).sort().map(u => `${esc_(u)}: <b>${c.units[u]}</b>`).join('<br>')}</td></tr>`;
  });
  h += `<tr class="tot"><td>الإجمالي <span class="en">Total</span></td><td class="c big">${grand}</td><td></td><td></td></tr></table>`;
  if (range) {
    h += `<h2>حسب اليوم <span class="en">By day</span></h2><table class="sum"><tr><th style="width:150px">اليوم</th><th class="c" style="width:56px">العمالة</th><th>المقاولين</th></tr>`;
    Object.keys(days).sort().forEach(d => {
      const c = {};
      days[d].forEach(v => (crew[v.id] || []).forEach(m => c[m.contractor] = (c[m.contractor] || 0) + m.count));
      const tot = Object.keys(c).reduce((s, x) => s + c[x], 0);
      h += `<tr><td>${dayAr_(d)}</td><td class="c b">${tot}</td><td class="sm">${Object.keys(c).sort((a, b) => c[b] - c[a]).map(x => `${esc_(x)} <b>${c[x]}</b>`).join(' · ')}</td></tr>`;
    });
    h += `</table>`;
  }
  const noCrew = visits.filter(v => !(crew[v.id] || []).length);
  if (noCrew.length && !range) h += `<p class="foot">وحدات من غير عمالة: ${noCrew.map(v => esc_(v.unit)).join('، ')}</p>`;
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
/* Every Thursday: one PDF per unit with that unit's photos of the last 7 days and a short, professional
   progress text (rewritten by AI from what the engineers typed). Emailed to "Weekly report to", CC "Weekly report cc".
   Big weeks are made in several rounds (Apps Script stops a run after 6 minutes), then one email batch goes out. */

const WEEK_KEY = 'WEEKLY_RUN';
const WEEK_BUDGET_MS = 200000;          // stop a round at ~3.5 min and continue 1 min later
const MAIL_MAX_BYTES = 15 * 1024 * 1024; // attachments per email

/** Scheduled on the weekly day. */
function weeklyJob() { weeklyStart_(fmt_(new Date(), 'yyyy-MM-dd'), true); }

/** Continues a weekly run that did not finish in one round (set by weeklyWork_). */
function weeklyContinue() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'weeklyContinue') ScriptApp.deleteTrigger(t); });
  weeklyWork_();
}

function weeklyStart_(endDay, send, onlyTo) {
  const startDay = fmt_(new Date(parseDay_(endDay).getTime() - 6 * 86400000), 'yyyy-MM-dd');
  const units = [];
  closedVisits_(v => v.date >= startDay && v.date <= endDay).forEach(v => { if (units.indexOf(v.unit) < 0) units.push(v.unit); });
  units.sort();
  if (!units.length) return { units: 0 };
  PropertiesService.getScriptProperties().setProperty(WEEK_KEY, JSON.stringify({ end: endDay, start: startDay, send: !!send, onlyTo: onlyTo || '', units: units, files: {} }));
  return weeklyWork_();
}

function weeklyWork_() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { busy: true };
  try {
    const raw = props.getProperty(WEEK_KEY);
    if (!raw) return { units: 0 };
    const st = JSON.parse(raw);
    const t0 = Date.now();
    // safety net: if this round gets killed (6-minute limit), the next round still starts by itself
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'weeklyContinue') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('weeklyContinue').timeBased().after(7 * 60 * 1000).create();
    for (const u of st.units) {
      if (st.files[u]) continue;
      if (Date.now() - t0 > WEEK_BUDGET_MS) {
        props.setProperty(WEEK_KEY, JSON.stringify(st));
        ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'weeklyContinue') ScriptApp.deleteTrigger(t); });
        ScriptApp.newTrigger('weeklyContinue').timeBased().after(60 * 1000).create();
        return { units: Object.keys(st.files).filter(k => st.files[k] !== '-').length, pending: true };
      }
      st.files[u] = '-';                                   // if this unit crashes the run, skip it next round
      props.setProperty(WEEK_KEY, JSON.stringify(st));
      const t1 = Date.now();
      let f = null;
      try { f = weeklyUnitPdf_(u, st.end); } catch (e) { console.error('weekly ' + u + ': ' + msg_(e)); }
      if (f) { try { publishWeeklyToClickUp_(u, f, st.end); } catch (e) { console.error('clickup ' + u + ': ' + msg_(e)); } }
      console.log(`weekly ${u}: ${f ? 'ok' : 'failed'} in ${Math.round((Date.now() - t1) / 1000)}s`);
      st.files[u] = f ? f.getId() : '-';
      props.setProperty(WEEK_KEY, JSON.stringify(st));
    }
    props.deleteProperty(WEEK_KEY);
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'weeklyContinue') ScriptApp.deleteTrigger(t); });
    return weeklyFinish_(st);
  } finally { lock.releaseLock(); }
}

function weeklyFinish_(st) {
  const S = settings_();
  const period = `${st.start} → ${st.end}`;
  const clients = {};
  rows_(T.units).forEach(r => clients[low_(r['Unit'])] = { name: str_(r['Client name']), email: str_(r['Client email']) });
  const made = [];
  st.units.forEach(u => {
    const id = st.files[u];
    if (!id || id === '-') return;
    try { made.push({ unit: u, client: clients[u.toLowerCase()] || {}, file: DriveApp.getFileById(id) }); } catch (e) {}
  });
  let mgrPdf = null;
  try { mgrPdf = managersWeeklyPdf_(st.start, st.end); } catch (e) { console.error('managers weekly: ' + msg_(e)); }
  const sentTo = {};
  if (st.send && made.length) {
    // onlyTo: a test / re-send to one address (no cc, never to clients)
    const to = st.onlyTo || S['Weekly report to'] || S['Weekly report copy to'];
    const cc = st.onlyTo ? '' : [S['Weekly report cc'], S['Weekly report to'] ? S['Weekly report copy to'] : ''].filter(Boolean).join(', ');
    if (to) {
      // split into emails of up to ~15 MB of PDFs
      const batches = [];
      let cur = [], size = 0;
      made.forEach(m => {
        const b = m.file.getSize();
        if (cur.length && size + b > MAIL_MAX_BYTES) { batches.push(cur); cur = []; size = 0; }
        cur.push(m); size += b;
      });
      if (cur.length) batches.push(cur);
      batches.forEach((list, i) => {
        const part = batches.length > 1 ? ` (${i + 1}/${batches.length})` : '';
        const att = list.map(m => m.file.getBlob());
        if (i === 0 && mgrPdf) att.unshift(mgrPdf.getBlob());
        const rowsH = list.map(m => `<tr><td style="padding:4px 10px;font-family:monospace">${esc_(m.unit)}</td><td style="padding:4px 10px">${esc_(m.client.name || '')}</td></tr>`).join('');
        MailApp.sendEmail({
          to: to, cc: cc,
          subject: `${S['Company name'] || 'O-Plus'} · Weekly progress reports · ${period} · ${made.length} units${part}`,
          htmlBody: `<div style="font-family:Arial,sans-serif;font-size:14px">
            <p><b>Weekly progress reports · ${period}</b>${part}</p>
            ${i === 0 && mgrPdf ? '<p><b>Internal:</b> "Managers weekly" PDF attached (visits, manpower, units not visited, issues). <b>Do not forward it to clients.</b></p>' : ''}
            <p>${list.length} client report(s) attached, one PDF per unit, ready to send to the client.</p>
            <table style="border-collapse:collapse;border:1px solid #ddd">${rowsH}</table>
            <p style="color:#666">All PDFs are also in Drive: ${REPORTS_ROOT} / Weekly / ${st.end}.</p></div>`,
          attachments: att
        });
        list.forEach(m => sentTo[m.unit] = [to, cc].filter(Boolean).join(', '));
      });
    }
    if (!st.onlyTo && yes_(S['Send weekly to clients'])) {
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
  made.forEach(m => logReport_('Weekly client', period, m.unit, m.file, st.send ? (sentTo[m.unit] || '') : '(made from menu, not emailed)'));
  if (mgrPdf) logReport_('Weekly managers', period, 'Managers', mgrPdf, st.send ? 'with the weekly email' : '(made from menu, not emailed)');
  return { units: made.length };
}

/** Client-facing weekly PDF for one unit: one A4 page in the O West brand.
    Header (logo, date, unit, client), stage tracker, progress ring with on-track badge, up to 3 site photos,
    elapsed vs planned days, contract dates, scope and quotation, latest updates (AI) and the project manager.
    Contract figures come live from the unit's ClickUp tasks (cuReportInfo_). No manpower, no internal notes. */
const BRAND = { navy: '#001641', peach: '#e9b88f', ink: '#1d2433', soft: '#fbf3ec', line: '#e2dcd5', mute: '#7a8194' };
const WK = { gold: '#f0b323', halo: '#fbe6b0', track: '#dfe3ea', grey: '#8a90a0' };
const STAGES = ['Initiation', 'Design', 'Quotation', 'Construction', 'Quality Inspection', 'Hand Over'];
const WK_PHOTOS = 3;          // site photos on the client report (newest of the week); 4 also works (2 rows of 2, smaller)
const ON_TRACK_SLACK = 10;   // actual % may be this many points under planned % and still read "On track"

function weeklyHtml_(S, unit, cl, startDay, endDay, list, photos, crew, info) {
  info = info || {};
  const C = BRAND;
  const days = groupBy_(list, v => v.date);
  const dayKeys = Object.keys(days).sort();
  // scope = what the client bought (ClickUp), never the engineers' photo tags
  const items = info.scope && info.scope.length ? info.scope : rows_(T.scope).filter(r => low_(r['Unit']) === unit.toLowerCase()).map(r => str_(r['Item'])).filter(Boolean)
    .filter((x, i, a) => a.indexOf(x) === i && !/^others?$|^general$/i.test(x));
  const ai = polishWeek_(unit, items, dayKeys.map(d => ({ date: d, notes: days[d].map(v => v.workDone).filter(Boolean) })));

  // up to WK_PHOTOS site photos: the newest landscape ones of the week (portrait only if there are not enough)
  let all = [];
  dayKeys.forEach(d => days[d].forEach(v => (photos[v.id] || []).forEach(p => all.push({ p: p, day: d }))));
  const recent = all.slice(-12);
  const imgs = images_(recent.map(x => x.p.fileId));
  const ok = recent.filter(x => imgs[x.p.fileId]);
  const wide = ok.filter(x => { const im = imgs[x.p.fileId]; return im.w && im.h && im.w >= im.h; });
  let pick = wide.slice(-WK_PHOTOS);
  if (pick.length < WK_PHOTOS) pick = pick.concat(ok.filter(x => pick.indexOf(x) < 0).slice(-(WK_PHOTOS - pick.length)));
  pick.sort((a, b) => ok.indexOf(a) - ok.indexOf(b));
  const photoDay = pick.length ? pick[pick.length - 1].day : endDay;
  // 1-2 photos: one row of 2; 3: one row of 3; 4: two rows of 2
  const cols = pick.length === 3 ? 3 : 2, FULL = 668, GAP = 12;
  const PW = Math.floor((FULL - GAP * (cols - 1)) / cols), PH = pick.length === 3 ? 215 : pick.length > 3 ? 128 : 190;
  const img = x => {
    const im = x && imgs[x.p.fileId];
    if (!im) return '';
    const s = im.w && im.h ? Math.min(PW / im.w, PH / im.h) : 0;
    const w = s ? Math.round(im.w * s) : PW, h = s ? Math.round(im.h * s) : 0;
    const pad = h ? Math.floor((PH - h) / 2) : 0;
    return `<div style="height:${pad}px"></div><img src="${im.src}" width="${w}"${h ? ` height="${h}"` : ''}>`;
  };
  let photoRows = '';
  for (let i = 0; i < pick.length; i += cols) {
    photoRows += `<table class="photos"><tr>` + Array.from({ length: cols }, (_, j) => {
      const x = pick[i + j];
      const pad = `padding-left:${j ? GAP / 2 : 0}px;padding-right:${j < cols - 1 ? GAP / 2 : 0}px`;
      return `<td class="p" style="width:${Math.round(100 / cols)}%;${pad}">${x ? `<div class="pbox">${img(x)}</div>` : ''}</td>`;
    }).join('') + '</tr></table>';
  }

  // stage tracker
  const stage = Math.max(0, STAGES.indexOf(info.stage || 'Construction'));
  const stageCells = STAGES.map((name, i) => {
    const kind = i < stage ? 'done' : i === stage ? (stage === STAGES.length - 1 ? 'done' : 'now') : 'next';
    const lineL = i === 0 ? '' : `<div class="ln" style="border-color:${i <= stage ? C.navy : WK.track}"></div>`;
    const lineR = i === STAGES.length - 1 ? '' : `<div class="ln" style="border-color:${i < stage ? C.navy : WK.track}"></div>`;
    return `<td class="st"><table class="stt"><tr><td class="side">${lineL}</td><td class="ico"><img src="${stageIcon_(kind)}" width="40" height="40"></td><td class="side">${lineR}</td></tr></table>
      <div class="stn" style="color:${kind === 'next' ? WK.grey : C.navy};font-weight:${kind === 'now' ? 700 : 500}">${esc_(name)}</div></td>`;
  }).join('');

  // progress and time
  const actual = info.poc == null ? null : Math.max(0, Math.min(100, Math.round(info.poc)));
  const t0 = info.start ? parseDay_(info.start).getTime() : null, t1 = info.finish ? parseDay_(info.finish).getTime() : null;
  const now = parseDay_(endDay).getTime();
  const planned = t0 && t1 && t1 > t0 ? Math.round((t1 - t0) / 86400000) : null;
  const elapsed = t0 ? Math.max(0, Math.round((now - t0) / 86400000)) : null;
  const plannedPct = planned ? Math.max(0, Math.min(100, elapsed / planned * 100)) : null;
  let badge = '';
  if (actual != null) badge = actual >= 100 ? 'Completed' : plannedPct == null || actual >= plannedPct - ON_TRACK_SLACK ? 'On track' : 'Behind plan';
  const barPct = planned ? Math.max(0, Math.min(100, Math.round(elapsed / planned * 100))) : 0;
  const mob = info.payment && info.start ? Math.round((parseDay_(info.start) - parseDay_(info.payment)) / 86400000) : null;

  const d_ = s => s ? esc_(dateEn_(s)) : '—';
  const updates = ai && ai.en.length ? ai.en.slice(0, 4).map(x => /[.!?]$/.test(x) ? x : x + '.').join(' ') : '';
  const money = n => n ? 'EGP ' + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '—';
  const unitName = info.unitName || unit;
  const client = info.client || cl.name || '';

  // O West brand: Poppins throughout (bold titles), peach kickers on navy, logos as in the action plan
  const INK = '#0b1a3a', MUTE = '#5a6378', PEACH = '#f3cfb3', LINE = '#e2dbd5', NAVY = '#001640';
  let h = `<html lang="en"><head><meta charset="utf-8"><style>
    ${fontsCss_()}
    @page { size: A4; margin: 0; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    html, body { margin: 0; padding: 0; }
    body { font-family: 'Poppins', Arial, sans-serif; font-size: 8.6pt; line-height: 1.5; color: ${INK}; }
    table { border-collapse: collapse; width: 100%; }
    td { vertical-align: top; padding: 0; }
    .serif { font-family: 'Liberation Serif', 'Times New Roman', serif; font-weight: bold; }
    .sheet { position: relative; height: 296mm; overflow: hidden; }
    .band td.b { background: ${NAVY}; color: #ffffff; padding: 12pt 34pt 14pt; }
    .kick { font-size: 7.5pt; letter-spacing: 2.2px; color: ${PEACH}; text-transform: uppercase; font-weight: 500; }
    .rule { border-top: 1px solid #2b3b63; margin: 12pt 0 10pt; height: 0; }
    .blab { font-size: 7pt; letter-spacing: 2px; color: ${PEACH}; text-transform: uppercase; }
    .unit { font-weight: 700; font-size: 25pt; letter-spacing: 0.5px; color: #ffffff; line-height: 1.15; margin-top: 4px; }
    .client { font-size: 10pt; color: #ffffff; font-weight: 500; text-align: right; margin-top: 6px; }
    .sub { font-size: 8.3pt; color: #d5d9e3; margin-top: 6px; }
    .sub b { color: #ffffff; font-weight: 500; }
    .body { padding: 0 34pt; }
    .stages { margin-top: 16pt; table-layout: fixed; }
    .st { text-align: center; }
    .stt td { vertical-align: middle; }
    .stt .side { width: 50%; }
    .stt .ico { width: 40px; }
    .ln { border-top: 3px solid; height: 0; }
    .stn { font-size: 7.6pt; margin-top: 5px; white-space: nowrap; }
    .lab { font-size: 7pt; letter-spacing: 1.8px; color: ${MUTE}; text-transform: uppercase; font-weight: 500; }
    .val { font-size: 9.5pt; color: ${INK}; margin-top: 3px; font-weight: 500; }
    .shead { margin-top: 13pt; }
    .shead td { border-bottom: 1.5px solid ${NAVY}; padding-bottom: 4px; vertical-align: bottom; }
    .shead h2 { font-size: 12.5pt; font-weight: 700; color: ${INK}; margin: 0; }
    .shead .r { font-size: 7.5pt; letter-spacing: 1.8px; color: ${MUTE}; text-transform: uppercase; font-weight: 500; text-align: right; }
    .photos { margin-top: 7pt; }
    .photos td.p { text-align: center; }
    .pbox { height: ${PH}px; background: #faf2ec; overflow: hidden; text-align: center; }
    .ph-empty { padding-top: ${Math.round(PH / 2) - 10}px; color: ${MUTE}; font-size: 9pt; }
    .card { background: #f4f5f8; border: 1px solid #e6e8ee; padding: 14pt 18pt; height: 104px; }
    .ring { position: relative; width: 104px; height: 104px; }
    .ring .pct { position: absolute; left: 0; right: 0; top: 39px; height: 28px; text-align: center; font-size: 17pt; font-weight: 700; color: ${INK}; line-height: 28px; }
    .badge { display: inline-block; border: 1px solid ${C.navy}; color: ${INK}; font-weight: 500; font-size: 7.6pt; padding: 2px 12px; border-radius: 10px; margin-top: 8px; background: #ffffff; }
    .badge.late { border-color: #c23a28; color: #c23a28; }
    .big { font-weight: 700; font-size: 24pt; color: ${INK}; line-height: 1.1; }
    .of { font-size: 8.6pt; color: ${MUTE}; }
    .bar { background: ${WK.track}; height: 7px; margin: 9px 0 7px; overflow: hidden; }
    .bar div { background: ${NAVY}; height: 7px; }
    .small { font-size: 7pt; color: ${MUTE}; }
    .sv { font-size: 8.6pt; color: ${INK}; font-weight: 500; }
    .facts { margin-top: 12pt; border-top: 1px solid ${LINE}; }
    .facts td { padding: 7pt 8px 7pt 0; border-bottom: 1px solid ${LINE}; }
    .utext { font-size: 9pt; line-height: 1.75; color: ${INK}; margin-top: 8pt; }
    .pm { margin-top: 12pt; background: #faf2ec; border-left: 3px solid ${PEACH}; }
    .pm td { padding: 10pt 16pt; vertical-align: middle; font-size: 8.6pt; }
    .foot { position: absolute; left: 34pt; right: 34pt; bottom: 12pt; }
    .foot td { font-size: 7.5pt; color: ${MUTE}; border-top: 1px solid ${LINE}; padding-top: 6pt; }
  </style></head><body><div class="sheet">
  <table class="band"><tr><td class="b">
    <table><tr>
      <td style="vertical-align:middle"><img src="data:image/png;base64,${LOGO_OWEST}" height="24"><span style="display:inline-block;width:1px;height:24px;background:#5a6a90;margin:0 14px;vertical-align:top"></span><img src="data:image/png;base64,${LOGO_OPLUS}" height="24"></td>
      <td style="vertical-align:middle;text-align:right"><div class="kick" style="text-align:right">Weekly report · ${d_(endDay).toUpperCase()}</div></td>
    </tr></table>
    <div class="rule"></div>
    <table><tr>
      <td><div class="blab">Unit</div><div class="unit">${esc_(unitName)}</div></td>
      <td style="text-align:right;vertical-align:bottom"><div class="blab" style="text-align:right">Client</div><div class="client">${esc_(client || '—')}</div></td>
    </tr></table>
  </td></tr></table>
  <div class="body">
  <table class="stages"><tr>${stageCells}</tr></table>
  <table class="shead"><tr><td><h2>Progress</h2></td><td class="r">As of ${d_(endDay).toUpperCase()}</td></tr></table>
  <table style="margin-top:8pt"><tr>
    <td style="width:50%;padding-right:6px"><div class="card"><table><tr>
      <td style="width:118px;vertical-align:middle"><div class="ring"><img src="${ringPng_(actual || 0)}" width="104" height="104"><div class="pct">${actual == null ? '—' : actual + '%'}</div></div></td>
      <td style="vertical-align:middle;padding-left:8px"><div class="lab">Actual progress</div>${badge ? `<div class="badge${badge === 'Behind plan' ? ' late' : ''}">${badge}</div>` : ''}</td>
    </tr></table></div></td>
    <td style="width:50%;padding-left:6px"><div class="card">
      <div class="lab">Elapsed duration</div>
      <div style="margin-top:4px"><span class="big">${elapsed == null ? '—' : elapsed}</span> <span class="of">${planned ? `of ${planned} planned days` : 'days'}</span></div>
      <div class="bar"><div style="width:${barPct}%"></div></div>
      <table><tr>
        <td><div class="small">Start</div><div class="sv">${d_(info.start)}</div></td>
        <td style="text-align:right"><div class="small">Contract finish</div><div class="sv">${d_(info.finish)}</div></td>
      </tr></table>
    </div></td>
  </tr></table>
  ${pick.length ? `<table class="shead"><tr><td><h2>Site photos</h2></td><td class="r">${d_(photoDay).toUpperCase()}</td></tr></table>${photoRows}` : ''}
  <table class="facts"><tr>
    <td style="width:20%"><div class="lab">Payment date</div><div class="val">${d_(info.payment)}</div></td>
    <td style="width:20%"><div class="lab">Mobilization</div><div class="val">${mob == null ? '—' : mob + ' days'}</div></td>
    <td style="width:20%"><div class="lab">Contract start</div><div class="val">${d_(info.start)}</div></td>
    <td style="width:20%"><div class="lab">Contract finish</div><div class="val">${d_(info.finish)}</div></td>
    <td style="width:20%"><div class="lab">Expected finish</div><div class="val" style="font-weight:700">${d_(info.expected || info.finish)}</div></td>
  </tr><tr>
    <td colspan="2"><div class="lab">Scope of work</div><div class="val">${esc_(items.join(', ') || '—')}</div></td>
    <td><div class="lab">Quotation ref.</div><div class="val">${esc_((info.quotes || []).join(', ') || '—')}</div></td>
    <td colspan="2"><div class="lab">Quotation amount</div><div class="val">${money(info.amount)}</div></td>
  </tr></table>
  ${updates ? `<table class="shead"><tr><td><h2>Latest updates</h2></td><td class="r">This week</td></tr></table><div class="utext">${esc_(updates)}</div>` : ''}
  <table class="pm"><tr>
    <td><span style="color:${MUTE}">Your project manager</span> &nbsp; <span style="color:${INK};font-weight:700;font-size:9.5pt">${esc_((info.pm || []).join(' & ') || '—')}</span></td>
    <td style="text-align:right;color:${INK}">Questions? Reply to this message.</td>
  </tr></table>
  </div>
  <div class="foot"><table><tr><td>O West by Orascom Development · ${esc_(unitName)}</td><td style="text-align:right">Weekly report · Page 1 of 1</td></tr></table></div>
  </div></body></html>`;
  return h;
}

/* ---------- report data from ClickUp ---------- */
// field IDs on the Main Tracking Sheet / Construction Team tasks
const CU_RPT = {
  poc: '640aa77a-6d8d-4b4d-b8a6-b4dfd74baa9f', payment: 'b2c472de-5d86-4e50-ad17-571055eb20e0',
  siteIn: 'e099ba01-7bd4-402b-9d0f-2263884b39b0', siteDue: '192b3218-0025-48f1-a418-488d0e87e5ee',
  quote: '432de28e-bb8f-4595-92ba-d2f8bd564df8', amount: 'f480622e-13eb-43c9-a141-f74a78d866e9',
  pm: 'f91d9a67-0ed7-4e6a-81c3-1b7bf673546f', cstatus: 'eb44d8ec-34af-462d-b01e-87f84b918f55',
  unitCode: '536dcc2d-7696-4813-88ed-e3f4966f793e'
};
// ClickUp task status -> stage on the tracker
const CU_STAGE = {
  'new case': 0, 'on hold': 0, 'client feedback': 1, 'design revision': 1, 'design execution': 1,
  'quotation issued': 2, 'cost revision': 2, 'site execution': 3, 'handed over': 5
};

/** Everything the client report needs from the unit's ClickUp tasks, read live (one small call per task).
    Several tasks for one unit (for example a finishing package and a pergola) are combined:
    scopes listed, amounts added, progress weighted by amount, earliest start and latest finish. */
function cuReportInfo_(unit) {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (!token) return null;
  const k = unit.toLowerCase();
  let ids = rows_(CU_TAB).filter(r => low_(normUnit_(r['Unit'])) === k && ['deleted', 'declined'].indexOf(low_(r['Status'])) < 0).map(r => str_(r['Task ID'])).filter(Boolean);
  if (!ids.length) ids = cuTaskIdsForUnit_(unit, token);
  if (!ids.length) return null;
  const got = UrlFetchApp.fetchAll(ids.map(id => ({ url: `https://api.clickup.com/api/v2/task/${id}`, headers: { Authorization: token }, muteHttpExceptions: true })));
  const tasks = [];
  got.forEach(r => {
    if (r.getResponseCode() !== 200) return;
    let t; try { t = JSON.parse(r.getContentText()); } catch (e) { return; }
    const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
    const v = id => str_(f[id] && f[id].value);
    const day = id => { const x = Number(v(id)); return x ? fmt_(new Date(x), 'yyyy-MM-dd') : ''; };
    const num = id => { const x = parseFloat(v(id)); return isNaN(x) ? null : x; };
    tasks.push({
      status: low_(t.status && t.status.status), unitCode: v(CU_RPT.unitCode), client: v(CU.f.client),
      type: dropName_(f[CU.f.requestType]), poc: num(CU_RPT.poc), amount: num(CU_RPT.amount) || 0,
      payment: day(CU_RPT.payment), start: day(CU_RPT.siteIn), finish: day(CU_RPT.siteDue),
      quote: v(CU_RPT.quote), pm: dropName_(f[CU_RPT.pm]), cstatus: low_(dropName_(f[CU_RPT.cstatus]))
    });
  });
  const live = tasks.filter(t => t.status !== 'declined');
  if (!live.length) return null;
  const site = live.filter(t => t.status === 'site execution');
  const main = site.length ? site : live;

  let stage;
  if (site.length) stage = site.every(t => t.cstatus === 'completed' || t.cstatus === 'handedover') ? 4 : 3;
  else if (live.every(t => t.status === 'handed over')) stage = 5;
  else stage = Math.max.apply(null, live.filter(t => t.status !== 'handed over').map(t => CU_STAGE[t.status] == null ? 0 : CU_STAGE[t.status]));

  const uniq = a => a.filter((x, i) => x && a.indexOf(x) === i);
  const withPoc = main.filter(t => t.poc != null);
  const wSum = withPoc.reduce((s, t) => s + (t.amount || 0), 0);
  const poc = !withPoc.length ? null : wSum > 0
    ? withPoc.reduce((s, t) => s + t.poc * (t.amount || 0), 0) / wSum
    : withPoc.reduce((s, t) => s + t.poc, 0) / withPoc.length;
  const min = a => a.filter(Boolean).sort()[0] || '', max = a => a.filter(Boolean).sort().slice(-1)[0] || '';
  return {
    unitName: (main.find(t => t.unitCode) || {}).unitCode || '',
    client: ((main.find(t => t.client) || live.find(t => t.client) || {}).client || '').replace(/\b[a-z]/g, c => c.toUpperCase()),
    stage: STAGES[stage], poc: poc,
    scope: (s => s.length > 1 ? s.filter(x => !/^others?$/i.test(x)) : s)(uniq(main.map(t => t.type).filter(x => CU.skipTypes.indexOf(low_(x)) < 0))),
    quotes: uniq(main.map(t => t.quote)), amount: main.reduce((s, t) => s + (t.amount || 0), 0),
    payment: min(main.map(t => t.payment)), start: min(main.map(t => t.start)), finish: max(main.map(t => t.finish)),
    pm: uniq(main.map(t => t.pm))
  };
}

/* ---------- small PNG images drawn in code (the PDF maker has no SVG or canvas) ---------- */
const PNG_CACHE = {};
function hex_(c) { return [parseInt(c.slice(1, 3), 16), parseInt(c.slice(3, 5), 16), parseInt(c.slice(5, 7), 16)]; }

/** Progress ring, 0..100 %, navy arc on a light track, clockwise from 12 o'clock. */
function ringPng_(pct) {
  const key = 'ring' + Math.round(pct);
  if (PNG_CACHE[key]) return PNG_CACHE[key];
  const S = 224, c = S / 2, ro = 108, ri = 84, end = Math.max(0, Math.min(100, pct)) / 100 * 2 * Math.PI;
  const navy = hex_(BRAND.navy), track = hex_(WK.track);
  return (PNG_CACHE[key] = rasterPng_(S, (x, y) => {
    const dx = x - c, dy = y - c, r = Math.sqrt(dx * dx + dy * dy);
    if (r < ri || r > ro) return null;
    let a = Math.atan2(dx, -dy); if (a < 0) a += 2 * Math.PI;
    return a <= end ? navy : track;
  }));
}

/** Stage tracker icon: done (navy with white tick), now (gold with halo and navy dot), next (white with grey ring). */
function stageIcon_(kind) {
  const key = 'st' + kind;
  if (PNG_CACHE[key]) return PNG_CACHE[key];
  const S = 88, c = S / 2;
  const navy = hex_(BRAND.navy), gold = hex_(WK.gold), halo = hex_(WK.halo), grey = hex_('#d5d9e2'), white = [255, 255, 255];
  const seg = (px, py, ax, ay, bx, by) => {   // distance from a point to a line segment
    const vx = bx - ax, vy = by - ay, t = Math.max(0, Math.min(1, ((px - ax) * vx + (py - ay) * vy) / (vx * vx + vy * vy)));
    const qx = ax + t * vx - px, qy = ay + t * vy - py;
    return Math.sqrt(qx * qx + qy * qy);
  };
  return (PNG_CACHE[key] = rasterPng_(S, (x, y) => {
    const r = Math.sqrt((x - c) * (x - c) + (y - c) * (y - c));
    if (kind === 'done') {
      if (r > 30) return null;
      return Math.min(seg(x, y, 30, 45, 39, 54), seg(x, y, 39, 54, 57, 35)) < 3.6 ? white : navy;
    }
    if (kind === 'now') {
      if (r > 42) return null;
      if (r > 32) return halo;
      return r < 9 ? navy : gold;
    }
    if (r > 30) return null;
    return r > 26 ? grey : white;
  }));
}

/** Draws an S x S image from fn(x, y) -> [r, g, b] or null (transparent), 4x4 samples per pixel, as a PNG data URL. */
function rasterPng_(S, fn) {
  const px = new Array(S * S * 4);
  for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
    let r = 0, g = 0, b = 0, n = 0;
    for (let sy = 0; sy < 4; sy++) for (let sx = 0; sx < 4; sx++) {
      const col = fn(x + (sx + 0.5) / 4, y + (sy + 0.5) / 4);
      if (col) { r += col[0]; g += col[1]; b += col[2]; n++; }
    }
    const i = (y * S + x) * 4;
    px[i] = n ? Math.round(r / n) : 0; px[i + 1] = n ? Math.round(g / n) : 0; px[i + 2] = n ? Math.round(b / n) : 0; px[i + 3] = Math.round(n / 16 * 255);
  }
  return 'data:image/png;base64,' + Utilities.base64Encode(pngBytes_(S, S, px).map(b => b > 127 ? b - 256 : b));
}

let CRC_T = null;
/** Minimal RGBA PNG encoder (uncompressed deflate blocks; the images are small). Returns bytes 0..255. */
function pngBytes_(w, h, px) {
  if (!CRC_T) { CRC_T = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; CRC_T[n] = c >>> 0; } }
  const crc = a => { let c = 0xFFFFFFFF; for (let i = 0; i < a.length; i++) c = CRC_T[(c ^ a[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const u32 = n => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
  const chunk = (type, data) => { const td = type.split('').map(ch => ch.charCodeAt(0)).concat(data); return u32(data.length).concat(td, u32(crc(td))); };
  const raw = [];
  for (let y = 0; y < h; y++) { raw.push(0); for (let x = 0; x < w * 4; x++) raw.push(px[y * w * 4 + x]); }
  const z = [0x78, 0x01];
  for (let i = 0; i < raw.length; i += 65535) {
    const n = Math.min(65535, raw.length - i);
    z.push(i + n >= raw.length ? 1 : 0, n & 255, n >>> 8, ~n & 255, (~n >>> 8) & 255);
    for (let j = 0; j < n; j++) z.push(raw[i + j]);
  }
  let a = 1, b = 0;
  for (let i = 0; i < raw.length; i++) { a = (a + raw[i]) % 65521; b = (b + a) % 65521; }
  z.push.apply(z, u32(((b << 16) | a) >>> 0));
  return [137, 80, 78, 71, 13, 10, 26, 10].concat(chunk('IHDR', u32(w).concat(u32(h), [8, 6, 0, 0, 0])), chunk('IDAT', z), chunk('IEND', []));
}

/* =========================== WEEKLY → CLICKUP (WhatsApp button) =========================== */
/* For each unit: share its weekly PDF (anyone with the link can view), then on every site-execution task of the unit
   (added to the private Construction Team list if missing) fill "Weekly Report" = PDF link and
   "WhatsApp Client" = wa.me link to the client's Mobile with the message ready. */

let CU_UNITS = null;
/** Unit -> its ClickUp task IDs, from the "ClickUp Tasks" tab (the full list API answer is too big, over 50 MB).
    Units missing from the tab are looked up by the "Unit Code" field. Client name/mobile are read task by task. */
function cuUnitTasks_() {
  if (CU_UNITS) return CU_UNITS;
  const out = {};
  rows_('ClickUp Tasks').forEach(r => {
    const u = normUnit_(r['Unit']), id = str_(r['Task ID']), st = low_(r['Status']);
    if (u && id && (!st || st === CU.status)) (out[u.toLowerCase()] = out[u.toLowerCase()] || []).push(id);
  });
  return (CU_UNITS = out);
}

function cuTaskIdsForUnit_(unit, token) {
  const k = unit.toLowerCase();
  const map = cuUnitTasks_();
  if (map[k] && map[k].length) return map[k];
  // fallback: tasks whose "Unit Code" field equals the unit
  const cf = encodeURIComponent(JSON.stringify([{ field_id: '536dcc2d-7696-4813-88ed-e3f4966f793e', operator: '=', value: unit }]));
  const res = UrlFetchApp.fetch(`https://api.clickup.com/api/v2/list/${CU.list}/task?statuses[]=${encodeURIComponent(CU.status)}&custom_fields=${cf}`, { headers: { Authorization: token }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  const ids = (JSON.parse(res.getContentText()).tasks || []).filter(t => normUnit_(t.name).toLowerCase() === k).map(t => t.id);
  if (ids.length) {
    const sh = sheet_('ClickUp Tasks', CU_COLS);
    ids.forEach(id => sh.appendRow([unit, id, CU.status]));
    map[k] = ids;
  }
  return ids;
}

/** "+20 12 75024002" / "01275024002" -> "201275024002" (WhatsApp format). */
function waNumber_(m) {
  let d = str_(m).replace(/\D/g, '');
  if (!d) return '';
  if (d.indexOf('00') === 0) d = d.slice(2);
  if (/^01\d{9}$/.test(d)) d = '2' + d;           // Egyptian local mobile
  else if (/^1\d{9}$/.test(d)) d = '20' + d;
  return d.length >= 10 ? d : '';
}

function publishWeeklyToClickUp_(unit, file, endDay) {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (!token) return;
  const ids = cuTaskIdsForUnit_(unit, token);
  if (!ids.length) { console.error('clickup ' + unit + ': no tasks found'); return; }
  // read client name + mobile from each task (one small call per task)
  const got = UrlFetchApp.fetchAll(ids.map(id => ({ url: `https://api.clickup.com/api/v2/task/${id}`, headers: { Authorization: token }, muteHttpExceptions: true })));
  const tasks = ids.map((id, i) => {
    const t = { id: id, client: '', mobile: '' };
    try {
      if (got[i].getResponseCode() === 200) {
        const j = JSON.parse(got[i].getContentText());
        (j.custom_fields || []).forEach(c => { if (c.id === CU.f.client) t.client = str_(c.value); if (c.id === CU.f.mobile) t.mobile = str_(c.value); });
      }
    } catch (e) {}
    return t;
  });
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const link = `https://drive.google.com/file/d/${file.getId()}/view`;
  const t0 = tasks.find(t => t.mobile) || tasks[0];
  const client = (tasks.find(t => t.client) || {}).client || '';
  const text = `Dear ${client || 'Client'},\nPlease find this week's progress report for your unit ${unit} (week ending ${dateEn_(endDay)}):\n${link}\n\nBest regards,\nO-Plus Extra Works · O West`;
  const num = waNumber_(t0.mobile);
  const wa = `https://wa.me/${num}?text=${encodeURIComponent(text)}`;
  const H = { Authorization: token, 'Content-Type': 'application/json' };
  const reqs = [];
  tasks.forEach(t => {
    reqs.push({ url: `https://api.clickup.com/api/v2/list/${CU.team.list}/task/${t.id}`, method: 'post', headers: H, muteHttpExceptions: true });
  });
  UrlFetchApp.fetchAll(reqs);   // add to Construction Team (already there = harmless error)
  const sets = [];
  tasks.forEach(t => {
    sets.push({ url: `https://api.clickup.com/api/v2/task/${t.id}/field/${CU.team.report}`, method: 'post', headers: H, payload: JSON.stringify({ value: link }), muteHttpExceptions: true });
    sets.push({ url: `https://api.clickup.com/api/v2/task/${t.id}/field/${CU.team.whatsapp}`, method: 'post', headers: H, payload: JSON.stringify({ value: wa }), muteHttpExceptions: true });
  });
  const res = UrlFetchApp.fetchAll(sets);
  const bad = res.filter(x => x.getResponseCode() !== 200);
  if (bad.length) console.error(`clickup ${unit}: ${bad.length} field updates failed: ${bad[0].getResponseCode()} ${bad[0].getContentText().slice(0, 200)}`);
}

/** Run from the editor to test one unit end to end: makes its weekly PDF and fills the ClickUp button. */
function testWeeklyOneUnit() {
  const unit = 'OW/RV6-11-VF', end = fmt_(new Date(), 'yyyy-MM-dd');
  const f = weeklyUnitPdf_(unit, end);
  if (!f) { Logger.log('No visits for ' + unit + ' in the last 7 days'); return; }
  publishWeeklyToClickUp_(unit, f, end);
  Logger.log('PDF: https://drive.google.com/file/d/' + f.getId() + '/view');
  Logger.log('ClickUp tasks for this unit: ' + (cuUnitTasks_()[unit.toLowerCase()] || []).length + ' (see Executions for any errors)');
}

/* =========================== MANAGERS WEEKLY (internal) =========================== */
/* Everything the client must not see: coverage gaps, idle days, manpower, raw notes, internal notes. */
const NEG_RE = /مفيش\s*(شغل|عمال|عماله|عمالة)|لا\s*يوجد|واقف|متوقف|توقف|مستني|منتظر|انتظار|تأخير|متأخر|مشكل|ناقص|نقص|مافيش|no\s*work|stopp|waiting|delay|missing|issue|problem/i;

function managersWeeklyPdf_(startDay, endDay) {
  const visits = closedVisits_(v => v.date >= startDay && v.date <= endDay);
  const crew = crewByVisit_(visits);
  const byUnit = groupBy_(visits, v => v.unit);
  const active = rows_(T.units).filter(r => str_(r['Unit']) && low_(r['Active']) !== 'no').map(r => ({ unit: str_(r['Unit']), client: str_(r['Client name']), project: str_(r['Project']) }));
  const notVisited = active.filter(u => !byUnit[u.unit]);
  const last = {};
  rows_(T.visits).forEach(r => { if (str_(r['Status']) !== 'Closed') return; const k = low_(r['Unit']), d = day_(r['Date']); if (d > (last[k] || '')) last[k] = d; });
  const workers = visits.reduce((s, v) => s + v.manpower, 0);
  const byC = {}, byEng = {};
  visits.forEach(v => {
    (crew[v.id] || []).forEach(m => byC[m.contractor] = (byC[m.contractor] || 0) + m.count);
    const e = byEng[v.by] = byEng[v.by] || { visits: 0, units: {}, days: {} };
    e.visits++; e.units[v.unit] = 1; e.days[v.date] = 1;
  });
  const flags = visits.filter(v => !v.manpower || NEG_RE.test(v.workDone) || v.notes);
  const C = BRAND;
  let h = `<html lang="en"><head><meta charset="utf-8"><style>
    @page { size: A4; margin: 11mm; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    body { font-family: Arial, sans-serif; font-size: 8.5pt; color: ${C.ink}; }
    table { border-collapse: collapse; width: 100%; }
    tr { page-break-inside: avoid; }
    .band td { background: ${C.navy}; color: #fff; padding: 12px 16px; }
    .kick { font-size: 7pt; letter-spacing: 2.5px; color: ${C.peach}; }
    .title { font-family: "Times New Roman", serif; font-weight: bold; font-size: 18pt; margin-top: 3px; }
    .kpi td { width: 20%; text-align: center; border: 1px solid ${C.line}; padding: 8px 4px; }
    .kpi .n { font-family: "Times New Roman", serif; font-size: 17pt; font-weight: bold; color: ${C.navy}; }
    .kpi .k { font-size: 7pt; letter-spacing: 1.5px; color: ${C.mute}; }
    .kpi .bad .n { color: #b3261e; }
    h2 { font-family: "Times New Roman", serif; font-size: 12.5pt; color: ${C.navy}; margin: 14px 0 4px; border-bottom: 1.5px solid ${C.navy}; padding-bottom: 3px; }
    .t th { background: ${C.soft}; color: ${C.navy}; text-align: left; font-size: 7.5pt; padding: 4px 5px; border-bottom: 1px solid ${C.line}; }
    .t td { padding: 4px 5px; border-bottom: 1px solid ${C.line}; vertical-align: top; }
    .ar { direction: rtl; text-align: right; }
    .r { text-align: right; } .b { font-weight: bold; } .red { color: #b3261e; font-weight: bold; } .mute { color: ${C.mute}; }
    .mono { font-family: "Courier New", monospace; }
  </style></head><body>
  <table class="band"><tr><td><img src="data:image/png;base64,${LOGO_OWEST}" height="22"> &nbsp; <img src="data:image/png;base64,${LOGO_OPLUS}" height="22">
    <div class="kick" style="margin-top:10px">INTERNAL · MANAGERS ONLY · DO NOT FORWARD</div>
    <div class="title">Site execution weekly review</div><div style="font-size:8.5pt;color:#c9cfdd">${esc_(dateEn_(startDay))} – ${esc_(dateEn_(endDay))}</div></td></tr></table>
  <div style="height:8px"></div>
  <table class="kpi"><tr>
    <td><div class="n">${Object.keys(byUnit).length}</div><div class="k">UNITS VISITED</div></td>
    <td class="${notVisited.length ? 'bad' : ''}"><div class="n">${notVisited.length}</div><div class="k">ACTIVE UNITS NOT VISITED</div></td>
    <td><div class="n">${visits.length}</div><div class="k">VISITS</div></td>
    <td><div class="n">${workers}</div><div class="k">WORKER-DAYS</div></td>
    <td class="${flags.length ? 'bad' : ''}"><div class="n">${flags.length}</div><div class="k">VISITS FLAGGED</div></td>
  </tr></table>`;

  h += `<h2>Flags · no workers, stoppages, waiting, internal notes</h2>`;
  h += flags.length ? `<table class="t"><tr><th style="width:62px">Date</th><th style="width:100px">Unit</th><th style="width:70px">Engineer</th><th style="width:40px" class="r">Workers</th><th>Work done (as typed)</th><th>Internal notes</th></tr>` +
    flags.sort((a, b) => (a.date + a.unit).localeCompare(b.date + b.unit)).map(v => `<tr><td>${esc_(v.date.slice(5))}</td><td class="mono">${esc_(v.unit)}</td><td>${esc_(v.by)}</td><td class="r ${v.manpower ? '' : 'red'}">${v.manpower}</td><td class="ar">${br_(v.workDone)}</td><td class="ar">${br_(v.notes)}</td></tr>`).join('') + '</table>'
    : '<p class="mute">Nothing flagged this week.</p>';

  h += `<h2>Active units not visited this week (${notVisited.length})</h2>`;
  h += notVisited.length ? `<table class="t"><tr><th style="width:130px">Unit</th><th>Client</th><th style="width:110px">Last visit</th></tr>` +
    notVisited.map(u => `<tr><td class="mono">${esc_(u.unit)}</td><td>${esc_(u.client)}</td><td>${esc_(last[u.unit.toLowerCase()] || 'never')}</td></tr>`).join('') + '</table>'
    : '<p class="mute">Every active unit was visited.</p>';

  h += `<h2>Manpower by contractor (worker-days)</h2><table class="t"><tr><th>Contractor</th><th class="r" style="width:80px">Worker-days</th></tr>` +
    Object.keys(byC).sort((a, b) => byC[b] - byC[a]).map(c => `<tr><td>${esc_(c)}</td><td class="r b">${byC[c]}</td></tr>`).join('') + '</table>';

  h += `<h2>Engineers</h2><table class="t"><tr><th>Engineer</th><th class="r">Visits</th><th class="r">Units</th><th class="r">Days on site</th></tr>` +
    Object.keys(byEng).sort().map(e => `<tr><td>${esc_(e)}</td><td class="r">${byEng[e].visits}</td><td class="r">${Object.keys(byEng[e].units).length}</td><td class="r">${Object.keys(byEng[e].days).length}</td></tr>`).join('') + '</table>';

  h += `<h2>All visits by unit</h2><table class="t"><tr><th style="width:100px">Unit</th><th style="width:44px">Date</th><th style="width:70px">Engineer</th><th style="width:36px" class="r">Wkrs</th><th>Contractors</th><th>Work done (as typed)</th></tr>`;
  Object.keys(byUnit).sort().forEach(u => byUnit[u].sort((a, b) => a.closed.localeCompare(b.closed)).forEach((v, i) => {
    const cs = groupBy_(crew[v.id] || [], m => m.contractor);
    h += `<tr><td class="mono">${i ? '' : esc_(u)}</td><td>${esc_(v.date.slice(5))}</td><td>${esc_(v.by)}</td><td class="r">${v.manpower}</td><td>${Object.keys(cs).map(c => esc_(c) + ' ' + cs[c].reduce((s, m) => s + m.count, 0)).join(', ')}</td><td class="ar">${br_(v.workDone)}</td></tr>`;
  }));
  h += '</table></body></html>';
  return savePdf_(h, reportsFolder_('Weekly', endDay), `Managers weekly ${endDay}.pdf`);
}

/** "Tue 10 Nov 2026" (or without weekday). */
function dateEn_(s, withDay) {
  const d = parseDay_(s);
  const wd = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getUTCDay()];
  const mo = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()];
  return (withDay ? wd + ' ' : '') + `${d.getUTCDate()} ${mo} ${d.getUTCFullYear()}`;
}

/** Brand fonts (Poppins + Liberation Serif), served from the website and embedded in the PDF. Arial if unreachable. */
const FONT_BASE = 'https://xowestx.github.io/owest-site-camera/fonts/';
function fontsCss_() {
  const list = [['Poppins', 400, 'Poppins-Regular'], ['Poppins', 500, 'Poppins-Medium'], ['Poppins', 700, 'Poppins-Bold'], ['Liberation Serif', 700, 'LiberationSerif-Bold']];
  const cache = CacheService.getScriptCache();
  return list.map(([fam, w, file]) => {
    let b64 = cache.get('font:' + file);
    if (!b64) {
      try {
        const r = UrlFetchApp.fetch(FONT_BASE + file + '.ttf', { muteHttpExceptions: true });
        if (r.getResponseCode() !== 200) return '';
        b64 = Utilities.base64Encode(r.getBlob().getBytes());
        try { cache.put('font:' + file, b64, 21600); } catch (e) {}
      } catch (e) { return ''; }
    }
    return `@font-face { font-family: '${fam}'; font-weight: ${w}; font-style: normal; src: url(data:font/ttf;base64,${b64}) format('truetype'); }`;
  }).join('\n');
}

const LOGO_OWEST = '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';
const LOGO_OPLUS = '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';

/** n items spread evenly through a list (keeps first and last). */
function spread_(list, n) {
  if (list.length <= n) return list.slice();
  const out = [];
  for (let i = 0; i < n; i++) out.push(list[Math.round(i * (list.length - 1) / (n - 1))]);
  return out;
}

/* =========================== AI writer (Gemini) =========================== */
/* Turns the engineers' quick notes into short professional client text, in Arabic and English.
   Needs a free Gemini API key (Script Property GEMINI_KEY). Without a key, or if the AI fails,
   the weekly PDF simply has photos only. */

let AI_DOWN = false;   // after one failure in a run, skip AI for the rest of that run (photos only)
function polishWeek_(unit, items, days) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key || AI_DOWN) return null;
  const notes = days.filter(d => d.notes.length);
  if (!notes.length) return null;
  const input = JSON.stringify({ unit: unit, scope: items, days: notes });
  const cache = CacheService.getScriptCache();
  const ck = 'ai:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input, Utilities.Charset.UTF_8));
  const hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const prompt = [
    'You write the weekly site progress report that a real-estate developer (Orascom, O West project) sends to a villa owner.',
    'Below are the site engineer\'s quick daily notes for one unit (often informal Egyptian Arabic, with typos).',
    'Rewrite them as clear, professional, polite client-facing text in English.',
    'Rules:',
    '- Use ONLY facts that are in the notes. Never add work, dates, percentages or promises that are not written.',
    '- Do not mention worker counts, contractor or subcontractor names, people\'s names, problems, delays or internal issues.',
    '- Leave out anything negative: "no work today", work stopped, waiting, no workers, missing or late materials, complaints. Do not soften it, just drop it.',
    '- If a day has only negative notes (for example "مفيش شغل النهارده"), leave that date out of "days" completely.',
    '- If the whole week has nothing positive to report, return empty lists.',
    '- Natural, confident professional English, like a project manager writing to a valued client. Past tense for completed work, present continuous for ongoing work.',
    '- summary: 2 to 5 short bullet points for the whole week, most important progress first.',
    '- days: one short sentence per date that has notes.',
    'Return JSON only, exactly: {"summary":{"en":["..."]},"days":[{"date":"YYYY-MM-DD","en":"..."}]}',
    '',
    'Data: ' + input
  ].join('\n');

  const S = settings_();
  const models = [S['AI model'] || 'gemini-3.8-flash', 'gemini-3.8-flash', 'gemini-flash-latest'].filter((m, i, a) => m && a.indexOf(m) === i);
  let lastErr = '';
  for (const model of models) {
    try {
      const call = () => UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, responseMimeType: 'application/json' } })
      });
      let res = call(), code = res.getResponseCode();
      if (code === 429 || code === 503) { Utilities.sleep(4000); res = call(); code = res.getResponseCode(); }   // busy: one quick retry
      if (code !== 200) { lastErr = model + ' ' + code + ': ' + res.getContentText().replace(/\s+/g, ' ').slice(0, 250); console.error('AI ' + lastErr); continue; }   // try the next model
      const j = JSON.parse(res.getContentText());
      const txt = (((j.candidates || [])[0] || {}).content || {}).parts;
      const t = (txt || []).map(p => p.text || '').join('').replace(/^```(json)?|```$/g, '').trim();
      const o = JSON.parse(t);
      const clean = a => (Array.isArray(a) ? a : []).map(x => str_(x)).filter(Boolean).slice(0, 6);
      const out = { ar: clean(o.summary && o.summary.ar), en: clean(o.summary && o.summary.en), days: {} };
      (o.days || []).forEach(d => { const k = str_(d.date); if (k) out.days[k] = { ar: str_(d.ar), en: str_(d.en) }; });
      try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {}
      aiStatus_('OK · ' + model);
      return out;
    } catch (e) { lastErr = model + ': ' + msg_(e); console.error('AI: ' + lastErr); }
  }
  AI_DOWN = true;
  aiStatus_('Error · ' + lastErr);
  return null;
}

function aiStatus_(text) { try { setSetting_('AI status', fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ' · ' + text); } catch (e) {} }

/** Daily (internal): tidy each visit's note into one short line + 2-4 bullets, same language, nothing hidden. */
function polishDay_(list) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  const notes = list.filter(v => v.workDone).map(v => ({ id: v.id, unit: v.unit, note: v.workDone }));
  if (!key || AI_DOWN || !notes.length) return {};
  const input = JSON.stringify(notes);
  const cache = CacheService.getScriptCache();
  const ck = 'aid:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input, Utilities.Charset.UTF_8));
  const hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const prompt = [
    'These are site engineers\' quick notes for construction site visits (Egyptian Arabic, typos, informal).',
    'For each visit, tidy the note for an INTERNAL daily report read by managers:',
    '- Keep every fact, including problems, stoppages and missing items. Do not add anything that is not written.',
    '- Write clear, short Arabic (simple Modern Standard Arabic), keep technical English words as they are (SOG, sandwich panel...).',
    '- "line": one short line summarising the visit (max 12 words).',
    '- "bullets": 1 to 4 short bullet points.',
    'Return JSON only: {"visits":[{"id":"...","line":"...","bullets":["..."]}]}',
    '', 'Data: ' + input
  ].join('\n');
  const o = aiJson_(key, prompt);
  if (!o) return {};
  const out = {};
  (o.visits || []).forEach(x => { if (x && x.id) out[str_(x.id)] = { line: str_(x.line), bullets: (Array.isArray(x.bullets) ? x.bullets : []).map(str_).filter(Boolean).slice(0, 4) }; });
  try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}

/** Calls Gemini with a prompt that must answer JSON; tries the models in order. null on failure. */
function aiJson_(key, prompt) {
  const S = settings_();
  const models = [S['AI model'] || 'gemini-3.8-flash', 'gemini-3.8-flash', 'gemini-flash-latest'].filter((m, i, a) => m && a.indexOf(m) === i);
  let lastErr = '';
  for (const model of models) {
    try {
      const call = () => UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, responseMimeType: 'application/json' } })
      });
      let res = call(), code = res.getResponseCode();
      if (code === 429 || code === 503) { Utilities.sleep(4000); res = call(); code = res.getResponseCode(); }
      if (code !== 200) { lastErr = model + ' ' + code + ': ' + res.getContentText().replace(/\s+/g, ' ').slice(0, 250); continue; }
      const j = JSON.parse(res.getContentText());
      const t = ((((j.candidates || [])[0] || {}).content || {}).parts || []).map(p => p.text || '').join('').replace(/^```(json)?|```$/g, '').trim();
      const o = JSON.parse(t);
      aiStatus_('OK · ' + model);
      return o;
    } catch (e) { lastErr = model + ': ' + msg_(e); }
  }
  AI_DOWN = true;
  aiStatus_('Error · ' + lastErr);
  return null;
}

function menuSetAiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Connect AI writer', 'Paste your Gemini API key (from aistudio.google.com > Get API key).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim();
  if (k.length < 20) return ui.alert('That does not look like an API key.');
  PropertiesService.getScriptProperties().setProperty('GEMINI_KEY', k);
  const t = polishWeek_('TEST', ['Landscape'], [{ date: '2026-01-01', notes: ['تم الردم بالكامل'] }]);
  ui.alert(t ? 'AI writer connected. Test: ' + (t.en[0] || t.ar[0] || 'OK') : 'Key saved, but the test failed. Check the key.');
}

/* =========================== shared report pieces =========================== */
/* Layout rules: tables only (the PDF converter is old-school), photos never cropped or stretched:
   each photo is scaled to fit inside a fixed box using its real width/height. */

const PH_BOX = { w: 330, h: 248 };   // one photo cell, 2 per row

function head_(title, titleEn, line1, line2) {
  return `<html dir="rtl" lang="ar"><head><meta charset="utf-8"><style>
    @page { size: A4; margin: 12mm 11mm; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    body { font-family: Arial, sans-serif; font-size: 10pt; color: #1d2433; direction: rtl; }
    table { border-collapse: collapse; width: 100%; }
    .top td { vertical-align: bottom; padding: 0 0 8px; }
    .brand { font-size: 9pt; letter-spacing: 3px; color: #8a6a4f; font-weight: bold; }
    .ttl { font-size: 20pt; font-weight: bold; color: #0d1b3e; }
    .ttl .en { font-size: 11pt; color: #6b7385; font-weight: normal; }
    .l1 { font-size: 11pt; font-weight: bold; color: #0d1b3e; text-align: left; }
    .l2 { font-size: 9.5pt; color: #6b7385; text-align: left; }
    .rule { height: 4px; background: #0d1b3e; border-bottom: 2px solid #e9b88f; margin-bottom: 12px; }
    .en { font-size: 8pt; color: #8a92a3; font-weight: normal; }
    .kpi { margin: 4px 0 14px; }
    .kpi td { width: 25%; text-align: center; border: 1px solid #e3e6ee; padding: 9px 4px; background: #f7f8fb; }
    .kpi .n { font-size: 18pt; font-weight: bold; color: #0d1b3e; }
    .kpi .k { font-size: 9pt; color: #4d5568; }
    h2 { font-size: 12pt; color: #0d1b3e; margin: 14px 0 6px; }
    .sum th { background: #eef0f5; color: #0d1b3e; text-align: right; padding: 6px; font-size: 9pt; border: 1px solid #e3e6ee; }
    .sum td { padding: 6px; border: 1px solid #e3e6ee; vertical-align: top; font-size: 9.5pt; }
    .sum .tot td { background: #f7f8fb; font-weight: bold; }
    tr { page-break-inside: avoid; }
    .c { text-align: center !important; }
    .b { font-weight: bold; }
    .big { font-size: 13pt; font-weight: bold; }
    .sm { font-size: 8.5pt; }
    .mono { font-family: "Courier New", monospace; direction: ltr; unicode-bidi: embed; }
    .foot { font-size: 8pt; color: #8a92a3; margin-top: 6px; }
    .page { page-break-before: always; }
    .uh td { vertical-align: middle; padding: 8px 10px; background: #f3f4f8; border-right: 6px solid #e9b88f; }
    .uh .ucode { font-size: 15pt; font-weight: bold; color: #0d1b3e; }
    .uh .sub { font-size: 9pt; color: #6b7385; margin-top: 2px; }
    .uh .meta { text-align: left; font-size: 9pt; color: #4d5568; }
    .meta-t { margin: 6px 0 10px; }
    .meta-t td { border: 1px solid #e3e6ee; padding: 5px 6px; text-align: center; font-size: 8.5pt; color: #6b7385; }
    .meta-t b { display: block; font-size: 10.5pt; color: #0d1b3e; }
    .ih { font-weight: bold; font-size: 10pt; color: #0d1b3e; margin: 10px 0 4px; border-bottom: 2px solid #e9b88f; padding-bottom: 2px; }
    .ih span { color: #8a92a3; font-weight: normal; }
    .grid td { width: 50%; padding: 4px; vertical-align: top; text-align: center; }
    .box { height: ${PH_BOX.h}px; background: #f3f4f8; border: 1px solid #e3e6ee; text-align: center; }
    .cap { font-size: 8.5pt; color: #4d5568; margin-top: 3px; text-align: center; }
    .cap .t { color: #8a92a3; }
    .two { margin-top: 10px; }
    .two td { vertical-align: top; border: 1px solid #e3e6ee; padding: 9px; }
    .lbl { font-weight: bold; color: #0d1b3e; margin-bottom: 6px; }
    .crew th { font-size: 8.5pt; color: #6b7385; text-align: right; border-bottom: 1px solid #e3e6ee; padding: 3px; }
    .crew td { font-size: 9.5pt; padding: 3px; border-bottom: 1px solid #f0f1f5; }
    .items { margin: 0 0 6px; }
    .chip { display: inline-block; background: #eef0f5; border: 1px solid #e3e6ee; padding: 2px 8px; font-size: 8.5pt; margin: 2px; }
  </style></head><body>
  <table class="top"><tr>
    <td><div class="brand">O WEST</div><div class="ttl">${title} <span class="en">${titleEn}</span></div></td>
    <td><div class="l1">${line1}</div><div class="l2">${line2}</div></td>
  </tr></table><div class="rule"></div>`;
}

function kpis_(list) {
  return '<table class="kpi"><tr>' + list.map(k => `<td><div class="n">${k[0]}</div><div class="k">${k[1]} <span class="en">${k[2]}</span></div></td>`).join('') + '</tr></table>';
}

function unitHead_(n, v, eng) {
  const items = v.items ? v.items.split(',').map(s => s.trim()).filter(Boolean) : [];
  const dur = minutes_(v.started, v.closed);
  return `<table class="uh"><tr><td><div class="ucode mono">${n} · ${esc_(v.unit)}</div><div class="sub">${esc_(v.project)}${items.length ? ' · ' + items.map(esc_).join(' · ') : ''}</div></td></tr></table>
    <table class="meta-t"><tr>
      <td>المهندس<b>${esc_(eng || v.by)}</b></td>
      <td>بداية<b class="mono">${hm_(v.started) || '—'}</b></td>
      <td>نهاية<b class="mono">${hm_(v.closed) || '—'}</b></td>
      <td>المدة<b>${dur != null ? dur + ' د' : '—'}</b></td>
      <td>صور<b>${v.photos}</b></td>
      <td>عمالة<b>${v.manpower || 0}</b></td>
    </tr></table>`;
}

function workAndCrew_(v, c, cl) {
  let crewHtml = '<div class="sm">مفيش عمالة.</div>';
  if (c.length) {
    // one readable line per contractor: "Arch for Construction: كهربائي 1 + مساعد 3"
    const byC = groupBy_(c, m => m.contractor);
    crewHtml = '<ul style="margin:0;padding-right:16px">' + Object.keys(byC).map(k => {
      const parts = byC[k].map(m => `${m.trade ? esc_(tradeAr_(m.trade)) + ' ' : ''}${m.count}`).join(' + ');
      const n = byC[k].reduce((s, m) => s + m.count, 0);
      return `<li><b>${esc_(k)}</b>: ${parts}${byC[k].length > 1 ? ` <span class="sm">(= ${n})</span>` : ''}</li>`;
    }).join('') + `</ul><div class="sm" style="margin-top:4px">الإجمالي: <b>${c.reduce((s, m) => s + m.count, 0)}</b></div>`;
  }
  const status = cl && cl.bullets && cl.bullets.length
    ? '<ul style="margin:0;padding-right:16px">' + cl.bullets.map(b => `<li>${esc_(b)}</li>`).join('') + '</ul>'
    : br_(v.workDone);
  return `<table class="two"><tr>
    <td style="width:55%"><div class="lbl">الحالة <span class="en">Status</span></div>${status}</td>
    <td><div class="lbl">العمالة <span class="en">Manpower</span></div>${crewHtml}</td>
  </tr></table>`;
}

/** Photos grouped by item, 2 per row, whole photo always visible. max = 0 means all. */
function photoGrid_(list, max) {
  if (!list.length) return '<p class="foot">مفيش صور.</p>';
  const ph = max ? list.slice(0, max) : list;
  const imgs = images_(ph.map(p => p.fileId));
  const groups = groupBy_(ph, p => p.item || GENERAL);
  const names = Object.keys(groups);
  let h = '';
  names.forEach(name => {
    const g = groups[name];
    if (names.length > 1) h += `<div class="ih">${esc_(name)} <span>(${g.length})</span></div>`;
    h += '<table class="grid">';
    for (let i = 0; i < g.length; i += 2) {
      h += '<tr>';
      for (let j = i; j < i + 2; j++) {
        const p = g[j];
        if (!p) { h += '<td></td>'; continue; }
        const im = imgs[p.fileId];
        let tag = '';
        if (im) {
          const s = im.w && im.h ? Math.min(PH_BOX.w / im.w, PH_BOX.h / im.h) : 0;
          const w = s ? Math.round(im.w * s) : PH_BOX.w, hh = s ? Math.round(im.h * s) : 0;
          const pad = hh ? Math.floor((PH_BOX.h - hh) / 2) : 0;
          tag = `<div style="height:${pad}px"></div><img src="${im.src}" width="${w}"${hh ? ` height="${hh}"` : ''}>`;
        }
        const cap = [p.note || (names.length > 1 ? '' : p.item), p.time ? `<span class="t">${p.time}</span>` : ''].filter(Boolean).join(' · ');
        h += `<td><div class="box">${tag}</div><div class="cap">${cap ? cap.replace(/^(?!<)/, '') : ''}</div></td>`;
      }
      h += '</tr>';
    }
    h += '</table>';
  });
  if (max && list.length > max) h += `<p class="foot">+ ${list.length - max} صور تانية في فولدر الوحدة على الدرايف.</p>`;
  return h;
}

/** Light JPEG previews (≈800px) with their real size, fetched in parallel. */
function images_(ids) {
  const out = {};
  if (!ids.length) return out;
  const token = ScriptApp.getOAuthToken();
  const reqs = ids.map(id => ({ url: 'https://drive.google.com/thumbnail?sz=w800&id=' + id, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }));
  let res = [];
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) {}
  ids.forEach((id, i) => {
    const r = res[i];
    let blob = null;
    if (r && r.getResponseCode() === 200 && /^image\//.test(String(r.getHeaders()['Content-Type'] || ''))) blob = r.getBlob();
    if (!blob) { try { blob = DriveApp.getFileById(id).getBlob(); } catch (e) {} }
    if (!blob) return;
    const bytes = blob.getBytes();
    const sz = jpegSize_(bytes) || {};
    out[id] = { src: 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(bytes), w: sz.w, h: sz.h };
  });
  return out;
}

/** Width/height from a JPEG header (Apps Script bytes are signed). */
function jpegSize_(b) {
  const u = i => b[i] & 0xFF;
  if (u(0) !== 0xFF || u(1) !== 0xD8) return null;
  let i = 2;
  while (i < b.length - 9) {
    if (u(i) !== 0xFF) { i++; continue; }
    const m = u(i + 1);
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: (u(i + 5) << 8) | u(i + 6), w: (u(i + 7) << 8) | u(i + 8) };
    i += 2 + ((u(i + 2) << 8) | u(i + 3));
  }
  return null;
}

function hm_(s) { return str_(s).slice(11, 16); }
function br_(s) { return esc_(s).replace(/\n/g, '<br>'); }
function timeRange_(v) { const a = hm_(v.started), b = hm_(v.closed); return a && b ? `${a}–${b}` : (b || a || '—'); }
function minutes_(a, b) {
  if (!a || !b) return null;
  const p = s => { const m = str_(s).match(/(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/); return m ? Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5]) : null; };
  const x = p(a), y = p(b);
  return x != null && y != null && y >= x ? Math.round((y - x) / 60000) : null;
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
    items: str_(r['Items']), by: str_(r['By']), started: dt_(r['Started']), closed: dt_(r['Closed']), photos: Number(r['Photos']) || 0,
    status: str_(r['Status']), workDone: str_(r['Work done']), manpower: Number(r['Manpower total']) || 0, notes: str_(r['Notes'])
  })).filter(v => v.id && v.status === 'Closed' && filter(v));
}

function photosByVisit_(visits) {
  const ids = {}; visits.forEach(v => ids[v.id] = 1);
  const out = {};
  rows_(T.photos).forEach(r => {
    if (!ids[r['Visit ID']] || r['Status'] !== 'OK') return;
    (out[r['Visit ID']] = out[r['Visit ID']] || []).push({ fileId: str_(r['File ID']), item: str_(r['Item']), note: str_(r['Note']), time: dt_(r['Uploaded at']).slice(11, 16) });
  });
  return out;
}

function crewByVisit_(visits) {
  const ids = {}; visits.forEach(v => ids[v.id] = 1);
  const out = {};
  rows_(T.manpower).forEach(r => {
    if (!ids[r['Visit ID']] || str_(r['Status']) === 'Deleted') return;
    (out[r['Visit ID']] = out[r['Visit ID']] || []).push({ contractor: str_(r['Contractor']), trade: str_(r['Trade']), count: Number(r['Count']) || 0, item: str_(r['Item']) });
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
    .addItem('Send weekly client reports now', 'menuSendWeekly')
    .addItem('Send weekly client reports to one email only', 'menuSendWeeklyTo')
    .addSeparator()
    .addItem('Delete selected photos', 'menuDeletePhotos')
    .addItem('Delete selected visits (photos + manpower)', 'menuDeleteVisits')
    .addSeparator()
    .addItem('Set manager password', 'menuSetPassword')
    .addItem('Connect ClickUp (API token)', 'menuSetClickUpToken')
    .addItem('Connect AI writer (Gemini key)', 'menuSetAiKey')
    .addItem('Sync from ClickUp now', 'menuSync')
    .addItem('Apply schedule from Settings', 'installTriggers')
    .addItem('Refresh lists in the app', 'menuRefresh')
    .addToUi();
}

function menuDailyToday() { reportAlert_(makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), false), 'visits'); }
function menuDailySend() { reportAlert_(makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), true), 'visits'); }
function menuWeekly() {
  const r = weeklyStart_(fmt_(new Date(), 'yyyy-MM-dd'), false);
  const ui = SpreadsheetApp.getUi();
  if (!r.units && !r.pending) return ui.alert('No closed visits in the last 7 days.');
  ui.alert(r.pending ? `Started: ${r.units} unit PDFs made so far. The rest are made in the background in the next few minutes (Drive > ${REPORTS_ROOT} > Weekly).` : `Done. ${r.units} unit PDFs in Drive > ${REPORTS_ROOT} > Weekly.`);
}
function menuSendWeeklyTo() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt('Send weekly reports to one email', 'Makes this week\'s unit reports and emails them ONLY to this address (no cc, no clients):', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const to = str_(a.getResponseText());
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) return ui.alert('That is not an email address.');
  const r = weeklyStart_(fmt_(new Date(), 'yyyy-MM-dd'), true, to);
  if (!r.units && !r.pending) return ui.alert('No closed visits in the last 7 days.');
  ui.alert(r.pending ? `Started: ${r.units} unit PDFs made so far. The rest are made in the background, then the email goes to ${to} (5–15 minutes).` : `Done. ${r.units} unit reports emailed to ${to}.`);
}
function menuSendWeekly() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Make this week\'s unit reports and email them now (Weekly report to / cc in Settings)?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const r = weeklyStart_(fmt_(new Date(), 'yyyy-MM-dd'), true);
  if (!r.units && !r.pending) return ui.alert('No closed visits in the last 7 days.');
  ui.alert(r.pending ? `Started: ${r.units} unit PDFs made so far. The rest are made in the background, then the email goes out (5–15 minutes).` : `Done. ${r.units} unit reports emailed.`);
}
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
    email: 'ecdc4bdb-8444-4074-b32c-cd4c36274f25',
    mobile: '1707ad3e-025b-457f-a4f3-cf1d7b6b1a7e'
  },
  // weekly report button lives only on the private Construction Team list
  team: { list: '901525752325', report: 'dfe137c2-147d-4810-8af3-439514d62fe4', whatsapp: '4180e90a-f5f6-4a54-b010-1aceb35fbd02' },
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
  if (r.pending) return SpreadsheetApp.getUi().alert(`Reading ClickUp tasks one by one: ${r.pending} left. Run "Sync from ClickUp now" again in a minute until it says done.`);
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
    // 1) bring the "ClickUp Tasks" tab up to date (the full list answer is over 50 MB, so we never ask for it):
    //    a) tasks changed since the last sync (small answer), b) a few stale rows re-read one by one
    const all = cuSyncTasks_(token);
    const pending = all.filter(t => !t.status).length;
    if (pending) {   // first fill of the task store: keep the current lists until every task has been read once
      setSetting_('Last ClickUp sync', fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ` · reading tasks, ${pending} left (runs again in 30 min, or run Sync again)`);
      return { units: 0, items: 0, contractors: 0, newContractors: 0, pending: pending };
    }
    const tasks = all.filter(t => low_(t.status) === CU.status);

    // 2) turn tasks into units / items / contractors
    const alias = aliases_();
    const knownC = {};
    rows_(T.contractors).forEach(r => knownC[low_(r['Contractor'])] = 1);
    const units = {}, scope = {}, uc = {}, freshC = [];
    tasks.forEach(t => {
      const unit = normUnit_(t.unit);
      if (!unit) return;
      const k = unit.toLowerCase();
      const u = units[k] = units[k] || { unit: unit, client: '', email: '' };
      u.client = u.client || t.client;
      u.email = u.email || t.email;

      const type = t.type;
      if (type && CU.skipTypes.indexOf(type.toLowerCase()) < 0) {
        let item = CU.typeNames[type.toLowerCase().trim()] || type.trim();
        if (item.toLowerCase() === 'others') {
          item = str_(t.desc).split('\n')[0].trim().slice(0, 40) || 'Other works';
        }
        (scope[k] = scope[k] || []).indexOf(item) < 0 && scope[k].push(item);
      }
      str_(t.contractor).split(/[,\/]/).forEach(c => {
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

/* ---------- ClickUp task store ("ClickUp Tasks" tab) ---------- */
const CU_TAB = 'ClickUp Tasks';
const CU_COLS = ['Unit', 'Task ID', 'Status', 'Client', 'Email', 'Mobile', 'Type', 'Description', 'Contractor', 'Updated', 'Checked'];

/** Compact row from a full ClickUp task object. */
function cuCompact_(t) {
  const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
  const v = id => str_(f[id] && f[id].value);
  return {
    unit: normUnit_(t.name) || str_(t.name), id: str_(t.id), status: str_(t.status && t.status.status),
    client: v(CU.f.client), email: v(CU.f.email), mobile: v(CU.f.mobile), type: dropName_(f[CU.f.requestType]),
    desc: v(CU.f.description).slice(0, 200), contractor: v(CU.f.contractor), updated: str_(t.date_updated), checked: String(Date.now())
  };
}

/** Updates the tab and returns every stored task (all statuses). Safe to run often. */
function cuSyncTasks_(token) {
  const sh = sheet_(CU_TAB, CU_COLS);
  if (sh.getRange(1, 1, 1, CU_COLS.length).getValues()[0].join('|') !== CU_COLS.join('|')) sh.getRange(1, 1, 1, CU_COLS.length).setValues([CU_COLS]).setFontWeight('bold');
  const byId = {};
  rows_(CU_TAB).forEach(r => {
    const id = str_(r['Task ID']); if (!id) return;
    byId[id] = { unit: str_(r['Unit']), id: id, status: str_(r['Status']), client: str_(r['Client']), email: str_(r['Email']), mobile: str_(r['Mobile']),
      type: str_(r['Type']), desc: str_(r['Description']), contractor: str_(r['Contractor']), updated: str_(r['Updated']), checked: str_(r['Checked']) };
  });
  const H = { Authorization: token };
  const props = PropertiesService.getScriptProperties();
  const started = Date.now();
  // a) changed since last sync (any status, so tasks that left "site execution" are caught too)
  const since = Number(props.getProperty('CU_LAST')) || (started - 3 * 86400000);
  let incOk = true;
  for (let page = 0; page < 20; page++) {
    const res = UrlFetchApp.fetch(`https://api.clickup.com/api/v2/list/${CU.list}/task?include_closed=true&subtasks=true&date_updated_gt=${since - 60000}&page=${page}`, { headers: H, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { incOk = false; console.error('clickup changes: ' + res.getResponseCode()); break; }
    let j; try { j = JSON.parse(res.getContentText()); } catch (e) { incOk = false; console.error('clickup changes: answer too big'); break; }
    (j.tasks || []).forEach(t => { byId[t.id] = cuCompact_(t); });
    if (j.last_page || !(j.tasks || []).length) break;
  }
  // b) re-read rows never read or not read for 24 h, one task per call, for at most ~2.5 minutes
  const stale = Object.keys(byId).filter(id => !byId[id].checked || started - Number(byId[id].checked) > 86400000)
    .sort((a, b) => Number(byId[a].checked || 0) - Number(byId[b].checked || 0));
  for (let i = 0; i < stale.length && Date.now() - started < 150000; i += 25) {
    const ids = stale.slice(i, i + 25);
    const res = UrlFetchApp.fetchAll(ids.map(id => ({ url: `https://api.clickup.com/api/v2/task/${id}`, headers: H, muteHttpExceptions: true })));
    let limited = false;
    res.forEach((x, n) => {
      const code = x.getResponseCode();
      if (code === 200) { try { byId[ids[n]] = cuCompact_(JSON.parse(x.getContentText())); } catch (e) {} }
      else if (code === 404) { byId[ids[n]].status = 'deleted'; byId[ids[n]].checked = String(Date.now()); }
      else if (code === 429) limited = true;
    });
    if (limited) Utilities.sleep(20000);
  }
  if (incOk) props.setProperty('CU_LAST', String(started));
  const list = Object.keys(byId).map(id => byId[id]).sort((a, b) => (a.unit + a.id).localeCompare(b.unit + b.id));
  rewrite_(sh, CU_COLS, list.map(t => [t.unit, t.id, t.status, t.client, t.email, t.mobile, t.type, t.desc, t.contractor, t.updated, t.checked]));
  CU_UNITS = null;
  return list;
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
