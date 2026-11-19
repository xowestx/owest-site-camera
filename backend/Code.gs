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
    if (b.action === 'pinLogin') return json_(Object.assign({ ok: true }, pinLogin_(b)));
    if (b.action === 'team') return json_({ ok: true, people: teamPeople_().filter(p => p.active).map(p => p.name) });
    if (/^(home|unitInfo|visitPhotos|mgr|comments|comment|reply|commentClose|commentReopen|changePin)$/.test(b.action)) {
      const me = requirePerson_(b.token);
      const fn = { home: home_, unitInfo: unitInfo_, visitPhotos: visitPhotos_, mgr: mgrHome_, comments: commentsList_,
        comment: addComment_, reply: replyComment_, commentClose: closeComment_, commentReopen: reopenComment_, changePin: changePin_ }[b.action];
      return json_(Object.assign({ ok: true }, fn(b, me)));
    }
    if (b.action === 'dashLogin') return json_(Object.assign({ ok: true }, dashLogin_(b)));
    if (b.action === 'dash') { requireDash_(b.token); return json_(Object.assign({ ok: true }, dash_(b))); }
    if (b.action === 'dashPhotos') { requireDash_(b.token); return json_(Object.assign({ ok: true }, dashPhotos_(b))); }
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

/* =========================== office dashboard (read only) =========================== */
/* dashboard.html on the website. Its own password (sheet menu: Set dashboard password); the manager code works too. */

function dashLogin_(b) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('DASH_HASH');
  if (!hash) throw new Error('No dashboard password yet. Sheet menu: O West Camera > Set dashboard password.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('dashFails') || 0);
  if (fails >= 8) throw new Error('Too many wrong tries. Wait 10 minutes.');
  const pw = str_(b.password);
  const okDash = sha_(pw + props.getProperty('DASH_SALT')) === hash;
  const okAdmin = props.getProperty('ADMIN_HASH') && sha_(pw + props.getProperty('ADMIN_SALT')) === props.getProperty('ADMIN_HASH');
  if (!okDash && !okAdmin) { cache.put('dashFails', String(fails + 1), 600); throw new Error('Wrong password.'); }
  cache.remove('dashFails');
  const exp = Date.now() + TOKEN_DAYS * 86400000;
  return { token: 'd' + exp + '.' + dashSign_('d' + exp), expires: exp };
}

function dashSign_(v) {
  const key = PropertiesService.getScriptProperties().getProperty('DASH_SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(v, key)).replace(/=+$/, '');
}

function requireDash_(token) {
  const t = str_(token);
  if (t.charAt(0) === 'd') {
    const i = t.indexOf('.'), exp = Number(t.slice(1, i));
    if (i > 0 && exp > Date.now() && dashSign_(t.slice(0, i)) === t.slice(i + 1)) return;
    const e = new Error('Please log in again.'); e.code = 'auth'; throw e;
  }
  requireAdmin_(t);
}

function menuSetDashPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Dashboard password', 'Password for the office dashboard (at least 6 characters).\nPhones logged in with the old one will need the new one.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText().trim();
  if (pw.length < 6) return ui.alert('Too short. Use at least 6 characters.');
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ DASH_SALT: salt, DASH_HASH: sha_(pw + salt), DASH_SECRET: Utilities.getUuid() + Utilities.getUuid() });
  ui.alert('Dashboard password saved.\nLink: https://xowestx.github.io/owest-site-camera/dashboard.html');
}

/** Everything the dashboard shows for a date range, in one small answer. */
function dash_(b) {
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const from = /^\d{4}-\d{2}-\d{2}$/.test(str_(b.from)) ? str_(b.from) : today;
  const to = /^\d{4}-\d{2}-\d{2}$/.test(str_(b.to)) ? str_(b.to) : from;
  const units = rows_(T.units).filter(r => str_(r['Unit']) && low_(r['Active']) !== 'no')
    .map(r => [str_(r['Unit']), str_(r['Client name']), str_(r['Project'])]);
  const all = closedVisits_(() => true);
  const last = {};
  all.forEach(v => { const k = v.unit; const t = v.closed || v.date; if (!last[k] || t > last[k][0]) last[k] = [t, v.by]; });
  const inRange = all.filter(v => v.date >= from && v.date <= to);
  const crew = crewByVisit_(inRange);
  const visits = inRange.map(v => ({
    id: v.id, d: v.date, u: v.unit, by: v.by, t1: (v.started || '').slice(11, 16), t2: (v.closed || '').slice(11, 16),
    items: v.items, work: v.workDone, mp: v.manpower, ph: v.photos,
    crew: (crew[v.id] || []).map(m => [m.contractor, m.trade, m.count, m.item])
  }));
  const trades = {};
  rows_(T.trades).forEach(r => { const en = str_(r['Trade']); if (en) trades[en] = str_(r['Arabic']); });
  return { now: fmt_(new Date(), 'yyyy-MM-dd HH:mm'), from: from, to: to, units: units, visits: visits, last: last, trades: trades };
}

/** Small thumbnails of one visit's photos (Drive files are private, so the server sends them). */
function dashPhotos_(b) {
  const id = str_(b.visitId);
  const ph = (photosByVisit_([{ id: id }])[id] || []).slice(0, 24);
  if (!ph.length) return { photos: [] };
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetchAll(ph.map(p => ({ url: 'https://drive.google.com/thumbnail?sz=w480&id=' + p.fileId, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true })));
  return {
    photos: ph.map((p, i) => {
      const r = res[i];
      const ok = r && r.getResponseCode() === 200 && /^image\//.test(String(r.getHeaders()['Content-Type'] || ''));
      return { src: ok ? 'data:image/jpeg;base64,' + Utilities.base64Encode(r.getBlob().getBytes()) : '', item: p.item, time: p.time, note: p.note,
        full: 'https://drive.google.com/file/d/' + p.fileId + '/view' };
    }).filter(p => p.src)
  };
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
function dailyJob() {
  // fresh ClickUp first, so the engineer split in the email is the latest one (sync normally takes a few seconds)
  try { syncClickUp_(); } catch (e) { console.error('sync before daily: ' + e); }
  makeDaily_(fmt_(new Date(), 'yyyy-MM-dd'), true);
}

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
        cc: S['Daily report cc'] || '',
        subject: `${S['Company name'] || 'O-Plus'} · Daily site report · ${day} · ${visits.length} units · ${total} workers`,
        htmlBody: dailyEmailBody_(day, visits, crew, made) + workloadHtml_(day),
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

/* ---------- per engineer: his units (ClickUp "Task Owner") vs what was photographed ---------- */

/** Same person written two ways: ClickUp "Seif" / site "SaifAllah Nawar", "Alaa" / "Alaa Negm", "Morcous" / "Morcos". */
const ENG_ALIAS = { seif: 'saifallah', saif: 'saifallah', morcous: 'morcos', morkos: 'morcos' };
function engKey_(name) { const w = low_(name).split(/\s+/)[0] || ''; return ENG_ALIAS[w] || w; }

/** { owner: { units:[...], visitedToday:{unit:[by...]}, notToday:[...], never:[...], othersUnits:{unit:[by]} } } */
function workload_(day) {
  const own = {};   // display owner name -> Set(units)
  rows_('ClickUp Tasks').forEach(r => {
    const u = normUnit_(r['Unit']); if (!u || low_(r['Status']) !== CU.status) return;
    str_(r['Owner']).split(/\s*,\s*/).forEach(o => {
      if (!o || o === '-' || low_(o) === 'real estate') return;
      (own[o] = own[o] || new Set()).add(u);
    });
  });
  const all = closedVisits_(() => true);
  const ever = {}, today = {};
  all.forEach(v => {
    const u = normUnit_(v.unit) || v.unit;
    ever[u] = true;
    if (v.date === day) (today[u] = today[u] || []).indexOf(v.by) < 0 && today[u].push(v.by);
  });
  const ownerOf = {}; Object.keys(own).forEach(o => own[o].forEach(u => (ownerOf[u] = ownerOf[u] || []).push(engKey_(o))));
  const out = {};
  Object.keys(own).sort().forEach(o => {
    const units = [...own[o]].sort(), k = engKey_(o);
    const w = { units: units, visited: [], notToday: [], never: [], elsewhere: [] };
    units.forEach(u => {
      if (today[u]) w.visited.push({ unit: u, by: today[u].filter(b => engKey_(b) !== k) });
      else { w.notToday.push(u); if (!ever[u]) w.never.push(u); }
    });
    // units he photographed today that are not his
    Object.keys(today).forEach(u => { if (today[u].some(b => engKey_(b) === k) && units.indexOf(u) < 0) w.elsewhere.push({ unit: u, owner: (ownerOf[u] || []).length ? Object.keys(own).filter(x => own[x].has(u)).join(', ') : 'no owner' }); });
    out[o] = w;
  });
  return out;
}

function workloadHtml_(day) {
  let W; try { W = workload_(day); } catch (e) { console.error(e); return ''; }
  const names = Object.keys(W);
  if (!names.length) return '<div style="font-family:Arial,sans-serif;font-size:13px;color:#888"><p>Engineer split: not available yet (run "Sync from ClickUp now").</p></div>';
  const short = u => esc_(u.replace(/^OW\//, ''));
  const list = a => a.length ? a.map(short).join(' · ') : '—';
  const rows = names.map(o => {
    const w = W[o];
    const vis = w.visited.map(x => short(x.unit) + (x.by.length ? ` <span style="color:#b26a00">(by ${esc_(x.by.join(', '))})</span>` : '')).join(' · ') || '—';
    const other = w.elsewhere.length ? `<div style="margin-top:4px"><b>Also photographed (not his):</b> ${w.elsewhere.map(x => short(x.unit) + ` <span style="color:#888">(${esc_(x.owner)})</span>`).join(' · ')}</div>` : '';
    return `<tr><td style="padding:10px;border-top:1px solid #ddd;vertical-align:top">
      <div style="font-size:15px"><b>${esc_(o)}</b> — ${w.visited.length} of ${w.units.length} units photographed today</div>
      <div style="margin-top:4px;color:#1f7a4d"><b>✔ Photographed today:</b> ${vis}</div>
      <div style="margin-top:4px;color:#b3261e"><b>✘ Not photographed today (${w.notToday.length}):</b> ${list(w.notToday)}</div>
      <div style="margin-top:4px;color:#555"><b>Never photographed since start (${w.never.length}):</b> ${list(w.never)}</div>${other}</td></tr>`;
  }).join('');
  return `<div style="font-family:Arial,sans-serif;font-size:13px"><p style="font-size:14px"><b>Engineers · their units (ClickUp Task Owner) · ${day}</b><br>
    <span style="color:#888">"by …" = photographed by another engineer. "Never" counts from the first day of the site camera.</span></p>
    <table style="border-collapse:collapse;width:100%;max-width:900px">${rows}</table></div>`;
}

/* ---------- "Workload" tab: share of the site work per engineer, rebuilt on every ClickUp sync ---------- */
/* Share = (0.6 × engineer's Amount ÷ all engineers' Amount + 0.4 × engineer's days ÷ all engineers' days) × 100.
   Days = ClickUp Duration (start → due, both days counted; none = 0). Real Estate and tasks with no Task Owner are left out. */
/** Sheet menu: rebuild the Workload tab from the stored ClickUp tasks (no ClickUp call). Shows any error. */
function menuWorkload() {
  const tasks = rows_(CU_TAB).filter(r => low_(r['Status']) === CU.status).map(r => ({
    unit: normUnit_(r['Unit']) || str_(r['Unit']), type: str_(r['Type']), owner: str_(r['Owner']), pm: str_(r['PM']),
    amount: str_(r['Amount']), poc: str_(r['POC']), days: str_(r['Days']) }));
  writeWorkloadTab_(tasks);
  SpreadsheetApp.getActive().getSheetByName('Workload').activate();
}

function writeWorkloadTab_(tasks) {
  const day = fmt_(new Date(), 'yyyy-MM-dd HH:mm');
  const band = s => s >= 30 ? 'VERY HIGH' : s >= 15 ? 'HIGH' : s >= 5 ? 'MEDIUM' : 'LOW';
  const num = x => { const n = Number(str_(x)); return str_(x) === '' || isNaN(n) ? null : n; };
  const notRead = tasks.filter(t => str_(t.days) === '').length;
  const by = {}, noOwner = [];
  let realEstate = 0;
  tasks.forEach(t => {
    const o = str_(t.owner);
    if (!o || o === '-') noOwner.push(t);
    else if (low_(o) === 'real estate') realEstate++;
    else (by[o] = by[o] || []).push(t);
  });
  const names = Object.keys(by);
  const V = {}, D = {};
  names.forEach(e => { V[e] = by[e].reduce((s, t) => s + (num(t.amount) || 0), 0); D[e] = by[e].reduce((s, t) => s + (num(t.days) || 0), 0); });
  const TV = names.reduce((s, e) => s + V[e], 0) || 1, TD = names.reduce((s, e) => s + D[e], 0) || 1;
  const S = {}; names.forEach(e => S[e] = (0.6 * V[e] / TV + 0.4 * D[e] / TD) * 100);
  const rank = names.slice().sort((a, b) => S[b] - S[a]);
  const all = [].concat(...names.map(e => by[e]));
  const W = 7, row = a => { const r = a.slice(0, W); while (r.length < W) r.push(''); return r; };
  const out = [], bold = [], heads = [], pct = [], money = [];
  const push = (a, kind) => { out.push(row(a)); const n = out.length; if (kind === 'b') bold.push(n); if (kind === 'h') heads.push(n); return n; };
  push(['Site engineers workload · updated ' + day + ' from ClickUp (Site Execution)'], 'b');
  if (notRead) push([`Still reading ${notRead} tasks from ClickUp: numbers complete after the next sync.`]);
  push([]);
  push(['Rank', 'Engineer', 'Units', 'Value (EGP)', 'Days', 'Workload share %', 'Density band'], 'h');
  rank.forEach((e, i) => { const n = push([i + 1, e, by[e].length, Math.round(V[e]), D[e], S[e] / 100, band(S[e])]); pct.push([n, 6]); money.push(n); });
  const tot = push(['', 'Total', all.length, Math.round(TV), TD, 1, ''], 'b'); pct.push([tot, 6]); money.push(tot);
  push([]);
  push(['Data gaps (engineers\' units)'], 'b');
  push(['Duration empty (0 days)', all.filter(t => !num(t.days)).length]);
  push(['Progress not entered', all.filter(t => num(t.poc) === null).length]);
  push(['PM empty', all.filter(t => !str_(t.pm)).length]);
  push(['Duration exactly 1 day', all.filter(t => num(t.days) === 1).length]);
  push(['Real Estate tasks (left out)', realEstate]);
  push(['Tasks with no Task Owner (left out, listed at the end)', noOwner.length]);
  push(['Share = (0.6 × value ÷ all value + 0.4 × days ÷ all days) × 100. Days = ClickUp Duration. Bands: ≥30 VERY HIGH · 15–30 HIGH · 5–15 MEDIUM · <5 LOW. Units = ClickUp tasks.']);
  const unitRows = list => list.slice().sort((a, b) => (a.unit + a.type).localeCompare(b.unit + b.type)).forEach(t => {
    const p = num(t.poc); const n = push([t.unit, t.type || '-', p === null ? 'not entered' : p / 100, str_(t.pm) || '-']); if (p !== null) pct.push([n, 3]);
  });
  rank.forEach((e, i) => {
    push([]); push([]);
    push([e], 'b');
    push(['Name', 'Units', 'Workload share %', 'Density band', 'Rank'], 'h');
    const n = push([e, by[e].length, S[e] / 100, band(S[e]), `${i + 1} of ${rank.length}`]); pct.push([n, 3]);
    push([]);
    push(['Unit ID', 'Request', 'Progress on site (%)', 'PM'], 'h');
    unitRows(by[e]);
  });
  push([]); push([]);
  push([`No Task Owner (${noOwner.length})`], 'b');
  push(['Unit ID', 'Request', 'Progress on site (%)', 'PM'], 'h');
  unitRows(noOwner);
  const sh = sheet_('Workload', row(['']));
  sh.clear();
  sh.getRange(1, 1, out.length, W).setValues(out);
  bold.forEach(n => sh.getRange(n, 1).setFontWeight('bold'));
  heads.forEach(n => sh.getRange(n, 1, 1, W).setFontWeight('bold').setBackground('#eef1f6'));
  pct.forEach(([n, c]) => sh.getRange(n, c).setNumberFormat('0.0%'));
  money.forEach(n => sh.getRange(n, 4, 1, 2).setNumberFormat('#,##0'));
  sh.getRange(1, 1).setFontSize(13);
  [12, 30, 18, 18, 10, 16, 14].forEach((w, i) => sh.setColumnWidth(i + 1, w * 8));
}

/* =========================== TEAM: engineers + managers (team.html) =========================== */
/* "Team" tab: Name (as on the site) | Role (Engineer / Manager) | ClickUp name (Task Owner) | Active | PIN.
   PINs are kept as hashes in Script Properties; the PIN column only says when one was set.
   ClickUp is the master: an engineer's units are the Site Execution tasks whose Task Owner is him. */
const TEAM_TAB = 'Team', TEAM_COLS = ['Name', 'Role', 'ClickUp name', 'Active', 'PIN'];
const COMMENTS_TAB = 'Comments';
const COMMENT_COLS = ['ID', 'Created', 'Unit', 'Visit ID', 'From', 'To', 'Text', 'Status', 'Parent', 'ClickUp', 'Updated', 'Closed by', 'Closed with visit'];

function teamPeople_() {
  const sh = sheet_(TEAM_TAB, TEAM_COLS);
  if (sh.getLastRow() < 2) {   // first use: seed from what we know
    sh.getRange(2, 1, 11, 4).setValues([
      ['Mohamed Emira', 'Manager', '', 'Yes'], ['Omar Sheir', 'Manager', '', 'Yes'], ['Morcos', 'Manager', 'Morcos', 'Yes'],
      ['Saad Ali', 'Manager', 'Saad Ali', 'Yes'], ['Admin', 'Manager', '', 'Yes'],
      ['Alaa Negm', 'Engineer', 'Alaa', 'Yes'], ['SaifAllah Nawar', 'Engineer', 'Seif', 'Yes'], ['David', 'Engineer', 'David', 'Yes'],
      ['Omar AbuGabal', 'Engineer', 'Omar AbuGabal', 'Yes'], ['Sakr', 'Engineer', 'Sakr', 'Yes'], ['Khaled Salah', 'Engineer', 'Khaled Salah', 'Yes']]);
  }
  return rows_(TEAM_TAB).filter(r => str_(r['Name'])).map(r => ({
    name: str_(r['Name']), manager: low_(r['Role']) === 'manager', cu: str_(r['ClickUp name']) || str_(r['Name']),
    active: low_(r['Active']) !== 'no' }));
}

/** Team login keys live apart from the manager password, so changing that password does not log the team out. */
function teamSecret_(k) {
  const props = PropertiesService.getScriptProperties();
  let v = props.getProperty(k);
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); props.setProperty(k, v); }
  return v;
}
function teamSign_(v) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(v, teamSecret_('TEAM_SECRET'))).replace(/=+$/, ''); }

function pinKey_(name) { return 'PIN_' + Utilities.base64EncodeWebSafe(low_(name)).replace(/=+$/, ''); }

function pinLogin_(b) {
  const name = str_(b.name), pin = str_(b.pin);
  const p = teamPeople_().find(x => x.active && low_(x.name) === low_(name));
  if (!p) throw new Error('الاسم ده مش في الفريق.');
  const cache = CacheService.getScriptCache(), ck = 'pinFails:' + low_(name);
  const fails = Number(cache.get(ck) || 0);
  if (fails >= 6) throw new Error('محاولات كتير غلط. استنى 10 دقايق.');
  const props = PropertiesService.getScriptProperties(), saved = props.getProperty(pinKey_(name));
  if (!saved) throw new Error('مفيش PIN متسجل للاسم ده. كلم الأدمن.');
  if (sha_(pin + '|' + low_(name) + '|' + teamSecret_('PIN_SALT')) !== saved) {
    cache.put(ck, String(fails + 1), 600); throw new Error('الـ PIN غلط.');
  }
  cache.remove(ck);
  const exp = Date.now() + 60 * 86400000, nm = Utilities.base64EncodeWebSafe(p.name).replace(/=+$/, '');
  const v = 'u' + exp + '.' + nm;
  return { token: v + '.' + teamSign_(v), name: p.name, manager: p.manager, expires: exp };
}

/** Who is calling. A manager token from the old "Extra reports" code also works, as Admin. */
function requirePerson_(token) {
  const t = str_(token), bad = () => { const e = new Error('ادخل تاني بالـ PIN.'); e.code = 'auth'; return e; };
  if (t.charAt(0) !== 'u') { requireAdmin_(t); return { name: 'Admin', manager: true, cu: 'Admin' }; }
  const parts = t.split('.');
  if (parts.length !== 3 || teamSign_(parts[0] + '.' + parts[1]) !== parts[2] || Number(parts[0].slice(1)) < Date.now()) throw bad();
  const name = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
  const p = teamPeople_().find(x => x.active && x.name === name);
  if (!p) throw bad();
  return p;
}

function menuSetPin() {
  const ui = SpreadsheetApp.getUi();
  const people = teamPeople_();
  const a = ui.prompt('Set a PIN', 'Name exactly as in the Team tab:\n' + people.map(p => p.name).join(', '), ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const p = people.find(x => low_(x.name) === low_(a.getResponseText()));
  if (!p) return ui.alert('Not in the Team tab.');
  const b = ui.prompt('PIN for ' + p.name, '4 to 8 digits.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const pin = b.getResponseText().trim();
  if (!/^\d{4,8}$/.test(pin)) return ui.alert('Use 4 to 8 digits.');
  savePin_(p.name, pin);
  ui.alert('PIN saved for ' + p.name + '.');
}

/** Gives a random 4-digit PIN to everyone who has none, and shows the list once. */
function menuMakePins() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const made = [];
  teamPeople_().filter(p => p.active && !props.getProperty(pinKey_(p.name))).forEach(p => {
    const pin = String(1000 + Math.floor(Math.random() * 9000)); savePin_(p.name, pin); made.push(p.name + ': ' + pin);
  });
  ui.alert(made.length ? 'New PINs (copy them now, they are not shown again):\n\n' + made.join('\n') : 'Everyone already has a PIN. Use "Set a PIN" to change one.');
}

/** Anyone can change his own PIN from team.html: old PIN, then the new one. */
function changePin_(b, me) {
  const props = PropertiesService.getScriptProperties();
  const hash = pin => sha_(pin + '|' + low_(me.name) + '|' + teamSecret_('PIN_SALT'));
  if (props.getProperty(pinKey_(me.name)) !== hash(str_(b.old))) throw new Error('الـ PIN القديم غلط.');
  const pin = str_(b.pin);
  if (!/^\d{4,8}$/.test(pin)) throw new Error('الـ PIN الجديد لازم يكون من 4 لـ 8 أرقام.');
  savePin_(me.name, pin);
  return { done: true };
}

function savePin_(name, pin) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(pinKey_(name), sha_(pin + '|' + low_(name) + '|' + teamSecret_('PIN_SALT')));
  const sh = sheet_(TEAM_TAB, TEAM_COLS), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (low_(v[i][0]) === low_(name)) sh.getRange(i + 1, 5).setValue('set ' + fmt_(new Date(), 'yyyy-MM-dd'));
}

/* ----- shared data ----- */

/** Site Execution units from ClickUp: unit -> {client, owners[], pms[], reqs[], poc[], mobile} */
function cuUnits_() {
  const out = {};
  rows_(CU_TAB).forEach(r => {
    if (low_(r['Status']) !== CU.status) return;
    const u = normUnit_(r['Unit']) || str_(r['Unit']); if (!u) return;
    const x = out[u] = out[u] || { unit: u, client: '', owners: [], pms: [], reqs: [], poc: [] };
    const add = (a, v) => { v = str_(v); if (v && v !== '-' && a.indexOf(v) < 0) a.push(v); };
    if (!x.client) x.client = str_(r['Client']);
    add(x.owners, r['Owner']); add(x.pms, r['PM']); add(x.reqs, r['Type']);
    if (str_(r['POC']) !== '') x.poc.push(Number(r['POC']));
  });
  Object.keys(out).forEach(u => { const p = out[u].poc; out[u].poc = p.length ? Math.round(p.reduce((a, b) => a + b, 0) / p.length) : null; });
  return out;
}

const isMine_ = (me, owners) => owners.some(o => low_(o) === low_(me.cu) || engKey_(o) === engKey_(me.name));

function visitsByUnit_() {
  const by = {};
  closedVisits_(() => true).forEach(v => { const u = normUnit_(v.unit) || v.unit; (by[u] = by[u] || []).push(v); });
  Object.keys(by).forEach(u => by[u].sort((a, b) => (b.closed || b.date).localeCompare(a.closed || a.date)));
  return by;
}

function commentRows_() {
  sheet_(COMMENTS_TAB, COMMENT_COLS);
  return rows_(COMMENTS_TAB).filter(r => str_(r['ID'])).map(r => ({
    id: str_(r['ID']), at: dt_(r['Created']), unit: str_(r['Unit']), visit: str_(r['Visit ID']), from: str_(r['From']), to: str_(r['To']),
    text: str_(r['Text']), status: str_(r['Status']), parent: str_(r['Parent']), updated: dt_(r['Updated']),
    closedBy: str_(r['Closed by']), closedVisit: str_(r['Closed with visit']) }));
}

/** Top-level comments with their replies, newest first. */
function threads_(filter) {
  const all = commentRows_(), reps = {};
  all.filter(c => c.parent).forEach(c => (reps[c.parent] = reps[c.parent] || []).push({ from: c.from, text: c.text, at: c.at }));
  return all.filter(c => !c.parent && filter(c)).map(c => Object.assign(c, { replies: reps[c.id] || [] }))
    .sort((a, b) => (b.updated || b.at).localeCompare(a.updated || a.at));
}

const vShort_ = v => ({ id: v.id, d: v.date, t: (v.closed || '').slice(11, 16), by: v.by, items: v.items, ph: v.photos, work: v.workDone, mp: v.manpower });

/* ----- engineer home (a manager can open anyone's with b.as) ----- */
function home_(b, me) {
  const who = me.manager && b.as ? (teamPeople_().find(p => p.name === b.as) || me) : me;
  const cu = cuUnits_(), vis = visitsByUnit_(), today = fmt_(new Date(), 'yyyy-MM-dd');
  const days = d => Math.round((new Date(today) - new Date(d)) / 86400000);
  const units = Object.keys(cu).filter(u => isMine_(who, cu[u].owners)).map(u => {
    const v = (vis[u] || [])[0];
    return { unit: u, client: cu[u].client, reqs: cu[u].reqs, poc: cu[u].poc, pms: cu[u].pms, last: v ? vShort_(v) : null, ago: v ? days(v.date) : null };
  }).sort((a, b) => (a.ago == null ? 9999 : a.ago) - (b.ago == null ? 9999 : b.ago) || a.unit.localeCompare(b.unit));
  const mine = {}; units.forEach(u => mine[u.unit] = 1);
  const elsewhere = [];
  Object.keys(vis).forEach(u => vis[u].forEach(v => { if (!mine[u] && engKey_(v.by) === engKey_(who.name) && days(v.date) <= 14) elsewhere.push(Object.assign(vShort_(v), { unit: u })); }));
  elsewhere.sort((a, b) => (b.d + b.t).localeCompare(a.d + a.t));
  const forMe = threads_(c => low_(c.to) === low_(who.name));
  return { me: { name: me.name, manager: me.manager }, who: who.name, units: units, elsewhere: elsewhere.slice(0, 40),
    open: forMe.filter(c => c.status !== 'closed'), closed: forMe.filter(c => c.status === 'closed').slice(0, 30) };
}

/* ----- one unit ----- */
function unitInfo_(b, me) {
  const u = normUnit_(b.unit) || str_(b.unit);
  const cu = cuUnits_()[u] || { unit: u, client: '', owners: [], pms: [], reqs: [], poc: null };
  const vs = (visitsByUnit_()[u] || []).slice(0, 30);
  if (!me.manager && !isMine_(me, cu.owners) && !vs.some(v => engKey_(v.by) === engKey_(me.name))) throw new Error('الوحدة دي مش من وحداتك.');
  const crew = crewByVisit_(vs);
  return { me: { name: me.name, manager: me.manager }, unit: u, client: cu.client, owners: cu.owners, pms: cu.pms, reqs: cu.reqs, poc: cu.poc,
    visits: vs.map(v => Object.assign(vShort_(v), { crew: (crew[v.id] || []).map(c => [c.contractor, c.trade, c.count]) })),
    comments: threads_(c => c.unit === u) };
}

function visitPhotos_(b, me) { return dashPhotos_({ visitId: b.visitId }); }

/* ----- manager home ----- */
function mgrHome_(b, me) {
  if (!me.manager) throw new Error('للمديرين بس.');
  const cu = cuUnits_(), vis = visitsByUnit_(), today = fmt_(new Date(), 'yyyy-MM-dd');
  const d0 = new Date(today); d0.setDate(d0.getDate() - ((d0.getDay() + 1) % 7)); const weekFrom = fmt_(d0, 'yyyy-MM-dd');  // week starts Saturday
  const days = d => Math.round((new Date(today) - new Date(d)) / 86400000);
  // workload share, same rule as the Workload tab
  const tasks = rows_(CU_TAB).filter(r => low_(r['Status']) === CU.status);
  const V = {}, D = {};
  tasks.forEach(r => { const o = str_(r['Owner']); if (!o || o === '-' || low_(o) === 'real estate') return;
    V[o] = (V[o] || 0) + (Number(r['Amount']) || 0); D[o] = (D[o] || 0) + (Number(r['Days']) || 0); });
  const TV = Object.keys(V).reduce((s, k) => s + V[k], 0) || 1, TD = Object.keys(D).reduce((s, k) => s + D[k], 0) || 1;
  const team = teamPeople_();
  const eng = Object.keys(V).map(o => {
    const p = team.find(x => low_(x.cu) === low_(o)) || { name: o, cu: o };
    const us = Object.keys(cu).filter(u => cu[u].owners.indexOf(o) >= 0);
    const st = u => { const v = (vis[u] || [])[0]; return v ? days(v.date) : null; };
    const todayVisits = Object.keys(vis).reduce((n, u) => n + vis[u].filter(v => v.date === today && engKey_(v.by) === engKey_(p.name)).length, 0);
    return { cu: o, name: p.name, share: Math.round((0.6 * V[o] / TV + 0.4 * D[o] / TD) * 1000) / 10, units: us.length,
      today: us.filter(u => st(u) === 0).length, week: us.filter(u => (vis[u] || []).some(v => v.date >= weekFrom)).length,
      never: us.filter(u => st(u) == null).length, stale: us.filter(u => st(u) != null && st(u) >= 3).length, visitsToday: todayVisits };
  }).sort((a, b) => b.share - a.share);
  const owned = Object.keys(cu).filter(u => cu[u].owners.some(o => V[o] != null));
  const todayAll = [].concat(...Object.keys(vis).map(u => vis[u].filter(v => v.date === today).map(v => Object.assign(v, { u: u }))));
  const alerts = [];
  const idle = eng.filter(e => !e.visitsToday).map(e => e.name);
  if (idle.length) alerts.push(['r', idle.length + ' مهندس من غير ولا زيارة النهارده: ' + idle.join(' · ')]);
  const never = owned.filter(u => !vis[u]).length; if (never) alerts.push(['r', never + ' وحدة Site Execution عمرها ما اتزارت']);
  const other = todayAll.filter(v => cu[v.u] && cu[v.u].owners.length && !isMine_({ name: v.by, cu: v.by }, cu[v.u].owners));
  if (other.length) alerts.push(['o', other.length + ' زيارة النهارده على وحدات مهندس تاني']);
  const weak = todayAll.filter(v => str_(v.workDone).replace(/[\s.\-_,،]/g, '').length < 4);
  if (weak.length) alerts.push(['o', weak.length + ' زيارة النهارده من غير وصف للشغل: ' + weak.map(v => v.u.replace(/^OW\//, '')).join(' · ')]);
  const th = threads_(() => true), open = th.filter(c => c.status !== 'closed');
  const late = open.filter(c => days(c.at.slice(0, 10)) >= 1 && !c.replies.length);
  if (late.length) alerts.push(['o', late.length + ' كومنت مفتوح من أكتر من يوم من غير رد']);
  return { me: { name: me.name, manager: true }, now: fmt_(new Date(), 'yyyy-MM-dd HH:mm'),
    kpi: { units: owned.length, today: owned.filter(u => (vis[u] || []).some(v => v.date === today)).length,
      week: owned.filter(u => (vis[u] || []).some(v => v.date >= weekFrom)).length,
      visits: todayAll.length, workers: todayAll.reduce((s, v) => s + v.manpower, 0), idle: idle.length, open: open.length },
    alerts: alerts, eng: eng,
    units: Object.keys(cu).sort().map(u => { const v = (vis[u] || [])[0]; return [u, cu[u].client, cu[u].owners.join(', '), v ? v.date : '', v ? v.by : '']; }) };
}

function commentsList_(b, me) {
  if (!me.manager) throw new Error('للمديرين بس.');
  return { comments: threads_(() => true).slice(0, 300) };
}

/* ----- writing comments (also written on the unit's ClickUp task) ----- */
function addComment_(b, me) {
  if (!me.manager) throw new Error('الكومنتات للمديرين بس.');
  const u = normUnit_(b.unit) || str_(b.unit), text = str_(b.text);
  if (!u || !text) throw new Error('اكتب الكومنت.');
  const cu = cuUnits_()[u] || { owners: [] };
  const team = teamPeople_();
  const to = cu.owners.map(o => (team.find(p => low_(p.cu) === low_(o)) || {}).name).filter(Boolean)[0] || '';
  const id = 'C-' + fmt_(new Date(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
  const v = b.visitId ? closedVisits_(x => x.id === str_(b.visitId))[0] : null;
  const cuText = `📌 Site comment from ${me.name}${to ? ' to ' + to : ''}${v ? ` (visit ${v.date} by ${v.by})` : ''}:\n${text}`;
  const cid = cuComment_(u, cuText);
  const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
  sheet_(COMMENTS_TAB, COMMENT_COLS).appendRow([id, now, u, v ? v.id : '', me.name, to, text, 'open', '', cid, now, '', '']);
  return { id: id, to: to, clickup: !!cid };
}

function replyComment_(b, me) {
  const c = commentRows_().find(x => x.id === str_(b.id) && !x.parent);
  if (!c) throw new Error('الكومنت ده مش موجود.');
  if (!me.manager && low_(c.to) !== low_(me.name)) throw new Error('الكومنت ده مش ليك.');
  const text = str_(b.text); if (!text) throw new Error('اكتب الرد.');
  const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
  const cid = cuComment_(c.unit, `↩️ ${me.name}: ${text}`);
  sheet_(COMMENTS_TAB, COMMENT_COLS).appendRow([c.id + '-r' + Utilities.getUuid().slice(0, 4), now, c.unit, '', me.name, '', text, '', c.id, cid, now, '', '']);
  setComment_(c.id, { Status: me.manager ? c.status : 'replied', Updated: now });
  return { done: true };
}

function closeComment_(b, me) {
  const c = commentRows_().find(x => x.id === str_(b.id) && !x.parent);
  if (!c) throw new Error('الكومنت ده مش موجود.');
  if (!me.manager && low_(c.to) !== low_(me.name)) throw new Error('الكومنت ده مش ليك.');
  const v = b.visitId ? closedVisits_(x => x.id === str_(b.visitId))[0] : null;
  const now = fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss');
  cuComment_(c.unit, `✅ Closed by ${me.name}${v ? ` with a new visit (${v.photos} photos)` : ''}.`);
  setComment_(c.id, { Status: 'closed', Updated: now, 'Closed by': me.name, 'Closed with visit': v ? v.id : '' });
  return { done: true };
}

function reopenComment_(b, me) {
  if (!me.manager) throw new Error('للمديرين بس.');
  setComment_(str_(b.id), { Status: 'open', Updated: fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss') });
  return { done: true };
}

function setComment_(id, vals) {
  const sh = sheet_(COMMENTS_TAB, COMMENT_COLS), v = sh.getDataRange().getValues(), h = v[0];
  for (let i = 1; i < v.length; i++) if (str_(v[i][0]) === id) {
    Object.keys(vals).forEach(k => { const c = h.indexOf(k); if (c >= 0) sh.getRange(i + 1, c + 1).setValue(vals[k]); });
    return;
  }
}

/** Posts a comment on the unit's first Site Execution task in ClickUp. Returns the ClickUp comment id, or '' if it could not. */
function cuComment_(unit, text) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN'); if (!token) return '';
    const ids = cuTaskIdsForUnit_(unit, token); if (!ids.length) return '';
    const r = UrlFetchApp.fetch(`https://api.clickup.com/api/v2/task/${ids[0]}/comment`, { method: 'post', contentType: 'application/json',
      headers: { Authorization: token }, payload: JSON.stringify({ comment_text: text, notify_all: false }), muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) { console.error('ClickUp comment ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200)); return ''; }
    return String(JSON.parse(r.getContentText()).id || '');
  } catch (e) { console.error('ClickUp comment: ' + e); return ''; }
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

function weeklyStart_(endDay, send) {
  const startDay = fmt_(new Date(parseDay_(endDay).getTime() - 6 * 86400000), 'yyyy-MM-dd');
  const units = [];
  closedVisits_(v => v.date >= startDay && v.date <= endDay).forEach(v => { if (units.indexOf(v.unit) < 0) units.push(v.unit); });
  units.sort();
  if (!units.length) return { units: 0 };
  PropertiesService.getScriptProperties().setProperty(WEEK_KEY, JSON.stringify({ end: endDay, start: startDay, send: !!send, units: units, files: {} }));
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
    const to = S['Weekly report to'] || S['Weekly report copy to'];
    const cc = [S['Weekly report cc'], S['Weekly report to'] ? S['Weekly report copy to'] : ''].filter(Boolean).join(', ');
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
  made.forEach(m => logReport_('Weekly client', period, m.unit, m.file, st.send ? (sentTo[m.unit] || '') : '(made from menu, not emailed)'));
  if (mgrPdf) logReport_('Weekly managers', period, 'Managers', mgrPdf, st.send ? 'with the weekly email' : '(made from menu, not emailed)');
  return { units: made.length };
}

/** Client-facing weekly PDF for one unit, in the O West brand (navy band, logos, peach labels).
    Page 1: unit, client, period, scope, progress highlights (AI) and the latest site photo.
    Then a photo log by day (a picked set, spread over the week). No manpower, no internal notes. */
const WK_MAX_PHOTOS = 8, WK_MAX_PER_DAY = 4;
const BRAND = { navy: '#001641', peach: '#e9b88f', ink: '#1d2433', soft: '#fbf3ec', line: '#e2dcd5', mute: '#7a8194' };

function weeklyHtml_(S, unit, cl, startDay, endDay, list, photos, crew) {
  const days = groupBy_(list, v => v.date);
  const dayKeys = Object.keys(days).sort();
  // scope = what the client bought (ClickUp), never the engineers' photo tags
  const items = rows_(T.scope).filter(r => low_(r['Unit']) === unit.toLowerCase()).map(r => str_(r['Item'])).filter(Boolean)
    .filter((x, i, a) => a.indexOf(x) === i && !/^others?$|^general$/i.test(x));
  const ai = polishWeek_(unit, items, dayKeys.map(d => ({ date: d, notes: days[d].map(v => v.workDone).filter(Boolean) })));

  // pick photos: spread evenly through each day, at most WK_MAX_PER_DAY a day and WK_MAX_PHOTOS a week
  const perDay = {};
  const quota = Math.max(1, Math.min(WK_MAX_PER_DAY, Math.floor(WK_MAX_PHOTOS / Math.max(1, dayKeys.length))));
  let total = 0;
  dayKeys.forEach(d => {
    let ph = [];
    days[d].forEach(v => ph = ph.concat(photos[v.id] || []));
    perDay[d] = { pick: spread_(ph, quota) };
    total += ph.length;
  });
  const lastDay = dayKeys[dayKeys.length - 1];
  const allPicked = [].concat(...dayKeys.map(d => perDay[d].pick));
  const imgs = images_(allPicked.map(p => p.fileId));
  const img = (p, W, H) => {
    const im = p && imgs[p.fileId];
    if (!im) return '';
    const s = im.w && im.h ? Math.min(W / im.w, H / im.h) : 0;
    const w = s ? Math.round(im.w * s) : W, h = s ? Math.round(im.h * s) : 0;
    const pad = h ? Math.floor((H - h) / 2) : 0;
    return `<div style="height:${pad}px"></div><img src="${im.src}" width="${w}"${h ? ` height="${h}"` : ''}>`;
  };
  let hero = allPicked.slice(-1)[0], heroDay = lastDay;
  dayKeys.forEach(d => perDay[d].pick.forEach(p => { const im = imgs[p.fileId]; if (im && im.w > im.h) { hero = p; heroDay = d; } }));
  const co = S['Company name'] || 'O-Plus';
  const C = BRAND;

  let h = `<html lang="en"><head><meta charset="utf-8"><style>
    ${fontsCss_()}
    @page { size: A4; margin: 16mm 15mm 16mm; }
    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    html, body { margin: 0; padding: 0; }
    body { font-family: 'Poppins', Arial, sans-serif; font-size: 9.5pt; line-height: 1.6; color: ${C.ink}; }
    table { border-collapse: collapse; width: 100%; }
    tr, .keep { page-break-inside: avoid; }
    .ar { direction: rtl; text-align: right; font-family: Arial, sans-serif; }
    .serif { font-family: 'Liberation Serif', 'Times New Roman', serif; font-weight: bold; }
    .band td { background: ${C.navy}; color: #ffffff; padding: 22px 26px 26px; vertical-align: top; }
    .kick { font-size: 7.5pt; letter-spacing: 2.5px; color: ${C.peach}; text-transform: uppercase; font-weight: 500; }
    .unit { font-family: 'Liberation Serif', 'Times New Roman', serif; font-weight: bold; font-size: 28pt; line-height: 1.15; color: #ffffff; margin-top: 10px; }
    .client { font-size: 13pt; color: #dfe3ec; margin-top: 6px; }
    .rev { font-size: 7.5pt; letter-spacing: 2px; color: #c9cfdd; text-align: right; }
    .meta td { padding: 18px 4px 16px 0; vertical-align: top; border-bottom: 1px solid ${C.line}; }
    .lab { font-size: 6.8pt; letter-spacing: 2px; color: ${C.mute}; text-transform: uppercase; font-weight: 500; }
    .val { font-size: 9.5pt; color: ${C.navy}; margin-top: 4px; }
    .sec { margin-top: 30px; }
    .sec td { border-bottom: 1.5px solid ${C.navy}; padding-bottom: 6px; vertical-align: bottom; }
    .sec h2 { font-family: 'Liberation Serif', 'Times New Roman', serif; font-size: 15pt; color: ${C.navy}; margin: 0; }
    .sec .r { font-size: 10pt; color: ${C.mute}; text-align: right; }
    .chip { display: inline-block; border: 1px solid ${C.peach}; color: #8a5a35; padding: 3px 12px; font-size: 8pt; margin: 4px 6px 0 0; }
    .hl td { vertical-align: top; padding: 14px 10px 0 0; font-size: 10pt; line-height: 1.8; }
    .hl ul { margin: 0; }
    .hl li { margin-bottom: 6px; }
    .hero { text-align: center; height: 300px; background: ${C.soft}; margin-top: 14px; }
    .day { margin-top: 20px; }
    .day td { padding: 0 0 6px; border-bottom: 1px solid ${C.line}; }
    .day .d { font-weight: 500; color: ${C.navy}; font-size: 10.5pt; }
    .day .n { font-size: 9pt; color: ${C.mute}; }
    .line td { padding: 10px 10px 4px 0; font-size: 9.5pt; line-height: 1.7; vertical-align: top; color: #33405a; }
    .grid { margin-top: 6px; }
    .grid td { width: 33.3%; padding: 5px; vertical-align: top; text-align: center; }
    .box { height: 196px; background: ${C.soft}; text-align: center; }
    .page { page-break-before: always; }
    .sheet { position: relative; height: 262mm; overflow: hidden; page-break-after: always; }
    .sheet.last { page-break-after: auto; }
    .foot { position: absolute; left: 0; right: 0; bottom: 0; }
    .foot td { font-size: 7pt; letter-spacing: 1px; color: ${C.mute}; border-top: 1px solid ${C.line}; padding-top: 6px; }
  </style></head><body><div class="sheet" id="p1">
  <table class="band"><tr><td>
    <table><tr>
      <td style="padding:0;background:${C.navy}"><img src="data:image/png;base64,${LOGO_OWEST}" height="24"> &nbsp;&nbsp; <img src="data:image/png;base64,${LOGO_OPLUS}" height="24"></td>
      <td class="rev" style="padding:0;background:${C.navy}">WEEK ENDING ${esc_(dateEn_(endDay)).toUpperCase()}</td>
    </tr></table>
    <div style="height:34px"></div>
    <div class="kick">Weekly Progress Report</div>
    <div class="unit">${esc_(unit)}</div>
    ${cl.name ? `<div class="client">${esc_(cl.name)}</div>` : ''}
  </td></tr></table>
  <table class="meta"><tr>
    <td style="width:30%"><div class="lab">Client</div><div class="val">${esc_(cl.name || '—')}</div></td>
    <td style="width:24%"><div class="lab">Unit</div><div class="val">${esc_(unit)}</div></td>
    <td style="width:26%"><div class="lab">Period</div><div class="val">${esc_(dateEn_(startDay))} – ${esc_(dateEn_(endDay))}</div></td>
    <td><div class="lab">Prepared by</div><div class="val">${esc_(co)} Extra Works</div></td>
  </tr></table>
`;
  if (ai && ai.en.length) {
    h += `<table class="sec"><tr><td><h2>Progress highlights</h2></td><td class="r">This week</td></tr></table>
      <table class="hl"><tr><td><ul style="padding-left:16px">${ai.en.map(x => `<li>${esc_(x)}</li>`).join('')}</ul></td></tr></table>`;
  }
  if (hero) {
    h += `<table class="sec"><tr><td><h2>Latest site photo</h2></td><td class="r">${esc_(dateEn_(heroDay))}</td></tr></table>
      <div class="hero keep">${img(hero, 650, 298)}</div>`;
  }

  // photo pages: laid out by hand so the footer always sits at the bottom of each page
  const foot = `<table class="foot"><tr><td>O West by Orascom Development · ${esc_(unit)}</td><td style="text-align:right">${esc_(co)} Extra Works · Week ending ${esc_(dateEn_(endDay))}</td></tr></table>`;
  const CAP = 250, HEAD = 18, LINE = 9, ROW = 56, TITLE = 14;   // mm
  const pages = [];
  let cur = '', used = 0;
  const newPage = () => { if (cur) pages.push(cur); cur = `<table class="sec" style="margin-top:0"><tr><td><h2>Site progress photos</h2></td><td class="r">Newest first</td></tr></table>`; used = TITLE; };
  newPage();
  dayKeys.slice().reverse().forEach(d => {              // newest day first
    const pd = perDay[d];
    if (!pd.pick.length) return;
    const line = ai && ai.days[d];
    const head = (cont) => `<table class="day"><tr><td><span class="d">${esc_(dateEn_(d, true))}</span>${cont ? ' <span class="n">(continued)</span>' : ''}</td></tr></table>` +
      (!cont && line && line.en ? `<table class="line"><tr><td>${esc_(line.en)}</td></tr></table>` : '');
    const headH = HEAD + (line && line.en ? LINE : 0);
    if (used + headH + ROW > CAP) newPage();
    cur += head(false); used += headH;
    for (let i = 0; i < pd.pick.length; i += 3) {
      if (used + ROW > CAP) { newPage(); cur += head(true); used += HEAD; }
      cur += '<table class="grid"><tr>' + [pd.pick[i], pd.pick[i + 1], pd.pick[i + 2]].map(p => p ? `<td><div class="box">${img(p, 208, 194)}</div></td>` : '<td></td>').join('') + '</tr></table>';
      used += ROW;
    }
  });
  if (cur) pages.push(cur);
  h += foot + '</div>';
  pages.forEach((pg, i) => h += `<div class="sheet${i === pages.length - 1 ? ' last' : ''}">${pg}${foot}</div>`);
  return h + '</body></html>';
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
/** Settings "AI text" = Off: reports go out with photos only, no AI call at all. */
function aiOff_() { return low_(settings_()['AI text']) === 'off'; }

function polishWeek_(unit, items, days) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key || AI_DOWN || aiOff_()) return null;
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
  if (!key || AI_DOWN || aiOff_() || !notes.length) return {};
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
  while (old.hasNext()) {
    const f = old.next();
    // a PDF already shared by link may be in a client's WhatsApp: never delete it, keep it working
    if (f.getSharingAccess() === DriveApp.Access.ANYONE_WITH_LINK) continue;
    f.setTrashed(true);
  }
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
    .addSeparator()
    .addItem('Delete selected photos', 'menuDeletePhotos')
    .addItem('Delete selected visits (photos + manpower)', 'menuDeleteVisits')
    .addSeparator()
    .addItem('Set manager password', 'menuSetPassword')
    .addItem('Set dashboard password', 'menuSetDashPassword')
    .addItem('Rebuild Workload tab now', 'menuWorkload')
    .addSeparator()
    .addItem('Team: give PINs to everyone without one', 'menuMakePins')
    .addItem('Team: set one PIN', 'menuSetPin')
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
    mobile: '1707ad3e-025b-457f-a4f3-cf1d7b6b1a7e',
    owner: 'd7255ee0-9352-4cb6-8e1e-6d412405d400',  // "Task Owner" = site engineer in charge of the unit
    pm: 'f91d9a67-0ed7-4e6a-81c3-1b7bf673546f',
    amount: 'f480622e-13eb-43c9-a141-f74a78d866e9',
    poc: '640aa77a-6d8d-4b4d-b8a6-b4dfd74baa9f'      // "Actual POC %" = progress on site
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
  if (r.busy) return SpreadsheetApp.getUi().alert('A sync is already running. Try again in 3 minutes.');
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
  // The long ClickUp reading must NOT hold the script lock: photo uploads and visit closing wait on it.
  // A cache flag stops two syncs running together; the lock is taken only for the short sheet writes (step 3).
  const cache = CacheService.getScriptCache();
  if (cache.get('syncRunning')) return { units: 0, items: 0, contractors: 0, newContractors: 0, pending: 0, busy: true };
  cache.put('syncRunning', '1', 330);
  let lock = null;
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
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
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
    try { writeWorkloadTab_(tasks); setSetting_('Workload tab', fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ' · OK'); }
    catch (e) { console.error('workload tab: ' + e.stack); setSetting_('Workload tab', fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ' · ERROR: ' + e.message); }
    clearBoot_();
    return { units: Object.keys(units).length, items: sRows.length, contractors: cRows.length, newContractors: freshC.length };
  } finally { if (lock) lock.releaseLock(); cache.remove('syncRunning'); }
}

/* ---------- ClickUp task store ("ClickUp Tasks" tab) ---------- */
const CU_TAB = 'ClickUp Tasks';
const CU_COLS = ['Unit', 'Task ID', 'Status', 'Client', 'Email', 'Mobile', 'Type', 'Description', 'Contractor', 'Updated', 'Checked', 'Owner', 'PM', 'Amount', 'POC', 'Days'];

/** Compact row from a full ClickUp task object. */
function cuCompact_(t) {
  const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
  const v = id => str_(f[id] && f[id].value);
  return {
    unit: normUnit_(t.name) || str_(t.name), id: str_(t.id), status: str_(t.status && t.status.status),
    client: v(CU.f.client), email: v(CU.f.email), mobile: v(CU.f.mobile), type: dropName_(f[CU.f.requestType]),
    desc: v(CU.f.description).slice(0, 200), contractor: v(CU.f.contractor), updated: str_(t.date_updated), checked: String(Date.now()),
    owner: dropName_(f[CU.f.owner]) || '-',  // '-' = read, no owner set
    pm: dropName_(f[CU.f.pm]), amount: v(CU.f.amount), poc: v(CU.f.poc),
    days: String(cuDays_(t.start_date, t.due_date))   // '' only on rows not read yet by this version
  };
}

/** ClickUp "Duration": whole days from start date to due date, both counted (Cairo dates). 0 when either is missing. */
function cuDays_(start, due) {
  if (!start || !due) return 0;
  const d = ms => { const s = Utilities.formatDate(new Date(Number(ms)), 'Africa/Cairo', 'yyyy-MM-dd'); return Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)); };
  return Math.round((d(due) - d(start)) / 86400000) + 1;
}

/** Updates the tab and returns every stored task (all statuses). Safe to run often. */
function cuSyncTasks_(token) {
  const sh = sheet_(CU_TAB, CU_COLS);
  if (sh.getRange(1, 1, 1, CU_COLS.length).getValues()[0].join('|') !== CU_COLS.join('|')) sh.getRange(1, 1, 1, CU_COLS.length).setValues([CU_COLS]).setFontWeight('bold');
  const byId = {};
  rows_(CU_TAB).forEach(r => {
    const id = str_(r['Task ID']); if (!id) return;
    byId[id] = { unit: str_(r['Unit']), id: id, status: str_(r['Status']), client: str_(r['Client']), email: str_(r['Email']), mobile: str_(r['Mobile']),
      type: str_(r['Type']), desc: str_(r['Description']), contractor: str_(r['Contractor']), updated: str_(r['Updated']), checked: str_(r['Checked']), owner: str_(r['Owner']),
      pm: str_(r['PM']), amount: str_(r['Amount']), poc: str_(r['POC']), days: str_(r['Days']) };
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
  //    (rows without an owner value yet come first, so the engineer split fills in on the first sync)
  //    (rows not read by this version yet — no Days value — come first, so new columns fill in on the first sync)
  const stale = Object.keys(byId).filter(id => byId[id].days === '' || !byId[id].checked || started - Number(byId[id].checked) > 86400000)
    .sort((a, b) => (byId[a].days === '' ? 0 : 1) - (byId[b].days === '' ? 0 : 1) || Number(byId[a].checked || 0) - Number(byId[b].checked || 0));
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
  rewrite_(sh, CU_COLS, list.map(t => [t.unit, t.id, t.status, t.client, t.email, t.mobile, t.type, t.desc, t.contractor, t.updated, t.checked, t.owner || '', t.pm || '', t.amount || '', t.poc || '', t.days == null ? '' : t.days]));
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
