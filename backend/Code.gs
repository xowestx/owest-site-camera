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
  put(`Manpower (${v.total}): ` + (v.crew.length ? v.crew.map(m => `${m.contractor}${m.trade ? ' – ' + m.trade : ''} ${m.count}`).join(' · ') : 'none'));
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
  h += `<h2>ملخص اليوم <span class="en">Summary</span></h2><table class="sum">
    <tr><th class="c" style="width:22px">#</th><th style="width:118px">الوحدة <span class="en">Unit</span></th><th style="width:78px">الوقت <span class="en">Time</span></th>
    <th style="width:120px">البنود <span class="en">Items</span></th><th class="c" style="width:46px">العمالة</th><th>الشغل <span class="en">Work done</span></th></tr>`;
  list.forEach((v, i) => {
    h += `<tr><td class="c">${i + 1}</td><td class="mono b">${esc_(v.unit)}</td><td class="mono">${timeRange_(v)}</td>
      <td class="sm">${esc_(v.items)}</td><td class="c b">${v.manpower || '—'}</td><td>${br_(v.workDone)}</td></tr>`;
  });
  h += `</table><p class="foot">الأوقات بتوقيت القاهرة · Cairo time</p>`;
  list.forEach((v, i) => {
    h += `<div class="page">${unitHead_(i + 1, v, eng)}`;
    h += photoGrid_(photos[v.id] || [], 0);
    h += workAndCrew_(v, crew[v.id] || []);
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
  const nPh = list.reduce((s, v) => s + (photos[v.id] || []).length, 0);

  let h = head_('تقرير الأسبوع', 'Weekly Progress Report', `${startDay} → ${endDay}`, `<span class="mono">${esc_(unit)}</span>${cl.name ? ' · ' + esc_(cl.name) : ''}`);
  h += kpis_([[dayKeys.length, 'أيام شغل', 'Working days'], [workerDays, 'عامل × يوم', 'Worker-days'], [nPh, 'صور', 'Photos']]);
  if (items.length) h += `<p class="items">${items.map(i => `<span class="chip">${esc_(i)}</span>`).join(' ')}</p>`;
  h += `<h2>ملخص الأسبوع <span class="en">This week</span></h2><table class="sum">
    <tr><th style="width:130px">اليوم <span class="en">Day</span></th><th class="c" style="width:56px">العمالة</th><th>الشغل اللي اتعمل <span class="en">Work done</span></th></tr>`;
  dayKeys.forEach(d => {
    const vs = days[d];
    h += `<tr><td>${dayAr_(d)}</td><td class="c b">${vs.reduce((s, v) => s + v.manpower, 0) || '—'}</td><td>${vs.map(v => br_(v.workDone)).join('<br>')}</td></tr>`;
  });
  h += `</table>`;
  dayKeys.forEach(d => {
    const vs = days[d];
    let ph = [];
    vs.forEach(v => ph = ph.concat(photos[v.id] || []));
    const tot = vs.reduce((s, v) => s + v.manpower, 0);
    h += `<div class="page"><table class="uh"><tr><td><div class="ucode">${dayAr_(d)}</div><div class="sub mono">${esc_(unit)}</div></td>
      <td class="meta">${tot ? `<b>${tot}</b> عامل · ` : ''}<b>${ph.length}</b> صور</td></tr></table>`;
    h += photoGrid_(ph, 8);
    h += `<table class="two"><tr><td><div class="lbl">الحالة <span class="en">Status</span></div>${vs.map(v => br_(v.workDone)).join('<br>')}</td></tr></table>`;
    h += `</div>`;
  });
  return h + '</body></html>';
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

function workAndCrew_(v, c) {
  let crewHtml = '<div class="sm">مفيش عمالة.</div>';
  if (c.length) {
    crewHtml = '<table class="crew"><tr><th>المقاول</th><th>الحرفة</th><th class="c">العدد</th></tr>' +
      c.map(m => `<tr><td>${esc_(m.contractor)}</td><td>${esc_(m.trade ? tradeAr_(m.trade) : '—')}</td><td class="c b">${m.count}</td></tr>`).join('') +
      `<tr><td class="b">الإجمالي</td><td></td><td class="c b">${c.reduce((s, m) => s + m.count, 0)}</td></tr></table>`;
  }
  return `<table class="two"><tr>
    <td style="width:55%"><div class="lbl">الحالة <span class="en">Status / Work done</span></div>${br_(v.workDone)}</td>
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
