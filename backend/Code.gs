/**
 * O West Site Camera — backend API (Google Apps Script, bound to the "O West Site Camera" sheet).
 *
 * The camera page lives on GitHub Pages. This script:
 *   GET  ?action=units            -> active units, each with its purchased items (Scope tab)
 *   POST {action:"upload", ...}   -> saves one photo to Drive, adds a row to Photos
 *   POST {action:"delete", ...}   -> trashes a photo the same visit uploaded, marks it Deleted
 *   POST {action:"close",  ...}   -> closes a visit (the "Done" button), adds a row to Visits
 *
 * Deploy as Web app:  Execute as = Me,  Who has access = Anyone.
 *
 * Tabs
 *   Units  (first tab) : Unit | Active | Project
 *   Scope              : Unit | Item                       (what each unit bought, from ClickUp)
 *   Photos             : one row per photo   (Status = OK / Deleted)
 *   Visits             : one row per closed visit
 *
 * Drive: O West Site Photos / <Unit> / <YYYY-MM-DD> / <Unit>_<date>_<time>_<Item>_<id>.jpg
 *
 * Fixing mistakes: in the sheet use the "O West Camera" menu
 *   - select rows in Photos  -> "Delete selected photos"
 *   - select rows in Visits  -> "Delete selected visits (all their photos)"
 *   Files go to Drive trash (recoverable for 30 days); rows stay, marked Deleted.
 */

const ROOT_FOLDER_NAME = 'O West Site Photos';
const TZ = 'Africa/Cairo';
const UNITS_SHEET_INDEX = 0;
const SCOPE = 'Scope';
const PHOTOS = 'Photos';
const VISITS = 'Visits';
const GENERAL = 'General';
const MAX_BASE64_CHARS = 8 * 1024 * 1024;

const PHOTO_HEADERS = ['Photo ID', 'Visit ID', 'Uploaded at', 'Date', 'Unit', 'Project', 'Item', 'Note', 'By', 'File name', 'Photo link', 'File ID', 'Status'];
const VISIT_HEADERS = ['Visit ID', 'Date', 'Unit', 'Project', 'Items', 'By', 'Started', 'Closed', 'Photos', 'Status', 'Folder link'];
// column numbers (1-based) used below
const P = { id: 1, visit: 2, unit: 5, fileId: 12, status: 13 };
const V = { id: 1, status: 10 };

/* =========================== web API =========================== */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'units';
    if (action === 'units') return json_({ ok: true, units: unitsCached_() });
    if (action === 'ping') return json_({ ok: true });
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'upload') return json_(Object.assign({ ok: true }, upload_(b)));
    if (b.action === 'delete') return json_(Object.assign({ ok: true }, deletePhoto_(b)));
    if (b.action === 'close') return json_(Object.assign({ ok: true }, closeVisit_(b)));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

/* =========================== units =========================== */

function readUnits_() {
  const ss = SpreadsheetApp.getActive();
  const rows = ss.getSheets()[UNITS_SHEET_INDEX].getDataRange().getValues().slice(1);
  const scope = {};
  const sc = ss.getSheetByName(SCOPE);
  if (sc) {
    sc.getDataRange().getValues().slice(1).forEach(r => {
      const u = String(r[0]).trim().toLowerCase(), item = String(r[1]).trim();
      if (!u || !item) return;
      scope[u] = scope[u] || [];
      if (scope[u].indexOf(item) < 0) scope[u].push(item);
    });
  }
  return rows
    .filter(r => String(r[0]).trim() && String(r[1]).trim().toLowerCase() !== 'no')
    .map(r => {
      const unit = String(r[0]).trim();
      return { unit: unit, project: String(r[2] || '').trim(), items: scope[unit.toLowerCase()] || [] };
    });
}

function unitsCached_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('units');
  if (hit) return JSON.parse(hit);
  const units = readUnits_();
  try { cache.put('units', JSON.stringify(units), 600); } catch (e) {}
  return units;
}

function findUnit_(name) {
  const k = String(name || '').trim().toLowerCase();
  const u = unitsCached_().find(x => x.unit.toLowerCase() === k);
  if (!u) throw new Error('This unit is not in the active list. Reload the page.');
  return u;
}

/** Any edit to Units or Scope refreshes the list the app sees. */
function onEdit(e) {
  try {
    const name = e.range.getSheet().getName();
    const first = SpreadsheetApp.getActive().getSheets()[UNITS_SHEET_INDEX].getName();
    if (name === first || name === SCOPE) CacheService.getScriptCache().remove('units');
  } catch (err) {}
}

/* =========================== upload =========================== */

function upload_(p) {
  const photoId = cleanId_(p.photoId, 'P');
  const visitId = cleanId_(p.visitId, 'V');
  const cache = CacheService.getScriptCache();

  // Same photo sent twice (retry after a slow network): return the first result.
  const done = cache.get('p:' + photoId);
  if (done) return JSON.parse(done);

  if (typeof p.base64 !== 'string' || !p.base64) throw new Error('No photo received.');
  if (p.base64.length > MAX_BASE64_CHARS) throw new Error('Photo is too large.');

  const u = findUnit_(p.unit);
  const item = u.items.indexOf(p.item) >= 0 ? p.item : GENERAL;
  const tag = item.replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
  const safeUnit = u.unit.replace(/[\/\\]+/g, '-').replace(/\s+/g, ' ').trim();
  const by = String(p.by || '').trim().slice(0, 60);
  const note = String(p.note || '').trim().slice(0, 300);

  const now = new Date();
  const day = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const time = Utilities.formatDate(now, TZ, 'HHmmss');
  const folder = dayFolder_(safeUnit, day);

  const name = `${safeUnit}_${day}_${time}_${tag}_${photoId.slice(-5)}.jpg`;
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.base64), 'image/jpeg', name));
  file.setDescription([u.unit, item, note, by && ('By ' + by), 'Photo ' + photoId, 'Visit ' + visitId].filter(Boolean).join(' | '));

  sheet_(PHOTOS, PHOTO_HEADERS).appendRow([
    photoId, visitId, Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), day,
    u.unit, u.project, item, note, by, name, file.getUrl(), file.getId(), 'OK'
  ]);

  const out = { photoId: photoId, name: name };
  try { cache.put('p:' + photoId, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}

/** <root>/<unit>/<day>, with the folder id cached so most uploads skip the folder search. */
function dayFolder_(safeUnit, day) {
  const cache = CacheService.getScriptCache();
  const key = 'f:' + safeUnit + ':' + day;
  const id = cache.get(key);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const again = cache.get(key);
    if (again) { try { return DriveApp.getFolderById(again); } catch (e) {} }
    const f = child_(child_(rootFolder_(), safeUnit), day);
    cache.put(key, f.getId(), 21600);
    return f;
  } finally {
    lock.releaseLock();
  }
}

/* =========================== delete & close =========================== */

/** From the app: only a photo of the same visit, while the visit is open. */
function deletePhoto_(p) {
  const photoId = cleanId_(p.photoId, 'P');
  const visitId = cleanId_(p.visitId, 'V');
  const sh = sheet_(PHOTOS, PHOTO_HEADERS);
  const cell = sh.getRange(1, P.id, sh.getLastRow(), 1).createTextFinder(photoId).matchEntireCell(true).findNext();
  if (!cell) throw new Error('Photo not found.');
  const row = sh.getRange(cell.getRow(), 1, 1, PHOTO_HEADERS.length).getValues()[0];
  if (row[P.visit - 1] !== visitId) throw new Error('This photo belongs to another visit.');
  if (row[P.status - 1] !== 'OK') return { photoId: photoId, already: true };
  trashFile_(row[P.fileId - 1]);
  sh.getRange(cell.getRow(), P.status).setValue('Deleted');
  CacheService.getScriptCache().remove('p:' + photoId);
  return { photoId: photoId };
}

function closeVisit_(p) {
  const visitId = cleanId_(p.visitId, 'V');
  const u = findUnit_(p.unit);
  const vs = sheet_(VISITS, VISIT_HEADERS);
  if (vs.getLastRow() > 1 && vs.getRange(1, V.id, vs.getLastRow(), 1).createTextFinder(visitId).matchEntireCell(true).findNext()) {
    return { visitId: visitId, already: true };
  }
  // Count what actually landed for this visit.
  const ps = sheet_(PHOTOS, PHOTO_HEADERS);
  let count = 0; const items = [];
  if (ps.getLastRow() > 1) {
    ps.getRange(2, 1, ps.getLastRow() - 1, PHOTO_HEADERS.length).getValues().forEach(r => {
      if (r[P.visit - 1] === visitId && r[P.status - 1] === 'OK') {
        count++;
        if (items.indexOf(r[6]) < 0) items.push(r[6]);
      }
    });
  }
  const now = new Date();
  const day = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const safeUnit = u.unit.replace(/[\/\\]+/g, '-').replace(/\s+/g, ' ').trim();
  let link = '';
  try { link = dayFolder_(safeUnit, day).getUrl(); } catch (e) {}
  vs.appendRow([
    visitId, day, u.unit, u.project, items.join(', '), String(p.by || '').trim().slice(0, 60),
    String(p.startedAt || '').slice(0, 19).replace('T', ' '), Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
    count, 'Closed', link
  ]);
  return { visitId: visitId, count: count, items: items };
}

/* =========================== sheet menu (for you) =========================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('O West Camera')
    .addItem('Delete selected photos', 'menuDeletePhotos')
    .addItem('Delete selected visits (all their photos)', 'menuDeleteVisits')
    .addSeparator()
    .addItem('Refresh unit list in the app', 'menuRefreshUnits')
    .addToUi();
}

function menuDeletePhotos() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== PHOTOS) return ui.alert('Open the Photos tab and select the rows to delete.');
  const rows = selectedRows_(sh);
  if (!rows.length) return ui.alert('Select at least one photo row.');
  if (ui.alert(`Delete ${rows.length} photo row(s)? Files go to Drive trash.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let n = 0;
  rows.forEach(r => {
    const v = sh.getRange(r, 1, 1, PHOTO_HEADERS.length).getValues()[0];
    if (v[P.status - 1] !== 'OK') return;
    trashFile_(v[P.fileId - 1]);
    sh.getRange(r, P.status).setValue('Deleted');
    n++;
  });
  ui.alert(`${n} photo(s) deleted.`);
}

function menuDeleteVisits() {
  const ui = SpreadsheetApp.getUi();
  const vs = SpreadsheetApp.getActiveSheet();
  if (vs.getName() !== VISITS) return ui.alert('Open the Visits tab and select the visit rows to delete.');
  const rows = selectedRows_(vs);
  if (!rows.length) return ui.alert('Select at least one visit row.');
  if (ui.alert(`Delete ${rows.length} visit(s) and all their photos? Files go to Drive trash.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ids = rows.map(r => vs.getRange(r, V.id).getValue());
  const ps = sheet_(PHOTOS, PHOTO_HEADERS);
  let n = 0;
  if (ps.getLastRow() > 1) {
    const vals = ps.getRange(2, 1, ps.getLastRow() - 1, PHOTO_HEADERS.length).getValues();
    vals.forEach((v, i) => {
      if (ids.indexOf(v[P.visit - 1]) >= 0 && v[P.status - 1] === 'OK') {
        trashFile_(v[P.fileId - 1]);
        ps.getRange(i + 2, P.status).setValue('Deleted');
        n++;
      }
    });
  }
  rows.forEach(r => vs.getRange(r, V.status).setValue('Deleted'));
  ui.alert(`${rows.length} visit(s) and ${n} photo(s) deleted.`);
}

function menuRefreshUnits() {
  CacheService.getScriptCache().remove('units');
  SpreadsheetApp.getUi().alert('Done. The app will load the new list on next open.');
}

/* =========================== helpers =========================== */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function msg_(err) { return String(err && err.message || err); }

function cleanId_(v, prefix) {
  const s = String(v || '');
  if (!new RegExp('^' + prefix + '-[A-Za-z0-9-]{6,40}$').test(s)) throw new Error('Bad request. Reload the page.');
  return s;
}

function trashFile_(id) {
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}

function selectedRows_(sh) {
  const out = [];
  const list = sh.getActiveRangeList();
  if (!list) return out;
  list.getRanges().forEach(rg => {
    for (let r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) if (r > 1 && out.indexOf(r) < 0) out.push(r);
  });
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

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const existing = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const f = existing.hasNext() ? existing.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_ID', f.getId());
  return f;
}

function child_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Run once from the editor after pasting: approves permissions, creates the tabs and the Drive folder. */
function setup() {
  rootFolder_();
  sheet_(PHOTOS, PHOTO_HEADERS);
  sheet_(VISITS, VISIT_HEADERS);
  CacheService.getScriptCache().remove('units');
}
