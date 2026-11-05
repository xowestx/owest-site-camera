/**
 * O West Site Camera — backend API (Google Apps Script, bound to the "O West Site Camera" sheet).
 *
 * The camera page itself lives on GitHub Pages. This script only:
 *   GET  ?action=units      -> active unit list from the first tab
 *   POST {action:"upload"}  -> saves the photo to Drive and logs it
 *
 * Deploy as Web app:  Execute as = Me,  Who has access = Anyone.
 * Photos land in YOUR Drive:
 *   O West Site Photos / <Unit> / <YYYY-MM-DD> / <Unit>_<YYYY-MM-DD>_<HHmmss>_<Category>_<NN>.jpg
 */

const ROOT_FOLDER_NAME = 'O West Site Photos';
const TZ = 'Africa/Cairo';
const UNITS_SHEET_INDEX = 0;              // first tab = Unit | Active | Project
const LOG_SHEET_NAME = 'Log';
const MAX_BASE64_CHARS = 12 * 1024 * 1024; // ~9 MB image cap
const CATEGORIES = ['General', 'Kitchen', 'AC', 'Plumbing', 'Finishing', 'Snag'];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'units';
    if (action === 'units') return json_({ ok: true, units: getUnits() });
    if (action === 'ping') return json_({ ok: true });
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'upload') return json_(Object.assign({ ok: true }, uploadPhoto(body)));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Active units. Rows with Active = "No" are hidden. */
function getUnits() {
  const sh = SpreadsheetApp.getActive().getSheets()[UNITS_SHEET_INDEX];
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => String(r[0]).trim() && String(r[1]).trim().toLowerCase() !== 'no')
    .map(r => ({ unit: String(r[0]).trim(), project: String(r[2] || '').trim() }));
}

function uploadPhoto(p) {
  if (!p || typeof p.base64 !== 'string' || !p.base64) throw new Error('No photo received.');
  if (p.base64.length > MAX_BASE64_CHARS) throw new Error('Photo is too large.');

  const match = getUnits().find(u => u.unit.toLowerCase() === String(p.unit || '').trim().toLowerCase());
  if (!match) throw new Error('This unit is no longer in the list. Reload the page.');

  const category = CATEGORIES.indexOf(p.category) >= 0 ? p.category : 'General';
  const by = String(p.by || '').trim().slice(0, 60);
  const note = String(p.note || '').trim().slice(0, 300);
  const safeUnit = match.unit.replace(/[\/\\]+/g, '-').replace(/\s+/g, ' ').trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = new Date();
    const day = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, TZ, 'HHmmss');

    const unitFolder = child_(rootFolder_(), safeUnit);
    const dayFolder = child_(unitFolder, day);

    let n = 1;
    const it = dayFolder.getFiles();
    while (it.hasNext()) { it.next(); n++; }

    const name = `${safeUnit}_${day}_${time}_${category}_${String(n).padStart(2, '0')}.jpg`;
    const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), 'image/jpeg', name);
    const file = dayFolder.createFile(blob);
    if (note || by) file.setDescription([note, by && ('By: ' + by)].filter(Boolean).join(' | '));

    logSheet_().appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
      match.unit, match.project, category, note, by, name, file.getUrl(), dayFolder.getUrl()
    ]);
    return { name: name, n: n };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* deleted: recreate */ }
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

function logSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(['Uploaded at', 'Unit', 'Project', 'Category', 'Note', 'By', 'File name', 'Photo link', 'Folder link']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Run once from the editor to approve permissions and create the Drive folder + Log tab. */
function setup() {
  const f = rootFolder_();
  logSheet_();
  Logger.log('Photos folder: ' + f.getUrl());
}
