/**
 * O-Plus POP check point (add this as a second file next to Code.gs in the portal's Apps Script).
 *
 * Flow:
 * 1. Sales ticks "Paid" on the unit's task in ClickUp (Main Tracking Sheet).
 * 2. popWatch (every few minutes) sees it, posts a comment on the task and emails the Sales person a link
 *    to the check point page (portal/pop.html).
 * 3. On the page Sales uploads the signed quotation (it fills the totals and cheque schedule),
 *    fills anything missing and presses Send.
 *    Sales can add other attachments too (proof of payment, signed papers, saved Outlook emails).
 * 4. The POP email goes out from the Sales person's own Outlook (Power Automate flow, subject tag [POP-SEND]),
 *    the client gets the portal welcome from Info.ew, and "Welcome Mail" is ticked in ClickUp.
 *
 * One-time setup: run popSetup() from the editor. It starts the 10-minute check. Only tasks paid from now on
 * count: the Payment date must be on or after the setup day (or, with no payment date, the task is new).
 * For reading quotations: in the editor, Services (+) → Drive API → Add.
 */

const POP_TAB = 'Portal POP';
const POP_COLS = ['Task ID', 'Unit', 'Status', 'Token', 'Client', 'Client email', 'Sales', 'Notified at', 'Reminded at',
  'Quote file ID', 'Quote data', 'Form', 'Sent at', 'Sent by', 'Extra files'];
const POP_EXTRA_MAX = 18 * 1024 * 1024; // all attachments together must fit in one email
const POP_LIST = '1201250000001676'; // ClickUp: O-Plus › Projects Pipeline & Teams › Main Tracking Sheet
const POP_PAGE = PORTAL_URL + 'pop.html';
const POP_TAG = '[POP-SEND]';
const POP_OLD = 'Old', POP_WAIT = 'Waiting', POP_SENT = 'Sent', POP_SKIP = 'Skipped';
const POP_REMIND_MS = 24 * 3600 * 1000;
// ClickUp "Sales Engineer" option → the person and the Outlook mailbox the POP goes out from.
const POP_SALES = {
  esheiry: { name: 'Mohamed El Esheiry', email: 'mohamed.elesheiry@orascomre.com' },
  yara: { name: 'Yara Ahmed', email: 'yara.ahmed@orascomre.com' },
  dina: { name: 'Dina Moez', email: 'dina.moez@orascomre.com' },
  ghandour: { name: 'Ahmed ElGhandour', email: 'Ahmed.ElGhandour@orascomre.com' }
};
const POP_F = {
  paid: '328ef6d8-e0e2-4e8f-a0bd-72801392be1e', welcome: 'c6f79353-244d-4fd6-a86a-99655ee3a4b9',
  client: '75747da4-250c-42e7-8d65-98d9d503060f', email: 'ecdc4bdb-8444-4074-b32c-cd4c36274f25',
  unit: '536dcc2d-7696-4813-88ed-e3f4966f793e', quote: '432de28e-bb8f-4595-92ba-d2f8bd564df8',
  amount: 'f480622e-13eb-43c9-a141-f74a78d866e9', payDate: 'b2c472de-5d86-4e50-ad17-571055eb20e0',
  sales: '72351678-fe4e-4f94-8eef-e41fcab03889', pm: 'f91d9a67-0ed7-4e6a-81c3-1b7bf673546f',
  type: '8f185b3c-db17-4dac-b577-d377a949db9f'
};

/* ---------------- setup and the 10-minute check ---------------- */

function popSetup() {
  sheet_(POP_TAB, POP_COLS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('POP_START')) props.setProperty('POP_START', String(Date.now()));
  props.setProperty('POP_CHECKED', String(Date.now() - 3600000));
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'popWatch'))
    ScriptApp.newTrigger('popWatch').timeBased().everyMinutes(10).create();
  popWatch();
  try { SpreadsheetApp.getUi().alert('POP check point is on. From now on, ticking Paid in ClickUp starts the POP.'); } catch (e) {}
}

function popWatch() {
  const props = PropertiesService.getScriptProperties();
  const start = Number(props.getProperty('POP_START') || 0);
  if (!props.getProperty('CLICKUP_TOKEN') || !start) return; // popSetup() was not run yet
  const since = Number(props.getProperty('POP_CHECKED') || start) - 15 * 60000, checkedAt = Date.now();
  const startDay = fmt_(new Date(start - 3 * 86400000), 'yyyy-MM-dd'); // payments up to 3 days before setup count
  const rows = rows_(POP_TAB);
  const byId = {};
  rows.forEach(r => byId[str_(r['Task ID'])] = r);
  cuPaidTasks_(since).forEach(t => {
    if (byId[t.id]) return;
    const f = cuFields_(t), pay = f.date(POP_F.payDate);
    const isNew = pay ? pay >= startDay : Number(t.date_created) >= start;
    if (!isNew) return; // paid before the check point started
    if (f.check(POP_F.welcome)) { // welcome already sent by hand
      sheet_(POP_TAB, POP_COLS).appendRow([t.id, cuUnit_(t), POP_SKIP, '', '', '', '', '', '', '', '', '', '', '']);
      return;
    }
    const token = newToken_(), sales = popSales_(f.drop(POP_F.sales));
    sheet_(POP_TAB, POP_COLS).appendRow([t.id, cuUnit_(t), POP_WAIT, token, f.text(POP_F.client), f.text(POP_F.email), sales ? sales.name : '', now_(), '', '', '', '', '', '']);
    popNotify_(t.id, cuUnit_(t), token, sales, false);
    byId[t.id] = 1;
  });
  props.setProperty('POP_CHECKED', String(checkedAt));
  // one reminder a day later
  const now = Date.now();
  rows.filter(r => str_(r['Status']) === POP_WAIT && !str_(r['Reminded at'])).forEach(r => {
    const at = parseTs_(dt_(r['Notified at']));
    if (!at || now - at.getTime() < POP_REMIND_MS) return;
    popNotify_(str_(r['Task ID']), str_(r['Unit']), str_(r['Token']), popSalesByName_(str_(r['Sales'])), true);
    const ref = findRow_(POP_TAB, POP_COLS, 'Token', str_(r['Token']));
    if (ref) setCells_(ref, { 'Reminded at': now_() });
  });
}

function popNotify_(taskId, unit, token, sales, reminder) {
  const link = POP_PAGE + '#t=' + token;
  const head = reminder ? `Reminder: the POP for ${unit} is still waiting.` : `${unit} is marked Paid.`;
  try {
    cuPost_(`/task/${taskId}/comment`, { comment_text: `${head}\nComplete the POP check point (upload the signed quotation, check the data, press Send):\n${link}`, notify_all: true });
  } catch (e) { console.error(e); }
  if (sales && sales.email) {
    // through the Info.ew relay: arrives inside the company at once, with no "external email" banner
    sendClientMail_({
      to: sales.email, subject: `${reminder ? 'Reminder: ' : ''}Complete the POP · ${unit}`.replace(/\|/g, '/'),
      htmlBody: mailHtml_('O-PLUS · POP CHECK POINT', head, [
        `Dear ${sales.name.split(' ')[0]},`,
        'Upload the signed quotation, check the data, then press Send. The POP goes out from your Outlook and the client gets the welcome email with the portal link.'
      ], { label: 'Open the check point', url: link })
    });
  }
}

/* ---------------- API for pop.html (by token) ---------------- */

function popApi_(a, b) {
  popCols_();
  const ref = findRow_(POP_TAB, POP_COLS, 'Token', str_(b.t));
  if (!ref || str_(b.t).length < 20) throw new Error('This link is not valid.');
  if (a === 'pop.get') return popView_(ref);
  if (str_(ref.obj['Status']) !== POP_WAIT) throw new Error('This POP was already sent.');
  if (a === 'pop.upload') return popUpload_(ref, b);
  if (a === 'pop.send') return popSend_(ref, b);
  if (a === 'pop.extra') return popExtra_(ref, b);
  if (a === 'pop.extraDel') return popExtraDel_(ref, b);
  throw new Error('Unknown action.');
}

/** Current data: ClickUp now, plus the quotation and what Sales typed. */
function popData_(ref) {
  const t = cuGet_(`/task/${str_(ref.obj['Task ID'])}`), f = cuFields_(t);
  const sales = popSales_(f.drop(POP_F.sales));
  let q = {}, form = {};
  try { q = JSON.parse(str_(ref.obj['Quote data']) || '{}'); } catch (e) {}
  try { form = JSON.parse(str_(ref.obj['Form']) || '{}'); } catch (e) {}
  const amount = f.num(POP_F.amount);
  const cu = {
    unit: cuUnit_(t), client: f.text(POP_F.client), email: f.text(POP_F.email),
    product: f.drop(POP_F.type), amount: amount ? amount.toFixed(2) : '',
    payDate: f.date(POP_F.payDate, 'dd-MM-yyyy'), am: sales ? sales.name : f.drop(POP_F.sales), pm: f.drop(POP_F.pm), se: 'N/A',
    quoteNo: f.text(POP_F.quote)
  };
  const d = Object.assign({}, cu, q.fields || {}, form);
  return { cu: cu, d: d, q: q, sales: sales, status: str_(ref.obj['Status']), hasFile: !!str_(ref.obj['Quote file ID']) };
}

function popMissing_(x) {
  const d = x.d, m = [];
  if (!x.hasFile) m.push('Signed quotation (PDF)');
  if (!/^[^@\s|]+@[^@\s|]+\.[^@\s|]+$/.test(d.email || '')) m.push('Client email');
  if (!d.client) m.push('Client name');
  if (!d.amount) m.push('Amount');
  if (!d.payDate) m.push('Payment date');
  if (!d.product) m.push('Product');
  if (!d.pm) m.push('PM');
  if (!x.sales || !x.sales.email) m.push('Sales person (Sales Engineer in ClickUp)');
  return m;
}

function popView_(ref) {
  const x = popData_(ref);
  return { status: x.status, data: x.d, fromClickUp: x.cu, quote: x.q, hasFile: x.hasFile, sales: x.sales, missing: popMissing_(x), sentAt: dt_(ref.obj['Sent at']),
    extras: popExtras_(ref).map(e => ({ id: e.id, name: e.name, size: e.size })) };
}

/* Other attachments (proof of payment, signed papers, saved Outlook emails ...): they go with the POP only, never to the client. */
function popCols_() { // adds columns that came after the tab was made
  const sh = sheet_(POP_TAB, POP_COLS), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  POP_COLS.forEach(c => { if (h.indexOf(c) < 0) { h.push(c); sh.getRange(1, h.length).setValue(c).setFontWeight('bold'); } });
}
function popExtras_(ref) { try { return JSON.parse(str_(ref.obj['Extra files']) || '[]'); } catch (e) { return []; } }
function popExtra_(ref, b) {
  const data = str_(b.data), name = str_(b.name).replace(/[\\/|]/g, '-').slice(0, 120) || 'attachment';
  if (!data || data.length > DOC_MAX_B64) throw new Error('Each file can be up to 15 MB.');
  const bytes = Utilities.base64Decode(data), list = popExtras_(ref);
  const quote = str_(ref.obj['Quote file ID']) ? DriveApp.getFileById(str_(ref.obj['Quote file ID'])).getSize() : 0;
  const used = list.reduce((n, e) => n + Number(e.size || 0), quote);
  if (used + bytes.length > POP_EXTRA_MAX) throw new Error('All the attachments together must stay under 18 MB (one email). Remove a file or send a smaller one.');
  const type = str_(b.type) || (/\.msg$/i.test(name) ? 'application/vnd.ms-outlook' : /\.eml$/i.test(name) ? 'message/rfc822' : 'application/octet-stream');
  const unit = str_(ref.obj['Unit']);
  const file = folderPath_([PORTAL_ROOT, safeUnit_(unit), 'POP']).createFile(Utilities.newBlob(bytes, type, name));
  list.push({ id: file.getId(), name: name, size: bytes.length });
  setCells_(ref, { 'Extra files': JSON.stringify(list) }); ref.obj['Extra files'] = JSON.stringify(list);
  return popView_(ref);
}
function popExtraDel_(ref, b) {
  const list = popExtras_(ref), keep = list.filter(e => e.id !== str_(b.id));
  list.filter(e => e.id === str_(b.id)).forEach(e => { try { DriveApp.getFileById(e.id).setTrashed(true); } catch (x) {} });
  setCells_(ref, { 'Extra files': JSON.stringify(keep) }); ref.obj['Extra files'] = JSON.stringify(keep);
  return popView_(ref);
}

function popUpload_(ref, b) {
  const data = str_(b.data);
  if (!data || data.length > DOC_MAX_B64) throw new Error('Upload a PDF up to 15 MB.');
  const unit = str_(ref.obj['Unit']);
  const blob = Utilities.newBlob(Utilities.base64Decode(data), 'application/pdf', str_(b.name) || `${safeUnit_(unit)} quotation.pdf`);
  const file = folderPath_([PORTAL_ROOT, safeUnit_(unit), 'Quotation']).createFile(blob);
  const q = popParse_(popPdfText_(blob));
  setCells_(ref, { 'Quote file ID': file.getId(), 'Quote data': JSON.stringify(q) });
  ref.obj['Quote file ID'] = file.getId(); ref.obj['Quote data'] = JSON.stringify(q);
  return popView_(ref);
}

function popSend_(ref, b) {
  // save what Sales typed (only the editable fields)
  const keep = ['client', 'email', 'product', 'amount', 'payDate', 'se', 'pm', 'note', 'quoteNo', 'total', 'dpPct', 'dpAmount', 'months', 'monthly'];
  const form = {};
  keep.forEach(k => { if (b.form && b.form[k] != null && str_(b.form[k]) !== '') form[k] = str_(b.form[k]).replace(/\|/g, '/'); });
  setCells_(ref, { 'Form': JSON.stringify(form) });
  ref.obj['Form'] = JSON.stringify(form);
  const x = popData_(ref), d = x.d, miss = popMissing_(x);
  if (miss.length) throw new Error('Still missing: ' + miss.join(', '));
  const file = DriveApp.getFileById(str_(ref.obj['Quote file ID']));

  // 1) POP from the Sales person's Outlook. The ready POP goes to the Sales person from Info.ew (inside the
  //    company, so it arrives at once); their Power Automate flow sends it on to finance and deletes it.
  const subject = `${d.unit} ${d.product} POP`.replace(/\|/g, '/');
  const files = [file.getBlob()].concat(popExtras_(ref).map(e => DriveApp.getFileById(e.id).getBlob().setName(e.name)));
  sendClientMail_({ to: x.sales.email, subject: `${POP_TAG} ${subject}`, htmlBody: popHtml_(d, x.sales), attachments: files });

  // 2) client portal + welcome email from Info.ew
  inviteClient_(d.unit, d.client, d.email, d.pm || x.sales.name, true, false, { product: d.product, sales: x.sales.name, pm: d.pm });

  // 3) ClickUp: Welcome Mail ✓ and a note on the task
  try { cuPost_(`/task/${str_(ref.obj['Task ID'])}/field/${POP_F.welcome}`, { value: true }); } catch (e) { console.error(e); }
  try { cuPost_(`/task/${str_(ref.obj['Task ID'])}/comment`, { comment_text: `POP sent by ${x.sales.name}. The client portal welcome email went to ${maskEmail_(d.email)}.`, notify_all: false }); } catch (e) { console.error(e); }

  setCells_(ref, { 'Status': POP_SENT, 'Sent at': now_(), 'Sent by': x.sales.name, 'Client': d.client, 'Client email': d.email });
  log_(d.unit, 'POP sent', str_(d.quoteNo), '', d.product, `Amount ${d.amount} EGP · paid ${d.payDate}`, x.sales.name, '');
  ref.obj['Status'] = POP_SENT; ref.obj['Sent at'] = now_();
  return popView_(ref);
}

/** The POP body, same layout as the email Sales send today. */
function popHtml_(d, sales) {
  const th = 'border:1px solid #000;background:#FBE2D5;padding:2px 8px;text-align:center;font-weight:bold';
  const td = 'border:1px solid #000;padding:2px 8px;text-align:center';
  const cols = [['Unit number', d.unit], ['Product', d.product], ['Amount in EGP', Number(d.amount).toLocaleString('en-US', { minimumFractionDigits: 2 })], ['S.E', d.se || 'N/A'], ['AM', sales.name], ['Payment Date', d.payDate]];
  return `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">
<p>Dears,</p><p>I need your assistance to update the below payment on OWest.</p>
${d.note ? `<p>${esc_(d.note)}</p>` : ''}
<table style="border-collapse:collapse"><tr>${cols.map(c => `<td style="${th}">${esc_(c[0])}</td>`).join('')}</tr>
<tr>${cols.map(c => `<td style="${td}">${esc_(c[1])}</td>`).join('')}</tr></table>
<p>&nbsp;</p><p>Regards,</p><p><b>${esc_(sales.name)}</b><br>Account Manager<br><a href="mailto:${esc_(sales.email)}">${esc_(sales.email)}</a></p></div>`;
}

/* ---------------- quotation PDF → data ---------------- */

/** PDF text through a temporary Google Doc (needs the Drive API service). Empty text if it is not on. */
function popPdfText_(blob) {
  let id = '';
  try {
    try { id = Drive.Files.create({ name: 'pop-read', mimeType: MimeType.GOOGLE_DOCS }, blob, { ocrLanguage: 'en' }).id; }
    catch (e) { id = Drive.Files.insert({ title: 'pop-read', mimeType: MimeType.GOOGLE_DOCS }, blob, { ocr: true, ocrLanguage: 'en' }).id; }
    return DocumentApp.openById(id).getBody().getText();
  } catch (e) { console.error(e); return ''; }
  finally { if (id) try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
}

/** Reads the O-Plus quotation template. Anything it cannot find stays empty for Sales to type. */
function popParse_(text) {
  const t = String(text || '').replace(/\r/g, '');
  const num = s => s ? Number(String(s).replace(/,/g, '')) : null;
  const m = re => { const x = t.match(re); return x ? x[1].trim() : ''; };
  const out = { fields: {}, schedule: [], scope: [] };
  const title = m(/QUOTATION\s*\n+\s*([^\n]+)/i);
  const ref = m(/quotation\s+(\d{3,6})\b/i) || m(/REF\.?\s*NO\.?[\s\S]{0,300}?\b(\d{3,6})\b(?![\/\-])/i);
  const total = num(m(/Total incl\.?\s*VAT[\s\S]{0,160}?(\d{1,3}(?:,\d{3})+(?:\.\d+)?)/i));
  const dpPct = num(m(/DOWN PAYMENT\s*[·\-]?\s*(\d{1,3})\s*%/i));
  const months = num(m(/(\d{1,3})\s+equal monthly/i));
  const sch = t.slice(Math.max(0, t.search(/CHEQUE SCHEDULE/i)));
  const re = /(Down Payment[^\n\d]*(?:\(\d+%\))?|Payment\s+\d{1,3})\s+([\d,]+\.\d{2})\s+(\d{1,2}\/\d{1,2}\/\d{4})/gi;
  let r;
  while ((r = re.exec(sch))) out.schedule.push({ label: r[1].trim(), amount: num(r[2]), date: r[3] });
  const dp = out.schedule.find(s => /down/i.test(s.label));
  const inst = out.schedule.filter(s => !/down/i.test(s.label));
  const sc = t.match(/INCLUDED ITEMS([\s\S]*?)(?:O\s*-\s*P\s*L\s*U\s*S|O-PLUS ·|$)/i);
  if (sc) sc[1].split('\n').map(s => s.replace(/^\s*\d{1,2}\s+/, '').trim()).filter(s => s && !/^(NO\.?\s*SCOPE|NO\.|SCOPE|\d+)$/i.test(s)).forEach(s => out.scope.push(s));
  const f = out.fields;
  if (title) f.title = title;
  if (ref) f.quoteNo = ref;
  if (total) f.total = total.toFixed(2);
  if (dpPct != null) f.dpPct = String(dpPct);
  if (dp) f.dpAmount = dp.amount.toFixed(2);
  else if (total && dpPct != null) f.dpAmount = (total * dpPct / 100).toFixed(2);
  if (months || inst.length) f.months = String(months || inst.length);
  if (inst.length) f.monthly = inst[0].amount.toFixed(2);
  else if (total && f.dpAmount && months) f.monthly = ((total - Number(f.dpAmount)) / months).toFixed(2);
  if (inst.length) f.firstCheque = inst[0].date;
  return out;
}

/* ---------------- ClickUp helpers ---------------- */

function cuGet_(path) {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  const r = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, { headers: { Authorization: token }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('ClickUp ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText());
}
function cuPost_(path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  const r = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), headers: { Authorization: token }, muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('ClickUp ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText() || '{}');
}
/** Tasks on the Main Tracking Sheet changed since the given time (ms) with Paid ticked. */
function cuPaidTasks_(since) {
  const out = [];
  for (let page = 0; page < 20; page++) {
    const res = cuGet_(`/list/${POP_LIST}/task?include_closed=true&page=${page}&date_updated_gt=${Math.floor(since)}&order_by=updated`);
    (res.tasks || []).forEach(t => { if (cuFields_(t).check(POP_F.paid)) out.push(t); });
    if (res.last_page !== false || !(res.tasks || []).length) break;
  }
  return out;
}
function cuFields_(t) {
  const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
  const raw = id => f[id] && f[id].value;
  return {
    text: id => str_(raw(id)),
    num: id => { const x = parseFloat(raw(id)); return isNaN(x) ? 0 : x; },
    check: id => raw(id) === true || raw(id) === 'true',
    drop: id => dropName_(f[id]),
    date: (id, p) => { const x = Number(raw(id)); return x ? fmt_(new Date(x), p || 'yyyy-MM-dd') : ''; }
  };
}
function cuUnit_(t) { return cuFields_(t).text(POP_F.unit) || str_(t.name); }
function popSales_(option) {
  const k = low_(option);
  return POP_SALES[k] || Object.keys(POP_SALES).map(x => POP_SALES[x]).find(s => low_(s.name).indexOf(k) >= 0 && k) || null;
}
function popSalesByName_(name) { return Object.keys(POP_SALES).map(x => POP_SALES[x]).find(s => s.name === name) || null; }

/** Client page: the signed quotation data of a unit (latest POP sent), or null. */
function unitQuote_(unit) {
  const r = rows_(POP_TAB).filter(x => low_(x['Unit']) === low_(unit) && str_(x['Status']) === POP_SENT).pop();
  if (!r) return null;
  let q = {}, form = {};
  try { q = JSON.parse(str_(r['Quote data']) || '{}'); } catch (e) {}
  try { form = JSON.parse(str_(r['Form']) || '{}'); } catch (e) {}
  return { fields: Object.assign({}, q.fields || {}, form), schedule: q.schedule || [], scope: q.scope || [], fileId: str_(r['Quote file ID']) };
}
