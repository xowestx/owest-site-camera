/**
 * O-Plus Client Portal: backend (Google Apps Script bound to its OWN sheet, "O-Plus Client Portal").
 *
 * It is separate from O West Site Camera. It only READS the Site Camera sheet (Units, Engineers,
 * Photos, Reports, ClickUp Tasks) and never writes to it. Everything the portal records lives in
 * this script's own sheet.
 *
 * Pages (GitHub Pages, folder portal/):
 *   portal/#t=<token>   the client's private page for one unit
 *   portal/pm.html      the team page (team password, set from this sheet's menu)
 *
 * Every call is a POST with an "action":
 *   client (token):  portal.open · portal.code · portal.decide · portal.concern · portal.reply · portal.resolve · portal.img
 *   team (pmToken):  login · portal.pm.home · portal.pm.unit · portal.pm.invite · portal.pm.upload · portal.pm.item · portal.pm.reply · portal.img
 *
 * Tabs in this sheet (menu O-Plus Portal > Set up):
 *   Portal Clients   one row per unit that has a portal link
 *   Portal Items     scope items, date changes and selections sent for approval (one row per version)
 *   Portal Concerns  client concerns; Portal Messages holds the conversation
 *   Portal Log       the agreed record: every send, approval, rejection and reply. Protected: only the owner can edit.
 *   Settings         "Team emails": who is told about approvals, rejections and new concerns
 *
 * Deploy: Execute as = Me (the same Google account that owns the Site Camera sheet and photos),
 * Who has access = Anyone. Then put the /exec URL in portal/portal.js (API_URL).
 */

const CAMERA_SHEET_ID = '1kWngNEurLts1djgYRbjmoUFhyCHm6i1vqpx5--D1Ngc';
const TZ = 'Africa/Cairo';
const MAX_BASE64_CHARS = 8 * 1024 * 1024;
const TOKEN_DAYS = 30;

/* =========================== web API =========================== */

function doGet() { return json_({ ok: true, app: 'O-Plus Client Portal' }); }

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'login') return json_(Object.assign({ ok: true }, login_(b)));
    return json_(Object.assign({ ok: true }, portal_(b)));
  } catch (err) { return json_({ ok: false, error: msg_(err), code: err && err.code }); }
}

const PT = { clients: 'Portal Clients', items: 'Portal Items', concerns: 'Portal Concerns', messages: 'Portal Messages', snags: 'Portal Snags', certs: 'Portal Handover', log: 'Portal Log' };
const PORTAL_URL = 'https://xowestx.github.io/owest-site-camera/portal/';
const PORTAL_ASSETS = 'https://xowestx.github.io/owest-site-camera/portal/assets/';
const PORTAL_ROOT = 'O West Client Portal';
const PC_HEADERS = ['Unit', 'Client name', 'Client email', 'Token', 'Invited at', 'Invited by', 'Last opened'];
const PI_HEADERS = ['Item ID', 'Unit', 'Type', 'Title', 'Details', 'Photo IDs', 'Version', 'Replaces', 'Status', 'Sent at', 'Sent by', 'Decided at', 'Decided by', 'Reason'];
const PCN_HEADERS = ['Concern ID', 'Unit', 'Subject', 'Status', 'Opened at', 'Updated at'];
const PMS_HEADERS = ['Concern ID', 'Unit', 'At', 'From', 'Name', 'Text', 'Photo ID'];
const PSN_HEADERS = ['Snag ID', 'Unit', 'No.', 'Location', 'Description', 'Photo IDs', 'Status', 'Reported at', 'Reported by', 'From',
  'Fixed at', 'Fixed by', 'Fix note', 'Fix photo IDs', 'Closed at', 'Closed by', 'Reopen note'];
const PH_HEADERS = ['Cert ID', 'Unit', 'Status', 'Handover date', 'Details', 'Notes', 'Issued at', 'Issued by', 'Signed at', 'Signed by', 'Signed with', 'Decline note', 'PDF ID'];
const PL_HEADERS = ['At', 'Unit', 'Event', 'Ref', 'Version', 'Title', 'Detail', 'By', 'Confirmed with'];
const P_TYPES = ['Selection', 'Scope of work', 'Date change'];
const S_PENDING = 'Sent for approval', S_OK = 'Approved', S_NO = 'Rejected', S_OLD = 'Replaced';
const C_OPEN = 'Under review', C_REPLIED = 'Replied', C_DONE = 'Resolved';
const SN_OPEN = 'Open', SN_FIXED = 'Fixed', SN_CLOSED = 'Closed';
const H_WAIT = 'Sent for signature', H_SIGNED = 'Signed', H_DECLINED = 'Declined';
const SNAG_MAIL_GAP = 1800; // seconds: one email per unit per half hour while a snag walk-through is going on

/* =========================== router =========================== */

function portal_(b) {
  const a = str_(b.action);
  if (a === 'portal.img') return portalImg_(b);
  if (a.indexOf('portal.pm.') === 0) {
    requireAdmin_(b.pmToken);
    if (a === 'portal.pm.home') return pmHome_();
    if (a === 'portal.pm.unit') return pmUnit_(str_(b.unit));
    if (a === 'portal.pm.invite') return pmInvite_(b);
    if (a === 'portal.pm.upload') return pmUpload_(b);
    if (a === 'portal.pm.item') return pmItem_(b);
    if (a === 'portal.pm.reply') return pmReply_(b);
    if (a === 'portal.pm.snag') return pmSnag_(b);
    if (a === 'portal.pm.snagfix') return pmSnagFix_(b);
    if (a === 'portal.pm.dash') return pmDash_();
    if (a === 'portal.pm.cert') return pmCert_(b);
    if (a === 'portal.pm.certpdf') return certPdf_(unitInfo_(str_(b.unit)).unit, b.certId);
  } else {
    const c = clientByToken_(b.token);
    if (a === 'portal.open') { touchClient_(c); return clientView_(c); }
    if (a === 'portal.code') return clientCode_(c, b);
    if (a === 'portal.decide') return clientDecide_(c, b);
    if (a === 'portal.concern') return clientConcern_(c, b);
    if (a === 'portal.reply') return clientReply_(c, b);
    if (a === 'portal.resolve') return clientResolve_(c, b);
    if (a === 'portal.snag') return clientSnag_(c, b);
    if (a === 'portal.snag.close') return clientSnagClose_(c, b);
    if (a === 'portal.snag.reopen') return clientSnagReopen_(c, b);
    if (a === 'portal.cert.sign') return clientCertSign_(c, b);
    if (a === 'portal.cert.decline') return clientCertDecline_(c, b);
    if (a === 'portal.cert.pdf') return certPdf_(c.unit, b.certId);
  }
  throw new Error('Unknown portal action.');
}

/* =========================== client side =========================== */

function clientByToken_(token) {
  const t = str_(token);
  if (t.length < 20) { const e = new Error('This link is not valid. Please ask your project manager for a new one.'); e.code = 'link'; throw e; }
  const r = rows_(PT.clients).find(x => str_(x['Token']) === t);
  if (!r) { const e = new Error('This link is no longer active. Please ask your project manager for a new one.'); e.code = 'link'; throw e; }
  return { unit: str_(r['Unit']), name: str_(r['Client name']), email: str_(r['Client email']), pm: str_(r['Invited by']), token: t };
}

function touchClient_(c) {
  const cache = CacheService.getScriptCache(), k = 'pseen:' + c.token;
  if (cache.get(k)) return;
  cache.put(k, '1', 3600);
  const ref = findRow_(PT.clients, PC_HEADERS, 'Token', c.token);
  if (ref) setCells_(ref, { 'Last opened': now_() });
}

function clientView_(c) {
  const u = unitInfo_(c.unit);
  return {
    unit: c.unit, project: u.project, client: c.name || u.client, pm: c.pm,
    email: maskEmail_(c.email), status: unitStatus_(c.unit),
    items: unitItems_(c.unit).filter(x => x.status !== S_OLD),
    concerns: unitConcerns_(c.unit), snags: unitSnags_(c.unit), certs: unitCerts_(c.unit),
    history: unitLog_(c.unit)
  };
}

function clientCode_(c, b) {
  const item = b.certId ? openCert_(c.unit, b.certId) : openItem_(c.unit, b.itemId);
  if (!c.email) throw new Error('We have no email for you yet. Please ask your project manager to add it.');
  const cache = CacheService.getScriptCache();
  const sent = Number(cache.get('pcodes:' + c.token) || 0);
  if (sent >= 6) throw new Error('Too many codes requested. Please wait an hour and try again.');
  cache.put('pcodes:' + c.token, String(sent + 1), 3600);
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('pcode:' + c.token, JSON.stringify({ code: code, itemId: item.id, tries: 0 }), 600);
  const verb = b.certId ? 'sign' : str_(b.decision) === 'reject' ? 'request changes to' : 'approve';
  MailApp.sendEmail({
    to: c.email, name: 'O West Extra Works',
    subject: `Your confirmation code: ${code} · ${c.unit}`,
    htmlBody: mailHtml_('CONFIRMATION CODE', `Your code is ${code}`,
      [`You asked to ${verb} "${item.title}" for unit ${c.unit}.`, `Enter this code in the portal within 10 minutes. If this was not you, please ignore this email.`])
  });
  return { sentTo: maskEmail_(c.email) };
}

function clientDecide_(c, b) {
  const decision = str_(b.decision);
  if (decision !== 'approve' && decision !== 'reject') throw new Error('Choose approve or request changes.');
  const reason = str_(b.reason).slice(0, 2000);
  if (decision === 'reject' && !reason) throw new Error('Please add a reason.');
  checkCode_(c, b.itemId, b.code);
  withLock_(() => {
    const item = openItem_(c.unit, b.itemId);
    const ref = findRow_(PT.items, PI_HEADERS, 'Item ID', item.id);
    const status = decision === 'approve' ? S_OK : S_NO;
    setCells_(ref, { 'Status': status, 'Decided at': now_(), 'Decided by': c.name || 'Client', 'Reason': reason });
    log_(c.unit, status, item.id, item.version, item.title, reason, c.name || 'Client', 'Email code · ' + maskEmail_(c.email));
    notifyTeam_(`${c.unit} · ${status}: ${item.title}`, `${status.toUpperCase()} BY CLIENT`, `${item.title}`,
      [`${c.name || 'The client'} ${decision === 'approve' ? 'approved' : 'rejected'} version ${item.version} of "${item.title}".`].concat(reason ? ['Reason: ' + reason] : []));
  });
  return clientView_(c);
}

function clientConcern_(c, b) {
  const subject = str_(b.subject).slice(0, 200), text = str_(b.text).slice(0, 4000);
  if (!subject && !text) throw new Error('Please write your concern.');
  const photo = b.photo ? savePortalPhoto_(c.unit, b.photo, 'concern') : '';
  withLock_(() => {
    const id = 'C' + fmt_(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    sheet_(PT.concerns, PCN_HEADERS).appendRow([id, c.unit, subject || text.slice(0, 60), C_OPEN, now_(), now_()]);
    sheet_(PT.messages, PMS_HEADERS).appendRow([id, c.unit, now_(), 'client', c.name || 'Client', text || subject, photo]);
    log_(c.unit, 'New concern', id, '', subject || text.slice(0, 60), text, c.name || 'Client', '');
    notifyTeam_(`${c.unit} · New concern: ${subject || 'from the client'}`, 'NEW CONCERN', subject || 'Concern from the client', [text]);
  });
  return clientView_(c);
}

function clientReply_(c, b) {
  const text = str_(b.text).slice(0, 4000);
  if (!text) throw new Error('Please write a reply.');
  withLock_(() => {
    const ref = concernRef_(c.unit, b.concernId);
    if (str_(ref.obj['Status']) === C_DONE) throw new Error('This concern is resolved. Please open a new one.');
    sheet_(PT.messages, PMS_HEADERS).appendRow([ref.obj['Concern ID'], c.unit, now_(), 'client', c.name || 'Client', text, '']);
    setCells_(ref, { 'Status': C_OPEN, 'Updated at': now_() });
    notifyTeam_(`${c.unit} · Client replied: ${ref.obj['Subject']}`, 'CLIENT REPLY', str_(ref.obj['Subject']), [text]);
  });
  return clientView_(c);
}

function clientResolve_(c, b) {
  withLock_(() => {
    const ref = concernRef_(c.unit, b.concernId);
    setCells_(ref, { 'Status': C_DONE, 'Updated at': now_() });
    log_(c.unit, 'Concern resolved', ref.obj['Concern ID'], '', str_(ref.obj['Subject']), '', c.name || 'Client', '');
  });
  return clientView_(c);
}

/* =========================== snag list (handover) =========================== */
/* Open -> Fixed (team, with after photos) -> Closed (client confirms). "Still not fixed" sends it back to Open. */

function addSnag_(unit, b, by, from, photoIds) {
  const location = str_(b.location).slice(0, 120), text = str_(b.text).slice(0, 2000);
  if (!text) throw new Error('Please describe the snag.');
  let no = 0;
  withLock_(() => {
    no = rows_(PT.snags).filter(r => low_(r['Unit']) === unit.toLowerCase()).reduce((m, r) => Math.max(m, Number(r['No.'] || 0)), 0) + 1;
    const id = 'S' + fmt_(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    sheet_(PT.snags, PSN_HEADERS).appendRow([id, unit, no, location, text, photoIds.join(','), SN_OPEN, now_(), by, from, '', '', '', '', '', '', '']);
    log_(unit, 'Snag reported', id, '', `#${no}${location ? ' · ' + location : ''}`, text, by, '');
  });
  return no;
}

/** One email per unit per half hour, so a walk-through of 30 snags is not 30 emails. */
function snagMailDue_(key) {
  const cache = CacheService.getScriptCache();
  if (cache.get(key)) return false;
  cache.put(key, '1', SNAG_MAIL_GAP);
  return true;
}

function clientSnag_(c, b) {
  const photos = (b.photos || []).slice(0, 4).map(p => savePortalPhoto_(c.unit, p, 'snag'));
  const no = addSnag_(c.unit, b, c.name || 'Client', 'client', photos);
  if (snagMailDue_('psnagmail:' + c.unit))
    notifyTeam_(`${c.unit} · Client is adding snags`, 'SNAG LIST', `${c.name || 'The client'} reported snag #${no}`,
      [`${str_(b.location) ? str_(b.location) + ': ' : ''}${str_(b.text)}`, 'More may follow. Open the unit in the team portal to see the full list.']);
  return clientView_(c);
}

function clientSnagClose_(c, b) {
  withLock_(() => {
    const ref = snagRef_(c.unit, b.snagId);
    if (str_(ref.obj['Status']) !== SN_FIXED) throw new Error('This snag is not marked as fixed yet. Please reload.');
    setCells_(ref, { 'Status': SN_CLOSED, 'Closed at': now_(), 'Closed by': c.name || 'Client' });
    log_(c.unit, 'Snag closed', ref.obj['Snag ID'], '', `#${ref.obj['No.']}${str_(ref.obj['Location']) ? ' · ' + str_(ref.obj['Location']) : ''}`, 'Client confirmed the fix', c.name || 'Client', '');
  });
  return clientView_(c);
}

function clientSnagReopen_(c, b) {
  const note = str_(b.note).slice(0, 2000);
  if (!note) throw new Error('Please tell us what is still wrong.');
  let title = '';
  withLock_(() => {
    const ref = snagRef_(c.unit, b.snagId);
    if (str_(ref.obj['Status']) !== SN_FIXED) throw new Error('This snag is not marked as fixed yet. Please reload.');
    title = `#${ref.obj['No.']}${str_(ref.obj['Location']) ? ' · ' + str_(ref.obj['Location']) : ''}`;
    setCells_(ref, { 'Status': SN_OPEN, 'Reopen note': note });
    log_(c.unit, 'Snag reopened', ref.obj['Snag ID'], '', title, note, c.name || 'Client', '');
  });
  notifyTeam_(`${c.unit} · Snag ${title} is still not fixed`, 'SNAG REOPENED', title, [`${c.name || 'The client'}: ${note}`]);
  return clientView_(c);
}

function pmSnag_(b) {
  const unit = unitInfo_(str_(b.unit)).unit;
  addSnag_(unit, b, str_(b.by) || 'O West team', 'team', (b.photoIds || []).map(str_).filter(Boolean).slice(0, 4));
  return pmUnit_(unit);
}

function pmSnagFix_(b) {
  const unit = unitInfo_(str_(b.unit)).unit, by = str_(b.by) || 'O West team';
  const note = str_(b.note).slice(0, 2000), photos = (b.photoIds || []).map(str_).filter(Boolean).slice(0, 4);
  if (!photos.length) throw new Error('Add at least one photo of the fix.');
  withLock_(() => {
    const ref = snagRef_(unit, b.snagId);
    if (str_(ref.obj['Status']) !== SN_OPEN) throw new Error('This snag is not open. Please reload.');
    setCells_(ref, { 'Status': SN_FIXED, 'Fixed at': now_(), 'Fixed by': by, 'Fix note': note, 'Fix photo IDs': photos.join(',') });
    log_(unit, 'Snag fixed', ref.obj['Snag ID'], '', `#${ref.obj['No.']}${str_(ref.obj['Location']) ? ' · ' + str_(ref.obj['Location']) : ''}`, note, by, '');
  });
  const c = rows_(PT.clients).find(r => low_(r['Unit']) === unit.toLowerCase());
  if (c && str_(c['Client email']) && snagMailDue_('psnagfixmail:' + unit)) {
    MailApp.sendEmail({
      to: str_(c['Client email']), name: 'O West Extra Works',
      subject: `Snags fixed, please check · ${unit}`,
      htmlBody: mailHtml_('SNAG LIST', 'Fixed and ready for your check', [
        `Dear ${str_(c['Client name']) || 'Client'},`,
        `${by} has fixed snags in unit ${unit} and added photos of each fix. Please check them in your portal and confirm, or tell us what is still not right.`
      ], { label: 'Check the snag list', url: PORTAL_URL + '#t=' + str_(c['Token']) })
    });
  }
  return pmUnit_(unit);
}

/* =========================== handover certificate =========================== */
/* The PM issues it once every snag is closed. The client signs it with the emailed code, or
   declines with a note. On signing, a branded PDF is saved to the unit's Drive folder and emailed. */

function unitCerts_(unit) {
  const k = unit.toLowerCase();
  return rows_(PT.certs).filter(r => low_(r['Unit']) === k).map(r => ({
    id: str_(r['Cert ID']), status: str_(r['Status']), date: day_(r['Handover date']), details: str_(r['Details']), notes: str_(r['Notes']),
    issuedAt: dt_(r['Issued at']), issuedBy: str_(r['Issued by']), signedAt: dt_(r['Signed at']), signedBy: str_(r['Signed by']),
    signedWith: str_(r['Signed with']), declineNote: str_(r['Decline note']), pdf: !!str_(r['PDF ID'])
  })).reverse();
}

function openCert_(unit, id) {
  const ref = findRow_(PT.certs, PH_HEADERS, 'Cert ID', str_(id));
  if (!ref || low_(ref.obj['Unit']) !== unit.toLowerCase()) throw new Error('This certificate was not found.');
  if (str_(ref.obj['Status']) !== H_WAIT) throw new Error('This certificate was already signed or replaced. Please reload.');
  return { id: str_(ref.obj['Cert ID']), title: 'Handover certificate', ref: ref };
}

function pmCert_(b) {
  const u = unitInfo_(str_(b.unit)), by = str_(b.by) || 'O West team';
  const date = str_(b.date).slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Choose the handover date.');
  const open = unitSnags_(u.unit).filter(x => x.status !== SN_CLOSED).length;
  if (open) throw new Error(`Close all snags first: ${open} still open or waiting for the client's check.`);
  const c = rows_(PT.clients).find(r => low_(r['Unit']) === u.unit.toLowerCase());
  if (!c) throw new Error('Send the portal invitation first.');
  withLock_(() => {
    rows_(PT.certs).forEach(r => { // one certificate waits at a time, and a signed one is final
      if (low_(r['Unit']) !== u.unit.toLowerCase()) return;
      if (str_(r['Status']) === H_WAIT) throw new Error('A certificate is already waiting for the client.');
      if (str_(r['Status']) === H_SIGNED) throw new Error('This unit is already handed over.');
    });
    const id = 'H' + fmt_(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    sheet_(PT.certs, PH_HEADERS).appendRow([id, u.unit, H_WAIT, date, str_(b.details).slice(0, 3000), str_(b.notes).slice(0, 3000), now_(), by, '', '', '', '', '']);
    log_(u.unit, 'Handover certificate sent', id, '', 'Handover ' + date, '', by, '');
  });
  if (str_(c['Client email'])) {
    MailApp.sendEmail({
      to: str_(c['Client email']), name: 'O West Extra Works',
      subject: `Your handover certificate is ready to sign · ${u.unit}`,
      htmlBody: mailHtml_('HANDOVER', 'Your unit is ready for handover', [
        `Dear ${str_(c['Client name']) || 'Client'},`,
        `All works and snags for ${u.unit} are complete. Please review the handover certificate in your portal and sign it. We will email you a 6-digit code to confirm.`
      ], { label: 'Review and sign', url: PORTAL_URL + '#t=' + str_(c['Token']) })
    });
  }
  return pmUnit_(u.unit);
}

function clientCertSign_(c, b) {
  checkCode_(c, b.certId, b.code);
  let cert;
  withLock_(() => {
    cert = openCert_(c.unit, b.certId);
    setCells_(cert.ref, { 'Status': H_SIGNED, 'Signed at': now_(), 'Signed by': c.name || 'Client', 'Signed with': 'Email code · ' + maskEmail_(c.email) });
    log_(c.unit, 'Handover certificate signed', cert.id, '', 'Handover ' + day_(cert.ref.obj['Handover date']), '', c.name || 'Client', 'Email code · ' + maskEmail_(c.email));
  });
  let pdf = null;
  try {
    pdf = makeCertPdf_(c.unit, cert.id);
    const attach = [pdf.getBlob()];
    MailApp.sendEmail({ to: c.email, name: 'O West Extra Works', subject: `Handover certificate · ${c.unit}`, attachments: attach,
      htmlBody: mailHtml_('HANDOVER · SIGNED', 'Welcome to your new home', [`Dear ${c.name || 'Client'},`, `Thank you. Your signed handover certificate for ${c.unit} is attached. It is also saved in your portal.`]) });
    const to = settings_()['Team emails'];
    if (to) MailApp.sendEmail({ to: to, name: 'O-Plus Client Portal', subject: `${c.unit} · Handover certificate signed`, attachments: attach,
      htmlBody: mailHtml_('HANDOVER · SIGNED', c.unit, [`${c.name || 'The client'} signed the handover certificate.`], { label: 'Open the team portal', url: PORTAL_URL + 'pm.html' }) });
  } catch (e) { console.error('certificate pdf: ' + msg_(e)); }
  return clientView_(c);
}

function clientCertDecline_(c, b) {
  const note = str_(b.note).slice(0, 2000);
  if (!note) throw new Error('Please tell us what is missing.');
  withLock_(() => {
    const cert = openCert_(c.unit, b.certId);
    setCells_(cert.ref, { 'Status': H_DECLINED, 'Decline note': note });
    log_(c.unit, 'Handover certificate declined', cert.id, '', '', note, c.name || 'Client', '');
  });
  notifyTeam_(`${c.unit} · Client did not sign the handover`, 'HANDOVER NOT SIGNED', c.unit, [`${c.name || 'The client'}: ${note}`, 'Fix what is missing, then issue a new certificate.']);
  return clientView_(c);
}

/** Same code rules as approvals: 10 minutes, 5 wrong tries. */
function checkCode_(c, id, code) {
  const cache = CacheService.getScriptCache(), key = 'pcode:' + c.token;
  const st = JSON.parse(cache.get(key) || 'null');
  if (!st || st.itemId !== str_(id)) throw new Error('The code expired. Please request a new one.');
  if (str_(code) !== st.code) {
    st.tries++;
    if (st.tries >= 5) { cache.remove(key); throw new Error('Too many wrong codes. Please request a new one.'); }
    cache.put(key, JSON.stringify(st), 600);
    throw new Error('That code is not correct.');
  }
  cache.remove(key);
}

/** The signed PDF, as a data URL, for the unit's client or the team. */
function certPdf_(unit, id) {
  const ref = findRow_(PT.certs, PH_HEADERS, 'Cert ID', str_(id));
  if (!ref || low_(ref.obj['Unit']) !== unit.toLowerCase() || str_(ref.obj['Status']) !== H_SIGNED) throw new Error('This certificate is not signed yet.');
  const f = str_(ref.obj['PDF ID']) ? DriveApp.getFileById(str_(ref.obj['PDF ID'])) : makeCertPdf_(unit, id);
  return { name: `Handover certificate ${safeUnit_(unit)}.pdf`, data: 'data:application/pdf;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
}

function makeCertPdf_(unit, id) {
  const ref = findRow_(PT.certs, PH_HEADERS, 'Cert ID', str_(id));
  const r = ref.obj, u = unitInfo_(unit), st = unitStatus_(u.unit) || {};
  const client = rows_(PT.clients).find(x => low_(x['Unit']) === unit.toLowerCase()) || {};
  const items = unitItems_(u.unit).filter(x => x.status === S_OK);
  const snags = unitSnags_(u.unit);
  const kv = str_(r['Details']).split(/\n+/).map(l => l.trim()).filter(l => /^[^:]{1,40}:\s*\S/.test(l))
    .map(l => { const i = l.indexOf(':'); return `<tr><td class="k">${esc_(l.slice(0, i))}</td><td>${esc_(l.slice(i + 1).trim())}</td></tr>`; }).join('');
  const day = d => { const m = String(d || '').match(/^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?/); if (!m) return esc_(d); const M = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m[2] - 1]; return `${+m[3]} ${M} ${m[1]}` + (m[4] ? ` · ${m[4]}:${m[5]}` : ''); };
  const logo = name => { try { return 'data:image/png;base64,' + Utilities.base64Encode(UrlFetchApp.fetch(PORTAL_ASSETS + name).getContent()); } catch (e) { return ''; } };
  const mark = logo('oplus-mark-white.png'), lockup = logo('owest-lockup-white.png');
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    @page{size:A4;margin:0} body{margin:0;font-family:Poppins,Arial,sans-serif;color:#1A2238;font-size:11px}
    .band{background:#001641;padding:18px 36px;height:44px} .band img{height:44px} .band .r{float:right;height:20px;margin-top:12px}
    .wrap{padding:26px 36px} .eb{font-size:8.5px;font-weight:700;letter-spacing:2px;color:#001641}
    h1{font-size:24px;color:#001641;margin:4px 0 0} .bar{width:36px;height:3px;background:#F0B323;margin:8px 0 18px}
    table{width:100%;border-collapse:collapse} td{font-size:11.5px;padding:7px 0;border-bottom:1px solid #DCDFE6;vertical-align:top} td.k{width:34%;font-size:8.5px;font-weight:700;letter-spacing:1.5px;color:#5F677A;text-transform:uppercase;padding-top:9px}
    h2{font-size:9px;letter-spacing:2px;color:#001641;margin:20px 0 6px} ul{margin:0;padding-left:16px} li{margin:2px 0}
    .sign{margin-top:22px;border-left:3px solid #2F6E4E;background:#F3F5F9;padding:12px 16px} .sign b{color:#001641}
    .foot{position:fixed;bottom:0;left:0;right:0;border-top:1px solid #DCDFE6;padding:8px 36px;font-size:8.5px;color:#5F677A}
  </style></head><body>
  <div class="band">${mark ? `<img src="${mark}">` : '<b style="color:#fff">O-PLUS</b>'}${lockup ? `<img class="r" src="${lockup}">` : ''}</div>
  <div class="wrap">
    <div class="eb">O-PLUS · EXTRA WORKS · ${esc_(u.unit)}</div><h1>Handover Certificate</h1><div class="bar"></div>
    <table>
      <tr><td class="k">Unit</td><td>${esc_(u.unit)}${u.project ? ' · ' + esc_(u.project) : ''}</td></tr>
      <tr><td class="k">Client</td><td>${esc_(str_(client['Client name']) || u.client)}</td></tr>
      <tr><td class="k">Handover date</td><td>${day(day_(r['Handover date']))}</td></tr>
      <tr><td class="k">Project manager</td><td>${esc_(st.pm || str_(r['Issued by']))}</td></tr>
      ${kv}
    </table>
    <h2>APPROVED WORKS AND SELECTIONS</h2>
    ${items.length ? `<ul>${items.map(x => `<li>${esc_(x.title)} <span style="color:#5F677A">· ${esc_(x.type)} · v${x.version} · approved ${day(x.decidedAt)}</span></li>`).join('')}</ul>` : '<p style="color:#5F677A">As agreed in the contract.</p>'}
    <h2>SNAG LIST</h2>
    <p>${snags.length ? `${snags.length} snag${snags.length > 1 ? 's' : ''} recorded, all fixed and confirmed by the client${snags[0] && snags[0].closedAt ? ' (last on ' + day(snags.map(x => x.closedAt).sort().pop()) + ')' : ''}.` : 'No snags recorded.'}</p>
    ${str_(r['Notes']) ? `<h2>NOTES</h2><p style="white-space:pre-line">${esc_(r['Notes'])}</p>` : ''}
    <div class="sign"><b>Accepted by ${esc_(r['Signed by'])} · ${day(dt_(r['Signed at']))}</b><br>
      The client confirmed receipt of the unit and the works above. Confirmed with a one-time code sent to ${esc_(String(r['Signed with']).replace('Email code · ', ''))}.<br>
      <span style="color:#5F677A">Issued by ${esc_(r['Issued by'])} · ${day(dt_(r['Issued at']))} · Record ${esc_(r['Cert ID'])}</span></div>
  </div>
  <div class="foot"><b style="color:#001641">O-PLUS</b> · O West by Orascom Development · This certificate is part of the unit's record in the O-Plus client portal.</div>
  </body></html>`;
  const blob = Utilities.newBlob(html, 'text/html', 'cert.html').getAs('application/pdf').setName(`Handover certificate ${safeUnit_(unit)}.pdf`);
  const file = folderPath_([PORTAL_ROOT, safeUnit_(unit)]).createFile(blob);
  setCells_(ref, { 'PDF ID': file.getId() });
  return file;
}

/* =========================== management dashboard =========================== */
/* ClickUp is read by an hourly job into the "Portal Dashboard" tab (one row per task), so the
   dashboard opens fast and never calls ClickUp while someone waits. Amounts are read only to
   weight progress across a unit's tasks; no money is shown. */

const DASH_TAB = 'Portal Dashboard';
const DASH_COLS = ['Unit', 'Task ID', 'Status', 'Construction status', 'POC', 'Site in', 'Site due', 'PM', 'Weight', 'Checked'];
const DASH_FRESH_MS = 3 * 3600 * 1000;
const DASH_DONE = ['handed over', 'declined', 'deleted', 'closed', 'cancelled'];

/** Hourly (trigger made by Set up) and from the menu. Re-reads tasks older than 3 hours, for at most 4.5 minutes. */
function dashRefresh() {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (!token) return;
  const started = Date.now(), sh = sheet_(DASH_TAB, DASH_COLS);
  const byId = {};
  rows_(DASH_TAB).forEach(r => { const id = str_(r['Task ID']); if (id) byId[id] = DASH_COLS.map(c => r[c]); });
  const want = {};
  camRows_('ClickUp Tasks').forEach(r => { const id = str_(r['Task ID']), u = normUnit_(r['Unit']); if (id && u) want[id] = u; });
  Object.keys(byId).forEach(id => { if (!want[id]) delete byId[id]; });
  // a handed-over task changes rarely: re-read it once a day
  const age = id => !byId[id] ? Infinity : started - Number(byId[id][9] || 0);
  const due = id => age(id) > (byId[id] && DASH_DONE.indexOf(low_(byId[id][2])) >= 0 ? 8 * DASH_FRESH_MS : DASH_FRESH_MS);
  const ids = Object.keys(want).filter(due).sort((a, b) => age(b) - age(a));
  for (let i = 0; i < ids.length && Date.now() - started < 270000; i += 25) {
    const batch = ids.slice(i, i + 25);
    let limited = false;
    UrlFetchApp.fetchAll(batch.map(id => ({ url: `https://api.clickup.com/api/v2/task/${id}`, headers: { Authorization: token }, muteHttpExceptions: true }))).forEach((r, n) => {
      const id = batch[n], code = r.getResponseCode();
      if (code === 404) { byId[id] = [want[id], id, 'deleted', '', '', '', '', '', 0, Date.now()]; return; }
      if (code === 429) { limited = true; return; }
      if (code !== 200) return;
      let t; try { t = JSON.parse(r.getContentText()); } catch (e) { return; }
      const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
      const v = k => str_(f[k] && f[k].value), num = k => { const x = parseFloat(v(k)); return isNaN(x) ? '' : x; };
      const d = k => { const x = Number(v(k)); return x ? fmt_(new Date(x), 'yyyy-MM-dd') : ''; };
      byId[id] = [want[id], id, low_(t.status && t.status.status), low_(dropName_(f[P_CU.cstatus])), num(P_CU.poc),
        d(P_CU.siteIn), d(P_CU.siteDue), dropName_(f[P_CU.pm]), num(P_CU.amount) || 0, Date.now()];
    });
    if (limited) Utilities.sleep(20000);
  }
  const out = Object.keys(byId).map(id => byId[id]).sort((a, b) => (a[0] + a[1]).localeCompare(b[0] + b[1]));
  withLock_(() => {
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, DASH_COLS.length).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, DASH_COLS.length).setValues(out);
  });
  CacheService.getScriptCache().remove('pdash');
}

/** Planned % on a straight line from Site in to Site due. */
function plannedPoc_(start, end, today) {
  if (!start || !end || end <= start) return null;
  const t = d => new Date(d + 'T00:00:00').getTime();
  return Math.max(0, Math.min(100, Math.round((t(today) - t(start)) / (t(end) - t(start)) * 100)));
}

function pmDash_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('pdash');
  if (hit) return JSON.parse(hit);
  const today = fmt_(new Date(), 'yyyy-MM-dd'), dayMs = 86400000;
  const back = n => fmt_(new Date(Date.now() - n * dayMs), 'yyyy-MM-dd');
  const d7 = back(6), d14 = back(13);
  const tasks = groupBy_(rows_(DASH_TAB).filter(r => str_(r['Unit']) && DASH_DONE.slice(1).indexOf(low_(r['Status'])) < 0), r => str_(r['Unit']).toUpperCase());
  const lastSync = rows_(DASH_TAB).reduce((m, r) => Math.max(m, Number(r['Checked']) || 0), 0);

  // Site Camera (read only): last visit per unit, manpower per unit / contractor / day
  const lastVisit = {}, visits7 = {};
  camRows_('Visits').forEach(r => {
    const u = str_(r['Unit']).toUpperCase(), d = day_(r['Date']);
    if (!u || !d || low_(r['Status']) === 'deleted') return;
    if (!lastVisit[u] || d > lastVisit[u]) lastVisit[u] = d;
    if (d >= d7) visits7[u] = (visits7[u] || 0) + 1;
  });
  const mpUnit = {}, mpCont = {}, mpDay = {};
  camRows_('Manpower').forEach(r => {
    const d = day_(r['Date']), n = Number(r['Count']) || 0;
    if (!d || d < d14 || !n || low_(r['Status']) === 'deleted') return;
    mpDay[d] = (mpDay[d] || 0) + n;
    if (d < d7) return;
    const u = str_(r['Unit']).toUpperCase(), c = str_(r['Contractor']) || 'Not named';
    mpUnit[u] = (mpUnit[u] || 0) + n; mpCont[c] = (mpCont[c] || 0) + n;
  });

  // Portal: approvals, concerns, snags per unit
  const cnt = (rows, col, val) => { const o = {}; rows.forEach(r => { if (str_(r[col]) === val) { const u = str_(r['Unit']).toUpperCase(); o[u] = (o[u] || 0) + 1; } }); return o; };
  const pend = cnt(rows_(PT.items), 'Status', S_PENDING), rej = cnt(rows_(PT.items), 'Status', S_NO);
  const conc = cnt(rows_(PT.concerns), 'Status', C_OPEN), snagO = cnt(rows_(PT.snags), 'Status', SN_OPEN);

  const stageCount = [0, 0, 0, 0, 0, 0], units = [];
  Object.keys(tasks).forEach(u => {
    const live = tasks[u].map(r => ({ status: low_(r['Status']), cs: low_(r['Construction status']), poc: r['POC'] === '' ? null : Number(r['POC']),
      start: day_(r['Site in']), end: day_(r['Site due']), pm: str_(r['PM']), w: Number(r['Weight']) || 0 }));
    const site = live.filter(t => t.status === 'site execution');
    let stage;
    if (site.length) stage = site.every(t => t.cs === 'completed' || t.cs === 'handedover') ? 4 : 3;
    else if (live.every(t => t.status === 'handed over')) stage = 5;
    else stage = Math.max.apply(null, live.map(t => P_CU_STAGE[t.status] == null ? 0 : P_CU_STAGE[t.status]));
    stageCount[stage]++;
    if (stage < 3 || stage > 4) return; // the table is for units on site
    const withPoc = site.filter(t => t.poc != null), w = withPoc.reduce((a, t) => a + t.w, 0);
    const poc = !withPoc.length ? null : Math.round(w > 0 ? withPoc.reduce((a, t) => a + t.poc * t.w, 0) / w : withPoc.reduce((a, t) => a + t.poc, 0) / withPoc.length);
    const start = site.map(t => t.start).filter(Boolean).sort()[0] || '', end = site.map(t => t.end).filter(Boolean).sort().pop() || '';
    const planned = plannedPoc_(start, end, today);
    const late = end && end < today && stage === 3 ? Math.round((new Date(today) - new Date(end)) / dayMs) : 0;
    units.push({
      unit: u, stage: stage, poc: poc, planned: planned, gap: poc != null && planned != null ? poc - planned : null,
      start: start, end: end, late: late, pm: site.map(t => t.pm).filter((x, i, a) => x && a.indexOf(x) === i).join(', '),
      lastVisit: lastVisit[u] || '', visits7: visits7[u] || 0, manpower7: mpUnit[u] || 0,
      pending: pend[u] || 0, rejected: rej[u] || 0, concerns: conc[u] || 0, snags: snagO[u] || 0
    });
  });
  units.sort((a, b) => (b.late - a.late) || ((a.gap == null ? 0 : a.gap) - (b.gap == null ? 0 : b.gap)) || a.unit.localeCompare(b.unit));

  const days = []; for (let i = 13; i >= 0; i--) { const d = back(i); days.push({ date: d, count: mpDay[d] || 0 }); }
  const pms = groupBy_(units, x => x.pm || 'No PM');
  const out = {
    today: today, synced: lastSync ? fmt_(new Date(lastSync), 'yyyy-MM-dd HH:mm') : '', stages: stageCount, units: units,
    totals: {
      onSite: units.length, late: units.filter(x => x.late > 0).length, behind: units.filter(x => x.gap != null && x.gap <= -15).length,
      noVisit: units.filter(x => !x.lastVisit || x.lastVisit < d7).length, manpowerToday: mpDay[today] || 0,
      manpower7: days.slice(7).reduce((a, d) => a + d.count, 0), pending: Object.keys(pend).reduce((a, k) => a + pend[k], 0),
      concerns: Object.keys(conc).reduce((a, k) => a + conc[k], 0), snags: Object.keys(snagO).reduce((a, k) => a + snagO[k], 0)
    },
    manpowerDays: days,
    contractors: Object.keys(mpCont).map(c => ({ name: c, count: mpCont[c] })).sort((a, b) => b.count - a.count).slice(0, 12),
    pms: Object.keys(pms).map(p => ({ name: p, units: pms[p].length, late: pms[p].filter(x => x.late > 0).length,
      behind: pms[p].filter(x => x.gap != null && x.gap <= -15).length, open: pms[p].reduce((a, x) => a + x.concerns + x.snags + x.rejected, 0) }))
      .sort((a, b) => b.units - a.units)
  };
  try { cache.put('pdash', JSON.stringify(out), 600); } catch (e) {}
  return out;
}

/* =========================== team side =========================== */

function pmHome_() {
  const clients = {};
  rows_(PT.clients).forEach(r => clients[low_(r['Unit'])] = r);
  const items = rows_(PT.items), concerns = rows_(PT.concerns), snags = rows_(PT.snags);
  const units = camRows_('Units').filter(r => str_(r['Unit']) && low_(r['Active']) !== 'no').map(r => {
    const u = str_(r['Unit']), k = u.toLowerCase();
    const its = items.filter(x => low_(x['Unit']) === k), cns = concerns.filter(x => low_(x['Unit']) === k);
    const sns = snags.filter(x => low_(x['Unit']) === k);
    return {
      unit: u, project: str_(r['Project']), client: str_(r['Client name']), invited: !!clients[k],
      pending: its.filter(x => str_(x['Status']) === S_PENDING).length,
      rejected: its.filter(x => str_(x['Status']) === S_NO).length,
      open: cns.filter(x => str_(x['Status']) === C_OPEN).length,
      snags: sns.filter(x => str_(x['Status']) === SN_OPEN).length,
      snagsToCheck: sns.filter(x => str_(x['Status']) === SN_FIXED).length
    };
  });
  const actions = [];
  items.filter(x => str_(x['Status']) === S_NO).forEach(x => actions.push({
    kind: 'rejected', unit: str_(x['Unit']), id: str_(x['Item ID']), title: str_(x['Title']), at: dt_(x['Decided at']), detail: str_(x['Reason'])
  }));
  concerns.filter(x => str_(x['Status']) === C_OPEN).forEach(x => actions.push({
    kind: 'concern', unit: str_(x['Unit']), id: str_(x['Concern ID']), title: str_(x['Subject']), at: dt_(x['Updated at']), detail: ''
  }));
  const openSnags = groupBy_(snags.filter(x => str_(x['Status']) === SN_OPEN), x => str_(x['Unit']));
  Object.keys(openSnags).forEach(u => {
    const l = openSnags[u], oldest = l.map(x => dt_(x['Reported at'])).sort()[0];
    actions.push({ kind: 'snags', unit: u, id: '', title: `${l.length} open snag${l.length > 1 ? 's' : ''}`, at: oldest,
      detail: l.slice(0, 3).map(x => str_(x['Location']) || str_(x['Description']).slice(0, 30)).join(' · ') });
  });
  actions.sort((a, b) => a.at < b.at ? -1 : 1);
  return { units: units, actions: actions, engineers: engineers_() };
}

function pmUnit_(unit) {
  const u = unitInfo_(unit);
  const c = rows_(PT.clients).find(r => low_(r['Unit']) === unit.toLowerCase());
  return {
    unit: u.unit, project: u.project,
    client: (c && str_(c['Client name'])) || u.client,
    email: (c && str_(c['Client email'])) || u.email,
    invited: !!c, invitedAt: c ? dt_(c['Invited at']) : '', lastOpened: c ? dt_(c['Last opened']) : '',
    link: c ? PORTAL_URL + '#t=' + str_(c['Token']) : '', status: unitStatus_(u.unit),
    items: unitItems_(u.unit), concerns: unitConcerns_(u.unit), snags: unitSnags_(u.unit), certs: unitCerts_(u.unit), history: unitLog_(u.unit)
  };
}

function pmInvite_(b) {
  const u = unitInfo_(str_(b.unit));
  const by = str_(b.by) || 'O West team';
  const email = str_(b.email) || u.email;
  const name = str_(b.client) || u.client;
  if (b.send && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Add the client email first.');
  let token;
  withLock_(() => {
    const ref = findRow_(PT.clients, PC_HEADERS, 'Unit', u.unit);
    if (ref) {
      token = str_(ref.obj['Token']);
      if (b.newLink) token = newToken_();
      setCells_(ref, { 'Client name': name, 'Client email': email, 'Token': token, 'Invited by': by, 'Invited at': b.send ? now_() : ref.obj['Invited at'] });
    } else {
      token = newToken_();
      sheet_(PT.clients, PC_HEADERS).appendRow([u.unit, name, email, token, b.send ? now_() : '', by, '']);
    }
    if (b.newLink) log_(u.unit, 'New link made', '', '', '', 'The old link stopped working.', by, '');
  });
  const link = PORTAL_URL + '#t=' + token;
  if (b.send) {
    MailApp.sendEmail({
      to: email, name: 'O West Extra Works',
      subject: `Your O-Plus client portal · ${u.unit}`,
      htmlBody: mailHtml_('O-PLUS · CLIENT PORTAL', 'Your unit now has its own portal', [
        `Dear ${name || 'Client'},`,
        `From today, every scope item, date and material selection for ${u.unit} will be shared with you in one place. You can approve each one, request changes, and send us any concern.`,
        'The portal is the agreed record between you and O West. WhatsApp stays for quick messages.',
        'No password is needed: the button opens your portal directly, so please keep this email private. When you approve or reject an item, we email you a 6-digit code to confirm it is you.',
        `Your project manager: ${by}`
      ], { label: 'Open my portal', url: link })
    });
    log_(u.unit, 'Portal invitation sent', '', '', '', 'To ' + maskEmail_(email), by, '');
  }
  return Object.assign(pmUnit_(u.unit), { sent: !!b.send });
}

function pmUpload_(b) {
  return { id: savePortalPhoto_(unitInfo_(str_(b.unit)).unit, b.data, 'sample') };
}

function pmItem_(b) {
  const u = unitInfo_(str_(b.unit));
  const type = P_TYPES.indexOf(str_(b.type)) >= 0 ? str_(b.type) : P_TYPES[0];
  const title = str_(b.title).slice(0, 200), details = str_(b.details).slice(0, 4000);
  const by = str_(b.by) || 'O West team';
  if (!title) throw new Error('Add a title.');
  const photos = (b.photoIds || []).map(str_).filter(Boolean).slice(0, 8);
  let id, version = 1;
  withLock_(() => {
    if (b.replaces) {
      const old = findRow_(PT.items, PI_HEADERS, 'Item ID', str_(b.replaces));
      if (!old || low_(old.obj['Unit']) !== u.unit.toLowerCase()) throw new Error('The item to revise was not found.');
      version = Number(old.obj['Version'] || 1) + 1;
      setCells_(old, { 'Status': S_OLD });
    }
    id = 'I' + fmt_(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    sheet_(PT.items, PI_HEADERS).appendRow([id, u.unit, type, title, details, photos.join(','), version, str_(b.replaces), S_PENDING, now_(), by, '', '', '']);
    log_(u.unit, S_PENDING, id, version, title, type, by, '');
  });
  const c = rows_(PT.clients).find(r => low_(r['Unit']) === u.unit.toLowerCase());
  let emailed = false;
  if (c && str_(c['Client email'])) {
    MailApp.sendEmail({
      to: str_(c['Client email']), name: 'O West Extra Works',
      subject: `For your approval: ${title} · ${u.unit}`,
      htmlBody: mailHtml_(`${type.toUpperCase()} · VERSION ${version}`, title, [
        `Dear ${str_(c['Client name']) || 'Client'},`,
        `${by} has sent you an item to review for unit ${u.unit}. Please open your portal to approve it or request changes.`
      ], { label: 'Review in my portal', url: PORTAL_URL + '#t=' + str_(c['Token']) })
    });
    emailed = true;
  }
  return Object.assign(pmUnit_(u.unit), { emailed: emailed });
}

function pmReply_(b) {
  const unit = unitInfo_(str_(b.unit)).unit;
  const text = str_(b.text).slice(0, 4000), by = str_(b.by) || 'O West team';
  if (!text) throw new Error('Write a reply.');
  const photo = b.photo ? savePortalPhoto_(unit, b.photo, 'reply') : '';
  let subject = '';
  withLock_(() => {
    const ref = concernRef_(unit, b.concernId);
    subject = str_(ref.obj['Subject']);
    sheet_(PT.messages, PMS_HEADERS).appendRow([ref.obj['Concern ID'], unit, now_(), 'team', by, text, photo]);
    setCells_(ref, { 'Status': C_REPLIED, 'Updated at': now_() });
    log_(unit, 'Team replied', ref.obj['Concern ID'], '', subject, text, by, '');
  });
  const c = rows_(PT.clients).find(r => low_(r['Unit']) === unit.toLowerCase());
  if (c && str_(c['Client email'])) {
    MailApp.sendEmail({
      to: str_(c['Client email']), name: 'O West Extra Works',
      subject: `Reply to your concern: ${subject} · ${unit}`,
      htmlBody: mailHtml_('REPLY TO YOUR CONCERN', subject, [`${by}: ${text}`], { label: 'Open my portal', url: PORTAL_URL + '#t=' + str_(c['Token']) })
    });
  }
  return pmUnit_(unit);
}

/* =========================== photos =========================== */

function savePortalPhoto_(unit, data, kind) {
  const b64 = str_(data).replace(/^data:image\/\w+;base64,/, '');
  if (!b64 || b64.length > MAX_BASE64_CHARS) throw new Error('The photo is too large.');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', `${safeUnit_(unit)}_${kind}_${fmt_(new Date(), 'yyyyMMdd_HHmmss')}.jpg`);
  return folderPath_([PORTAL_ROOT, safeUnit_(unit)]).createFile(blob).getId();
}

/** A photo is served only to the client of its unit, or to the team. */
function portalImg_(b) {
  const id = str_(b.id);
  let unit = '';
  if (b.pmToken) requireAdmin_(b.pmToken);
  else unit = clientByToken_(b.token).unit.toLowerCase();
  if (unit) {
    const mine = rows_(PT.items).some(r => low_(r['Unit']) === unit && str_(r['Photo IDs']).split(',').indexOf(id) >= 0) ||
      rows_(PT.messages).some(r => low_(r['Unit']) === unit && str_(r['Photo ID']) === id) ||
      rows_(PT.snags).some(r => low_(r['Unit']) === unit && (str_(r['Photo IDs']) + ',' + str_(r['Fix photo IDs'])).split(',').indexOf(id) >= 0);
    if (!mine && !sitePhotos_(unit).some(p => p.id === id)) throw new Error('Photo not found.');
  }
  const f = DriveApp.getFileById(id);
  return { id: id, data: 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) };
}

/* =========================== unit status (ClickUp + Site Camera) =========================== */

const P_CU = {
  poc: '640aa77a-6d8d-4b4d-b8a6-b4dfd74baa9f', siteDue: '192b3218-0025-48f1-a418-488d0e87e5ee',
  pm: 'f91d9a67-0ed7-4e6a-81c3-1b7bf673546f', cstatus: 'eb44d8ec-34af-462d-b01e-87f84b918f55',
  amount: 'f480622e-13eb-43c9-a141-f74a78d866e9', siteIn: 'e099ba01-7bd4-402b-9d0f-2263884b39b0'
};
const P_STAGES = ['Initiation', 'Design', 'Quotation', 'Construction', 'Quality Inspection', 'Hand Over'];
const P_CU_STAGE = { 'new case': 0, 'on hold': 0, 'client feedback': 1, 'design revision': 1, 'design execution': 1, 'quotation issued': 2, 'cost revision': 2, 'site execution': 3, 'handed over': 5 };

/** Stage, progress %, handover date and PM from the unit's ClickUp tasks (cached 3 h),
    plus the newest site photos and weekly report from Site Camera. Never fails the page. */
function unitStatus_(unit) {
  const out = { stage: null, stageName: '', poc: null, handover: '', pm: '', photos: [], report: null };
  try { Object.assign(out, portalClickUp_(unit) || {}); } catch (e) { console.error(e); }
  try { out.photos = sitePhotos_(unit.toLowerCase()).slice(0, 4); } catch (e) {}
  try {
    const r = camRows_('Reports').filter(x => str_(x['Report']) === 'Weekly client' && low_(x['For']) === unit.toLowerCase()).pop();
    if (r) out.report = { url: str_(r['PDF link']), period: str_(r['Period']) };
  } catch (e) {}
  return out;
}

function portalClickUp_(unit) {
  const cache = CacheService.getScriptCache(), key = 'pcu:' + unit.toLowerCase();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_TOKEN');
  if (!token) return null;
  const k = unit.toLowerCase();
  let ids = camRows_('ClickUp Tasks').filter(r => low_(normUnit_(r['Unit'])) === k && ['deleted', 'declined'].indexOf(low_(r['Status'])) < 0).map(r => str_(r['Task ID'])).filter(Boolean);
  if (!ids.length) return null;
  const tasks = [];
  UrlFetchApp.fetchAll(ids.slice(0, 10).map(id => ({ url: `https://api.clickup.com/api/v2/task/${id}`, headers: { Authorization: token }, muteHttpExceptions: true }))).forEach(r => {
    if (r.getResponseCode() !== 200) return;
    let t; try { t = JSON.parse(r.getContentText()); } catch (e) { return; }
    const f = {}; (t.custom_fields || []).forEach(c => f[c.id] = c);
    const v = id => str_(f[id] && f[id].value);
    const num = id => { const x = parseFloat(v(id)); return isNaN(x) ? null : x; };
    const due = Number(v(P_CU.siteDue));
    tasks.push({ status: low_(t.status && t.status.status), poc: num(P_CU.poc), amount: num(P_CU.amount) || 0,
      due: due ? fmt_(new Date(due), 'yyyy-MM-dd') : '', pm: dropName_(f[P_CU.pm]), cstatus: low_(dropName_(f[P_CU.cstatus])) });
  });
  const live = tasks.filter(t => t.status !== 'declined');
  if (!live.length) return null;
  const site = live.filter(t => t.status === 'site execution');
  const main = site.length ? site : live;
  let stage;
  if (site.length) stage = site.every(t => t.cstatus === 'completed' || t.cstatus === 'handedover') ? 4 : 3;
  else if (live.every(t => t.status === 'handed over')) stage = 5;
  else stage = Math.max.apply(null, live.map(t => P_CU_STAGE[t.status] == null ? 0 : P_CU_STAGE[t.status]));
  const withPoc = main.filter(t => t.poc != null);
  const w = withPoc.reduce((s, t) => s + t.amount, 0);
  const poc = !withPoc.length ? null : Math.round(w > 0 ? withPoc.reduce((s, t) => s + t.poc * t.amount, 0) / w : withPoc.reduce((s, t) => s + t.poc, 0) / withPoc.length);
  const pms = main.map(t => t.pm).filter((x, i, a) => x && a.indexOf(x) === i);
  const out = { stage: stage, stageName: P_STAGES[stage], poc: poc, handover: main.map(t => t.due).filter(Boolean).sort().pop() || '', pm: pms.join(', ') };
  try { cache.put(key, JSON.stringify(out), 10800); } catch (e) {}
  return out;
}

/** Newest site photos of a unit from Site Camera (unit name in lower case). */
function sitePhotos_(k) {
  return camRows_('Photos').filter(r => low_(r['Unit']) === k && str_(r['Status']) !== 'Deleted' && str_(r['File ID']))
    .map(r => ({ id: str_(r['File ID']), date: day_(r['Date']), item: str_(r['Item']) }))
    .reverse().slice(0, 12);
}

/* =========================== reading =========================== */

function unitInfo_(unit) {
  const k = low_(unit);
  const r = camRows_('Units').find(x => low_(x['Unit']) === k);
  if (!r) throw new Error('Unit not found: ' + unit);
  return { unit: str_(r['Unit']), project: str_(r['Project']), client: str_(r['Client name']), email: str_(r['Client email']) };
}

function unitItems_(unit) {
  const k = unit.toLowerCase();
  return rows_(PT.items).filter(r => low_(r['Unit']) === k).map(r => ({
    id: str_(r['Item ID']), type: str_(r['Type']), title: str_(r['Title']), details: str_(r['Details']),
    photos: str_(r['Photo IDs']).split(',').filter(Boolean), version: Number(r['Version'] || 1), replaces: str_(r['Replaces']),
    status: str_(r['Status']), sentAt: dt_(r['Sent at']), sentBy: str_(r['Sent by']),
    decidedAt: dt_(r['Decided at']), decidedBy: str_(r['Decided by']), reason: str_(r['Reason'])
  })).reverse();
}

function unitConcerns_(unit) {
  const k = unit.toLowerCase();
  const msgs = groupBy_(rows_(PT.messages).filter(r => low_(r['Unit']) === k), r => str_(r['Concern ID']));
  return rows_(PT.concerns).filter(r => low_(r['Unit']) === k).map(r => ({
    id: str_(r['Concern ID']), subject: str_(r['Subject']), status: str_(r['Status']),
    openedAt: dt_(r['Opened at']), updatedAt: dt_(r['Updated at']),
    messages: (msgs[str_(r['Concern ID'])] || []).map(m => ({ from: str_(m['From']), name: str_(m['Name']), at: dt_(m['At']), text: str_(m['Text']), photo: str_(m['Photo ID']) }))
  })).sort((a, b) => a.updatedAt < b.updatedAt ? 1 : -1);
}

function unitSnags_(unit) {
  const k = unit.toLowerCase(), ids = v => str_(v).split(',').filter(Boolean);
  return rows_(PT.snags).filter(r => low_(r['Unit']) === k).map(r => ({
    id: str_(r['Snag ID']), no: Number(r['No.'] || 0), location: str_(r['Location']), text: str_(r['Description']),
    photos: ids(r['Photo IDs']), status: str_(r['Status']), reportedAt: dt_(r['Reported at']), reportedBy: str_(r['Reported by']), from: str_(r['From']),
    fixedAt: dt_(r['Fixed at']), fixedBy: str_(r['Fixed by']), fixNote: str_(r['Fix note']), fixPhotos: ids(r['Fix photo IDs']),
    closedAt: dt_(r['Closed at']), closedBy: str_(r['Closed by']), reopenNote: str_(r['Reopen note'])
  })).sort((a, b) => b.no - a.no);
}

function unitLog_(unit) {
  const k = unit.toLowerCase();
  return rows_(PT.log).filter(r => low_(r['Unit']) === k).map(r => ({
    at: dt_(r['At']), event: str_(r['Event']), version: str_(r['Version']), title: str_(r['Title']), detail: str_(r['Detail']), by: str_(r['By'])
  })).reverse().slice(0, 200);
}

function openItem_(unit, itemId) {
  const r = rows_(PT.items).find(x => str_(x['Item ID']) === str_(itemId) && low_(x['Unit']) === unit.toLowerCase());
  if (!r) throw new Error('This item was not found.');
  if (str_(r['Status']) !== S_PENDING) throw new Error('This item was already decided or replaced. Please reload.');
  return { id: str_(r['Item ID']), title: str_(r['Title']), version: Number(r['Version'] || 1) };
}

function snagRef_(unit, id) {
  const ref = findRow_(PT.snags, PSN_HEADERS, 'Snag ID', str_(id));
  if (!ref || low_(ref.obj['Unit']) !== unit.toLowerCase()) throw new Error('This snag was not found.');
  return ref;
}

function concernRef_(unit, id) {
  const ref = findRow_(PT.concerns, PCN_HEADERS, 'Concern ID', str_(id));
  if (!ref || low_(ref.obj['Unit']) !== unit.toLowerCase()) throw new Error('This concern was not found.');
  return ref;
}

/* =========================== writing =========================== */

function findRow_(name, headers, col, value) {
  const sh = sheet_(name, headers);
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim()), c = h.indexOf(col), want = low_(value);
  for (let i = 1; i < v.length; i++) {
    if (low_(v[i][c]) === want) { const obj = {}; h.forEach((k, j) => obj[k] = v[i][j]); return { sh: sh, row: i + 1, h: h, obj: obj }; }
  }
  return null;
}

function setCells_(ref, patch) {
  Object.keys(patch).forEach(k => { const c = ref.h.indexOf(k); if (c >= 0) ref.sh.getRange(ref.row, c + 1).setValue(patch[k]); });
}

function log_(unit, event, ref, version, title, detail, by, confirmed) {
  sheet_(PT.log, PL_HEADERS).appendRow([now_(), unit, event, ref, version, title, detail, by, confirmed]);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function newToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40); }
function now_() { return fmt_(new Date(), 'yyyy-MM-dd HH:mm'); }
function maskEmail_(e) {
  const m = str_(e).match(/^(.)(.*)(.)@(.+)$/);
  return m ? `${m[1]}***${m[3]}@${m[4]}` : str_(e);
}

/* =========================== email =========================== */

function notifyTeam_(subject, eyebrow, title, lines) {
  const S = settings_();
  const to = S['Team emails'];
  if (!to) return;
  try {
    MailApp.sendEmail({ to: to, name: 'O-Plus Client Portal', subject: subject,
      htmlBody: mailHtml_(eyebrow, title, lines, { label: 'Open the team portal', url: PORTAL_URL + 'pm.html' }) });
  } catch (e) { console.error(e); }
}

/** One branded email layout: navy band, eyebrow, title, amber bar, lines, one button. */
function mailHtml_(eyebrow, title, lines, button) {
  const p = lines.map(l => `<p style="margin:0 0 12px;font-size:14px;line-height:22px;color:#1A2238">${esc_(l)}</p>`).join('');
  const btn = button ? `<a href="${esc_(button.url)}" style="display:inline-block;background:#001641;color:#ffffff;text-decoration:none;font-weight:600;font-size:15px;padding:13px 26px;border-radius:8px;margin:6px 0 16px">${esc_(button.label)}</a>` : '';
  return `<div style="background:#F3F5F9;padding:16px 0;font-family:Poppins,Arial,sans-serif">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #DCDFE6">
    <tr><td style="background:#001641;padding:14px 20px">
      <img src="${PORTAL_ASSETS}oplus-mark-white.png" alt="O-Plus" height="36" style="vertical-align:middle">
      <img src="${PORTAL_ASSETS}owest-lockup-white.png" alt="O West by Orascom Development" height="18" style="vertical-align:middle;float:right;margin-top:9px">
    </td></tr>
    <tr><td style="padding:24px 22px 8px">
      <div style="font-size:10px;font-weight:700;letter-spacing:2px;color:#001641">${esc_(eyebrow)}</div>
      <div style="font-size:20px;line-height:27px;font-weight:700;color:#001641;margin-top:4px">${esc_(title)}</div>
      <div style="width:32px;height:3px;background:#F0B323;margin:8px 0 16px"></div>
      ${p}${btn}
      <p style="margin:0 0 12px;font-size:13px;color:#5F677A">Kind regards,<br>O West Extra Works</p>
    </td></tr>
    <tr><td style="border-top:1px solid #DCDFE6;padding:10px 22px;font-size:11px;color:#5F677A"><b style="color:#001641">O-PLUS</b> · O West by Orascom Development</td></tr>
  </table></div>`;
}

/* =========================== setup =========================== */

/** Run once (or from the menu): makes the portal tabs and protects the record. */
function portalSetup() {
  SpreadsheetApp.getActive().setSpreadsheetTimeZone(TZ);
  sheet_(PT.clients, PC_HEADERS); sheet_(PT.items, PI_HEADERS); sheet_(PT.concerns, PCN_HEADERS); sheet_(PT.messages, PMS_HEADERS); sheet_(PT.snags, PSN_HEADERS); sheet_(PT.certs, PH_HEADERS);
  const log = sheet_(PT.log, PL_HEADERS);
  if (!log.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
    const p = log.protect().setDescription('Client portal record: written by the portal only');
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
  }
  folderPath_([PORTAL_ROOT]);
  if (!('Team emails' in settings_())) sheet_('Settings', ['Setting', 'Value']).appendRow(['Team emails', '']);
  sheet_(DASH_TAB, DASH_COLS);
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dashRefresh'))
    ScriptApp.newTrigger('dashRefresh').timeBased().everyHours(1).create();
  MailApp.getRemainingDailyQuota();
  try { SpreadsheetApp.getUi().alert('Client portal is ready. Next: set the team password, connect ClickUp, and fill "Team emails" in Settings.'); } catch (e) {}
}


/* =========================== team access =========================== */

function login_(b) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_HASH');
  if (!hash) throw new Error('No team password yet. Set it from the sheet menu: O-Plus Portal > Set team password.');
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
    const e = new Error('Please sign in with the team password.'); e.code = 'auth'; throw e;
  }
}

function sign_(v) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(v, key)).replace(/=+$/, '');
}
function sha_(v) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8)); }

/* =========================== sheet menu =========================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('O-Plus Portal')
    .addItem('Set up (run once)', 'portalSetup')
    .addItem('Set team password', 'menuSetPassword')
    .addItem('Connect ClickUp (API token)', 'menuSetClickUpToken')
    .addItem('Refresh dashboard now', 'dashRefresh')
    .addToUi();
}

function menuSetPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Team password', 'Type the password PMs will use on the team page (at least 6 characters).\nPhones signed in with the old password will need the new one.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText().trim();
  if (pw.length < 6) return ui.alert('Too short. Use at least 6 characters.');
  const props = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid();
  props.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: sha_(pw + salt), ADMIN_SECRET: Utilities.getUuid() + Utilities.getUuid() });
  ui.alert('Team password saved.');
}

function menuSetClickUpToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Connect ClickUp', 'Paste a ClickUp personal API token (starts with pk_). It is used only to read progress, handover date and PM.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim();
  if (!/^pk_/.test(t)) return ui.alert('That does not look like a ClickUp token (it should start with pk_).');
  PropertiesService.getScriptProperties().setProperty('CLICKUP_TOKEN', t);
  ui.alert('ClickUp connected.');
}

/* =========================== helpers =========================== */

/** Read-only access to the Site Camera sheet. */
let CAM = null;
function camRows_(name) {
  CAM = CAM || SpreadsheetApp.openById(CAMERA_SHEET_ID);
  const sh = CAM.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return toObjects_(sh.getDataRange().getValues());
}
function engineers_() { return camRows_('Engineers').filter(r => str_(r['Name']) && low_(r['Active']) !== 'no').map(r => str_(r['Name'])); }

function rows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return toObjects_(sh.getDataRange().getValues());
}
function toObjects_(v) {
  const h = v[0].map(x => String(x).trim());
  return v.slice(1).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
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
function settings_() {
  const out = {};
  rows_('Settings').forEach(r => { if (str_(r['Setting'])) out[str_(r['Setting'])] = str_(r['Value']); });
  return out;
}
function folderPath_(parts) {
  const props = PropertiesService.getScriptProperties();
  let f = null;
  const rid = props.getProperty('ROOT:' + parts[0]);
  if (rid) { try { f = DriveApp.getFolderById(rid); } catch (e) {} }
  if (!f) { const it = DriveApp.getFoldersByName(parts[0]); f = it.hasNext() ? it.next() : DriveApp.createFolder(parts[0]); props.setProperty('ROOT:' + parts[0], f.getId()); }
  for (let i = 1; i < parts.length; i++) { const it = f.getFoldersByName(parts[i]); f = it.hasNext() ? it.next() : f.createFolder(parts[i]); }
  return f;
}
function normUnit_(name) { const n = str_(name).replace(/\s+/g, '').toUpperCase(); return /^OW\//.test(n) ? n : ''; }
function dropName_(f) {
  if (!f || f.value == null || f.value === '') return '';
  const opts = (f.type_config && f.type_config.options) || [];
  const v = f.value;
  const o = opts.find(x => x.id === v || x.orderindex === v || String(x.orderindex) === String(v));
  return o ? str_(o.name) : (typeof v === 'object' && v.name ? str_(v.name) : '');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function msg_(e) { return String(e && e.message || e); }
function str_(v) { return v == null ? '' : String(v).trim(); }
function low_(v) { return str_(v).toLowerCase(); }
function esc_(s) { return str_(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function day_(v) { return v instanceof Date ? fmt_(v, 'yyyy-MM-dd') : str_(v).slice(0, 10); }
function dt_(v) { return v instanceof Date ? fmt_(v, 'yyyy-MM-dd HH:mm') : str_(v); }
function safeUnit_(u) { return u.replace(/[\/\\]+/g, '-').replace(/\s+/g, ' ').trim(); }
function groupBy_(list, f) { const o = {}; list.forEach(x => (o[f(x)] = o[f(x)] || []).push(x)); return o; }
