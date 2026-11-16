// O-Plus portal role-play demo: shared action log for all phones.
// Stand-alone Apps Script. It stores only demo actions (no client data).
// Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.

function doGet(e) {
  var p = (e && e.parameter) || {};
  var props = PropertiesService.getScriptProperties();
  var out;
  if (p.op === 'add') {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var n = Number(props.getProperty('n') || 0);
      var d = String(p.d || '');
      if (d.length > 8000) throw new Error('too long');
      var act = JSON.parse(d);
      if (act.a === 'reset') {
        props.deleteAllProperties();
        props.setProperty('epoch', String(Date.now()));
        n = 0;
      }
      props.setProperty('a' + n, JSON.stringify(act));
      props.setProperty('n', String(n + 1));
      out = { ok: true, n: n + 1 };
    } finally {
      lock.releaseLock();
    }
  } else {
    var all = props.getProperties();
    var total = Number(all.n || 0);
    var since = Math.max(0, Number(p.since || 0));
    var items = [];
    for (var i = since; i < total; i++) {
      if (all['a' + i]) items.push(JSON.parse(all['a' + i]));
    }
    out = { n: total, epoch: all.epoch || '0', items: items };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
