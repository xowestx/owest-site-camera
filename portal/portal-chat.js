/* O-Plus Client Portal: the WhatsApp-style chat, shared by the client page and the team page.
   Text, photos, voice notes, quick replies, reply-to, sent/seen ticks, day separators, live refresh. */

const Chat = (() => {
  let S = null; // the open chat: { o: options, c: chat from server, pending: [], replyTo, rec, timer }
  const FILE = {}; // file id -> data URL (photos and voice notes, fetched once)
  const T = (en, ar) => (S && S.o.L ? S.o.L(en, ar) : en);
  const hhmm = at => { const m = String(at || "").match(/ (\d{2}):(\d{2})/); if (!m) return ""; const h = +m[1]; return `${h % 12 || 12}:${m[2]} ${AR() ? (h < 12 ? "ص" : "م") : (h < 12 ? "AM" : "PM")}`; };
  const dayKey = at => String(at || "").slice(0, 10);
  const isoDay = d => new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  function dayLabel(k) {
    const today = isoDay(new Date()), y = isoDay(new Date(Date.now() - 864e5));
    return k === today ? T("Today", "اليوم") : k === y ? T("Yesterday", "أمس") : when(k, true);
  }
  const secs = n => `${Math.floor(n / 60)}:${String(Math.round(n) % 60).padStart(2, "0")}`;
  const ICON = {
    back: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 6l-6 6 6 6"/></svg>`,
    plus: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>`,
    send: `<svg viewBox="0 0 24 24" fill="currentColor"><path d="M3.4 20.4l17.4-7.5c.8-.3.8-1.5 0-1.8L3.4 3.6c-.7-.3-1.4.3-1.2 1l1.7 6.1 9.1 1.3-9.1 1.3-1.7 6.1c-.2.7.5 1.3 1.2 1z"/></svg>`,
    mic: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0014 0M12 18v3"/></svg>`,
    trash: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>`,
    play: `<svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>`,
    pause: `<svg viewBox="0 0 24 24" fill="currentColor"><path d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>`,
    one: `<svg viewBox="0 0 18 12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6.5l3.5 3.5L13 2"/></svg>`,
    two: `<svg viewBox="0 0 18 12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M1 6.5l3.5 3.5L12 2"/><path d="M7 9l1 1L15.5 2"/></svg>`,
    clock: `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="8" cy="8" r="6"/><path d="M8 5v3l2 1.5"/></svg>`,
    reply: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 9V5l-7 7 7 7v-4.1c5 0 8.5 1.6 11 5.1-1-5-4-10-11-11z"/></svg>`,
    photo: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-8 8"/></svg>`,
    cam: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>`,
    bell: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 16V11a6 6 0 0112 0v5l2 2H4z"/><path d="M10 21h4"/></svg>`
  };

  /** Open the chat inside `root`. o = { side, title, avatar, L, call(payload), quick: [..], onBack, statusLine(c) } */
  async function open(root, o) {
    close();
    S = { o, c: null, pending: [], replyTo: "", rec: null, timer: null, root };
    document.body.classList.add("chatmode");
    root.innerHTML = `<div class="chatwrap" id="chatwrap">
      <header class="chathead"><button class="cb" id="chBack" aria-label="${T("Back", "رجوع")}">${ICON.back}</button>
        <span class="av">${o.avatar || ""}</span><span class="who"><b>${esc(o.title)}</b><span id="chStatus">${T("Loading…", "جاري التحميل…")}</span></span></header>
      <div class="chatlog" id="clog" aria-live="polite"></div>
      <footer class="chatfoot">
        <div class="quote" id="chQuote" hidden></div>
        <div class="quick" id="chQuick">${(o.quick || []).map(q => `<button class="chip" data-q="${esc(q)}" dir="auto">${esc(q)}</button>`).join("")}</div>
        <div class="composer" id="chComp">
          <button class="cb" id="chPlus" aria-label="${T("Attach", "إرفاق")}">${ICON.plus}</button>
          <textarea id="chText" rows="1" dir="auto" placeholder="${T("Message", "اكتب رسالة")}" aria-label="${T("Message", "اكتب رسالة")}"></textarea>
          <button class="cb go" id="chGo" aria-label="${T("Record a voice note", "تسجيل رسالة صوتية")}">${ICON.mic}</button>
        </div>
        <div class="recbar" id="chRec" hidden><button class="cb" id="chRecX" aria-label="${T("Cancel", "إلغاء")}">${ICON.trash}</button>
          <span class="rdot"></span><span id="chRecT">0:00</span><span class="small muted">${T("Recording…", "جاري التسجيل…")}</span>
          <button class="cb go" id="chRecSend" aria-label="${T("Send voice note", "إرسال")}">${ICON.send}</button></div>
        <input type="file" id="chFile" accept="image/*" multiple hidden><input type="file" id="chCam" accept="image/*" capture="environment" hidden>
      </footer></div>`;
    $("chBack").onclick = () => { close(); o.onBack && o.onBack(); };
    const ta = $("chText");
    ta.oninput = () => { ta.style.height = "auto"; ta.style.height = Math.min(120, ta.scrollHeight) + "px"; sync(); };
    ta.onkeydown = e => { if (e.key === "Enter" && !e.shiftKey && matchMedia("(pointer:fine)").matches) { e.preventDefault(); $("chGo").click(); } };
    $("chGo").onclick = () => ta.value.trim() ? sendText() : startRec();
    $("chPlus").onclick = menu;
    $("chFile").onchange = e => sendPhotos([...e.target.files]);
    $("chCam").onchange = e => sendPhotos([...e.target.files]);
    root.querySelectorAll("[data-q]").forEach(b => b.onclick = () => { ta.value = b.dataset.q; ta.oninput(); ta.focus(); });
    $("chRecX").onclick = () => stopRec(false);
    $("chRecSend").onclick = () => stopRec(true);
    sync();
    await refresh(true);
    S && (S.timer = setInterval(() => { if (!document.hidden && !(S && S.busy)) refresh(false); }, 6000));
  }

  function close() {
    if (!S) return;
    clearInterval(S.timer);
    if (S.rec) try { S.rec.r.stop(); S.rec.stream.getTracks().forEach(t => t.stop()); } catch (e) {}
    document.body.classList.remove("chatmode");
    S = null;
  }

  function sync() {
    const has = $("chText") && $("chText").value.trim();
    if (!$("chGo")) return;
    $("chGo").innerHTML = has ? ICON.send : ICON.mic;
    $("chGo").setAttribute("aria-label", has ? T("Send", "إرسال") : T("Record a voice note", "تسجيل رسالة صوتية"));
    $("chQuick").hidden = !!has || !(S.o.quick || []).length;
  }

  async function refresh(first) {
    const s = S;
    try {
      const j = await s.o.call({ action: "get", after: first || !s.c ? 0 : s.c.total });
      if (s !== S) return;
      merge(j.chat, first);
    } catch (e) { if (first && s === S) $("clog").innerHTML = `<div class="center">${esc(e.message)}</div>`; }
  }

  function merge(c, all) {
    const prev = S.c;
    if (all || !prev || c.from === 0) S.c = c;
    else { S.c = Object.assign({}, c, { msgs: prev.msgs.concat(c.msgs.filter(m => !prev.msgs.some(p => p.id === m.id))) }); }
    S.c.total = S.c.msgs.length;
    const grew = !prev || S.c.msgs.length !== prev.msgs.length;
    const sig = S.c.msgs.length + "|" + JSON.stringify(S.c.read) + "|" + JSON.stringify(S.c.rules) + "|" + S.pending.length;
    if (all || sig !== S.sig) { S.sig = sig; draw(grew); }
    S.o.onSeen && S.o.onSeen(S.c);
  }

  /** Draw the whole conversation (it is small: one unit's chat). */
  function draw(scroll) {
    const log = $("clog"); if (!log) return;
    const nearEnd = log.scrollHeight - log.scrollTop - log.clientHeight < 120;
    const c = S.c, side = S.o.side;
    const seenAt = side === "client" ? c.read.team : c.read.client;
    $("chStatus").textContent = S.o.statusLine ? S.o.statusLine(c) : "";
    const byId = {}; c.msgs.forEach(m => byId[m.id] = m);
    let html = "", lastDay = "", lastSeenMine = null;
    c.msgs.forEach(m => { if (m.from === side && seenAt && seenAt >= m.at) lastSeenMine = m.id; });
    if (!c.msgs.length && !S.pending.length) html = `<div class="chempty">${S.o.empty || ""}</div>`;
    c.msgs.forEach(m => {
      const d = dayKey(m.at);
      if (d !== lastDay) { html += `<div class="dsep"><span>${dayLabel(d)}</span></div>`; lastDay = d; }
      html += bubble(m, m.from === side, m.from === side ? (seenAt && seenAt >= m.at ? "seen" : "sent") : "", byId);
      if (m.id === lastSeenMine && m === c.msgs.filter(x => x.from === side).slice(-1)[0])
        html += `<div class="seenline">${side === "client" ? T("Seen", "شوهدت") + (c.read.teamBy ? " · " + esc(c.read.teamBy) : "") : T("Seen by the client", "شاهدها العميل")} · ${hhmm(seenAt)}</div>`;
    });
    S.pending.forEach(p => html += bubble(p, true, p.failed ? "failed" : "wait", byId));
    log.innerHTML = html;
    log.querySelectorAll("[data-img]").forEach(loadImg);
    log.querySelectorAll("[data-voice]").forEach(v => v.onclick = () => play(v));
    log.querySelectorAll("[data-reply]").forEach(b => b.onclick = e => { e.stopPropagation(); setReply(b.dataset.reply); });
    log.querySelectorAll("[data-retry]").forEach(b => b.onclick = () => retry(b.dataset.retry));
    log.querySelectorAll("[data-open]").forEach(i => i.onclick = () => i.src && lightbox(i.src));
    if (scroll || nearEnd) log.scrollTop = log.scrollHeight;
  }

  function bubble(m, mine, state, byId) {
    const q = m.replyTo && byId[m.replyTo];
    const quote = q ? `<div class="bq"><b>${esc(q.from === S.o.side ? T("You", "أنت") : q.name)}</b><span dir="auto">${esc(preview(q))}</span></div>` : "";
    let body = "";
    if (m.kind === "photo") body = `<img class="bimg" data-open ${m.local ? `src="${m.local}"` : `data-img="${esc(m.file)}"`} alt="${T("Photo", "صورة")}">${m.text ? `<div class="bt" dir="auto">${esc(m.text)}</div>` : ""}`;
    else if (m.kind === "voice") body = `<button class="voice" ${m.file ? `data-voice="${esc(m.file)}"` : "disabled"} aria-label="${T("Play voice note", "تشغيل الرسالة الصوتية")}"><span class="pp">${ICON.play}</span><span class="track"><i></i></span><span class="dur">${secs(m.secs || 0)}</span></button>`;
    else if (m.kind === "urgent") body = `<div class="burg">${ICON.bell}<b>${T("Urgent: please call me", "عاجل: برجاء الاتصال بي")}</b></div>${m.text ? `<div class="bt" dir="auto">${esc(m.text)}</div>` : ""}`;
    else body = `<div class="bt" dir="auto">${linkify(esc(m.text))}</div>`;
    const tick = state === "seen" ? `<span class="tk seen" aria-label="${T("Seen", "شوهدت")}">${ICON.two}</span>` : state === "sent" ? `<span class="tk" aria-label="${T("Sent", "أُرسلت")}">${ICON.one}</span>`
      : state === "wait" ? `<span class="tk" aria-label="${T("Sending", "جاري الإرسال")}">${ICON.clock}</span>` : state === "failed" ? `<button class="retry" data-retry="${esc(m.id)}">${T("Not sent · tap to retry", "لم تُرسل · اضغط لإعادة المحاولة")}</button>` : "";
    const who = mine ? "" : `<div class="bn">${esc(m.name)}</div>`;
    return `<div class="bub ${mine ? "me" : "them"}${m.kind === "urgent" ? " urgent" : ""}" data-id="${esc(m.id)}">${who}${quote}${body}
      <div class="bm">${hhmm(m.at)}${tick}</div>${m.local ? "" : `<button class="rb" data-reply="${esc(m.id)}" aria-label="${T("Reply", "رد")}">${ICON.reply}</button>`}</div>`;
  }
  const preview = m => m.kind === "voice" ? T("Voice note", "رسالة صوتية") + " · " + secs(m.secs || 0) : m.kind === "photo" ? T("Photo", "صورة") + (m.text ? " · " + m.text : "") : m.kind === "urgent" ? T("Urgent: please call me", "عاجل: برجاء الاتصال بي") : String(m.text || "").slice(0, 90);
  const linkify = h => h.replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');

  function setReply(id) {
    const m = S.c.msgs.find(x => x.id === id); if (!m) return;
    S.replyTo = id;
    const q = $("chQuote"); q.hidden = false;
    q.innerHTML = `<span><b>${esc(m.from === S.o.side ? T("You", "أنت") : m.name)}</b><span dir="auto">${esc(preview(m))}</span></span><button class="cb" aria-label="${T("Cancel reply", "إلغاء الرد")}">×</button>`;
    q.querySelector("button").onclick = clearReply;
    $("chText").focus();
  }
  function clearReply() { if (!S) return; S.replyTo = ""; $("chQuote").hidden = true; }

  async function loadImg(el) {
    const id = el.dataset.img;
    if (FILE[id]) { el.src = FILE[id]; return; }
    try { const j = await S.o.call({ action: "file", id }); FILE[id] = j.data; el.src = j.data; } catch (e) { el.alt = T("Photo unavailable", "الصورة غير متاحة"); }
  }

  let audio = null, audioBtn = null;
  async function play(btn) {
    const id = btn.dataset.voice;
    if (audioBtn === btn && audio) { if (audio.paused) audio.play(); else audio.pause(); return; }
    if (audio) { audio.pause(); audioBtn && (audioBtn.querySelector(".pp").innerHTML = ICON.play); }
    audioBtn = btn; btn.querySelector(".pp").innerHTML = ICON.clock;
    try {
      if (!FILE[id]) FILE[id] = (await S.o.call({ action: "file", id })).data;
      audio = new Audio(FILE[id]);
      const bar = btn.querySelector(".track i"), dur = btn.querySelector(".dur"), total = dur.textContent;
      audio.ontimeupdate = () => { if (audio.duration && isFinite(audio.duration)) bar.style.width = (audio.currentTime / audio.duration * 100) + "%"; dur.textContent = secs(audio.currentTime); };
      audio.onplay = () => btn.querySelector(".pp").innerHTML = ICON.pause;
      audio.onpause = () => btn.querySelector(".pp").innerHTML = ICON.play;
      audio.onended = () => { bar.style.width = "0"; dur.textContent = total; btn.querySelector(".pp").innerHTML = ICON.play; };
      await audio.play();
    } catch (e) { btn.querySelector(".pp").innerHTML = ICON.play; toast(T("This voice note cannot play on this phone.", "لا يمكن تشغيل الرسالة الصوتية على هذا الجهاز.")); }
  }

  function lightbox(src) {
    const d = document.createElement("div"); d.className = "lightbox"; d.innerHTML = `<img src="${src}" alt=""><button class="cb" aria-label="${T("Close", "إغلاق")}">×</button>`;
    d.onclick = () => d.remove(); document.body.appendChild(d);
  }

  function menu() {
    const s = sheet(`<div class="eyebrow">${T("Send", "إرسال")}</div><div class="attach">
      <button id="aPh">${ICON.photo}<span>${T("Photos", "صور")}</span></button>
      <button id="aCam">${ICON.cam}<span>${T("Camera", "الكاميرا")}</span></button>
      ${S.o.side === "client" ? `<button id="aUrg" class="u">${ICON.bell}<span>${T("Urgent call", "مكالمة عاجلة")}</span></button>` : ""}
    </div><div class="stack"><button class="btn ghost" id="aX">${T("Cancel", "إلغاء")}</button></div>`);
    s.querySelector("#aX").onclick = closeSheet;
    s.querySelector("#aPh").onclick = () => { closeSheet(); $("chFile").click(); };
    s.querySelector("#aCam").onclick = () => { closeSheet(); $("chCam").click(); };
    const u = s.querySelector("#aUrg"); if (u) u.onclick = urgent;
  }

  function urgent() {
    const s = sheet(`<div class="eyebrow" style="color:var(--problem)">${T("Urgent", "عاجل")}</div>
      <div class="title" style="font-size:17px">${T("Ask your PM to call you within 30 minutes", "اطلب من مدير المشروع الاتصال بك خلال 30 دقيقة")}</div>
      <p class="small muted">${T("Use this for something that cannot wait, such as a leak or a safety problem. Your PM and their manager are alerted now.", "استخدم هذا لأمر لا يحتمل الانتظار، مثل تسريب مياه أو مشكلة أمان. سيتم تنبيه مدير المشروع ومديره فوراً.")}</p>
      <label class="f" for="ug">${T("What happened? (optional)", "ماذا حدث؟ (اختياري)")}</label><textarea class="f" id="ug" rows="3" dir="auto"></textarea>
      <div class="stack"><button class="btn red" id="ugGo">${T("Send urgent request", "إرسال طلب عاجل")}</button><button class="btn ghost" id="ugX">${T("Cancel", "إلغاء")}</button></div>`);
    s.querySelector("#ugX").onclick = closeSheet;
    s.querySelector("#ugGo").onclick = () => { const t = $("ug").value.trim(); closeSheet(); send({ kind: "urgent", text: t }); };
  }

  function sendText() {
    const ta = $("chText"), text = ta.value.trim(); if (!text) return;
    ta.value = ""; ta.oninput();
    send({ kind: "text", text });
  }

  async function sendPhotos(files) {
    $("chFile").value = ""; $("chCam").value = "";
    const caption = $("chText").value.trim(); if (caption) { $("chText").value = ""; $("chText").oninput(); }
    for (let i = 0; i < Math.min(files.length, 6); i++) {
      try { const data = await compress(files[i]); await send({ kind: "photo", data, text: i === 0 ? caption : "" }, data); }
      catch (e) { toast(T("That photo could not be read.", "تعذرت قراءة الصورة.")); }
    }
  }

  /** Optimistic send: the bubble shows a clock until the server answers, then a tick. */
  async function send(p, local) {
    if (!S) return;
    const id = "tmp" + Date.now() + Math.random().toString(36).slice(2, 6);
    const now = new Date(), at = isoDay(now) + " " + String(now.getHours()).padStart(2, "0") + ":" + String(now.getMinutes()).padStart(2, "0") + ":00";
    const m = Object.assign({ id, at, from: S.o.side, name: "", local: local || "", replyTo: S.replyTo, payload: p }, p);
    S.pending.push(m); clearReply(); draw(true);
    await push(m);
  }
  async function push(m) {
    const s = S; s.busy = true;
    try {
      const j = await s.o.call(Object.assign({ action: "send", replyTo: m.replyTo, after: 0 }, m.payload));
      if (s !== S) return;
      s.pending = s.pending.filter(x => x !== m);
      merge(j.chat, true);
    } catch (e) {
      if (s !== S) return;
      m.failed = true; draw(true); toast(e.message);
    } finally { s.busy = false; }
  }
  function retry(id) { const m = S.pending.find(x => x.id === id); if (m) { m.failed = false; draw(true); push(m); } }

  /* ---------- voice notes ---------- */
  async function startRec() {
    if (!navigator.mediaDevices || !window.MediaRecorder) return toast(T("Voice notes need a newer browser on this phone.", "الرسائل الصوتية تحتاج متصفحاً أحدث على هذا الجهاز."));
    let stream;
    try { stream = await navigator.mediaDevices.getUserMedia({ audio: true }); }
    catch (e) { return toast(T("Please allow the microphone to record a voice note.", "برجاء السماح باستخدام الميكروفون لتسجيل رسالة صوتية.")); }
    const type = ["audio/mp4", "audio/webm;codecs=opus", "audio/webm", "audio/ogg;codecs=opus"].find(t => MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(t)) || "";
    const r = new MediaRecorder(stream, type ? { mimeType: type, audioBitsPerSecond: 32000 } : undefined);
    const chunks = [], t0 = Date.now();
    r.ondataavailable = e => e.data.size && chunks.push(e.data);
    S.rec = { r, stream, chunks, t0, keep: false };
    r.start(250);
    $("chComp").hidden = true; $("chQuick").hidden = true; $("chRec").hidden = false;
    S.rec.tick = setInterval(() => {
      const s = (Date.now() - t0) / 1000; $("chRecT").textContent = secs(s);
      if (s >= 180) stopRec(true);
    }, 250);
  }
  function stopRec(keep) {
    const rec = S && S.rec; if (!rec) return;
    S.rec = null; clearInterval(rec.tick);
    $("chComp").hidden = false; $("chRec").hidden = true; sync();
    const secsLen = Math.max(1, Math.round((Date.now() - rec.t0) / 1000));
    rec.r.onstop = () => {
      rec.stream.getTracks().forEach(t => t.stop());
      if (!keep) return;
      if (secsLen < 1) return;
      const blob = new Blob(rec.chunks, { type: (rec.r.mimeType || "audio/webm").split(";")[0] });
      const fr = new FileReader();
      fr.onload = () => send({ kind: "voice", data: fr.result, secs: secsLen });
      fr.readAsDataURL(blob);
    };
    rec.r.stop();
  }

  return { open, close, isOpen: () => !!S, refresh: () => S && refresh(false), ICON };
})();
