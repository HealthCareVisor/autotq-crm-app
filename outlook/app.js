// AutoTQ Merge — Outlook compose task pane. Design: DESIGN.md. Logic: merge.js.
// Sends each personalized copy through the CRM's send-invoice `notify` action, from the signed-in teammate's own mailbox.
import * as M from "./merge.js";

const LOCAL = /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
const QS = new URLSearchParams(location.search);
const SUPA_URL = (LOCAL && QS.get("api")) || "https://nrbjtqczntrwqoheurpt.supabase.co";
const SUPA_KEY = "sb_publishable_nJhBA0zj3h11qzDWeIU0VA_1LrjKOp-";
const GAP_MS = (LOCAL && QS.get("gap") !== null ? Math.max(0, +QS.get("gap") || 0) : 15) * 1000;
const CAP_MAX = 50;
const SKEY = "autotq_merge_session_v1", PKEY = "autotq_merge_prefs_v1";
const TEAM_EMAILS = { "max@goldenhourmedical.com": "max", "hannah@goldenhourmedical.com": "hannah", "john@goldenhourmedical.com": "john",
  "jamie@goldenhourmedical.com": "jamie", "kellan@goldenhourmedical.com": "kellan", "johnmorrison@goldenhourmedical.com": "johnm",
  "billing@goldenhourmedical.com": "laresa", "timford@goldenhourmedical.com": "tim" };
const OPT_OUT = "If you'd rather not get emails like this from me, just reply and let me know.";

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const uid = () => Math.random().toString(36).slice(2, 9);
const ICON = {
  refresh: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg>',
  left: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m15 18-6-6 6-6"/></svg>',
  right: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m9 18 6-6-6-6"/></svg>',
};

let prefs = {};
try { prefs = JSON.parse(localStorage.getItem(PKEY) || "{}"); } catch { prefs = {}; }
const savePrefs = () => { try { localStorage.setItem(PKEY, JSON.stringify(prefs)); } catch { /* private mode */ } };

const S = {
  office: false, session: null, me: "", uid: "",
  login: { step: "email", email: prefs.email || "", mode: "code", msg: "", bad: false },
  email: { subject: "", html: "", text: "" }, manualText: prefs.manualText || "",
  src: prefs.src || "list", lists: [], listsLoaded: false, listId: "", pasteText: "", fileName: "", filePeople: [],
  people: [], loadingPeople: false, suppressed: new Set(), allowed: new Set(),
  fallbackThere: prefs.fallbackThere !== false, optOut: prefs.optOut !== false, optOutLine: prefs.optOutLine || OPT_OUT,
  cap: Math.min(CAP_MAX, +prefs.cap || CAP_MAX),
  merges: [], merge: null, idx: 0, rows: [], ready: [], fields: [], unknown: [], copyWarn: [], cid: false, hash: "", sentToday: 0, left: CAP_MAX,
  testedHash: "", testMsg: null, busy: "", confirm: false,
  run: null, // {total, done, failed, nextAt, paused, stopped, msg}
};

try { S.session = JSON.parse(localStorage.getItem(SKEY) || "null"); } catch { S.session = null; }
const saveSession = (s) => { S.session = s; try { if (s) localStorage.setItem(SKEY, JSON.stringify(s)); else localStorage.removeItem(SKEY); } catch { /* */ } };

/* ---------------- Supabase ---------------- */
async function supa(path, opts = {}) {
  const h = Object.assign({ apikey: SUPA_KEY, "Content-Type": "application/json" }, opts.headers || {});
  if (S.session?.access_token) h.Authorization = "Bearer " + S.session.access_token;
  const res = await fetch(SUPA_URL + path, Object.assign({}, opts, { headers: h }));
  if (res.status === 401 && S.session?.refresh_token && !opts._retried) {
    if (await refresh()) return supa(path, Object.assign({}, opts, { _retried: true }));
  }
  return res;
}
async function refresh() {
  try {
    const r = await fetch(SUPA_URL + "/auth/v1/token?grant_type=refresh_token", { method: "POST",
      headers: { apikey: SUPA_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ refresh_token: S.session.refresh_token }) });
    if (!r.ok) { saveSession(null); return false; }
    saveSession(await r.json()); return true;
  } catch { return false; }
}
const inList = (vals) => "(" + vals.map((v) => '"' + String(v).replace(/"/g, "") + '"').join(",") + ")";
async function records(coll, filters = "") {
  const out = [];
  for (let off = 0; ; off += 1000) {
    const r = await supa(`/rest/v1/records?collection=eq.${coll}&deleted=eq.false${filters}&select=id,data&order=id.asc&limit=1000&offset=${off}`);
    if (!r.ok) throw new Error(`Couldn't read ${coll} (${r.status})`);
    const j = await r.json(); out.push(...j);
    if (j.length < 1000) return out;
  }
}
async function recordsByIds(coll, ids) {
  const out = [];
  const u = [...new Set(ids.filter(Boolean).map(String))];
  for (let i = 0; i < u.length; i += 60) out.push(...await records(coll, "&id=in." + encodeURIComponent(inList(u.slice(i, i + 60)))));
  return out;
}
async function putRecords(rows) {
  const r = await supa("/rest/v1/records?on_conflict=collection,id", { method: "POST",
    headers: { Prefer: "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(rows.map((x) => ({ collection: x.collection, id: x.id, data: x.data, deleted: false, updated_by: S.uid }))) });
  if (!r.ok) throw new Error("Couldn't save to the CRM (" + r.status + ")");
}

/* ---------------- sign-in ---------------- */
function signedIn(s) {
  const email = String(s?.user?.email || "").toLowerCase();
  if (!TEAM_EMAILS[email]) { saveSession(null); S.login.msg = "That address isn't on the AutoTQ team."; S.login.bad = true; return false; }
  S.me = email; S.uid = TEAM_EMAILS[email]; prefs.email = email; savePrefs();
  return true;
}
async function sendCode() {
  const email = ($("liEmail")?.value || "").trim().toLowerCase();
  S.login.email = email;
  if (!TEAM_EMAILS[email]) return loginMsg("Use your @goldenhourmedical.com address.", true);
  loginMsg("Sending code…");
  try {
    const r = await fetch(SUPA_URL + "/auth/v1/otp", { method: "POST", headers: { apikey: SUPA_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ email, create_user: false }) });
    if (!r.ok) { const j = await r.json().catch(() => ({})); return loginMsg(j.msg || j.error_description || "Couldn't send the code. Try again.", true); }
    S.login.step = "code"; S.login.msg = "Code sent to " + email + "."; S.login.bad = false; render();
  } catch { loginMsg("Network problem. Try again.", true); }
}
async function verify() {
  const email = S.login.email, isPw = S.login.mode === "password";
  const body = isPw ? { email: ($("liEmail")?.value || email).trim().toLowerCase(), password: $("liPw")?.value || "" }
    : { type: "email", email, token: ($("liCode")?.value || "").replace(/\D/g, "") };
  if (!isPw && body.token.length < 6) return loginMsg("Enter the full code from the email.", true);
  loginMsg("Checking…");
  try {
    const r = await fetch(SUPA_URL + (isPw ? "/auth/v1/token?grant_type=password" : "/auth/v1/verify"), { method: "POST",
      headers: { apikey: SUPA_KEY, "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.access_token) return loginMsg(j.msg || j.error_description || (isPw ? "That password didn't work." : "That code didn't work. Check it or send a new one."), true);
    saveSession(j);
    if (!signedIn(j)) return render();
    S.login = { step: "email", email: S.me, mode: "code", msg: "", bad: false };
    await afterSignIn();
  } catch { loginMsg("Network problem. Try again.", true); }
}
function loginMsg(m, bad) { S.login.msg = m; S.login.bad = !!bad; const el = $("liMsg"); if (el) { el.textContent = m; el.className = "note " + (bad ? "bad" : "ok"); el.hidden = !m; } }
function signOut() {
  if (S.run && !S.run.stopped) S.run.stopped = true;
  if (S.session) fetch(SUPA_URL + "/auth/v1/logout", { method: "POST", headers: { apikey: SUPA_KEY, Authorization: "Bearer " + S.session.access_token } }).catch(() => {});
  saveSession(null); S.me = ""; S.uid = ""; S.lists = []; S.listsLoaded = false; S.merges = []; S.people = []; render();
}

/* ---------------- the draft ---------------- */
function readDraft() {
  if (!S.office) {
    const t = S.manualText || "";
    const subject = (t.match(/^Subject:\s*(.*)$/im) || [])[1] || "";
    const bodyText = t.replace(/^Subject:.*(\r?\n)?/im, "");
    S.email = { subject: subject.trim(), html: textToHtml(bodyText), text: bodyText };
    return Promise.resolve();
  }
  const it = Office.context.mailbox.item;
  return new Promise((res) => {
    it.subject.getAsync((a) => {
      it.body.getAsync(Office.CoercionType.Html, (b) => {
        S.email = { subject: a.status === "succeeded" ? String(a.value || "") : "", html: b.status === "succeeded" ? String(b.value || "") : "", text: "" };
        res();
      });
    });
  });
}
function textToHtml(t) {
  const paras = String(t || "").trim().split(/\r?\n\s*\r?\n/).filter(Boolean);
  return paras.map((p) => "<p>" + esc(p).replace(/\r?\n/g, "<br>") + "</p>").join("");
}

/* ---------------- people ---------------- */
async function loadLists() {
  try {
    const rows = await records("lists");
    S.lists = rows.map((r) => Object.assign({ id: r.id }, r.data)).filter((l) => l && !l.removed && (l.items || []).length)
      .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
  } catch (e) { S.lists = []; S.listErr = String(e.message || e); }
  S.listsLoaded = true;
}
async function loadMerges() {
  const since = new Date(Date.now() - 7 * 864e5).toISOString();
  try {
    const r = await supa(`/rest/v1/records?collection=eq.merges&deleted=eq.false&updated_at=gte.${encodeURIComponent(since)}&data-%3E%3Eby=eq.${encodeURIComponent(S.uid)}&select=id,data`);
    S.merges = r.ok ? (await r.json()).map((x) => x.data).filter(Boolean) : [];
  } catch { S.merges = []; }
}
function personFromItem(it) {
  return M.finishPerson({ key: it.pid || it.refId || it.leadId || it.email || it.name, name: it.name || "", company: it.company || "", title: it.role || "",
    email: it.email || "", city: it.city || "", state: String(it.state || "").trim().split(/\s+/)[0] || "",
    note: it.note || "", pid: it.pid || null, leadId: it.leadId || null, refId: it.refId || null, extra: {} });
}
async function pickList(id) {
  S.listId = id; S.people = []; S.allowed = new Set(); S.idx = 0; S.peopleErr = ""; clearDoneRun();
  const L = S.lists.find((l) => l.id === id);
  if (!L) { recompute(); return render(); }
  S.loadingPeople = true; render();
  let ppl = (L.items || []).map(personFromItem);
  try {
    // refresh from the live CRM record (current email, its "verify before sending" note)
    const [pr, lr, cr] = await Promise.all([
      recordsByIds("people", ppl.map((p) => p.pid)), recordsByIds("leads", ppl.map((p) => p.leadId)), recordsByIds("confLeads", ppl.map((p) => p.refId))]);
    const by = (rows) => new Map(rows.map((r) => [String(r.id), r.data || {}]));
    const P = by(pr), Ld = by(lr), C = by(cr);
    ppl = ppl.map((p) => {
      const d = (p.pid && P.get(String(p.pid))) || (p.leadId && Ld.get(String(p.leadId))) || (p.refId && C.get(String(p.refId)));
      if (!d) return p;
      return M.finishPerson(Object.assign(p, {
        email: d.email || p.email, title: p.title || d.role || d.title || "", city: p.city || d.city || "",
        note: [d.note, d.notes, d.status].filter(Boolean).join(" · "), emailNote: d.emailNote || "", cid: d.cid || null,
        conf: d.conf || "", company: p.company || d.company || "",
        name: d.name || p.name, first_name: "", last_name: "",
        pid: p.pid || d.refPid || null,
      }));
    });
  } catch (e) { S.peopleErr = "Showing the list as saved; couldn't refresh from the CRM (" + (e.message || e) + ")."; }
  S.people = ppl; await enrich(); S.loadingPeople = false; recompute(); render();
}
async function setTablePeople(ppl) {
  S.people = ppl; S.allowed = new Set(); S.idx = 0; S.loadingPeople = true; clearDoneRun(); renderPeople(); renderCopies(); renderBar();
  await enrich(true); S.loadingPeople = false; recompute(); renderPeople(); renderCopies(); renderBar(); renderEmailInfo();
}
// suppress list (Max + Hannah can read it) + CRM contacts by email for pasted / file rows (link the activity, pick up their flag)
async function enrich(matchContacts) {
  const emails = [...new Set(S.people.map((p) => p.email).filter(M.emailOk))];
  if (!emails.length) { S.suppressed = new Set(); return; }
  try { S.suppressed = new Set((await recordsByIds("suppress", emails)).map((r) => String(r.id).toLowerCase())); } catch { S.suppressed = new Set(); }
  if (!matchContacts) return;
  try {
    const found = [];
    for (let i = 0; i < emails.length; i += 60) found.push(...await records("people", "&data->>email=in." + encodeURIComponent(inList(emails.slice(i, i + 60)))));
    const byEmail = new Map(found.map((r) => [String(r.data?.email || "").toLowerCase(), Object.assign({ id: r.id }, r.data)]));
    S.people.forEach((p) => {
      const c = byEmail.get(p.email); if (!c) return;
      p.pid = c.id; p.cid = c.cid || null;
      p.note = [p.note, c.note].filter(Boolean).join(" · "); p.emailNote = c.emailNote || "";
      if (!p.title && c.role) p.title = c.role;
    });
  } catch { /* linking is a bonus; the send works without it */ }
}
function parsePaste(text) {
  const t = String(text || "").trim();
  if (!t) return [];
  const first = t.split(/\r?\n/, 1)[0];
  const hasHeader = /[a-z]/i.test(first) && !/@/.test(first);
  if (hasHeader) { const { headers, rows } = M.parseTable(t); return M.tableToPeople(headers, rows); }
  // no header: lines like "Dana Smith <dana@x.co>" or "dana@x.co, Dana Smith, Acme"
  return t.split(/\r?\n/).map((line, i) => {
    const em = (line.match(/[^\s<>(),;:"]+@[^\s<>(),;:"]+/) || [])[0] || "";
    const rest = line.replace(em, "").replace(/[<>"]/g, "").split(/[,\t;]/).map((x) => x.trim()).filter(Boolean);
    return M.finishPerson({ key: "p" + i, email: em, name: rest[0] || "", company: rest[1] || "", extra: {} });
  }).filter((p) => p.email || p.name);
}
async function readFile(file) {
  S.fileName = file.name; S.fileErr = "";
  try {
    if (/\.xlsx?$/i.test(file.name)) {
      const X = await import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs");
      const wb = X.read(await file.arrayBuffer());
      const rows = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: "" })
        .map((r) => r.map((c) => String(c ?? "").trim())).filter((r) => r.some(Boolean));
      S.filePeople = M.tableToPeople(rows[0] || [], rows.slice(1));
    } else {
      const { headers, rows } = M.parseTable(await file.text());
      S.filePeople = M.tableToPeople(headers, rows);
    }
  } catch (e) { S.filePeople = []; S.fileErr = "Couldn't read that file (" + (e.message || e) + ")."; }
  await setTablePeople(S.filePeople);
}

/* ---------------- derived state ---------------- */
const midnightIso = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d.toISOString(); };
const htmlOut = () => (S.optOut ? M.addOptOut(S.email.html, S.optOutLine) : S.email.html);
function sameEmail(m) { return m.hash === S.hash || (m.subject && m.subject.trim() === S.email.subject.trim()); }
function recompute() {
  S.fields = M.fieldsIn(S.email.subject, S.email.html);
  S.unknown = M.unknownFields(S.fields, S.people);
  S.copyWarn = M.copyCheck(S.email.subject + " " + S.email.html);
  S.cid = M.hasCidImages(S.email.html);
  S.hash = M.contentHash(S.email.subject, S.email.html, S.optOut ? S.optOutLine : "", S.fallbackThere ? 1 : 0);
  const sentBefore = new Set();
  S.merges.filter(sameEmail).forEach((m) => (m.items || []).forEach((it) => { if (it.status === "sent") sentBefore.add(String(it.email).toLowerCase()); }));
  if (S.merge) (S.merge.items || []).forEach((it) => { if (it.status === "sent") sentBefore.add(String(it.email).toLowerCase()); });
  S.rows = M.classify(S.people, { fields: S.fields, suppressed: S.suppressed, sentEmails: sentBefore, fallbackThere: S.fallbackThere, allowed: S.allowed });
  S.ready = S.rows.filter((r) => r.status === "ready");
  S.sentToday = M.sentToday(S.merge && !S.merges.includes(S.merge) ? S.merges.concat([S.merge]) : S.merges, S.me, midnightIso());
  S.left = Math.max(0, Math.min(S.cap, CAP_MAX) - S.sentToday);
  if (S.idx >= S.rows.length) S.idx = Math.max(0, S.rows.length - 1);
}
function copyFor(p) {
  return { to: p.email, subject: M.render(S.email.subject, p, { html: false, fallbackThere: S.fallbackThere }).out.replace(/\s+/g, " ").trim(),
    html: M.render(htmlOut(), p, { html: true, fallbackThere: S.fallbackThere }).out };
}
function blockers() {
  const b = [];
  if (!S.email.subject.trim()) b.push("The email needs a subject.");
  if (!M.textOf(S.email.html)) b.push("The email is empty.");
  if (S.unknown.length) b.push(S.unknown.map((n) => "{{" + n + "}}").join(", ") + (S.unknown.length > 1 ? " aren't fields" : " isn't a field") + " on this list.");
  if (S.cid) b.push("Pictures pasted into the email can't go out in merged copies yet. Take the picture out or link to it.");
  return b;
}

/* ---------------- sending ---------------- */
function clearDoneRun() { if (S.run && (S.run.finished || S.run.stopped)) S.run = null; S.confirm = false; }
async function sendOne(to, subject, html) {
  const r = await supa("/functions/v1/send-invoice", { method: "POST", body: JSON.stringify({ action: "notify", to, subject, html, from: S.me }) });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.ok) throw new Error(j.error || ("send failed (" + r.status + ")"));
  if (String(j.from || "").toLowerCase() !== S.me) { const e = new Error("It went out from " + (j.from || "another mailbox") + ", not " + S.me + "."); e.wrongFrom = true; throw e; }
  return j;
}
function ensureMerge() {
  if (S.merge && S.merge.hash === S.hash) return S.merge;
  const resume = S.merges.filter((m) => m.hash === S.hash).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))[0];
  if (resume) { S.merge = resume; return resume; }
  const L = S.lists.find((l) => l.id === S.listId);
  S.merge = { id: "mm_" + uid() + uid(), by: S.uid, from: S.me, subject: S.email.subject, preview: M.textOf(S.email.html).slice(0, 400), hash: S.hash,
    optOut: S.optOut ? S.optOutLine : "", fallbackThere: S.fallbackThere, createdAt: new Date().toISOString(),
    source: { kind: S.src, listId: S.src === "list" ? S.listId : "", name: S.src === "list" ? (L?.name || "") : S.src === "file" ? S.fileName : "pasted" },
    items: [], tests: [], sent: 0, failed: 0 };
  S.merges.push(S.merge);
  return S.merge;
}
function mergeItem(m, p) {
  let it = m.items.find((x) => x.email === p.email);
  if (!it) { it = { email: p.email, name: p.name || "", pid: p.pid || null, refId: p.refId || null, leadId: p.leadId || null, status: "", why: "" }; m.items.push(it); }
  return it;
}
async function saveMerge() {
  const m = S.merge; if (!m) return;
  m.updatedAt = new Date().toISOString();
  m.sent = m.items.filter((i) => i.status === "sent").length; m.failed = m.items.filter((i) => i.status === "failed").length;
  await putRecords([{ collection: "merges", id: m.id, data: m }]);
}
// CRM Home "TQbot this week" card: one botLog record per batch (merge) per local day, written after the sends went out.
// Resuming the same batch later that day updates the same record (n and names are that day's totals for the batch).
const localDay = (iso) => { const d = new Date(iso), z = (n) => String(n).padStart(2, "0"); return d.getFullYear() + "-" + z(d.getMonth() + 1) + "-" + z(d.getDate()); };
async function logBot(m, sentNow) {
  const days = [...new Set(sentNow.map((it) => localDay(it.at)))];
  if (!days.length) return;
  m.botLog = m.botLog || {};
  const rows = days.map((day) => {
    const sent = m.items.filter((it) => it.status === "sent" && it.at && localDay(it.at) === day).sort((a, b) => String(a.at).localeCompare(String(b.at)));
    const id = m.botLog[day] || (m.botLog[day] = "bl_" + uid() + uid());
    return { collection: "botLog", id, data: { id, by: S.uid, at: sent[sent.length - 1].at, day, kind: "email_sent", src: "mailmerge",
      n: sent.length, names: sent.map((it) => it.name || it.email).slice(0, 200), extra: m.subject } };
  });
  rows.push({ collection: "merges", id: m.id, data: m });
  await putRecords(rows);
}
function activityFor(p, subject, at) {
  const a = { id: uid(), type: "email", note: "Emailed: " + subject + " · mail merge", date: at, by: S.uid, norm: 2, merge: S.merge.id };
  if (p.pid) { a.pid = p.pid; a.cid = p.cid || null; }
  else if (p.refId) Object.assign(a, { refId: p.refId, name: p.name, company: p.company || "", conf: p.conf || "", pid: null });
  else if (p.leadId) Object.assign(a, { leadId: p.leadId, name: p.name, company: p.company || "", conf: p.conf || "", pid: null });
  else return null;
  return a;
}
async function sendTest() {
  const cur = S.rows[S.idx], p = cur && cur.status === "ready" ? cur : S.ready[0];
  if (!p || S.busy) return;
  S.busy = "test"; S.testMsg = null; renderBar();
  const c = copyFor(p);
  try {
    await sendOne(S.me, "[Test] " + c.subject, c.html);
    S.testedHash = S.hash;
    const m = ensureMerge(); m.tests.push({ at: new Date().toISOString(), to: S.me, email: p.email });
    S.testMsg = { ok: true, text: "Test sent to " + S.me + " (" + (p.name || p.email) + "'s copy). Check it, then send." };
    saveMerge().catch(() => {});
  } catch (e) {
    S.testMsg = { ok: false, text: "Test not sent: " + (e.message || e) };
  }
  S.busy = ""; recompute(); renderBar();
}
async function runBatch() {
  S.confirm = false;
  const queue = S.ready.slice(0, S.left);
  if (!queue.length) return renderBar();
  const m = ensureMerge();
  S.rows.forEach((p) => { const it = mergeItem(m, p); if (it.status !== "sent") { it.status = p.status; it.why = p.why || ""; } });
  S.run = { total: queue.length, done: 0, failed: 0, nextAt: 0, paused: false, stopped: false, msg: "" };
  await saveMerge().catch(() => {});
  let failsInRow = 0;
  const sentNow = [];
  for (let i = 0; i < queue.length; i++) {
    if (S.run.stopped || S.run.paused) break;
    const p = queue[i], it = mergeItem(m, p), c = copyFor(p);
    if (it.status === "sent") continue;
    S.run.current = p.name || p.email; renderBar();
    try {
      await sendOne(p.email, c.subject, c.html);
      it.status = "sent"; it.at = new Date().toISOString(); it.err = ""; S.run.done++; failsInRow = 0; sentNow.push(it);
      const a = activityFor(p, c.subject, it.at); if (a) it.act = a.id;
      const rows = [{ collection: "merges", id: m.id, data: m }]; if (a) rows.push({ collection: "activities", id: a.id, data: a });
      m.updatedAt = it.at; m.sent = m.items.filter((x) => x.status === "sent").length;
      await putRecords(rows).catch((e) => { S.run.msg = "Sent, but the CRM log didn't save: " + (e.message || e); });
    } catch (e) {
      it.status = "failed"; it.err = String(e.message || e).slice(0, 200); S.run.failed++; failsInRow++;
      saveMerge().catch(() => {});
      if (e.wrongFrom) { S.run.stopped = true; S.run.msg = "Stopped. " + it.err; break; }
      if (failsInRow >= 2) { S.run.stopped = true; S.run.msg = "Stopped after two failed sends in a row. Last error: " + it.err; break; }
    }
    recompute(); renderPeople(); renderBar();
    if (i < queue.length - 1 && !S.run.stopped && !S.run.paused) await wait(GAP_MS);
  }
  S.run.finished = !S.run.paused; S.run.current = "";
  if (sentNow.length) await logBot(m, sentNow).catch((e) => { S.run.msg = (S.run.msg ? S.run.msg + " " : "") + "The CRM Home log didn't save: " + (e.message || e); });
  recompute(); render();
}
function wait(ms) {
  return new Promise((res) => {
    S.run.nextAt = Date.now() + ms;
    const t = setInterval(() => {
      if (!S.run || S.run.paused || S.run.stopped || Date.now() >= S.run.nextAt) { clearInterval(t); S.run.nextAt = 0; renderBar(); res(); return; }
      renderBar();
    }, 250);
  });
}

/* ---------------- rendering ---------------- */
function render() {
  const app = $("app");
  if (!S.session || !S.me) { app.innerHTML = signinHTML(); const f = $(S.login.step === "code" ? "liCode" : "liEmail"); if (f) f.focus(); return; }
  app.innerHTML = '<div class="top"><img src="assets/icon-32.png" alt=""><b>AutoTQ Merge</b><div class="me">' + esc(S.me) +
    '<br><button class="lnk" data-act="signout">Sign out</button></div></div>' +
    '<section class="card" id="emailCard"></section><section class="card" id="peopleCard"></section>' +
    '<section class="card" id="copiesCard"></section><details class="card fold" id="optsCard"' + (S.optsOpen ? " open" : "") + '></details>' +
    '<div class="bar"><div class="in" id="bar"></div></div>';
  renderEmail(); renderPeople(); renderCopies(); renderOpts(); renderBar();
}
function signinHTML() {
  const L = S.login, code = L.step === "code", pw = L.mode === "password";
  return '<div class="signin"><div class="top"><img src="assets/icon-32.png" alt=""><b>AutoTQ Merge</b></div><div class="card">' +
    '<label class="lbl" for="liEmail">Your email</label><input class="f" id="liEmail" type="email" autocomplete="username" value="' + esc(L.email) + '"' + (code ? " readonly" : "") + ">" +
    (code ? '<label class="lbl" for="liCode">Code from the email</label><input class="f" id="liCode" inputmode="numeric" autocomplete="one-time-code">' +
      '<button class="btn pri" data-act="verify">Sign in</button><div class="mt sub"><button class="lnk" data-act="resend">Send a new code</button> · <button class="lnk" data-act="restart">Use another address</button></div>'
      : pw ? '<label class="lbl" for="liPw">Password</label><input class="f" id="liPw" type="password" autocomplete="current-password">' +
        '<button class="btn pri" data-act="verify">Sign in</button><div class="mt sub"><button class="lnk" data-act="usecode">Email me a code instead</button></div>'
      : '<button class="btn pri" data-act="sendcode">Email me a code</button><div class="mt sub"><button class="lnk" data-act="usepw">Use my CRM password</button></div>') +
    '<div id="liMsg" class="note ' + (L.bad ? "bad" : "ok") + '"' + (L.msg ? "" : " hidden") + ">" + esc(L.msg) + "</div></div></div>";
}
function renderEmail() {
  const el = $("emailCard"); if (!el) return;
  el.innerHTML = '<div class="ch"><h2>Email</h2><div class="r">' + (S.office ? '<button class="ibtn" data-act="reread" title="Read the draft again" aria-label="Read the draft again">' + ICON.refresh + "</button>" : "") + "</div></div>" +
    (S.office ? "" : '<textarea class="f" id="manual" rows="8" placeholder="Subject: Quick question, {{first_name}}&#10;&#10;Hi {{first_name}},">' + esc(S.manualText) + "</textarea>") +
    '<div id="emailInfo"></div>';
  renderEmailInfo();
}
function renderEmailInfo() {
  const el = $("emailInfo"); if (!el) return;
  const subj = S.email.subject.trim();
  const fields = S.fields.map((f) => {
    const bad = S.unknown.includes(f.name);
    return '<span class="chip ' + (bad ? "bad" : "ok") + '">' + esc(f.name) + (f.fallback ? " | " + esc(f.fallback) : "") + "</span>";
  }).join("");
  el.innerHTML = (S.office ? '<div class="subj">' + (subj ? esc(subj) : '<span class="muted">No subject yet</span>') + "</div>" : "") +
    '<div class="chips">' + (fields || '<span class="chip">No {{fields}}: everyone gets the same email</span>') + "</div>" +
    blockers().map((b) => '<div class="note bad">' + esc(b) + "</div>").join("") +
    (S.copyWarn.length ? '<div class="note warn">Off our copy rules: ' + S.copyWarn.map((c) => esc(c.phrase) + " (" + c.why + ")").join(", ") + "</div>" : "");
}
function renderPeople() {
  const el = $("peopleCard"); if (!el) return;
  const seg = (k, t) => '<button class="' + (S.src === k ? "on" : "") + '" data-act="src" data-k="' + k + '">' + t + "</button>";
  const keepPaste = $("paste") && S.src === "paste";
  let body = "";
  if (S.src === "list") {
    body = !S.listsLoaded ? '<div class="empty">Loading your CRM lists…</div>'
      : !S.lists.length ? '<div class="empty">No saved lists in the CRM yet.' + (S.listErr ? " " + esc(S.listErr) : "") + "</div>"
      : '<select class="f" id="listSel"><option value="">Pick a CRM list…</option>' + S.lists.map((l) => '<option value="' + esc(l.id) + '"' + (l.id === S.listId ? " selected" : "") + ">" +
        esc(l.name || "Untitled list") + " (" + (l.items || []).length + ")" + "</option>").join("") + "</select>";
  } else if (S.src === "file") {
    body = '<input class="f" id="fileIn" type="file" accept=".csv,.tsv,.txt,.xlsx,.xls">' + (S.fileName ? '<div class="sub mt">' + esc(S.fileName) + "</div>" : "") +
      (S.fileErr ? '<div class="note bad">' + esc(S.fileErr) + "</div>" : "");
  }
  el.innerHTML = '<div class="ch"><h2>People</h2></div><div class="seg">' + seg("list", "CRM list") + seg("paste", "Paste") + seg("file", "CSV / Excel") + "</div>" +
    (S.src === "paste" ? '<div id="pasteWrap"></div>' : body) + '<div id="pcounts"></div>';
  if (S.src === "paste") {
    const w = $("pasteWrap");
    w.innerHTML = '<textarea class="f" id="paste" rows="5" placeholder="Email, First name, Company&#10;dana@acme.com, Dana, Acme"></textarea>';
    $("paste").value = S.pasteText;
    if (keepPaste) $("paste").focus();
  }
  renderCounts();
}
function renderCounts() {
  const el = $("pcounts"); if (!el) return;
  if (S.loadingPeople) { el.innerHTML = '<div class="empty mt">Checking each person…</div>'; return; }
  if (!S.people.length) { el.innerHTML = S.peopleErr ? '<div class="note warn">' + esc(S.peopleErr) + "</div>" : ""; return; }
  const held = S.rows.filter((r) => r.status === "held");
  const sent = held.filter((r) => r.why === "already sent").length;
  el.innerHTML = '<div class="counts"><span class="chip ok">' + S.ready.length + " ready</span>" +
    (held.length - sent ? '<span class="chip warn">' + (held.length - sent) + " held</span>" : "") +
    (sent ? '<span class="chip">' + sent + " already sent</span>" : "") + "</div>" +
    (S.peopleErr ? '<div class="note warn">' + esc(S.peopleErr) + "</div>" : "") +
    (held.length ? '<details class="held"' + (S.heldOpen ? " open" : "") + ' id="heldFold"><summary>Who\'s held and why</summary>' + held.map((r) => {
      const i = S.rows.indexOf(r);
      return '<div class="hrow"><div class="who"><div><b>' + esc(r.name || r.email || "No name") + "</b></div><div class=\"sub\">" + esc(r.email || r.company || "") + "</div></div>" +
        '<span class="chip ' + (r.why === "already sent" ? "" : "warn") + '">' + esc(r.why) + "</span>" +
        (r.canAllow ? '<button class="btn sm" data-act="allow" data-i="' + i + '">Send anyway</button>' : "") + "</div>";
    }).join("") + "</details>" : "") +
    (S.allowed.size ? '<div class="sub mt">' + S.allowed.size + ' ticked in by you · <button class="lnk" data-act="unallow">Hold them again</button></div>' : "");
}
function renderCopies() {
  const el = $("copiesCard"); if (!el) return;
  if (!S.rows.length) { el.innerHTML = '<div class="ch"><h2>Copies</h2></div><div class="empty">Pick people to see each copy.</div>'; return; }
  const p = S.rows[S.idx], c = copyFor(p);
  el.innerHTML = '<div class="ch"><h2>Copies</h2><div class="r pv"><button class="ibtn" data-act="prev" aria-label="Previous copy">' + ICON.left +
    '</button><span class="n">' + (S.idx + 1) + " of " + S.rows.length + '</span><button class="ibtn" data-act="next" aria-label="Next copy">' + ICON.right + "</button></div></div>" +
    '<div class="to">To: ' + esc(p.name ? p.name + " <" + (p.email || "no email") + ">" : p.email || "no email") +
    (p.status === "held" ? ' <span class="chip warn">' + esc(p.why) + "</span>" : "") + "</div>" +
    '<div class="subj">' + esc(c.subject || "(no subject)") + '</div><iframe class="body" sandbox="" title="This copy" id="pvFrame"></iframe>';
  $("pvFrame").srcdoc = '<base target="_blank"><style>body{font:14px/1.5 -apple-system,Segoe UI,Arial,sans-serif;color:#1A1916;margin:10px}</style>' + c.html;
}
function renderOpts() {
  const el = $("optsCard"); if (!el) return;
  el.innerHTML = "<summary>Options</summary>" +
    '<label class="row"><input type="checkbox" id="optThere"' + (S.fallbackThere ? " checked" : "") + '> No first name: write "there"</label>' +
    '<label class="row"><input type="checkbox" id="optOut"' + (S.optOut ? " checked" : "") + "> Opt-out line at the end</label>" +
    (S.optOut ? '<input class="f mt" id="optLine" value="' + esc(S.optOutLine) + '">' : "") +
    '<label class="lbl" for="optCap">Most I send in a day (up to 50)</label><input class="f" id="optCap" type="number" min="1" max="50" value="' + S.cap + '" style="width:90px">';
}
function renderBar() {
  const el = $("bar"); if (!el) return;
  const R = S.run;
  if (R && !R.finished && !R.stopped) {
    const pct = Math.round(100 * (R.done + R.failed) / Math.max(1, R.total));
    const secs = R.nextAt ? Math.max(0, Math.ceil((R.nextAt - Date.now()) / 1000)) : 0;
    el.innerHTML = '<div class="prog"><i style="width:' + pct + '%"></i></div><div class="meta"><span>' +
      (R.paused ? "Paused · " + R.done + " of " + R.total + " sent" : "Sent " + R.done + " of " + R.total + (R.failed ? " · " + R.failed + " failed" : "") +
        (R.current && !secs ? " · sending to " + esc(R.current) : secs ? " · next in " + secs + " s" : "")) +
      "</span><span>Keep this pane open</span></div>" + (R.msg ? '<div class="note warn" style="margin:0 0 8px">' + esc(R.msg) + "</div>" : "") +
      '<div class="btns">' + (R.paused ? '<button class="btn pri" data-act="resume">Keep sending</button>' : '<button class="btn" data-act="pause">Pause</button>') + "</div>";
    return;
  }
  const b = blockers(), n = Math.min(S.ready.length, S.left);
  const tested = S.testedHash && S.testedHash === S.hash;
  const canTest = !b.length && S.ready.length > 0 && !S.busy;
  const canSend = !b.length && tested && n > 0 && !S.busy;
  const meta = '<div class="meta"><span>Today ' + S.sentToday + " of " + Math.min(S.cap, CAP_MAX) + " sent</span><span>" + (!tested && canTest ? "Send yourself a test first" : "One every " + Math.round(GAP_MS / 1000) + " s") + "</span></div>";
  let note = "";
  if (R && (R.finished || R.stopped)) note = '<div class="note ' + (R.stopped ? "bad" : "ok") + '" style="margin:0 0 8px">' +
    (R.stopped ? esc(R.msg || "Stopped.") : "Done: " + R.done + " sent" + (R.failed ? ", " + R.failed + " failed" : "") + "." + (R.msg ? " " + esc(R.msg) : "")) + "</div>";
  else if (S.testMsg) note = '<div class="note ' + (S.testMsg.ok ? "ok" : "bad") + '" style="margin:0 0 8px">' + esc(S.testMsg.text) + "</div>";
  if (S.ready.length > S.left && S.ready.length) note += '<div class="note warn" style="margin:0 0 8px">' + (S.left ? S.left + " can go today; the other " + (S.ready.length - S.left) + " wait for tomorrow." : "Today's limit is used up. These go tomorrow.") + "</div>";
  if (S.confirm && canSend) {
    el.innerHTML = meta + note + '<div class="btns"><button class="btn" data-act="cancel">Cancel</button><button class="btn pri" data-act="go">Yes, send ' + n + " from " + esc(S.me.split("@")[0]) + "@</button></div>";
    return;
  }
  el.innerHTML = meta + note + '<div class="btns"><button class="btn" data-act="test"' + (canTest ? "" : " disabled") + ">" + (S.busy === "test" ? "Sending test…" : "Send a test to me") + "</button>" +
    '<button class="btn pri" data-act="send"' + (canSend ? "" : " disabled") + ' title="' + (tested ? "" : "Send yourself a test of this email first") + '">Send ' + (n || "") + "</button></div>";
}

/* ---------------- events ---------------- */
document.addEventListener("click", async (ev) => {
  const t = ev.target.closest("[data-act]"); if (!t) return;
  const a = t.dataset.act;
  if (a === "sendcode" || a === "resend") return sendCode();
  if (a === "verify") return verify();
  if (a === "usepw") { S.login.mode = "password"; S.login.email = $("liEmail")?.value || S.login.email; S.login.msg = ""; return render(); }
  if (a === "usecode") { S.login.mode = "code"; S.login.msg = ""; return render(); }
  if (a === "restart") { S.login.step = "email"; S.login.msg = ""; return render(); }
  if (a === "signout") return signOut();
  if (a === "reread") { await readDraft(); recompute(); renderEmailInfo(); renderCounts(); renderCopies(); renderBar(); return; }
  if (a === "src") { S.src = t.dataset.k; prefs.src = S.src; savePrefs(); S.people = []; S.listId = ""; S.idx = 0; S.allowed = new Set(); S.peopleErr = "";
    if (S.src === "paste" && S.pasteText) S.people = parsePaste(S.pasteText);
    if (S.src === "file" && S.filePeople.length) S.people = S.filePeople;
    recompute(); renderPeople(); renderCopies(); renderBar(); renderEmailInfo(); return; }
  if (a === "allow") { const r = S.rows[+t.dataset.i]; if (r) S.allowed.add(r.email); S.heldOpen = true; recompute(); renderCounts(); renderCopies(); renderBar(); return; }
  if (a === "unallow") { S.allowed = new Set(); recompute(); renderCounts(); renderCopies(); renderBar(); return; }
  if (a === "prev" || a === "next") { if (S.rows.length) S.idx = (S.idx + (a === "next" ? 1 : -1) + S.rows.length) % S.rows.length; renderCopies(); return; }
  if (a === "test") return sendTest();
  if (a === "send") { S.confirm = true; return renderBar(); }
  if (a === "cancel") { S.confirm = false; return renderBar(); }
  if (a === "go") return runBatch();
  if (a === "pause") { if (S.run) S.run.paused = true; return renderBar(); }
  if (a === "resume") { S.run = null; recompute(); S.confirm = true; return renderBar(); }
});
document.addEventListener("toggle", (ev) => {
  if (ev.target.id === "heldFold") S.heldOpen = ev.target.open;
  if (ev.target.id === "optsCard") S.optsOpen = ev.target.open;
}, true);
document.addEventListener("keydown", (ev) => {
  if (ev.key !== "Enter") return;
  if (ev.target.id === "liEmail" && S.login.mode === "code" && S.login.step === "email") { ev.preventDefault(); sendCode(); }
  else if (ev.target.id === "liCode" || ev.target.id === "liPw") { ev.preventDefault(); verify(); }
});
let pasteTimer = null;
document.addEventListener("input", (ev) => {
  const id = ev.target.id;
  if (id === "paste") {
    S.pasteText = ev.target.value; clearTimeout(pasteTimer);
    pasteTimer = setTimeout(() => setTablePeople(parsePaste(S.pasteText)), 350);
  } else if (id === "manual") {
    S.manualText = ev.target.value; prefs.manualText = S.manualText; savePrefs();
    readDraft().then(() => { recompute(); renderEmailInfo(); renderCounts(); renderCopies(); renderBar(); });
  } else if (id === "optLine") {
    S.optOutLine = ev.target.value; prefs.optOutLine = S.optOutLine; savePrefs(); recompute(); renderCopies(); renderBar();
  }
});
document.addEventListener("change", async (ev) => {
  const id = ev.target.id;
  if (id === "listSel") return pickList(ev.target.value);
  if (id === "fileIn" && ev.target.files[0]) return readFile(ev.target.files[0]);
  if (id === "optThere") { S.fallbackThere = ev.target.checked; prefs.fallbackThere = S.fallbackThere; }
  else if (id === "optOut") { S.optOut = ev.target.checked; prefs.optOut = S.optOut; renderOpts(); }
  else if (id === "optCap") { S.cap = Math.max(1, Math.min(CAP_MAX, Math.round(+ev.target.value || CAP_MAX))); ev.target.value = S.cap; prefs.cap = S.cap; }
  else return;
  savePrefs(); recompute(); renderCounts(); renderCopies(); renderBar();
});
// coming back to the pane after editing the draft: read it again
window.addEventListener("focus", () => { if (S.office && S.me && !(S.run && !S.run.finished && !S.run.stopped)) { readDraft().then(() => { recompute(); renderEmailInfo(); renderCounts(); renderCopies(); renderBar(); }); } });

/* ---------------- boot ---------------- */
async function afterSignIn() {
  render();
  await Promise.all([readDraft(), loadLists(), loadMerges()]);
  if (S.src === "paste" && S.pasteText) S.people = parsePaste(S.pasteText);
  recompute(); render();
}
async function boot(info) {
  S.office = !!(info && window.Office && info.host === Office.HostType.Outlook && Office.context?.mailbox?.item?.body);
  if (S.office && !S.login.email) S.login.email = String(Office.context.mailbox.userProfile?.emailAddress || "").toLowerCase();
  if (S.session && S.session.access_token) {
    // a stored session: check it's still good (refreshes if needed)
    const r = await supa("/auth/v1/user").catch(() => null);
    const u = r && r.ok ? await r.json().catch(() => null) : null;
    if (u && signedIn({ user: u })) return afterSignIn();
    saveSession(null);
  }
  render();
}
let booted = false;
const bootOnce = (info) => { if (booted) return; booted = true; boot(info); };
if (window.Office && Office.onReady) Office.onReady(bootOnce); else bootOnce(null);
setTimeout(() => bootOnce(null), 4000); // office.js didn't answer (opened outside Office, or it failed to load)
