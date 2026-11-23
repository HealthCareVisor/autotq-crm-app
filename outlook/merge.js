// AutoTQ Merge — pure logic (no DOM, no network). Tested by merge_test.ts.
// Person = {key, email, name, first_name, last_name, company, title, city, state, note, emailNote,
//           pid, cid, leadId, refId, conf, extra:{}}

export const KNOWN_FIELDS = ["first_name", "last_name", "name", "company", "title", "email", "city", "state"];

// {{field}} or {{field|fallback}}; Outlook may wrap part of a placeholder in formatting tags, so tags inside are ignored
const PH = /\{\{((?:[^{}<]|<[^>]*>)*?)\}\}/g;
const stripTags = (s) => String(s).replace(/<[^>]*>/g, "");
const decode = (s) => String(s).replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function parsePh(inner) {
  const raw = decode(stripTags(inner));
  const bar = raw.indexOf("|");
  const name = (bar < 0 ? raw : raw.slice(0, bar)).trim().toLowerCase().replace(/\s+/g, "_");
  const fallback = bar < 0 ? null : raw.slice(bar + 1).trim();
  return { name, fallback };
}

export function fieldsIn(subject, html) {
  const out = [], seen = new Set();
  for (const src of [String(subject || ""), String(html || "")]) {
    for (const m of src.matchAll(PH)) {
      const f = parsePh(m[1]);
      if (!f.name || seen.has(f.name)) continue;
      seen.add(f.name); out.push(f);
    }
  }
  return out;
}

export function valueOf(person, name) {
  const p = person || {};
  const v = KNOWN_FIELDS.includes(name) ? p[name] : (p.extra || {})[name];
  return v == null ? "" : String(v).trim();
}

export function render(tpl, person, opts = {}) {
  const missing = [];
  const out = String(tpl || "").replace(PH, (_m, inner) => {
    const f = parsePh(inner);
    let v = valueOf(person, f.name);
    if (!v && f.fallback != null && f.fallback !== "") v = f.fallback;
    if (!v && f.name === "first_name" && opts.fallbackThere) v = "there";
    if (!v) { if (!missing.includes(f.name)) missing.push(f.name); return ""; }
    return opts.html ? esc(v) : v;
  });
  return { out, missing };
}

const HONORIFIC = /^(dr|mr|mrs|ms|miss|mx|prof|sgt|lt|capt|chief|officer)\.?$/i;
const fixCase = (s) => {
  s = String(s || "").trim();
  if (!s) return s;
  if (s === s.toUpperCase() || s === s.toLowerCase())
    return s.toLowerCase().replace(/(^|[\s\-'’])([a-z])/g, (_m, a, b) => a + b.toUpperCase());
  return s;
};

export function splitName(full) {
  let s = String(full || "").replace(/\s+/g, " ").trim();
  if (!s) return { first: "", last: "" };
  if (s.includes(",")) { const [l, f] = s.split(",", 2).map((x) => x.trim()); if (f) s = f + " " + l; }
  const parts = s.split(" ").filter((w) => w && !HONORIFIC.test(w));
  const first = fixCase(parts[0] || ""), last = fixCase(parts.slice(1).join(" "));
  return { first, last };
}

export function parseTable(text) {
  const t = String(text || "").replace(/^﻿/, "");
  const firstLine = t.split(/\r?\n/, 1)[0] || "";
  const d = (firstLine.match(/\t/g) || []).length > (firstLine.match(/,/g) || []).length ? "\t" : ",";
  const rows = []; let row = [], cell = "", q = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (q) {
      if (c === '"') { if (t[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"' && cell === "") q = true;
    else if (c === d) { row.push(cell); cell = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && t[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += c;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  const clean = rows.map((r) => r.map((x) => x.trim())).filter((r) => r.some((x) => x !== ""));
  return { headers: clean[0] || [], rows: clean.slice(1) };
}

const ALIAS = {
  email: ["email", "e-mail", "email address", "e-mail address", "mail", "work email"],
  first_name: ["first_name", "first name", "firstname", "first", "given name"],
  last_name: ["last_name", "last name", "lastname", "last", "surname", "family name"],
  name: ["name", "full name", "fullname", "contact", "contact name", "person"],
  company: ["company", "company name", "organization", "organisation", "org", "account", "employer", "business"],
  title: ["title", "job title", "role", "position"],
  city: ["city", "town"],
  state: ["state", "st", "province", "region"],
};
const snake = (h) => String(h).trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");

export function tableToPeople(headers, rows) {
  const map = headers.map((h) => {
    const k = String(h).trim().toLowerCase();
    for (const [f, names] of Object.entries(ALIAS)) if (names.includes(k)) return f;
    return "x:" + snake(h);
  });
  return rows.map((r, i) => {
    const p = { key: "row" + i, extra: {} };
    map.forEach((f, j) => {
      const v = String(r[j] ?? "").trim();
      if (!v) return;
      if (f.startsWith("x:")) { if (f.length > 2) p.extra[f.slice(2)] = v; } else if (!p[f]) p[f] = v;
    });
    finishPerson(p);
    return p;
  });
}

export function finishPerson(p) {
  p.extra = p.extra || {};
  p.email = String(p.email || "").trim().toLowerCase();
  if (p.name && (!p.first_name || !p.last_name)) {
    const s = splitName(p.name);
    if (!p.first_name) p.first_name = s.first;
    if (!p.last_name) p.last_name = s.last;
  }
  p.first_name = fixCase(p.first_name || "");
  p.last_name = fixCase(p.last_name || "");
  p.name = [p.first_name, p.last_name].filter(Boolean).join(" ") || String(p.name || "").trim();
  if (p.state && String(p.state).trim().length === 2) p.state = String(p.state).trim().toUpperCase();
  if (!p.key) p.key = p.email || p.name;
  return p;
}

export const emailOk = (s) => /^[^@\s<>(),;:"]+@[^@\s<>(),;:"]+\.[a-z]{2,}$/i.test(String(s || "").trim());

export function emailFlag(text) {
  const t = String(text || "");
  if (/verify before sending/i.test(t)) return "verify before sending";
  if (/company inbox/i.test(t)) return "company inbox";
  return "";
}

// people → each with status "ready" | "held", why (exact reason text), canAllow (Max may tick it back in)
export function classify(people, opts) {
  const o = Object.assign({ fields: [], suppressed: new Set(), sentEmails: new Set(), fallbackThere: true, allowed: new Set() }, opts || {});
  const seen = new Set();
  return people.map((p0) => {
    const p = Object.assign({}, p0, { status: "ready", why: "", canAllow: false });
    const e = String(p.email || "").trim().toLowerCase();
    const hold = (why, canAllow = false) => { p.status = "held"; p.why = why; p.canAllow = canAllow; return p; };
    if (!e) return hold("no email");
    if (!emailOk(e)) return hold("not a valid email");
    if (seen.has(e)) return hold("duplicate");
    seen.add(e);
    if (o.suppressed.has(e)) return hold("unsubscribed or bounced");
    if (o.sentEmails.has(e)) return hold("already sent");
    for (const f of o.fields) {
      if (valueOf(p, f.name)) continue;
      if (f.fallback) continue;
      if (f.name === "first_name" && o.fallbackThere) continue;
      return hold("missing " + f.name.replace(/_/g, " "));
    }
    const flag = emailFlag([p.note, p.emailNote, p.flag].filter(Boolean).join(" "));
    if (flag && !o.allowed.has(e)) return hold(flag, true);
    return p;
  });
}

export function unknownFields(fields, people) {
  const extra = new Set();
  (people || []).forEach((p) => Object.keys((p && p.extra) || {}).forEach((k) => extra.add(k)));
  return fields.map((f) => f.name).filter((n) => !KNOWN_FIELDS.includes(n) && !extra.has(n));
}

// house style: claims retired in the 2026-10-17 / 2026-11-16 copy passes + banned buzzwords
const COPY_RULES = [
  [/revolutioni[sz]ing/i, "revolutionizing", "buzzword"],
  [/cutting[- ]edge/i, "cutting-edge", "buzzword"],
  [/empowering/i, "empowering", "buzzword"],
  [/game[- ]chang/i, "game-changing", "buzzword"],
  [/92\s*%\s*first[- ]try/i, "92% first try", "retired claim"],
  [/>\s*84\s*%|84\s*%\s*(of\s+)?(tourniquets\s+)?misapplied/i, ">84% misapplied", "retired claim"],
  [/\b2\s*x\b|\btwice as (fast|quick)/i, "2x", "retired claim"],
];
export function copyCheck(text) {
  const t = decode(stripTags(String(text || "")));
  return COPY_RULES.filter(([re]) => re.test(t)).map(([, phrase, why]) => ({ phrase, why }));
}

// the visible text of an html body (for the empty-email check)
export const textOf = (html) => decode(stripTags(String(html || "").replace(/<(style|script|head)[^>]*>[\s\S]*?<\/\1>/gi, ""))).replace(/\s+/g, " ").trim();

export const hasCidImages = (html) => /<img[^>]+src\s*=\s*["']?cid:/i.test(String(html || ""));

export function addOptOut(html, line) {
  const h = String(html || ""), l = String(line || "").trim();
  if (!l) return h;
  const p = '<p style="color:#62656B;font-size:12px;margin-top:16px">' + esc(l) + "</p>";
  const at = h.search(/<div[^>]*id\s*=\s*["']?Signature/i);
  if (at >= 0) return h.slice(0, at) + p + h.slice(at);
  const body = h.search(/<\/body>/i);
  return body >= 0 ? h.slice(0, body) + p + h.slice(body) : h + p;
}

export function sentToday(merges, from, sinceIso) {
  const f = String(from || "").toLowerCase();
  let n = 0;
  for (const m of merges || []) {
    if (String(m.from || "").toLowerCase() !== f) continue;
    for (const it of m.items || []) if (it.status === "sent" && it.at && it.at >= sinceIso) n++;
  }
  return n;
}

// a short stable hash of the email content (the test send unlocks the batch only for this exact content)
export function contentHash(...parts) {
  let h = 2166136261;
  const s = parts.map(String).join("\u0001");
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return (h >>> 0).toString(36);
}
