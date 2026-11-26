"use strict";

/* ------------------------------------------------------------------ *
 * State & storage
 * ------------------------------------------------------------------ */

const DEFAULT_TEMPLATES = [
  {
    subject: "You're invited, {{name}}!",
    body: "Hi {{name}},\n\nI'd love for you to join us at our upcoming event.\n\nDate: ...\nPlace: ...\n\nPlease reply to this email to let me know if you can make it.\n\nBest regards,\n...",
  },
  {
    subject: "Invitation: join us at the event",
    body: "Hello {{name}},\n\nWe're hosting an event and it would be great to see you there.\n\nDetails: ...\n\nJust reply to confirm.\n\nThanks,\n...",
  },
];

const state = {
  recipients: [],      // {id, name, email, fields, status, template, error, sentAt}
  fileName: "",
  templates: [],       // {id, subject, body}
  settings: { minDelay: 2, maxDelay: 6, dryRun: false, testEmail: "", clientId: "", accountEmail: "" },
  running: false,
  paused: false,
  stopRequested: false,
};

const hasChromeStorage = typeof chrome !== "undefined" && chrome.storage && chrome.storage.local;

async function loadStored() {
  if (hasChromeStorage) return chrome.storage.local.get(["recipients", "fileName", "templates", "settings"]);
  try { return JSON.parse(localStorage.getItem("invite-sender") || "{}"); } catch { return {}; }
}

let saveTimer = null;
function save(immediate = false) {
  clearTimeout(saveTimer);
  const doSave = () => {
    const data = {
      recipients: state.recipients,
      fileName: state.fileName,
      templates: state.templates,
      settings: state.settings,
    };
    if (hasChromeStorage) return chrome.storage.local.set(data);
    try { localStorage.setItem("invite-sender", JSON.stringify(data)); } catch {}
  };
  if (immediate) return doSave();
  saveTimer = setTimeout(doSave, 300);
}

const uid = () => Math.random().toString(36).slice(2, 10);
const $ = (sel) => document.querySelector(sel);

/* ------------------------------------------------------------------ *
 * Timer (worker based, so delays are not throttled in background tabs)
 * ------------------------------------------------------------------ */

let timerWorker = null;
try { timerWorker = new Worker("timer-worker.js"); } catch { timerWorker = null; }
let timerSeq = 0;
const timerWaiters = new Map();
if (timerWorker) {
  timerWorker.onmessage = (e) => {
    const resolve = timerWaiters.get(e.data.id);
    if (resolve) { timerWaiters.delete(e.data.id); resolve(); }
  };
}
function sleep(ms) {
  return new Promise((resolve) => {
    if (!timerWorker) return setTimeout(resolve, ms);
    const id = ++timerSeq;
    timerWaiters.set(id, resolve);
    timerWorker.postMessage({ id, ms });
  });
}

/* ------------------------------------------------------------------ *
 * Recipient file parsing
 * ------------------------------------------------------------------ */

const EMAIL_RE = /^[^\s@<>(),;:"\[\]]+@[^\s@<>(),;:"\[\]]+\.[^\s@<>(),;:"\[\]]{2,}$/;
const EMAIL_HEADER_RE = /^(e-?mail|mail|email address|e-mail address|邮箱|邮件|電子郵件|电子邮件|邮箱地址)$/i;
const NAME_HEADER_RE = /^(name|full ?name|first ?name|姓名|名字|名称|名稱)$/i;

function cleanEmail(v) {
  let s = String(v ?? "").trim();
  const m = s.match(/<([^>]+)>/);          // "John <john@x.com>"
  if (m) s = m[1].trim();
  return s.replace(/^mailto:/i, "");
}

async function parseFile(file) {
  const isText = /\.(csv|txt)$/i.test(file.name);
  let wb;
  if (isText) {
    const text = await file.text();
    wb = XLSX.read(text, { type: "string", raw: true });
  } else {
    const buf = await file.arrayBuffer();
    wb = XLSX.read(buf, { type: "array" });
  }
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false, raw: false });
  const nonEmpty = rows.filter((r) => r.some((c) => String(c).trim() !== ""));
  if (!nonEmpty.length) throw new Error("The file is empty.");

  // Header row present unless the first row already contains an email address.
  const first = nonEmpty[0].map((c) => String(c).trim());
  const hasHeader = !first.some((c) => EMAIL_RE.test(cleanEmail(c)));
  const width = Math.max(...nonEmpty.map((r) => r.length));
  const headers = [];
  for (let i = 0; i < width; i++) {
    const h = hasHeader ? first[i] : "";
    headers.push(h || `column${i + 1}`);
  }
  const data = hasHeader ? nonEmpty.slice(1) : nonEmpty;

  // Find the email column: by header, else the column with most email-looking values.
  let emailCol = headers.findIndex((h) => EMAIL_HEADER_RE.test(h.trim()));
  if (emailCol < 0) emailCol = headers.findIndex((h) => /mail|邮箱/i.test(h));
  if (emailCol < 0) {
    let best = -1, bestCount = 0;
    for (let i = 0; i < width; i++) {
      const count = data.filter((r) => EMAIL_RE.test(cleanEmail(r[i]))).length;
      if (count > bestCount) { best = i; bestCount = count; }
    }
    emailCol = best;
  }
  if (emailCol < 0) throw new Error("Couldn't find a column with email addresses.");

  let nameCol = headers.findIndex((h, i) => i !== emailCol && NAME_HEADER_RE.test(h.trim()));
  if (nameCol < 0) nameCol = headers.findIndex((h, i) => i !== emailCol && /name|姓名|名字/i.test(h));
  if (nameCol < 0) nameCol = headers.findIndex((_, i) => i !== emailCol);

  const seen = new Set();
  const recipients = data.map((r) => {
    const fields = {};
    headers.forEach((h, i) => { fields[h] = String(r[i] ?? "").trim(); });
    const email = cleanEmail(r[emailCol]);
    const name = nameCol >= 0 ? String(r[nameCol] ?? "").trim() : "";
    let status = "pending", error = "";
    if (!EMAIL_RE.test(email)) { status = "invalid"; error = "Not a valid email address"; }
    else if (seen.has(email.toLowerCase())) { status = "duplicate"; error = "Same address appears earlier in the list"; }
    else seen.add(email.toLowerCase());
    return { id: uid(), name, email, fields, status, template: "", error, sentAt: null };
  }).filter((r) => r.email || r.name);

  return {
    recipients,
    info: {
      emailHeader: headers[emailCol],
      nameHeader: nameCol >= 0 ? headers[nameCol] : null,
      hasHeader,
    },
  };
}

/* ------------------------------------------------------------------ *
 * Templates
 * ------------------------------------------------------------------ */

function renderTemplate(text, r) {
  const map = {};
  for (const [k, v] of Object.entries(r.fields || {})) map[k.trim().toLowerCase()] = v;
  map.name = r.name;
  map.email = r.email;
  return text.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (whole, key) => {
    const k = key.toLowerCase();
    return k in map ? map[k] : whole;
  });
}

function usableTemplates() {
  return state.templates
    .map((t, i) => ({ ...t, label: `Template ${i + 1}` }))
    .filter((t) => t.subject.trim() && t.body.trim());
}

function pickTemplate() {
  const list = usableTemplates();
  return list[Math.floor(Math.random() * list.length)];
}

function renderTemplates() {
  const host = $("#templates");
  host.innerHTML = "";
  const tplNode = $("#template-card");
  state.templates.forEach((t, i) => {
    const node = tplNode.content.firstElementChild.cloneNode(true);
    node.dataset.id = t.id;
    node.querySelector(".tpl-title").textContent = `Template ${i + 1}`;
    node.querySelector(".tpl-subject").value = t.subject;
    node.querySelector(".tpl-body").value = t.body;
    host.appendChild(node);
  });
  lockWhileRunning();
}

$("#templates").addEventListener("input", (e) => {
  const card = e.target.closest(".tpl");
  if (!card) return;
  const t = state.templates.find((x) => x.id === card.dataset.id);
  if (!t) return;
  if (e.target.classList.contains("tpl-subject")) t.subject = e.target.value;
  if (e.target.classList.contains("tpl-body")) t.body = e.target.value;
  save();
});

$("#templates").addEventListener("click", (e) => {
  if (!e.target.classList.contains("tpl-delete")) return;
  const card = e.target.closest(".tpl");
  const t = state.templates.find((x) => x.id === card.dataset.id);
  if ((t.subject || t.body) && !confirm("Delete this template?")) return;
  state.templates = state.templates.filter((x) => x.id !== card.dataset.id);
  renderTemplates();
  save();
});

$("#add-template").addEventListener("click", () => {
  state.templates.push({ id: uid(), subject: "", body: "" });
  renderTemplates();
  save();
  const cards = document.querySelectorAll(".tpl-subject");
  cards[cards.length - 1]?.focus();
});

$("#export-templates").addEventListener("click", () => {
  const data = state.templates.map(({ subject, body }) => ({ subject, body }));
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "invite-templates.json";
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
});

$("#import-templates").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    if (!Array.isArray(data)) throw new Error("Expected a list of {subject, body}.");
    const imported = data
      .filter((t) => t && typeof t === "object")
      .map((t) => ({ id: uid(), subject: String(t.subject ?? t.title ?? ""), body: String(t.body ?? t.content ?? "") }));
    if (!imported.length) throw new Error("No templates found in the file.");
    const replace = state.templates.every((t) => !t.subject && !t.body) ||
      confirm(`Replace your ${state.templates.length} current template(s) with the ${imported.length} imported one(s)?\n\nCancel = add them to the existing ones.`);
    state.templates = replace ? imported : state.templates.concat(imported);
    renderTemplates();
    save();
    log(`Imported ${imported.length} template(s).`, "ok");
  } catch (err) {
    alert("Couldn't import templates: " + err.message);
  }
});

/* ------------------------------------------------------------------ *
 * Recipients UI
 * ------------------------------------------------------------------ */

const STATUS_LABEL = {
  pending: "Pending",
  sending: "Sending…",
  sent: "Sent",
  failed: "Failed",
  invalid: "Invalid",
  duplicate: "Duplicate",
  skipped: "Skipped",
  unknown: "Check Sent folder",
};

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function rowAction(r) {
  if (r.status === "pending") return `<button class="ghost small" data-act="skip" data-id="${r.id}">Skip</button>`;
  if (["failed", "skipped", "unknown"].includes(r.status)) return `<button class="ghost small" data-act="queue" data-id="${r.id}">Queue</button>`;
  if (r.status === "sent") return `<button class="ghost small" data-act="resend" data-id="${r.id}">Send again</button>`;
  return "";
}

function renderRecipients() {
  const has = state.recipients.length > 0;
  $("#table-wrap").hidden = !has;
  $("#recipient-summary").hidden = !has;
  $("#clear-recipients").hidden = !has;
  $("#file-name").textContent = has ? state.fileName : "No file loaded";

  $("#recipient-rows").innerHTML = state.recipients.map((r, i) => `
    <tr data-id="${r.id}">
      <td>${i + 1}</td>
      <td class="who">${escapeHtml(r.name) || '<span class="muted">(no name)</span>'}<small>${escapeHtml(r.email)}</small></td>
      <td><span class="badge ${r.status}">${STATUS_LABEL[r.status] || r.status}</span>${r.template ? `<small class="tpl-used">${escapeHtml(r.template)}</small>` : ""}${r.error ? `<span class="err">${escapeHtml(r.error)}</span>` : ""}</td>
      <td>${rowAction(r)}</td>
    </tr>`).join("");

  const c = counts();
  const parts = [`<strong>${state.recipients.length}</strong> rows`, `${c.pending} pending`];
  if (c.sent) parts.push(`${c.sent} sent`);
  if (c.failed) parts.push(`${c.failed} failed`);
  if (c.unknown) parts.push(`${c.unknown} to check`);
  if (c.invalid) parts.push(`${c.invalid} invalid`);
  if (c.duplicate) parts.push(`${c.duplicate} duplicate`);
  if (c.skipped) parts.push(`${c.skipped} skipped`);
  $("#recipient-summary").innerHTML = parts.join(" · ");
  $("#retry-failed").hidden = !c.failed;
  renderProgress();
  lockWhileRunning();
}

function counts() {
  const c = { pending: 0, sending: 0, sent: 0, failed: 0, invalid: 0, duplicate: 0, skipped: 0, unknown: 0 };
  for (const r of state.recipients) c[r.status] = (c[r.status] || 0) + 1;
  return c;
}

$("#recipient-rows").addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-act]");
  if (!btn || state.running) return;
  const r = state.recipients.find((x) => x.id === btn.dataset.id);
  if (!r) return;
  if (btn.dataset.act === "skip") { r.status = "skipped"; r.error = ""; }
  if (btn.dataset.act === "queue") { r.status = "pending"; r.error = ""; }
  if (btn.dataset.act === "resend") {
    if (!confirm(`${r.email} was already sent an email. Queue it again?`)) return;
    r.status = "pending"; r.error = "";
  }
  renderRecipients();
  save();
});

$("#retry-failed").addEventListener("click", () => {
  if (!confirm("Queue all failed rows again?\n\nTip: check your Gmail Sent folder first — in rare cases a message counted as failed may still have gone out.")) return;
  for (const r of state.recipients) if (r.status === "failed") { r.status = "pending"; r.error = ""; }
  renderRecipients();
  save();
});

$("#file-input").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;
  if (state.recipients.some((r) => r.status === "sent") &&
      !confirm("Loading a new file replaces the current list, including its sent/failed history. Continue?")) return;
  try {
    const { recipients, info } = await parseFile(file);
    state.recipients = recipients;
    state.fileName = file.name;
    renderRecipients();
    save(true);
    const c = counts();
    log(`Loaded ${file.name}: ${recipients.length} rows (email column "${info.emailHeader}"${info.nameHeader ? `, name column "${info.nameHeader}"` : ", no name column"}). ${c.pending} ready to send.`, "ok");
    if (!info.nameHeader) log("No name column found, so {{name}} will be empty.", "warn");
  } catch (err) {
    alert("Couldn't read the file: " + err.message);
    log("Couldn't read the file: " + err.message, "error");
  }
});

$("#clear-recipients").addEventListener("click", () => {
  if (!confirm("Remove the recipient list (and its sent/failed history)?")) return;
  state.recipients = [];
  state.fileName = "";
  renderRecipients();
  save(true);
});

/* ------------------------------------------------------------------ *
 * Settings UI
 * ------------------------------------------------------------------ */

function renderSettings() {
  $("#min-delay").value = state.settings.minDelay;
  $("#max-delay").value = state.settings.maxDelay;
  $("#dry-run").checked = state.settings.dryRun;
  $("#test-email").value = state.settings.testEmail || "";
}

function readSettings() {
  const s = state.settings;
  s.minDelay = Math.max(0, parseFloat($("#min-delay").value) || 0);
  s.maxDelay = Math.max(0, parseFloat($("#max-delay").value) || 0);
  s.dryRun = $("#dry-run").checked;
  s.testEmail = $("#test-email").value.trim();
  save();
}
["#min-delay", "#max-delay", "#dry-run", "#test-email"].forEach((sel) =>
  $(sel).addEventListener("change", readSettings));

function randomDelaySeconds() {
  const lo = Math.min(state.settings.minDelay, state.settings.maxDelay);
  const hi = Math.max(state.settings.minDelay, state.settings.maxDelay);
  return Math.round((lo + Math.random() * (hi - lo)) * 10) / 10;
}

/* ------------------------------------------------------------------ *
 * Log & progress
 * ------------------------------------------------------------------ */

function log(msg, kind = "") {
  const li = document.createElement("li");
  const time = new Date().toLocaleTimeString();
  li.textContent = `[${time}] ${msg}`;
  if (kind) li.className = kind;
  const ol = $("#log");
  ol.appendChild(li);
  ol.scrollTop = ol.scrollHeight;
}

function setStatus(text) { $("#status-line").textContent = text; }

function renderProgress() {
  const c = counts();
  const total = c.pending + c.sending + c.sent + c.failed + c.unknown;
  const done = c.sent;
  $("#progress-bar").style.width = total ? `${(done / total) * 100}%` : "0";
  $("#progress-text").textContent = total
    ? `Sent ${done} of ${total}` + (c.failed ? ` · ${c.failed} failed` : "") + (c.pending ? ` · ${c.pending} remaining` : "")
    : "Nothing queued yet.";
}

function lockWhileRunning() {
  const busy = state.running;
  $("#start").disabled = busy;
  $("#pause").disabled = !busy;
  $("#stop").disabled = !busy;
  $("#pause").textContent = state.paused ? "Resume" : "Pause";
  $("#send-test").disabled = busy;
  $("#file-input").disabled = busy;
  $("#clear-recipients").disabled = busy;
  $("#retry-failed").disabled = busy;
  $("#import-templates").disabled = busy;
  $("#add-template").disabled = busy;
  document.querySelectorAll(".tpl-delete, #recipient-rows button").forEach((b) => (b.disabled = busy));
  $("#dry-run").disabled = busy;
}

/* ------------------------------------------------------------------ *
 * Google sign-in (OAuth) for the Gmail API
 * ------------------------------------------------------------------ */

const hasIdentity = typeof chrome !== "undefined" && !!chrome.identity;
const REDIRECT_URI = hasIdentity ? chrome.identity.getRedirectURL() : "https://dlibgmnmglameeebdkhpelifjmkaimml.chromiumapp.org/";
const SCOPES = "https://www.googleapis.com/auth/gmail.send openid email";

// Access token lives in session storage (cleared when Chrome closes).
const auth = { token: null, expiresAt: 0 };

async function loadAuth() {
  try {
    const s = await chrome.storage.session.get("auth");
    if (s.auth) Object.assign(auth, s.auth);
  } catch {}
}
function saveAuth() {
  try { chrome.storage.session.set({ auth }); } catch {}
}

async function authorize({ interactive, chooseAccount = false }) {
  const clientId = state.settings.clientId.trim();
  if (!clientId) throw new Error("Finish the one-time setup first (paste your Client ID at the top).");
  const p = new URLSearchParams({
    client_id: clientId,
    response_type: "token",
    redirect_uri: REDIRECT_URI,
    scope: SCOPES,
    include_granted_scopes: "true",
  });
  if (chooseAccount) p.set("prompt", "select_account");
  else if (!interactive) p.set("prompt", "none");
  if (!chooseAccount && state.settings.accountEmail) p.set("login_hint", state.settings.accountEmail);

  let redirect;
  try {
    redirect = await chrome.identity.launchWebAuthFlow({
      url: "https://accounts.google.com/o/oauth2/v2/auth?" + p.toString(),
      interactive,
    });
  } catch (err) {
    throw new Error(interactive ? `Google sign-in didn't finish: ${err.message}` : "silent sign-in failed");
  }
  const frag = new URLSearchParams(new URL(redirect).hash.slice(1));
  if (frag.get("error")) throw new Error(`Google sign-in error: ${frag.get("error")}. See the setup steps in the README.`);
  const token = frag.get("access_token");
  if (!token) throw new Error("Google didn't return an access token.");
  auth.token = token;
  auth.expiresAt = Date.now() + (Number(frag.get("expires_in")) || 3600) * 1000 - 60000;
  saveAuth();

  try {
    const r = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", { headers: { Authorization: "Bearer " + token } });
    if (r.ok) {
      const info = await r.json();
      if (info.email) state.settings.accountEmail = info.email;
    }
  } catch {}
  save();
  renderAccount();
  return token;
}

// Valid token, refreshing silently when it expires (falls back to a Google popup).
async function getToken() {
  if (auth.token && Date.now() < auth.expiresAt) return auth.token;
  try {
    return await authorize({ interactive: false });
  } catch {
    return authorize({ interactive: true });
  }
}

function renderAccount() {
  const box = $("#account-box");
  const s = state.settings;
  if (!hasIdentity) {
    box.className = "account-box bad";
    box.innerHTML = `Preview mode: open this from the extension icon in Chrome to send.`;
    return;
  }
  if (!s.clientId || box.dataset.editing === "1") {
    box.className = "account-box bad setup";
    box.innerHTML = `
      <strong>One-time setup</strong> (about 10 minutes, free). Follow “Gmail API setup” in the README.
      When it asks for the redirect URI, use:
      <div class="copy-row"><code id="redirect-uri">${escapeHtml(REDIRECT_URI)}</code><button class="small" id="copy-redirect">Copy</button></div>
      <label>Client ID
        <input type="text" id="client-id" placeholder="xxxxxxxx.apps.googleusercontent.com" value="${escapeHtml(s.clientId)}">
      </label>
      <div class="row"><button class="primary small" id="save-client-id">Save</button>
      ${s.clientId ? `<button class="ghost small" id="cancel-client-id">Cancel</button>` : ""}</div>`;
    return;
  }
  const connected = auth.token && Date.now() < auth.expiresAt && s.accountEmail;
  if (!s.accountEmail) {
    box.className = "account-box bad";
    box.innerHTML = `Gmail isn't connected yet.<div class="row"><button class="primary small" id="connect">Connect Gmail account</button>
      <button class="ghost small" id="edit-client-id">Setup</button></div>`;
    return;
  }
  box.className = "account-box";
  box.innerHTML = `Sending from <strong>${escapeHtml(s.accountEmail)}</strong>${connected ? "" : " <small>(will reconnect when you start)</small>"}
    <div class="row"><button class="ghost small" id="switch-account" ${state.running ? "disabled" : ""}>Switch account</button>
    <button class="ghost small" id="edit-client-id" ${state.running ? "disabled" : ""}>Setup</button></div>`;
}

document.addEventListener("click", async (e) => {
  const id = e.target.id;
  const box = $("#account-box");
  if (id === "copy-redirect") {
    try { await navigator.clipboard.writeText(REDIRECT_URI); e.target.textContent = "Copied"; } catch {}
  } else if (id === "save-client-id") {
    const v = $("#client-id").value.trim();
    if (!/\.apps\.googleusercontent\.com$/.test(v)) { alert("That doesn't look like a Client ID. It ends with .apps.googleusercontent.com"); return; }
    if (v !== state.settings.clientId) { auth.token = null; saveAuth(); }
    state.settings.clientId = v;
    delete box.dataset.editing;
    save(true);
    renderAccount();
  } else if (id === "cancel-client-id") {
    delete box.dataset.editing;
    renderAccount();
  } else if (id === "edit-client-id") {
    box.dataset.editing = "1";
    renderAccount();
  } else if (id === "connect" || id === "switch-account") {
    try {
      await authorize({ interactive: true, chooseAccount: true });
      log(`Connected to ${state.settings.accountEmail}.`, "ok");
    } catch (err) {
      alert(err.message);
      log(err.message, "error");
    }
  }
});

/* ------------------------------------------------------------------ *
 * Gmail API sending
 * ------------------------------------------------------------------ */

function b64utf8(s) {
  const bytes = new TextEncoder().encode(s);
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
const b64url = (s) => b64utf8(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
const oneLine = (s) => String(s ?? "").replace(/[\r\n]+/g, " ").trim();

function encodeHeader(s) {
  return /^[\x20-\x7e]*$/.test(s) ? s : `=?UTF-8?B?${b64utf8(s)}?=`;
}

function buildMime(to, name, subject, body) {
  const cleanName = oneLine(name).replace(/["\\]/g, "");
  let toHeader = to;
  if (cleanName) {
    toHeader = /^[\x20-\x7e]*$/.test(cleanName) ? `"${cleanName}" <${to}>` : `${encodeHeader(cleanName)} <${to}>`;
  }
  const text = String(body).replace(/\r?\n/g, "\r\n");
  return [
    `To: ${toHeader}`,
    `Subject: ${encodeHeader(oneLine(subject))}`,
    "MIME-Version: 1.0",
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: base64",
    "",
    b64utf8(text).replace(/.{76}/g, "$&\r\n"),
  ].join("\r\n");
}

async function sendOne(to, name, subject, body) {
  const raw = b64url(buildMime(to, name, subject, body));
  for (let attempt = 0; attempt < 2; attempt++) {
    const token = await getToken();
    let r;
    try {
      r = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
        method: "POST",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ raw }),
      });
    } catch (err) {
      throw new Error("Network error, couldn't reach Gmail: " + err.message);
    }
    if (r.status === 401 && attempt === 0) { auth.token = null; continue; } // expired: refresh once
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch {}
      if (/has not been used|is disabled/i.test(msg)) msg = "The Gmail API isn't enabled in your Google Cloud project (setup step 2).";
      throw new Error(msg);
    }
    return r.json();
  }
  throw new Error("Google rejected the sign-in. Click “Switch account” and connect again.");
}

/* ------------------------------------------------------------------ *
 * Main loop
 * ------------------------------------------------------------------ */

async function waitWhilePaused() {
  let shown = false;
  while (state.paused && !state.stopRequested) {
    if (!shown) { setStatus("Paused."); shown = true; }
    await sleep(200);
  }
}

async function countdown(seconds) {
  const end = Date.now() + seconds * 1000;
  while (!state.stopRequested) {
    if (state.paused) { await waitWhilePaused(); continue; }
    const left = end - Date.now();
    if (left <= 0) break;
    setStatus(`Waiting ${seconds}s before the next email… ${(left / 1000).toFixed(1)}s left`);
    await sleep(Math.min(200, left));
  }
}

function checkTemplates() {
  readSettings();
  if (!usableTemplates().length) { alert("Add at least one template with both a subject and a message."); return false; }
  const empty = state.templates.length - usableTemplates().length;
  if (empty && !confirm(`${empty} template(s) are missing a subject or message and will be ignored. Continue?`)) return false;
  return true;
}

// Makes sure we have a working Gmail connection before sending.
async function checkConnected() {
  if (!hasIdentity) { alert("Open this from the extension icon in Chrome to send."); return false; }
  if (!state.settings.clientId) { alert("Finish the one-time setup first (paste your Client ID at the top)."); return false; }
  try {
    if (!state.settings.accountEmail) await authorize({ interactive: true, chooseAccount: true });
    else await getToken();
    return true;
  } catch (err) {
    alert(err.message);
    log(err.message, "error");
    return false;
  }
}

async function start() {
  if (state.running) return;
  if (!checkTemplates()) return;
  const queue = state.recipients.filter((r) => r.status === "pending");
  if (!queue.length) { alert(state.recipients.length ? "No pending recipients left." : "Load a recipient file first."); return; }
  const s = state.settings;
  if (!s.dryRun && !(await checkConnected())) return;

  const lo = Math.min(s.minDelay, s.maxDelay), hi = Math.max(s.minDelay, s.maxDelay);
  const msg = s.dryRun
    ? `Dry run: go through ${queue.length} recipient(s) without sending anything?`
    : `Send ${queue.length} email(s) from ${s.accountEmail}?\n\nRandom wait between emails: ${lo}–${hi} seconds.\nKeep this side panel open until it finishes.`;
  if (!confirm(msg)) return;

  state.running = true;
  state.paused = false;
  state.stopRequested = false;
  lockWhileRunning();
  renderAccount();
  log(`${s.dryRun ? "Dry run" : `Sending from ${s.accountEmail}`} started: ${queue.length} recipient(s).`);

  let consecutiveFailures = 0;
  let processed = 0;
  for (let i = 0; i < queue.length; i++) {
    await waitWhilePaused();
    if (state.stopRequested) break;
    const r = queue[i];
    if (r.status !== "pending") continue; // changed meanwhile

    const tpl = pickTemplate();
    const subject = renderTemplate(tpl.subject, r);
    const body = renderTemplate(tpl.body, r);
    r.status = "sending";
    r.template = tpl.label;
    r.error = "";
    renderRecipients();
    save(true);
    setStatus(`${s.dryRun ? "Simulating" : "Sending"} ${i + 1} of ${queue.length}: ${r.email}`);

    try {
      if (s.dryRun) await sleep(300);
      else await sendOne(r.email, r.name, subject, body);
      r.status = "sent";
      r.sentAt = new Date().toISOString();
      consecutiveFailures = 0;
      log(`${s.dryRun ? "[dry run] " : ""}Sent to ${r.name ? r.name + " " : ""}<${r.email}> using ${tpl.label}.`, "ok");
    } catch (err) {
      r.status = "failed";
      r.error = err.message;
      consecutiveFailures++;
      log(`Failed for <${r.email}>: ${err.message}`, "error");
    }
    processed++;
    renderRecipients();
    save(true);

    if (consecutiveFailures >= 3) {
      log("Stopped after 3 failures in a row. Fix the error above, then press Start again.", "error");
      break;
    }
    const more = queue.slice(i + 1).some((x) => x.status === "pending");
    if (more && !state.stopRequested) await countdown(randomDelaySeconds());
  }

  const stopped = state.stopRequested;
  if (s.dryRun) {
    // A dry run shouldn't leave rows marked as sent.
    for (const r of queue) if (r.status === "sent") { r.status = "pending"; r.template = ""; r.sentAt = null; }
  }
  state.running = false;
  state.paused = false;
  state.stopRequested = false;
  renderRecipients();
  renderAccount();
  save(true);
  const c = counts();
  const summary = s.dryRun
    ? `Dry run finished (${processed} processed). Nothing was sent.`
    : `${stopped ? "Stopped" : "Finished"}. Sent ${c.sent}, failed ${c.failed}, remaining ${c.pending}.`;
  setStatus(summary);
  log(summary, c.failed ? "warn" : "ok");
}

$("#start").addEventListener("click", start);
$("#pause").addEventListener("click", () => {
  if (!state.running) return;
  state.paused = !state.paused;
  log(state.paused ? "Paused." : "Resumed.");
  if (!state.paused) setStatus("Resuming…");
  lockWhileRunning();
});
$("#stop").addEventListener("click", () => {
  if (!state.running) return;
  state.stopRequested = true;
  state.paused = false;
  setStatus("Stopping after the current email…");
  log("Stop requested.");
  lockWhileRunning();
});

$("#send-test").addEventListener("click", async () => {
  if (state.running) return;
  readSettings();
  const to = state.settings.testEmail;
  if (!EMAIL_RE.test(to)) { alert("Type the address the test email should go to (e.g. your own)."); return; }
  if (!usableTemplates().length) { alert("Add at least one template with both a subject and a message."); return; }
  if (!(await checkConnected())) return;
  const sample = state.recipients.find((r) => r.status !== "invalid") || { name: "Test Person", email: to, fields: {} };
  const fake = { ...sample, email: to };
  const tpl = pickTemplate();
  if (!confirm(`Send ${tpl.label} from ${state.settings.accountEmail} to ${to}, filled in with the details of "${sample.name || sample.email}"?`)) return;

  state.running = true;
  lockWhileRunning();
  setStatus(`Sending test email to ${to}…`);
  try {
    await sendOne(to, sample.name, renderTemplate(tpl.subject, fake), renderTemplate(tpl.body, fake));
    log(`Test email sent to ${to} using ${tpl.label}.`, "ok");
    setStatus("Test email sent. Check the inbox.");
  } catch (err) {
    log(`Test email failed: ${err.message}`, "error");
    setStatus("Test email failed.");
  } finally {
    state.running = false;
    lockWhileRunning();
  }
});

/* ------------------------------------------------------------------ *
 * Init
 * ------------------------------------------------------------------ */

(async function init() {
  const stored = await loadStored();
  state.templates = Array.isArray(stored.templates) && stored.templates.length
    ? stored.templates
    : DEFAULT_TEMPLATES.map((t) => ({ id: uid(), ...t }));
  state.settings = { ...state.settings, ...(stored.settings || {}) };
  delete state.settings.account; // from v1.0
  state.recipients = Array.isArray(stored.recipients) ? stored.recipients : [];
  state.fileName = stored.fileName || "";
  if (hasIdentity) await loadAuth();

  // If the panel was closed mid-send we can't know whether that email went out.
  let interrupted = 0;
  for (const r of state.recipients) {
    if (r.status === "sending") {
      r.status = "unknown";
      r.error = "Interrupted while sending. Check Gmail's Sent folder, then Queue or Skip it.";
      interrupted++;
    }
  }
  renderTemplates();
  renderSettings();
  renderRecipients();
  renderAccount();
  save(true);
  if (interrupted) log(`${interrupted} email(s) were interrupted last time. Check your Sent folder.`, "warn");
  if (!hasIdentity) log("Running outside the extension: preview only, sending is disabled.", "warn");
})();
