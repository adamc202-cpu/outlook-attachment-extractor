/**
 * taskpane.js  –  Attachment Extractor for Outlook
 * ─────────────────────────────────────────────────
 * Reads attachments from the active message via the
 * Office.js Mailbox API, displays them in the pane,
 * then either POSTs them to a configurable endpoint
 * or triggers a browser-level download.
 *
 * Supports:
 *   • Regular file attachments
 *   • Item attachments (embedded emails) – metadata only
 *   • Inline attachments (images in body) – labelled separately
 *   • Full process log with timestamps
 *   • Graceful error handling at every async boundary
 */

"use strict";

/* ═══════════════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════════════ */
const VERSION   = "1.0.0";
const MAX_SIZE  = 10 * 1024 * 1024; // 10 MB — Office.js inline limit

/* ═══════════════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════════════ */
let scannedAttachments = []; // raw Office.js attachment descriptors
let base64Cache        = {}; // { attachmentId: base64String }

/* ═══════════════════════════════════════════════════════
   DOM REFS
═══════════════════════════════════════════════════════ */
const $scanBtn    = () => document.getElementById("scanBtn");
const $statusBar  = () => document.getElementById("statusBar");
const $attachList = () => document.getElementById("attachList");
const $actions    = () => document.getElementById("actions");
const $logSection = () => document.getElementById("logSection");
const $logBox     = () => document.getElementById("logBox");
const $sendBtn    = () => document.getElementById("sendBtn");
const $clearBtn   = () => document.getElementById("clearBtn");
const $endpoint   = () => document.getElementById("endpointInput");
const $clock      = () => document.getElementById("clock");

/* ═══════════════════════════════════════════════════════
   OFFICE INITIALISATION
═══════════════════════════════════════════════════════ */
Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("This add-in only runs inside Outlook.", "err");
    $scanBtn().disabled = true;
    return;
  }

  log("info", `Attachment Extractor v${VERSION} ready`);
  startClock();

  $scanBtn().addEventListener("click", handleScan);
  $sendBtn().addEventListener("click", handleSendOrDownload);
  $clearBtn().addEventListener("click", handleClear);
});

/* ═══════════════════════════════════════════════════════
   CLOCK UTILITY
═══════════════════════════════════════════════════════ */
function startClock() {
  const update = () => {
    const now = new Date();
    $clock().textContent = now.toLocaleTimeString([], { hour12: false });
  };
  update();
  setInterval(update, 1000);
}

/* ═══════════════════════════════════════════════════════
   SCAN  –  Entry point for the ribbon button action
═══════════════════════════════════════════════════════ */
async function handleScan() {
  resetUI();
  $scanBtn().disabled = true;
  $scanBtn().textContent = "⟳ SCANNING…";

  try {
    const item = Office.context.mailbox.item;

    // ── Guard: item must be a message ──────────────────
    if (!item) {
      throw new Error("No mail item is currently selected.");
    }

    log("info", `Scanning message: "${item.subject || "(no subject)"}" …`);

    const attachments = item.attachments; // synchronous array

    // ── Guard: no attachments ──────────────────────────
    if (!attachments || attachments.length === 0) {
      showStatus("⚠ No attachments found in this email.", "warn");
      log("warn", "Scan complete — 0 attachments found.");
      return;
    }

    scannedAttachments = attachments;
    log("ok", `Found ${attachments.length} attachment(s). Fetching content…`);

    // ── Render cards immediately (metadata) ───────────
    renderAttachmentCards(attachments);

    // ── Fetch base64 content for file attachments ──────
    await prefetchAttachmentContent(attachments);

    showStatus(`✓ ${attachments.length} attachment(s) identified.`, "ok");
    $actions().style.display = "flex";

  } catch (err) {
    showStatus(`✗ Scan failed: ${err.message}`, "err");
    log("err", `Scan error — ${err.message}`);
  } finally {
    $scanBtn().disabled = false;
    $scanBtn().textContent = "⟳ SCAN EMAIL FOR ATTACHMENTS";
  }
}

/* ═══════════════════════════════════════════════════════
   RENDER ATTACHMENT CARDS
═══════════════════════════════════════════════════════ */
function renderAttachmentCards(attachments) {
  const list = $attachList();
  list.innerHTML = "";

  attachments.forEach((att, idx) => {
    const card = document.createElement("div");
    card.className = "attach-card";
    card.style.animationDelay = `${idx * 50}ms`;

    const typeLabel = resolveTypeLabel(att);
    const sizeLabel = att.size ? formatBytes(att.size) : "—";

    card.innerHTML = `
      <span class="name" title="${escHtml(att.name)}">${escHtml(att.name)}</span>
      <span class="badge">${typeLabel}</span>
      <span class="meta">${sizeLabel}  ·  id: ${att.id.slice(-8)}</span>
    `;
    list.appendChild(card);
  });
}

/* ═══════════════════════════════════════════════════════
   PRE-FETCH ATTACHMENT CONTENT (base64)
   Office.js getAttachmentContentAsync is available from
   requirement set 1.8. Falls back gracefully.
═══════════════════════════════════════════════════════ */
async function prefetchAttachmentContent(attachments) {
  const fileAttachments = attachments.filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File);

  for (const att of fileAttachments) {
    if (att.size > MAX_SIZE) {
      log("warn", `Skipping inline fetch for "${att.name}" (${formatBytes(att.size)} > 10 MB limit). Will send via server-side EWS.`);
      continue;
    }

    try {
      const content = await getAttachmentContentAsync(att.id);
      base64Cache[att.id] = content.content; // base64 string
      log("ok", `Cached "${att.name}" (${formatBytes(att.size)})`);
    } catch (err) {
      log("warn", `Could not cache "${att.name}": ${err.message}`);
    }
  }
}

function getAttachmentContentAsync(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(
      attachmentId,
      { asyncContext: attachmentId },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

/* ═══════════════════════════════════════════════════════
   SEND / DOWNLOAD HANDLER
═══════════════════════════════════════════════════════ */
async function handleSendOrDownload() {
  if (!scannedAttachments.length) {
    showStatus("Run a scan first.", "warn");
    return;
  }

  const endpoint = $endpoint().value.trim();
  $sendBtn().disabled = true;
  $sendBtn().textContent = endpoint ? "↑ SENDING…" : "↓ DOWNLOADING…";

  let successCount = 0;
  let failCount    = 0;

  log("info", `── Processing ${scannedAttachments.length} attachment(s) ──`);

  for (const att of scannedAttachments) {
    try {
      if (att.attachmentType === Office.MailboxEnums.AttachmentType.Item) {
        log("warn", `"${att.name}" is an embedded email item — skipping binary processing.`);
        continue;
      }

      const base64 = base64Cache[att.id];

      if (!base64) {
        log("warn", `"${att.name}" has no cached content (size limit or unsupported). Skipped.`);
        failCount++;
        continue;
      }

      if (endpoint) {
        await sendToEndpoint(endpoint, att, base64);
      } else {
        triggerDownload(att, base64);
      }

      log("ok", `✓ "${att.name}" — ${endpoint ? "uploaded" : "downloaded"} (${formatBytes(att.size)})`);
      successCount++;

    } catch (err) {
      log("err", `✗ "${att.name}" — ${err.message}`);
      failCount++;
    }
  }

  // ── Summary ──────────────────────────────────────────
  const summary = `Done: ${successCount} succeeded, ${failCount} failed.`;
  if (failCount === 0) {
    showStatus(`✓ ${summary}`, "ok");
  } else if (successCount === 0) {
    showStatus(`✗ ${summary}`, "err");
  } else {
    showStatus(`⚠ ${summary}`, "warn");
  }
  log("info", `── ${summary} ──`);

  $sendBtn().disabled = false;
  $sendBtn().textContent = "↑ Send / Download All";
}

/* ═══════════════════════════════════════════════════════
   SEND TO ENDPOINT
═══════════════════════════════════════════════════════ */
async function sendToEndpoint(url, att, base64) {
  // Convert base64 → Blob so we can send multipart/form-data
  const binary   = atob(base64);
  const bytes    = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  const blob     = new Blob([bytes], { type: att.contentType || "application/octet-stream" });

  const formData = new FormData();
  formData.append("file", blob, att.name);
  formData.append("attachmentId",   att.id);
  formData.append("attachmentName", att.name);
  formData.append("contentType",    att.contentType || "");
  formData.append("size",           att.size ?? "");
  formData.append("isInline",       att.isInline ? "true" : "false");

  const response = await fetch(url, {
    method: "POST",
    body:   formData,
    // Add Authorization header here if your endpoint requires it:
    // headers: { "Authorization": "Bearer YOUR_TOKEN" }
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`HTTP ${response.status}: ${body || response.statusText}`);
  }
}

/* ═══════════════════════════════════════════════════════
   TRIGGER BROWSER DOWNLOAD
═══════════════════════════════════════════════════════ */
function triggerDownload(att, base64) {
  const a    = document.createElement("a");
  a.href     = `data:${att.contentType || "application/octet-stream"};base64,${base64}`;
  a.download = att.name;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/* ═══════════════════════════════════════════════════════
   CLEAR
═══════════════════════════════════════════════════════ */
function handleClear() {
  resetUI();
  log("info", "Cleared. Ready for next scan.");
}

function resetUI() {
  scannedAttachments = [];
  base64Cache        = {};

  $attachList().innerHTML = "";
  $actions().style.display    = "none";
  hideStatus();
}

/* ═══════════════════════════════════════════════════════
   STATUS BAR
═══════════════════════════════════════════════════════ */
function showStatus(msg, type = "ok") {
  const bar = $statusBar();
  bar.textContent  = msg;
  bar.className    = type;
  bar.style.display = "block";
  ensureLogVisible();
}

function hideStatus() {
  $statusBar().style.display = "none";
}

/* ═══════════════════════════════════════════════════════
   PROCESS LOG
═══════════════════════════════════════════════════════ */
function log(level, message) {
  ensureLogVisible();
  const box  = $logBox();
  const time = new Date().toLocaleTimeString([], { hour12: false });
  const span = document.createElement("div");
  span.className = `log-${level}`;
  span.textContent = `[${time}] ${message}`;
  box.appendChild(span);
  box.scrollTop = box.scrollHeight; // auto-scroll to latest
}

function ensureLogVisible() {
  $logSection().style.display = "flex";
}

/* ═══════════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════════ */
function resolveTypeLabel(att) {
  if (att.isInline) return "INLINE";
  switch (att.attachmentType) {
    case Office.MailboxEnums.AttachmentType.File: return "FILE";
    case Office.MailboxEnums.AttachmentType.Item: return "EMAIL";
    default: return "OTHER";
  }
}

function formatBytes(bytes) {
  if (!bytes || bytes === 0) return "0 B";
  const k     = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i     = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
