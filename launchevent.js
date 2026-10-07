/*
 * Send Check - Outlook add-in
 * Runs every time you click Send (Outlook "Smart Alerts") and shows a check:
 *   - all To / Cc / Bcc recipients, with outside-company addresses marked EXTERNAL
 *   - every attachment: [OK] password protected, [NO] not protected, [??] check manually
 * You then choose "Send" or "Don't Send".
 *
 * Works in: classic Outlook for Windows, New Outlook for Windows,
 *           Outlook on the web, and Outlook for Mac.
 */

// ----- Settings ---------------------------------------------------------------
// Your company email domain, e.g. "yourcompany.com". Leave "" to detect it from your address.
var MY_DOMAIN = "iriskpo.in";
// Attachments bigger than this are not scanned (marked [??]).
var MAX_SCAN_BYTES = 25 * 1024 * 1024;
// ------------------------------------------------------------------------------

var PUBLIC_DOMAINS = ["gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com",
  "msn.com", "yahoo.com", "yahoo.co.in", "icloud.com", "me.com", "rediffmail.com", "proton.me"];

/* ============================ byte helpers ============================ */

var B64_MAP = (function () {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var m = {};
  for (var i = 0; i < chars.length; i++) m[chars.charCodeAt(i)] = i;
  return m;
})();

function base64ToBytes(s) {
  s = String(s).replace(/[^A-Za-z0-9+\/]/g, "");
  var out = new Uint8Array(Math.floor(s.length * 6 / 8));
  var buf = 0, bits = 0, o = 0;
  for (var i = 0; i < s.length; i++) {
    buf = ((buf << 6) | B64_MAP[s.charCodeAt(i)]) & 0xFFFFFF;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out[o++] = (buf >> bits) & 0xFF;
    }
  }
  return o === out.length ? out : out.subarray(0, o);
}

function asciiBytes(str) {
  var b = new Uint8Array(str.length);
  for (var i = 0; i < str.length; i++) b[i] = str.charCodeAt(i) & 0xFF;
  return b;
}

function utf16Bytes(str) {
  var b = new Uint8Array(str.length * 2);
  for (var i = 0; i < str.length; i++) { b[i * 2] = str.charCodeAt(i) & 0xFF; b[i * 2 + 1] = 0; }
  return b;
}

function findBytes(hay, needle, from) {
  var n = needle.length, first = needle[0], last = hay.length - n;
  for (var i = from || 0; i <= last; i++) {
    if (hay[i] !== first) continue;
    var j = 1;
    while (j < n && hay[i + j] === needle[j]) j++;
    if (j === n) return i;
  }
  return -1;
}

/* ====================== password-protection checks ====================== */

var PROTECTED = 1, NOT_PROTECTED = 0, UNKNOWN = -1;
var SIG_ENCRYPTION_INFO = utf16Bytes("EncryptionInfo");
var SIG_PDF_ENCRYPT = asciiBytes("/Encrypt");
var SIG_ZIP_CENTRAL = new Uint8Array([0x50, 0x4B, 0x01, 0x02]);

function extOf(name) {
  var p = String(name).lastIndexOf(".");
  return p >= 0 ? String(name).slice(p + 1).toLowerCase() : "";
}

function zipStatus(b) {
  var total = 0, enc = 0, p = findBytes(b, SIG_ZIP_CENTRAL, 0);
  while (p >= 0 && p + 46 <= b.length) {
    var nameLen = b[p + 28] | (b[p + 29] << 8);
    var isDir = nameLen > 0 && p + 46 + nameLen <= b.length && b[p + 46 + nameLen - 1] === 0x2F;
    if (!isDir) {
      total++;
      if (b[p + 8] & 1) enc++;
    }
    p = findBytes(b, SIG_ZIP_CENTRAL, p + 4);
  }
  if (total === 0) return { code: UNKNOWN, note: "couldn't read the zip, check manually" };
  if (enc === total) return { code: PROTECTED, note: "password protected" };
  if (enc === 0) return { code: NOT_PROTECTED, note: "NOT password protected" };
  return { code: NOT_PROTECTED, note: "NOT fully protected (" + enc + " of " + total + " files inside)" };
}

/* Decide whether a file (name + raw bytes) is password protected. */
function classifyFile(name, b) {
  var ext = extOf(name);
  if (ext === "7z" || ext === "rar") return { code: UNKNOWN, note: "can't verify this format, check manually" };
  if (!b || b.length < 8) return { code: NOT_PROTECTED, note: "empty file, NOT protected" };

  var isOle = b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0;
  var hasEncInfo = isOle && findBytes(b, SIG_ENCRYPTION_INFO, 0) >= 0;

  switch (ext) {
    case "zip":
      return zipStatus(b);
    case "pdf":
      return findBytes(b, SIG_PDF_ENCRYPT, 0) >= 0
        ? { code: PROTECTED, note: "password protected" }
        : { code: NOT_PROTECTED, note: "NOT password protected" };
    case "docx": case "docm": case "dotx": case "xlsx": case "xlsm": case "xlsb":
    case "xltx": case "pptx": case "pptm":
      // Password-protected Office files are stored in an OLE container with an
      // "EncryptionInfo" stream; unprotected ones are plain zip files.
      return hasEncInfo
        ? { code: PROTECTED, note: "password protected" }
        : { code: NOT_PROTECTED, note: "NOT password protected" };
    case "doc": case "xls": case "ppt":
      return hasEncInfo
        ? { code: PROTECTED, note: "password protected" }
        : { code: UNKNOWN, note: "older Office format, check manually" };
    default:
      return { code: NOT_PROTECTED, note: "NOT protected (file type can't hold a password; zip it with one)" };
  }
}

/* ============================ Outlook helpers ============================ */

// Time limits so the pop-up always appears quickly and never hangs.
var STEP_TIMEOUT_MS = 2500;   // any single Outlook request
var TOTAL_TIMEOUT_MS = 4000;  // whole check; unfinished files become "check manually"

function withTimeout(promise, ms, fallback) {
  return new Promise(function (resolve) {
    var done = false;
    var t = setTimeout(function () { if (!done) { done = true; resolve(fallback); } }, ms);
    promise.then(function (v) { if (!done) { done = true; clearTimeout(t); resolve(v); } },
                 function () { if (!done) { done = true; clearTimeout(t); resolve(fallback); } });
  });
}

function call(fn, fallback) {
  // Wraps Outlook's callback-style "...Async" methods in a Promise with a time limit.
  var p = new Promise(function (resolve, reject) {
    try {
      fn(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(r && r.error);
      });
    } catch (e) { reject(e); }
  });
  return withTimeout(p, STEP_TIMEOUT_MS, fallback);
}

function domainOf(addr) {
  var p = String(addr || "").lastIndexOf("@");
  return p >= 0 ? String(addr).slice(p + 1).toLowerCase() : "";
}

function myDomain() {
  if (MY_DOMAIN) return MY_DOMAIN.toLowerCase();
  var d = domainOf(((Office.context.mailbox || {}).userProfile || {}).emailAddress);
  return PUBLIC_DOMAINS.indexOf(d) >= 0 ? "" : d;
}

function getRecipients(field) {
  if (!field || !field.getAsync) return Promise.resolve([]);
  return call(function (cb) { field.getAsync(cb); }, []).then(function (v) { return v || []; });
}

var CHECK_MANUALLY = { code: UNKNOWN, note: "check manually" };

function checkAttachment(item, att) {
  if (att.attachmentType === "item") return Promise.resolve({ code: NOT_PROTECTED, note: "attached email" });
  if (att.attachmentType === "cloud") return Promise.resolve({ code: UNKNOWN, note: "cloud link, check sharing" });
  if (att.size > MAX_SCAN_BYTES) return Promise.resolve({ code: UNKNOWN, note: "too large, check manually" });
  var ext = extOf(att.name);
  if (ext === "7z" || ext === "rar") return Promise.resolve(CHECK_MANUALLY);
  return call(function (cb) { item.getAttachmentContentAsync(att.id, cb); }, null).then(function (c) {
    if (!c || c.format !== "base64") return CHECK_MANUALLY;
    try { return classifyFile(att.name, base64ToBytes(c.content)); } catch (e) { return CHECK_MANUALLY; }
  });
}

/* ============================ the check itself ============================ */

function runCheck(item) {
  var recipientsP = Promise.all([getRecipients(item.to), getRecipients(item.cc), getRecipients(item.bcc)]);
  var attsP = call(function (cb) { item.getAttachmentsAsync(cb); }, []).then(function (list) {
    var atts = (list || []).filter(function (a) { return !a.isInline; });
    // Scan all files at the same time; anything not finished in time is "check manually".
    var scans = atts.map(function (a) { return withTimeout(checkAttachment(item, a), TOTAL_TIMEOUT_MS, CHECK_MANUALLY); });
    return Promise.all(scans).then(function (results) { return { atts: atts, results: results }; });
  });
  return Promise.all([recipientsP, attsP]).then(function (r) {
    return { domain: myDomain(), to: r[0][0], cc: r[0][1], bcc: r[0][2], atts: r[1].atts, results: r[1].results };
  });
}

function formatRecipients(list, dom, maxShown) {
  var ext = 0;
  var shown = list.map(function (r) {
    var a = r.emailAddress || r.displayName || "?";
    if (dom && domainOf(a) !== dom) { ext++; return a + " (EXTERNAL)"; }
    return a;
  });
  var extra = shown.length - maxShown;
  if (extra > 0) shown = shown.slice(0, maxShown).concat(["+" + extra + " more"]);
  return { text: shown.join(", "), external: ext };
}

/* Builds the pop-up text (Outlook allows at most 500 characters). */
function buildMessage(rep, maxPerField, shortNames) {
  var lines = [], unprotected = 0;

  lines.push("Have you included the correct recipients?");
  [["To", rep.to], ["Cc", rep.cc], ["Bcc", rep.bcc]].forEach(function (f) {
    if (f[1].length) lines.push(f[0] + ": " + formatRecipients(f[1], rep.domain, maxPerField).text);
  });
  if (!rep.to.length && !rep.cc.length && !rep.bcc.length) lines.push("(no recipients)");

  if (rep.atts.length) {
    lines.push("");
    rep.results.forEach(function (res) { if (res.code !== PROTECTED) unprotected++; });
    lines.push(unprotected ? "File Not Password Protected - please check:" : "All attachments are password protected:");
    rep.atts.forEach(function (a, i) {
      var res = rep.results[i];
      var name = a.name;
      if (shortNames && name.length > 30) name = name.slice(0, 27) + "...";
      if (res.code === PROTECTED) lines.push("✔ " + name + " - protected");
      else if (res.code === NOT_PROTECTED) lines.push("✘ " + name + " - NOT password protected");
      else lines.push("? " + name + " - " + res.note);
    });
  }
  return { text: lines.join("\n"), unprotected: unprotected };
}

function fitMessage(rep) {
  var LIMIT = 500, m;
  var tries = [[6, false], [3, false], [2, true], [1, true]];
  for (var i = 0; i < tries.length; i++) {
    m = buildMessage(rep, tries[i][0], tries[i][1]);
    if (m.text.length <= LIMIT) return m;
  }
  m.text = m.text.slice(0, LIMIT - 20) + "\n...(list shortened)";
  return m;
}

/* ============================ event handlers ============================ */

function onMessageSendHandler(event) {
  var finished = false;
  function finish(text) {
    if (finished) return;
    finished = true;
    try {
      event.completed({ allowEvent: false, errorMessage: text });   // shows "Send" / "Don't Send"
    } catch (e) {
      try { event.completed({ allowEvent: false, errorMessage: "Have you included the correct recipients? Please also check your attachments." }); } catch (e2) {}
    }
  }
  // Safety net: whatever happens, show the pop-up within 5 seconds.
  setTimeout(function () {
    finish("Have you included the correct recipients?\n\nSend Check couldn't finish in time - please check the recipients and that attachments are password protected.");
  }, 5000);

  try {
    runCheck(Office.context.mailbox.item).then(function (rep) {
      finish(fitMessage(rep).text);
    }, function () {
      finish("Have you included the correct recipients?\n\nSend Check couldn't read this email - please check the recipients and attachments yourself.");
    });
  } catch (e) {
    finish("Have you included the correct recipients?\n\nPlease check the recipients and attachments yourself.");
  }
}

/* Optional "Run send check" button in the compose ribbon: shows a short summary. */
function runCheckNow(event) {
  try {
    var item = Office.context.mailbox.item;
    runCheck(item).then(function (rep) {
      var m = buildMessage(rep, 1, true);
      var n = rep.to.length + rep.cc.length + rep.bcc.length;
      var summary = n + " recipient(s); " + rep.atts.length + " attachment(s), " + m.unprotected +
        " not password protected. Full check appears when you click Send.";
      item.notificationMessages.replaceAsync("sendcheck", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: summary.slice(0, 150), icon: "Icon.16x16", persistent: false
      }, function () { event.completed(); });
    }, function () { event.completed(); });
  } catch (e) { event.completed(); }
}

/* Register the handlers. Done both immediately and once Office is ready, so it always works. */
function registerHandlers() {
  try {
    if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
      Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
      Office.actions.associate("runCheckNow", runCheckNow);
    }
  } catch (e) {}
}
registerHandlers();
if (typeof Office !== "undefined" && Office.onReady) {
  try { Office.onReady(registerHandlers); } catch (e) {}
}

// For local testing in Node only.
if (typeof module !== "undefined" && module.exports) {
  module.exports = { classifyFile: classifyFile, base64ToBytes: base64ToBytes, buildMessage: buildMessage,
    fitMessage: fitMessage, runCheck: runCheck, onMessageSendHandler: onMessageSendHandler };
}
