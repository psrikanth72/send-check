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

function call(fn) {
  // Wraps Outlook's callback-style "...Async" methods in a Promise.
  return new Promise(function (resolve, reject) {
    fn(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
      else reject(r.error || new Error("Outlook request failed"));
    });
  });
}

function domainOf(addr) {
  var p = String(addr || "").lastIndexOf("@");
  return p >= 0 ? String(addr).slice(p + 1).toLowerCase() : "";
}

function myDomain(item) {
  if (MY_DOMAIN) return Promise.resolve(MY_DOMAIN.toLowerCase());
  var fallback = (Office.context.mailbox.userProfile || {}).emailAddress || "";
  var getFrom = item.from && item.from.getAsync
    ? call(function (cb) { item.from.getAsync(cb); }).then(function (f) { return (f && f.emailAddress) || fallback; },
                                                           function () { return fallback; })
    : Promise.resolve(fallback);
  return getFrom.then(function (addr) {
    var d = domainOf(addr);
    return PUBLIC_DOMAINS.indexOf(d) >= 0 ? "" : d;
  });
}

function getRecipients(field) {
  if (!field || !field.getAsync) return Promise.resolve([]);
  return call(function (cb) { field.getAsync(cb); }).then(function (v) { return v || []; }, function () { return []; });
}

function checkAttachment(item, att) {
  if (att.attachmentType === "item") return Promise.resolve({ code: NOT_PROTECTED, note: "attached email, can't hold a password" });
  if (att.attachmentType === "cloud") return Promise.resolve({ code: UNKNOWN, note: "cloud link, check who can open it" });
  if (att.size > MAX_SCAN_BYTES) return Promise.resolve({ code: UNKNOWN, note: "too large to scan, check manually" });
  var ext = extOf(att.name);
  if (ext === "7z" || ext === "rar") return Promise.resolve(classifyFile(att.name, null));
  return call(function (cb) { item.getAttachmentContentAsync(att.id, cb); }).then(function (c) {
    if (c.format !== "base64") return { code: UNKNOWN, note: "couldn't read it, check manually" };
    return classifyFile(att.name, base64ToBytes(c.content));
  }, function () {
    return { code: UNKNOWN, note: "couldn't read it, check manually" };
  });
}

function mentionsAttachment(item) {
  return call(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }).then(function (text) {
    text = String(text || "");
    var cut = text.search(/\r?\nFrom:/i);           // ignore the quoted earlier emails
    if (cut >= 0) text = text.slice(0, cut);
    return /attach|enclosed/i.test(text);
  }, function () { return false; });
}

/* ============================ the check itself ============================ */

function runCheck(item) {
  return Promise.all([
    myDomain(item),
    getRecipients(item.to), getRecipients(item.cc), getRecipients(item.bcc),
    call(function (cb) { item.getAttachmentsAsync(cb); }).catch(function () { return []; })
  ]).then(function (r) {
    var dom = r[0], to = r[1], cc = r[2], bcc = r[3];
    var atts = (r[4] || []).filter(function (a) { return !a.isInline; });
    return Promise.all(atts.map(function (a) { return checkAttachment(item, a); })).then(function (results) {
      var needMention = atts.length === 0 ? mentionsAttachment(item) : Promise.resolve(false);
      return needMention.then(function (mentioned) {
        return { domain: dom, to: to, cc: cc, bcc: bcc, atts: atts, results: results, mentionedButMissing: mentioned };
      });
    });
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

/* Builds the alert text. Outlook allows at most 500 characters. */
function buildMessage(rep, markdown, maxPerField) {
  var B = markdown ? "**" : "";
  var issues = 0, external = 0, lines = [];

  lines.push(B + "Check before sending" + B);
  lines.push("");
  var fields = [["To", rep.to], ["Cc", rep.cc], ["Bcc", rep.bcc]];
  fields.forEach(function (f) {
    if (!f[1].length) return;
    var fr = formatRecipients(f[1], rep.domain, maxPerField);
    external += fr.external;
    lines.push("- " + f[0] + ": " + fr.text);
  });
  if (!rep.to.length && !rep.cc.length && !rep.bcc.length) lines.push("- No recipients");

  lines.push("");
  if (!rep.atts.length) {
    lines.push(B + "Attachments:" + B + " none");
    if (rep.mentionedButMissing) { issues++; lines.push("- Your email mentions an attachment, but nothing is attached"); }
  } else {
    lines.push(B + "Attachments:" + B);
    rep.atts.forEach(function (a, i) {
      var res = rep.results[i];
      var tag = res.code === PROTECTED ? "[OK]" : res.code === NOT_PROTECTED ? "[NO]" : "[??]";
      if (res.code !== PROTECTED) issues++;
      lines.push("- " + tag + " " + a.name + " - " + res.note);
    });
  }

  if (external || issues) lines.push("");
  if (external) lines.push(external + " external recipient(s) - make sure they should get this.");
  if (issues) lines.push(B + issues + " item(s) need your attention." + B);

  return { text: lines.join("\n"), issues: issues, external: external };
}

function fitMessage(rep, markdown) {
  var LIMIT = 500;
  var tries = [5, 3, 2, 1];
  var m;
  for (var i = 0; i < tries.length; i++) {
    m = buildMessage(rep, markdown, tries[i]);
    if (m.text.length <= LIMIT) return m;
  }
  m.text = m.text.slice(0, LIMIT - 20) + "\n...(list shortened)";
  return m;
}

/* ============================ event handlers ============================ */

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  runCheck(item).then(function (rep) {
    var md = Office.context.requirements.isSetSupported("Mailbox", "1.15");
    var m = fitMessage(rep, md);
    var opts = { allowEvent: false };     // shows the alert with "Send" / "Don't Send"
    if (md) opts.errorMessageMarkdown = m.text; else opts.errorMessage = m.text;
    event.completed(opts);
  }).catch(function (e) {
    event.completed({
      allowEvent: false,
      errorMessage: "Send check couldn't finish (" + ((e && e.message) || "unknown error") +
        "). Please check the recipients and attachments yourself."
    });
  });
}

/* Optional "Run send check" button in the compose ribbon: shows a short summary. */
function runCheckNow(event) {
  var item = Office.context.mailbox.item;
  runCheck(item).then(function (rep) {
    var m = buildMessage(rep, false, 1);
    var nRec = rep.to.length + rep.cc.length + rep.bcc.length;
    var summary = nRec + " recipient(s), " + m.external + " external; " + rep.atts.length + " attachment(s), " +
      m.issues + " need attention. Full check appears when you click Send.";
    item.notificationMessages.replaceAsync("sendcheck", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: summary.slice(0, 150), icon: "Icon.16x16", persistent: false
    }, function () { event.completed(); });
  }).catch(function () { event.completed(); });
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("runCheckNow", runCheckNow);
}

// For local testing in Node only.
if (typeof module !== "undefined" && module.exports) {
  module.exports = { classifyFile: classifyFile, base64ToBytes: base64ToBytes, buildMessage: buildMessage,
    fitMessage: fitMessage, runCheck: runCheck, onMessageSendHandler: onMessageSendHandler };
}
