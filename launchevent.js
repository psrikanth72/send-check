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

/* ===================== remembering the original reply recipients ===================== */

// When a Reply / Reply all opens, the people already on it are saved for this compose session.
// Stored with item.sessionData (Outlook Mailbox 1.11+).
var ORIG_KEY = "sendcheck_original_recipients";
var FWD_KEY = "sendcheck_is_forward";   // set when the email was started with Forward

function addrKey(r) {
  return String((r && (r.emailAddress || r.displayName)) || "").trim().toLowerCase();
}

function allRecipients(item) {
  return Promise.all([getRecipients(item.to), getRecipients(item.cc), getRecipients(item.bcc)]);
}

function onNewComposeHandler(event) {
  var finished = false;
  function done() { if (!finished) { finished = true; try { event.completed(); } catch (e) {} } }
  setTimeout(done, 4000);
  try {
    var item = Office.context.mailbox.item;
    // A reply/reply-all already has a conversation and starts with recipients.
    // A new email has no conversation yet; a forward starts with no recipients.
    if (!item.conversationId || !item.sessionData || !item.sessionData.setAsync) { done(); return; }
    allRecipients(item).then(function (r) {
      var list = r[0].concat(r[1], r[2]).map(addrKey).filter(Boolean);
      if (!list.length) {
        // Part of an existing conversation but starting with nobody on it = a forward.
        item.sessionData.setAsync(FWD_KEY, "1", function () { done(); });
        return;
      }
      item.sessionData.setAsync(ORIG_KEY, JSON.stringify(list), function () { done(); });
    }, done);
  } catch (e) { done(); }
}

function getIsForward(item) {
  if (!item.sessionData || !item.sessionData.getAsync) return Promise.resolve(false);
  return call(function (cb) { item.sessionData.getAsync(FWD_KEY, cb); }, null).then(function (v) { return v === "1"; });
}

function getOriginalRecipients(item) {
  if (!item.sessionData || !item.sessionData.getAsync) return Promise.resolve(null);
  return call(function (cb) { item.sessionData.getAsync(ORIG_KEY, cb); }, null).then(function (v) {
    if (!v) return null;
    try { var a = JSON.parse(v); return Array.isArray(a) ? a : null; } catch (e) { return null; }
  });
}

/* ============ removing names added to a reply (OnMessageRecipientsChanged) ============ */

// On a Reply / Reply all that involves anyone outside iriskpo.in, any newly typed name is
// taken straight back out of To / Cc / Bcc, and a bar at the top of the email explains why.
// (Outlook doesn't let add-ins grey these boxes out, so this is the closest equivalent.)
var NOTICE_KEY = "sendcheck_noadd";

function onRecipientsChangedHandler(event) {
  var finished = false;
  function done() { if (!finished) { finished = true; try { event.completed(); } catch (e) {} } }
  setTimeout(done, 4000);
  try {
    var item = Office.context.mailbox.item;
    var dom = myDomain();
    Promise.all([getOriginalRecipients(item), getIsForward(item)]).then(function (st) {
      var orig = st[0], isForward = st[1];
      if (isForward) { removeExternalFromForward(item, dom, done); return; }
      if (!orig) { done(); return; }                     // a new email: nothing to enforce
      var origSet = {};
      orig.forEach(function (k) { origSet[k] = true; });
      var fields = [["to", item.to], ["cc", item.cc], ["bcc", item.bcc]];
      Promise.all(fields.map(function (f) { return getRecipients(f[1]); })).then(function (cur) {
        var everyone = cur[0].concat(cur[1], cur[2]);
        var involvesExternal = orig.some(function (k) { return isExternal(k, dom); }) ||
                               everyone.some(function (r) { return isExternal(r.emailAddress, dom); });
        var removed = [];
        var fixes = [];
        fields.forEach(function (f, i) {
          var keep = cur[i].filter(function (r) { return origSet[addrKey(r)]; });
          if (keep.length !== cur[i].length) {
            cur[i].forEach(function (r) { if (!origSet[addrKey(r)]) removed.push(r.emailAddress || r.displayName); });
            var list = keep.map(function (r) { return { displayName: r.displayName || r.emailAddress, emailAddress: r.emailAddress }; });
            fixes.push(call(function (cb) { f[1].setAsync(list, cb); }, null));
          }
        });
        if (!involvesExternal || !removed.length) { done(); return; }
        Promise.all(fixes).then(function () {
          var msg = "Removed " + removed.join(", ") + ": new recipients can't be added to a reply on an external email. Start a new email if they need this.";
          if (msg.length > 150) msg = "Removed " + removed.length + " added recipient(s): new recipients can't be added to a reply on an external email.";
          try {
            item.notificationMessages.replaceAsync(NOTICE_KEY, {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message: msg.slice(0, 150), icon: "Icon.16x16", persistent: false
            }, function () { done(); });
          } catch (e) { done(); }
        });
      }, done);
    }, done);
  } catch (e) { done(); }
}

// Forward: colleagues at iriskpo.in are fine, any external address is taken back out.
function removeExternalFromForward(item, dom, done) {
  var fields = [item.to, item.cc, item.bcc];
  Promise.all(fields.map(getRecipients)).then(function (cur) {
    var removed = [], fixes = [];
    fields.forEach(function (f, i) {
      var keep = cur[i].filter(function (r) { return !isExternal(r.emailAddress, dom); });
      if (keep.length !== cur[i].length) {
        cur[i].forEach(function (r) { if (isExternal(r.emailAddress, dom)) removed.push(r.emailAddress); });
        var list = keep.map(function (r) { return { displayName: r.displayName || r.emailAddress, emailAddress: r.emailAddress }; });
        fixes.push(call(function (cb) { f.setAsync(list, cb); }, null));
      }
    });
    if (!removed.length) { done(); return; }
    Promise.all(fixes).then(function () {
      var msg = "Removed " + removed.join(", ") + ": emails can't be forwarded outside iriskpo.in.";
      if (msg.length > 150) msg = "Removed " + removed.length + " external recipient(s): emails can't be forwarded outside iriskpo.in.";
      try {
        item.notificationMessages.replaceAsync(NOTICE_KEY, {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: msg.slice(0, 150), icon: "Icon.16x16", persistent: false
        }, function () { done(); });
      } catch (e) { done(); }
    });
  }, done);
}

/* ============================ the check itself ============================ */

function isExternal(addr, dom) {
  if (!dom) return false;
  var d = domainOf(addr);
  if (!d) return false;
  return !(d === dom || d.slice(-(dom.length + 1)) === "." + dom);   // subdomains count as internal
}

// "progress" lets the safety net know what was already found if the check runs out of time.
function runCheck(item, progress) {
  progress = progress || {};
  var dom = myDomain();
  var recipientsP = allRecipients(item).then(function (r) {
    progress.external = r[0].concat(r[1], r[2]).some(function (x) { return isExternal(x.emailAddress, dom); });
    return r;
  });
  var origP = getOriginalRecipients(item);
  var attsP = call(function (cb) { item.getAttachmentsAsync(cb); }, []).then(function (list) {
    var atts = (list || []).filter(function (a) { return !a.isInline; });
    progress.attachments = atts.length;
    // Scan all files at the same time; anything not finished in time is "check manually".
    var scans = atts.map(function (a) { return withTimeout(checkAttachment(item, a), TOTAL_TIMEOUT_MS, CHECK_MANUALLY); });
    return Promise.all(scans).then(function (results) { return { atts: atts, results: results }; });
  });
  var fwdP = getIsForward(item);
  return Promise.all([recipientsP, attsP, origP, fwdP]).then(function (r) {
    var rep = { domain: dom, to: r[0][0], cc: r[0][1], bcc: r[0][2], atts: r[1].atts, results: r[1].results,
                isReply: !!r[2], isForward: !!r[3], added: [] };
    if (r[2]) {
      var orig = {};
      r[2].forEach(function (k) { orig[k] = true; });
      rep.added = rep.to.concat(rep.cc, rep.bcc).filter(function (x) { return !orig[addrKey(x)]; });
    }
    return rep;
  });
}

function formatRecipients(list, dom, maxShown, externalOnly) {
  var ext = 0, shown = [];
  list.forEach(function (r) {
    var a = r.emailAddress || r.displayName || "?";
    if (isExternal(a, dom)) { ext++; shown.push(a + " (EXTERNAL)"); }
    else if (!externalOnly) shown.push(a);
  });
  var extra = shown.length - maxShown;
  if (extra > 0) shown = shown.slice(0, maxShown).concat(["+" + extra + " more"]);
  return { text: shown.join(", "), external: ext };
}

function shorten(name, shortNames) {
  return shortNames && name.length > 30 ? name.slice(0, 27) + "..." : name;
}

/* Decides what happens and builds the pop-up text (Outlook allows at most 500 characters).
   Blocked when the email goes outside iriskpo.in and:
     (a) it's a reply / reply all and recipients were added, or
     (b) an attachment isn't (or can't be confirmed as) password protected. */
function buildMessage(rep, maxPerField, shortNames) {
  var lines = [], unprotected = 0, external = 0;
  var all = rep.to.concat(rep.cc, rep.bcc);
  all.forEach(function (r) { if (isExternal(r.emailAddress, rep.domain)) external++; });
  rep.results.forEach(function (res) { if (res.code !== PROTECTED) unprotected++; });

  var addedBlock = external > 0 && rep.isReply && rep.added.length > 0;
  var forwardBlock = external > 0 && rep.isForward;
  var fileBlock = external > 0 && unprotected > 0 && !forwardBlock;   // forward message already says it all

  if (addedBlock || fileBlock || forwardBlock) {
    lines.push("EMAIL BLOCKED - this email is going outside iriskpo.in.");
    if (forwardBlock) {
      lines.push("");
      lines.push("Emails can't be forwarded outside iriskpo.in. Please remove:");
      lines.push(formatRecipients(all, rep.domain, maxPerField, true).text);
    }
    if (addedBlock) {
      lines.push("");
      lines.push("You can't add new recipients to a reply. Please remove:");
      lines.push(rep.added.map(function (r) {
        var a = r.emailAddress || r.displayName || "?";
        return a + (isExternal(a, rep.domain) ? " (EXTERNAL)" : "");
      }).slice(0, maxPerField).join(", ") + (rep.added.length > maxPerField ? " +" + (rep.added.length - maxPerField) + " more" : ""));
      lines.push("If they need this, start a new email instead.");
    }
    if (fileBlock) {
      lines.push("");
      lines.push("These attachments must be password protected:");
      rep.atts.forEach(function (a, i) {
        var res = rep.results[i];
        if (res.code === NOT_PROTECTED) lines.push("✘ " + shorten(a.name, shortNames));
        else if (res.code !== PROTECTED) lines.push("? " + shorten(a.name, shortNames) + " - could not be verified");
      });
      lines.push("Add a password (or use a password-protected ZIP), re-attach and send again.");
    }
    return { text: lines.join("\n"), unprotected: unprotected, external: external, blocked: true };
  }

  lines.push("Have you included the correct recipients?");
  [["To", rep.to], ["Cc", rep.cc], ["Bcc", rep.bcc]].forEach(function (f) {
    if (f[1].length) lines.push(f[0] + ": " + formatRecipients(f[1], rep.domain, maxPerField).text);
  });
  if (!all.length) lines.push("(no recipients)");

  if (rep.atts.length) {
    lines.push("");
    lines.push(unprotected ? "File Not Password Protected - please check:" : "All attachments are password protected:");
    rep.atts.forEach(function (a, i) {
      var res = rep.results[i], name = shorten(a.name, shortNames);
      if (res.code === PROTECTED) lines.push("✔ " + name + " - protected");
      else if (res.code === NOT_PROTECTED) lines.push("✘ " + name + " - NOT password protected");
      else lines.push("? " + name + " - " + res.note);
    });
  }
  return { text: lines.join("\n"), unprotected: unprotected, external: external, blocked: false };
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

/* The manifest sets SendMode="SoftBlock": an alert only has "Don't Send" unless we switch it
   back to a normal prompt (with "Send anyway") for emails that are allowed to go. */
function promptOptions(text) {
  try {
    if (Office.context.requirements.isSetSupported("Mailbox", "1.14") &&
        Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
      return { allowEvent: false, errorMessage: text,
               sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser };
    }
  } catch (e) {}
  return null;   // very old Outlook can't show "Send anyway", so allowed emails go without the reminder
}

function onMessageSendHandler(event) {
  var finished = false, progress = {};
  function done(opts) {
    if (finished) return;
    finished = true;
    try { event.completed(opts); } catch (e) { try { event.completed({ allowEvent: true }); } catch (e2) {} }
  }
  function prompt(text) { done(promptOptions(text) || { allowEvent: true }); }
  function block(text) { done({ allowEvent: false, errorMessage: text }); }

  // Safety net: whatever happens, respond within 5 seconds.
  setTimeout(function () {
    if (progress.external && progress.attachments > 0) {
      block("EMAIL BLOCKED - this email is going outside iriskpo.in with attachments, and Send Check couldn't confirm they are password protected. Please try again in a moment.");
    } else {
      prompt("Have you included the correct recipients?\n\nSend Check couldn't finish in time - please check the recipients and attachments yourself.");
    }
  }, 5000);

  try {
    runCheck(Office.context.mailbox.item, progress).then(function (rep) {
      var m = fitMessage(rep);
      if (m.blocked) block(m.text); else prompt(m.text);
    }, function () {
      prompt("Have you included the correct recipients?\n\nSend Check couldn't read this email - please check the recipients and attachments yourself.");
    });
  } catch (e) {
    prompt("Have you included the correct recipients?\n\nPlease check the recipients and attachments yourself.");
  }
}

/* Optional "Run send check" button in the compose ribbon: shows a short summary. */
function runCheckNow(event) {
  try {
    var item = Office.context.mailbox.item;
    runCheck(item).then(function (rep) {
      var m = buildMessage(rep, 1, true);
      var n = rep.to.length + rep.cc.length + rep.bcc.length;
      var summary = m.blocked
        ? "This email will be BLOCKED when you click Send. Click Send to see why."
        : n + " recipient(s); " + rep.atts.length + " attachment(s), " + m.unprotected + " not password protected.";
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
      Office.actions.associate("onNewComposeHandler", onNewComposeHandler);
      Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
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
    fitMessage: fitMessage, runCheck: runCheck, onMessageSendHandler: onMessageSendHandler,
    onNewComposeHandler: onNewComposeHandler, onRecipientsChangedHandler: onRecipientsChangedHandler };
}
