/**
 * EASY BUYS – CLOUD  (started as the Easy Buys Bakery cloud)
 * Runs inside the Easy Buys Google account (Google Apps Script).
 *  - Keeps every device in step: the shop tablet, other tablets and the owners' phones
 *    (staff, Time Clock punches, timesheets, leave, and every future module)
 *  - Devices list: who is linked, when they last synced; block a lost device
 *  - Invites for new devices: QR code, email or a short code (single use, 48 hours)
 *  - The main Easy Buys software on the office PC collects everything with its own key
 *  - Bakery: live dashboard, shift reports, purchase orders, photos and backups in Drive
 *
 * SETUP: change the two values below, Deploy → New deployment → Web app
 *        (Execute as: Me, Who has access: Anyone), then run setup() once. See CLOUD-SETUP.md.
 */

// 1) Paste the "Cloud key" shown in the app (Admin → Online & email). Only the first device needs it;
//    other devices are linked with invites.
const API_KEY = 'PASTE-THE-KEY-FROM-THE-APP-HERE';

// 2) Choose a password for opening the online dashboard.
const DASHBOARD_PASSWORD = 'choose-a-dashboard-password';
// 1.0.0 build 1: better – keep both in Project Settings › Script properties as API_KEY and DASHBOARD_PASSWORD (then the
// lines above can stay as they are; the Script properties win). See CLOUD-SETUP.md.
function apiKey_() { return PropertiesService.getScriptProperties().getProperty('API_KEY') || API_KEY; }

const ROOT_NAME = 'Easy Buys Bakery Cloud';   // existing folder name kept so nothing moves
const DATA_NAME = 'Easy Buys Data (do not edit by hand)';
const INVITE_HOURS = 48;
const SCRIPT_VERSION = 18;                // 1.0.0 build 10 (go-live, 2 Nov): sync under many phones – the write lock only around writes and short (busy → {busy, retryMs} at once instead of a hang), pulls read the rows in a few blocks instead of one by one, the devices list cached, a device's last-seen written at most every 2 minutes, refusals logged after the lock; job_lease (any device does the main-tablet jobs, one at a time); Viber posts once (once key).  1.0.0 build 9: Start fresh needs the owner's session (any device).  1.0.0 (owner, 2 Nov): readable photo names in Drive › Records › month › shop (Invoices / Cheques) and month (Deposits / Z reports / Payslips); tidy_records moves the old ones; payslip_save.  1.0.0 build 6: backups go up in resumable parts (backup_part / backup_done / backup_get_part).  1.0.0 build 5: "Get the End of day now" (link_pull_request; link_check / link_push answer pull: true to that shop's till PCs).  1.0.0 build 4: recharge machine wallets sync (recharge_sales, recharge_topups, recharge_setup).  1.0.0 build 1 (new app): devices sign every request with their pass (unknown / removed devices and wrong passes refused), PINs hashed (PIN_SALT) and never kept plain, role / PIN / locked-setting changes only from the owner's or admin's device, owner Viber channel for money alerts, setup_pack (the supplier set-up pack from Drive), clean_start (owner PIN checked here, cloud copy kept).  1.0.0 build 4: FCCC gas read by the cloud (fccc_lpg), bills read by AI (account and due date), Gemini model chain.  1.0.0: AI invoice reading (invoice_ai, GEMINI_KEY), page-by-page OCR, PDFs saved as PDFs, dashboard password lock.  v3.9.8: one Viber Channel per shop (each shop's alerts go to its own channel; the rest to the main one).  v3.9.5: Viber copies in Drive › Viber posts are binned after 30 days; Viber sets its own webhook. v3.7 / 1.0: documents (checked uploads, visibility, nightly backups, Windows file API); device control (pause, end date, remove, wipe, commands); v3.6: Start fresh
const SUPPORT_EMAIL = 'prasaddiptesh@gmail.com'; // "Send to support" problem reports (Admin & support)
const PULL_LIMIT = 800;
const MAIN_TABLES_PULL_ALL = true;   // the main software receives every table, including the change history

// ---------------------------------------------------------------- helpers
// The log spreadsheet: found once (its id kept), not searched for in Drive on every log line.
function logBook_() {
  if (REQ_.log) return REQ_.log;
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('LOG_ID');
  let ss = null;
  if (id && SpreadsheetApp.openById) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    const ssIt = root_().getFilesByName('Bakery Cloud Log');
    ss = ssIt.hasNext() ? SpreadsheetApp.open(ssIt.next()) : (function () {
      const created = SpreadsheetApp.create('Bakery Cloud Log');
      DriveApp.getFileById(created.getId()).moveTo(root_());
      return created;
    })();
    try { if (ss.getId) p.setProperty('LOG_ID', ss.getId()); } catch (e) { /* searched again next time */ }
  }
  REQ_.log = ss;
  return ss;
}
function json_(obj) {
  // 1.0.0 build 1 (hardening): every reply carries the cloud's own clock (ms since 1970, UTC) so a device can tell when
  // its clock is wrong (the app refuses money work then – the time is always Fiji time).
  if (obj && typeof obj === 'object' && !Array.isArray(obj) && obj.serverTime === undefined) obj.serverTime = Date.now();
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function root_() {
  const it = DriveApp.getFoldersByName(ROOT_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME);
}
function folder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function path_(names) {
  let f = root_();
  names.filter(String).forEach(function (n) { f = folder_(f, n); });
  return f;
}
// 1.0.0 build 1: a file is never replaced or binned – a second file with the same name gets " (2)", " (3)"… so a photo,
// report or backup already kept can't be swapped by sending another with its name.
function saveFile_(folder, fileName, base64, mime) {
  let name = String(fileName);
  if (folder.getFilesByName(name).hasNext()) {
    const dot = name.lastIndexOf('.');
    const stem = dot > 0 ? name.slice(0, dot) : name; const ext = dot > 0 ? name.slice(dot) : '';
    for (let i = 2; i < 500 && folder.getFilesByName(name).hasNext(); i++) name = stem + ' (' + i + ')' + ext;
  }
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, name);
  return folder.createFile(blob);
}
// Only for the cloud's own working files (the dashboard figures): replaced each time.
function replaceFile_(folder, fileName, base64, mime) {
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mime, fileName));
}
function log_(sheetName, row) {
  if (REQ_.locked) { later_(function () { log_(sheetName, row); }); return; } // 1.0.0 build 10: never inside the lock
  const ss = logBook_();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.appendRow([new Date()].concat((row || []).map(function (v) { return typeof v === 'string' ? esc_(v) : v; })));
  // 1.0.0 build 1: every refusal is a security event for the owner (wrong PINs count up separately: failCount_).
  if (sheetName === 'Refused' && String((row || [])[1]) !== 'user_login') secEvent_('refused', { device: String((row || [])[0] || ''), what: (row || []).slice(1).join(' – ') });
}
function cleanEmails_(list) {
  return (list || []).map(function (e) { return String(e).trim(); }).filter(function (e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); });
}

// ---------------------------------------------------------------- tablet → cloud
function doPost(e) {
  REQ_ = {}; // what this request has read already (a request never reads the same sheet twice)
  try {
    const body = JSON.parse(e.postData.contents);
    // Viber calls this address (the channel's webhook) to check it is alive, and with channel events: just say OK.
    if (body && body.event && !body.kind) return json_({ ok: true });
    // --- the main Easy Buys software (office PC) ---
    if (body.kind === 'main_pull' || body.kind === 'main_push' || body.kind === 'main_devices' || body.kind === 'main_link_check') return json_(mainApi_(body));
    if (body.kind === 'main_files' || body.kind === 'main_file_get') return json_(mainFiles_(body)); // v3.5 Windows software: full file mirror
    // --- 1.0.0 build 1: a till PC's own key (its Aronium records and its link check only) ---
    if (body.kind === 'link_push' || body.kind === 'link_check') return json_(linkApi_(body));
    // --- a new device joining with an invite (no key yet) ---
    if (body.kind === 'join') return json_(join_(body));
    // --- 1.0.0 build 1: email logins (off unless the owner switched them on) ---
    if (body.kind === 'email_login_start') return json_(emailLoginStart_(body));
    if (body.kind === 'email_login_finish') return json_(emailLoginFinish_(body));
    const who = auth_(body);
    if (!who.ok) return json_({ ok: false, error: who.error, blocked: !!who.blocked, waiting: !!who.waiting, lock: who.lock || null, wipe: !!who.wipe, removed: !!who.removed });
    // Only now (a known device): housekeeping that writes settings.
    try { viberUnshare_(); } catch (x) { /* never stops a request */ }
    // 1.0.0 build 1 (script 12): PINs in the cloud copy are hashed once (and every PIN that arrives is hashed).
    try { migratePins_(); } catch (x) { /* tried again on the next request */ }
    who.user = sessionUser_(body, who);
    if (body.kind === 'user_login') return json_(userLogin_(body, who));
    if (body.kind === 'link_pull_request') return json_(linkPullRequest_(body, who));
    if (body.kind === 'push') return json_(push_(body.records || [], who.deviceId, who));
    if (body.kind === 'pull') { const out = pull_(Number(body.since) || 0, who.deviceId, false, who, Number(body.limit) || 0); devicePulled_(who.deviceId, body, out);
      // 1.0.0 build 10: new phones waiting for the owner's OK – the owner's / an admin's device asks "Allow …?" at once
      if (who.admin || isBoss_(who)) out.waitingDevices = deviceRows_().filter(function (r) { return String(r[3]) === 'waiting'; }).map(function (r) { return { id: String(r[0]), name: String(r[1] || ''), kind: String(r[2] || ''), at: text_(r[5]) }; }).slice(0, 10); out.version = SCRIPT_VERSION; out.pinSalt = pinSalt_(body.pinSalt); out.admin = !!who.admin; out.user = who.user ? who.user.role : null; return json_(out); }
    if (body.kind === 'invite_create') return json_(needOwner_(who, body, 'Adding a device') || inviteCreate_(body, who));
    if (body.kind === 'invite_cancel') return json_(needOwner_(who, body, 'Cancelling an invite') || inviteCancel_(body, who));
    if (body.kind === 'devices') return json_({ ok: true, devices: devices_() });
    if (body.kind === 'device_status') return json_(needOwner_(who, body, 'Blocking or allowing a device') || deviceStatus_(body.id, body.status, who));
    if (body.kind === 'device_control') return json_(needOwner_(who, body, 'Changing a device') || deviceControl_(body, who));
    if (body.kind === 'problem_log') return json_(problemLogSave_(body, who));
    if (body.kind === 'problem_log_get') return json_(problemLogGet_(body.id));
    if (body.kind === 'cloud_settings') return json_((body.requireApproval === undefined && body.strictSign === undefined && body.linkCutover === undefined ? null : needOwner_(who, body, 'Changing the cloud settings')) || cloudSettings_(body));
    if (body.kind === 'fresh_start') return json_(needOwner_(who, body, 'Start fresh') || freshStart_(body)); // 1.0.0 build 9: the owner only (any device the owner is logged in on)
    if (body.kind === 'clean_start') return json_(needOwner_(who, body, 'Start clean') || cleanStart_(body, who));
    if (body.kind === 'tidy_records') return json_(needOwner_(who, body, 'Tidy Drive photos') || tidyRecords_(body.maxMs));
    if (body.kind === 'setup_pack') return json_(needOwner_(who, body, 'The supplier set-up pack') || setupPack_());
    if (body.kind === 'email_password_set') return json_(emailPasswordSet_(body, who));
    if (body.kind === 'security_email') {
      // where security alerts go: only the owner, logged in, with the PIN typed again
      if (!who.user || who.user.role !== 'owner' || !proofOk_(who.user.uid, body.ownerProof)) return json_({ ok: false, error: 'The owner must be logged in and type their PIN again.' });
      const se = cleanEmails_([body.email])[0];
      if (!se) return json_({ ok: false, error: 'That does not look like an email address.' });
      const before = secEmail_();
      props_().setProperty('SECURITY_EMAIL', se);
      secEvent_('locked_change', { device: who.deviceId, person: who.user.name, what: 'Security alerts now go to ' + se + ' (were ' + before + ')', refused: false });
      return json_({ ok: true, email: se });
    }
    if (body.kind === 'link_keys') return json_({ ok: true, tills: Object.keys(linkKeys_()).map(function (h) { const x = linkKeys_()[h]; return { source: x.source, at: x.at, by: x.by }; }) });
    if (body.kind === 'link_key') return json_(needOwner_(who, body, 'A till PC key') || linkKeyCreate_(body, who));
    if (['backup_list', 'backup_get', 'backup_get_part', 'myday_events', 'myday_out'].indexOf(body.kind) >= 0) { const no = needOwner_(who, body, body.kind.indexOf('backup') === 0 ? 'Backups' : 'The owner\'s calendar'); if (no) return json_(no); }
    switch (body.kind) {
      case 'ping':
        return json_({ ok: true, account: Session.getEffectiveUser().getEmail(), folder: root_().getUrl(), device: who.deviceId, version: SCRIPT_VERSION, pinSalt: pinSalt_(body.pinSalt), admin: !!who.admin, user: who.user ? who.user.role : null });

      case 'snapshot': {
        // 1.0.0 build 1: the online dashboard's figures come only from the main tablet (the one holding the cloud key).
        // 1.0.0 build 9 (owner, 2 Nov: "any device should be able to do it"): any linked device with the owner, an admin
        // or a manager logged in sends the dashboard (their device has the whole picture); a staff-only phone does not
        if (!who.viaKey && !isAdminDevice_(who.deviceId, false) && !(who.user && ['owner', 'admin', 'manager', 'supervisor'].indexOf(who.user.role) >= 0)) return json_({ ok: true, ignored: true, code: 'NOT_MAIN', note: 'Only the main tablet sends the dashboard figures.', version: SCRIPT_VERSION }); // 2 Nov: ignored, never an error – an error stopped older apps' whole sync
        const props = PropertiesService.getScriptProperties();
        const text = JSON.stringify(body.snapshot);
        // Script properties hold up to ~9 KB per value, so larger snapshots go to a Drive file.
        if (text.length < 8500) { props.setProperty('SNAPSHOT', text); props.deleteProperty('SNAPSHOT_FILE'); }
        else {
          const f = replaceFile_(root_(), 'dashboard-status.json', Utilities.base64Encode(text, Utilities.Charset.UTF_8), 'application/json');
          props.setProperty('SNAPSHOT_FILE', f.getId());
          props.deleteProperty('SNAPSHOT');
        }
        props.setProperty('SNAPSHOT_AT', new Date().toISOString());
        // Hand back production-plan changes made on the online dashboard that the tablet has not applied yet.
        const maxSeq = Number(props.getProperty('PLAN_SEQ') || '0');
        const since = Number(body.planSeq || 0) > maxSeq ? 0 : Number(body.planSeq || 0); // cloud program was reset → resend
        const edits = planEdits_().filter(function (x) { return x.seq > since; });
        return json_({ ok: true, planEdits: edits, version: SCRIPT_VERSION });
      }

      case 'report': {
        // script 16: Records › month › Shift reports › "2026-10-14 Shift report - Bakery - SH-0042" (older apps: the old place)
        const niceR = recordsFolder_(body, ['Shift reports']);
        const folder = niceR ? niceR.folder : path_(['Shift reports', body.month || 'Other']);
        const stemR = niceR ? niceR.name : body.fileName;
        const pdf = saveFile_(folder, stemR + '.pdf', body.pdf, MimeType.PDF);
        const xlsx = body.xlsx ? saveFile_(folder, stemR + '.xlsx', body.xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') : null;
        // 1.0.0 build 1: only to addresses the owner saved (Admin › Emails), and the device's text is shown as text, not HTML.
        const to = onlyAllowed_(cleanEmails_(body.emails), false);
        if (to.length) {
          mail_({
            to: to.join(','), subject: String(body.subject || '').slice(0, 200),
            htmlBody: '<p>The shift has been balanced and approved. The full report is attached (PDF and Excel) and saved in Google Drive.</p>' + asText_(body.html),
            attachments: [pdf.getBlob()].concat(xlsx ? [xlsx.getBlob()] : []),
            name: 'Easy Buys Bakery',
          });
        }
        log_('Reports', [body.shiftCode, body.fileName, to.join(', '), pdf.getUrl()]);
        return json_({ ok: true, url: pdf.getUrl(), emailed: to.length });
      }

      case 'po': {
        // script 16: Records › month › Purchase orders › "PO 0123 - Bijay - Week 2 (2026-11-09).pdf" (older apps: the old place)
        const niceP = recordsFolder_(body, ['Purchase orders']);
        const folder = niceP ? niceP.folder : path_(['Purchase orders', body.month || 'Other']);
        const pdf = saveFile_(folder, (niceP ? niceP.name : body.fileName) + '.pdf', body.pdf, MimeType.PDF);
        const to = onlyAllowed_(cleanEmails_(body.emails), true);
        if (to.length) {
          mail_({
            to: to.join(','), subject: String(body.subject || '').slice(0, 200),
            htmlBody: '<p>Stock has reached the minimum level. Purchase order attached for reordering'
              + (body.supplierEmail ? ' (supplier email: ' + asText_(body.supplierEmail) + ')' : '') + '.</p>' + asText_(body.html),
            attachments: [pdf.getBlob()], name: 'Easy Buys Bakery',
          });
        }
        log_('Purchase orders', [body.fileName, to.join(', '), pdf.getUrl()]);
        return json_({ ok: true, url: pdf.getUrl(), emailed: to.length });
      }

      case 'photo': {
        // 1.0.0 script 16: invoices, cheques / payments, deposits and Z reports go in Records › month › shop › kind with a
        // readable name (the app works it out: services/photoNames.js). Older apps (no niceName) keep the old place.
        const nice = recordsFolder_(body, ['Invoices', 'Cheques', 'Deposits', 'Z reports']);
        const folder = nice ? nice.folder : path_(['Photos', body.month || 'Other', body.refType || 'other']);
        const f = saveFile_(folder, nice ? nice.name : body.fileName, body.data, body.mime === 'application/pdf' ? MimeType.PDF : MimeType.JPEG);
        return json_({ ok: true, url: f.getUrl() });
      }

      case 'ocr': {
        // Read the text of invoice photos with Google's free OCR (needs Services › Drive API added once).
        // 1.0.0 each page on its own: one bad page no longer fails the others (its problem is sent back in pageErrors).
        if (typeof Drive === 'undefined' || !Drive.Files) return json_({ ok: true, off: true });
        const texts = [];
        const pageErrors = [];
        (body.files || []).slice(0, 6).forEach(function (f, i) {
          try {
            const t = ocr_(Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType || 'image/jpeg', 'page-' + (i + 1) + '.jpg'));
            if (t) texts.push(t); else pageErrors.push('Page ' + (i + 1) + ': no text found');
          } catch (err) {
            pageErrors.push('Page ' + (i + 1) + ': ' + String(err && err.message ? err.message : err).slice(0, 200));
          }
        });
        return json_({ ok: true, text: texts.join('\n'), pageErrors: pageErrors });
      }

      case 'invoice_ai': {
        // 1.0.0 an invoice (photos or a PDF) read by Google Gemini with the owner's own free key (Script properties ›
        // GEMINI_KEY). Answers { ok, invoice: {…}, model } or { ok: false, code: 'NO_AI_KEY' | 'AI_ERROR', error }.
        return json_(aiReadInvoice_(body.files || [], body.hints || {}));
      }

      case 'fccc_lpg':
        // 1.0.0 build 54: the newest FCCC LPG authorisation read here (the phone could not): see fcccLpg_ at the end.
        return json_(fcccLpg_());

      case 'file_put': {
        // v3.5 apps send an upload id + fingerprint: filed by type, checked, never duplicated (see filePut_).
        if (body.uploadId) return json_(filePut_(body, who));
        // Documents kept by the app (licences, applications, contracts…): Easy Buys › Documents › <section>.
        const sect = String(body.folder || 'other').replace(/[^A-Za-z0-9 _-]/g, '').slice(0, 40) || 'other';
        const name = String(body.fileName || 'document').replace(/[\\/]/g, '-').slice(0, 120);
        const f = saveFile_(path_(['Documents', sect]), name, body.data, String(body.mime || 'application/octet-stream'));
        return json_({ ok: true, id: f.getId(), url: f.getUrl() });
      }

      case 'file_get': {
        // v3.5: the role of the person logged in on the device is checked against the document visibility setting.
        if (body.user || fileRow_(String(body.id || ''))) return json_(fileGet_(body, who));
        // Only files kept under Easy Buys › Documents can be fetched.
        const file = DriveApp.getFileById(String(body.id || ''));
        const docs = path_(['Documents']).getId();
        let ok = false;
        const parents = file.getParents();
        while (parents.hasNext() && !ok) { const p = parents.next(); const up = p.getParents(); if (p.getId() === docs || (up.hasNext() && up.next().getId() === docs)) ok = true; }
        if (!ok) return json_({ ok: false, error: 'Not an Easy Buys document.' });
        return json_({ ok: true, data: Utilities.base64Encode(file.getBlob().getBytes()), mime: file.getMimeType(), name: file.getName() });
      }

      case 'backup': {
        if (body.tier) return json_(backupTiered_(body, who)); // v3.5: daily (30 days) / monthly (12 months) + manifest
        const folder = path_(['Backups']);
        saveFile_(folder, body.fileName, Utilities.base64Encode(body.data, Utilities.Charset.UTF_8), 'application/json');
        return json_({ ok: true });
      }

      // 1.0.0 speed: a backup sent in parts of a few MB (the app goes on where it stopped if it is cut off).
      case 'backup_part': return json_(backupPart_(body, who));
      case 'backup_done': return json_(backupDone_(body, who));
      case 'backup_get_part': return json_(backupGetPart_(body, who));

      case 'backup_list': {
        // v3.5: the backups kept in Drive › Backups, › Daily, › Monthly and › Manual (newest first) for Admin › Backup & restore.
        return json_({ ok: true, files: backupFiles_().slice(0, Number(body.limit) || 60) });
      }

      case 'backup_get': {
        // Only files kept in Easy Buys › Backups can be fetched this way.
        const file = DriveApp.getFileById(String(body.id || ''));
        if (!insideFolder_(file, path_(['Backups']).getId(), 2)) return json_({ ok: false, error: 'Not an Easy Buys backup.' });
        secEvent_('backup', { device: who.deviceId, person: who.user ? who.user.name : '', what: 'Backup downloaded: ' + file.getName(), refused: false });
        return json_({ ok: true, name: file.getName(), text: file.getBlob().getDataAsString('UTF-8') });
      }

      case 'viber_setup': {
        // Viber Channel "Post API": the token comes from the channel's Settings › Developer tools. Kept only here.
        // v7: the Viber token comes as viberToken – "token" is the device's own pass (it was overwritten before).
        const token = String(body.viberToken || '').trim();
        // 1.0.0 build 1: Viber tokens are locked – only from the owner's / admin's device.
        const na = needAdmin_(who, 'Connecting a Viber Channel');
        if (na) return json_(na);
        // v9: a shop's own channel (shopCode) or the main channel. 1.0.0 build 1: shopCode 'OWNER' = the owner / admin
        // channel, where business-wide money alerts go.
        const sfx = viberSuffix_(body.shopCode);
        const current = PropertiesService.getScriptProperties().getProperty('VIBER_TOKEN' + sfx);
        // Replacing or removing a token needs the current one (so a lost tablet cannot take the alerts over).
        if (current && String(body.oldToken || '').trim() !== current) return json_({ ok: false, error: 'A Viber Channel is already connected. Enter its current token to replace or remove it.' });
        if (!token) { PropertiesService.getScriptProperties().deleteProperty('VIBER_TOKEN' + sfx); PropertiesService.getScriptProperties().deleteProperty('VIBER_FROM' + sfx); return json_({ ok: true, message: 'Viber Channel disconnected.' }); }
        PropertiesService.getScriptProperties().setProperty('VIBER_TOKEN' + sfx, token);
        PropertiesService.getScriptProperties().deleteProperty('VIBER_FROM' + sfx);
        if (/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(String(body.selfUrl || ''))) PropertiesService.getScriptProperties().setProperty('VIBER_HOOK_URL', String(body.selfUrl));
        const t = viber_(sfx ? 'Easy Buys alerts for shop ' + sfx.slice(1) + ' are connected to this channel.' : 'Easy Buys alerts are connected to this channel.', body.shopCode);
        if (!t.ok) PropertiesService.getScriptProperties().deleteProperty('VIBER_TOKEN' + sfx);
        return json_(t.ok ? { ok: true, message: 'Connected – a test message was sent to the channel.' } : t);
      }

      case 'email': {
        // Emails to staff (welcome, signed contract). Sent from the Google account the cloud runs on.
        const to = cleanEmails_([body.to]);
        if (!to.length) return json_({ ok: false, error: 'No valid email address.' });
        // Only to people being onboarded (their email is in the synced onboarding records) – a device cannot email anyone else.
        if (!staffEmail_(to[0])) return json_({ ok: false, error: 'This address is not a staff member being onboarded.' });
        const atts = (body.attachments || []).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/pdf', a.name || 'document.pdf'); });
        mail_({ to: to.join(','), subject: String(body.subject || 'Easy Buys'), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to.join(', '), body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true });
      }

      case 'payslip_save': {
        // 1.0.0 script 16: a copy of a released payslip (PDF) in Records › month › Payslips. Nothing is replaced (" (2)").
        if (body.typeFolder !== 'Payslips') return json_({ ok: false, error: 'Only payslips are kept here.' });
        const nice = recordsFolder_({ niceName: body.fileName, monthFolder: body.monthFolder, shopFolder: '', typeFolder: 'Payslips' }, ['Payslips']);
        if (!nice || !/\.pdf$/i.test(nice.name)) return json_({ ok: false, error: 'No month or name for the payslip.' });
        const f = saveFile_(nice.folder, nice.name, body.data, MimeType.PDF);
        return json_({ ok: true, url: f.getUrl() });
      }

      case 'payslip_email': {
        // v3.5 payslips: only to an address saved in a person's pay details (or onboarding) – a device cannot email anyone else.
        const to = cleanEmails_([body.to]);
        if (!to.length) return json_({ ok: false, error: 'No valid email address.' });
        if (!payEmail_(to[0]) && !staffEmail_(to[0])) return json_({ ok: false, error: 'This address is not in anyone\'s pay details.' });
        const atts = (body.attachments || []).slice(0, 2).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/pdf', a.name || 'payslip.pdf'); });
        mail_({ to: to[0], subject: String(body.subject || 'Easy Buys payslip').slice(0, 200), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to[0], body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true });
      }

      case 'pack_email': {
        // 1.0 monthly accountant pack: only to the addresses the owner saved in Admin › Emails (synced setting).
        const to = prefEmails_('accountant_pack');
        if (!to.length) return json_({ ok: false, error: 'No accountant email saved in Admin › Emails.' });
        const atts = (body.attachments || []).slice(0, 3).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/octet-stream', a.name || 'pack.xlsx'); });
        mail_({ to: to.join(','), subject: String(body.subject || 'Easy Buys accountant pack').slice(0, 200), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to.join(', '), body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true, emailed: to.length });
      }

      case 'support_email': {
        // v3.5 "Send to support": problem reports go only to the app's support address (fixed here, not chosen by the device).
        mailEssential_({ to: SUPPORT_EMAIL, subject: String(body.subject || 'Easy Buys problem report').slice(0, 200),
          body: String(body.text || '').slice(0, 60000) + '\n\nDevice: ' + who.deviceId, name: 'Easy Buys' });
        log_('Emails', [SUPPORT_EMAIL, body.subject, 'support']);
        return json_({ ok: true });
      }

      case 'calendar': {
        // The roster week: one Google Calendar per shop ("Easy Buys Roster – S1"). The week is replaced each time it is sent.
        const tz = Session.getScriptTimeZone();
        const byShop = {};
        (body.shifts || []).forEach(function (x) { (byShop[x.shop] = byShop[x.shop] || []).push(x); });
        const from = new Date(body.weekStart + 'T00:00:00');
        const to = new Date(body.weekEnd + 'T23:59:59');
        let n = 0;
        // Clear this week in every roster calendar first (also shops with no shifts now).
        CalendarApp.getAllOwnedCalendars().filter(function (c) { return c.getName().indexOf('Easy Buys Roster – ') === 0; }).forEach(function (c) {
          c.getEvents(from, to).forEach(function (ev) { if (ev.getTag('eb') === '1' && ev.getStartTime() >= from) ev.deleteEvent(); });
        });
        Object.keys(byShop).forEach(function (shop) {
          const name = 'Easy Buys Roster – ' + shop;
          const cal = CalendarApp.getCalendarsByName(name)[0] || CalendarApp.createCalendar(name, { timeZone: tz });
          byShop[shop].forEach(function (x) {
            const st = new Date(x.date + 'T' + x.start + ':00');
            let en = new Date(x.date + 'T' + x.end + ':00');
            if (en <= st) en = new Date(en.getTime() + 86400000);
            const ev = cal.createEvent(x.name + (x.note ? ' – ' + x.note : ''), st, en);
            ev.setTag('eb', '1');
            n += 1;
          });
        });
        return json_({ ok: true, events: n });
      }

      case 'myday_events':
        // v3.5 My day: the Easy Buys account's own calendar → meetings for one day (read only).
        return json_(mydayEvents_(String(body.date || '')));

      case 'myday_out':
        // v3.5 My day: the app tasks of one day → the separate "Easy Buys – My day" calendar (made if missing).
        return json_(mydayOut_(String(body.date || ''), body.items || [], String(body.calendar || '')));

      case 'viber': {
        // v7: a post can carry photos (invoice, cheque stub) and a PDF (purchase order). See viberFiles_.
        // 1.0.0 build 1: owner: true = a business-wide money alert – it goes ONLY to the owner / admin channel when there is
        // one (never to a shop's channel, where the staff are); without an owner channel, to the main channel as before.
        const ch = body.owner ? (viberKey_('OWNER') ? 'OWNER' : '') : body.shopCode;
        // 1.0.0 build 10: the automatic posts carry a "once" key – whichever device makes them, each goes once.
        // review: the check and the claim are one step inside the write lock (two devices asking at the same moment used
        // to both see "not posted yet"); the post itself goes outside the lock; a failed post lets the key go again.
        // forget: the key of an alert that is sorted now (so it is posted again if it comes back).
        const once = body.once ? String(body.once).slice(0, 200) : '';
        const forget = body.forget ? String(body.forget).slice(0, 200) : '';
        if (once) {
          const mine = withLock_(function () { if (viberOnceSeen_(once) || viberOnceClaimed_(once)) return false; viberOnceClaim_(once, true); return true; }, 5000);
          if (!mine) return json_({ ok: true, skipped: true, note: 'Already posted (from another device).' });
        }
        let vr;
        try { vr = Array.isArray(body.files) && body.files.length ? viberFiles_(String(body.text || '').slice(0, 7000), body.files.slice(0, 4), ch) : viber_(String(body.text || '').slice(0, 7000), ch); }
        finally {
          if (once) { if (vr && vr.ok) viberOnceKeep_(once); viberOnceClaim_(once, false); }
        }
        if (forget && vr && vr.ok) { try { withLock_(function () { viberOnceForget_(forget); }, 5000); } catch (e) { /* kept 15 days at most */ } }
        return json_(vr);
      }

      case 'job_lease':
        return json_(jobLease_(body, who));

      default:
        return json_({ ok: false, error: 'Unknown request: ' + body.kind });
    }
  } catch (err) {
    // 1.0.0 build 10: the write lock was not free within a few seconds – answer at once (the app tries again in a moment)
    // instead of hanging until the app gives up ("Aborted").
    if (err && err.busy) return json_(busyReply_());
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally { flushLater_(); }
}
function busyReply_() { return { ok: false, busy: true, retryMs: 1500 + Math.floor(Math.random() * 2500), error: 'The cloud is busy – trying again.' }; }

// v3.5: payslip addresses are the ones saved in Payroll › Pay details (synced pay_profiles).
function payEmail_(addr) {
  const a = String(addr).trim().toLowerCase();
  const sh = data_().getSheetByName('T_pay_profiles');
  if (!sh) return false;
  return values_(sh).some(function (r) { try { return String(JSON.parse(r[4] || '{}').email || '').trim().toLowerCase() === a; } catch (e) { return false; } });
}

// Addresses saved in Admin › Emails for one kind of email (from the synced shared settings), when it is not Off.
function prefEmails_(type) {
  const sh = data_().getSheetByName('T_shared_settings');
  if (!sh) return [];
  let out = [];
  values_(sh).forEach(function (r) {
    if (String(r[0]) !== 'email_prefs') return;
    try {
      const row = JSON.parse(r[4] || '{}');
      const prefs = typeof row.value === 'string' ? JSON.parse(row.value) : row.value;
      const p = prefs && prefs[type];
      if (p && p.mode && p.mode !== 'off') out = cleanEmails_(p.to || []);
    } catch (e) { /* ignore */ }
  });
  return out;
}

// 1.0.0 build 1: emails only go to addresses the owner saved in Admin › Emails (any kind of email that is not Off) – a
// purchase order also to the suppliers' own addresses. Anything else is dropped.
function onlyAllowed_(list, suppliers) {
  const ok = {};
  ['shift_report', 'bakery_po', 'accountant_pack', 'po', 'reports', 'payslips'].concat(allPrefTypes_()).forEach(function (t) { prefEmails_(t).forEach(function (a) { ok[a.toLowerCase()] = true; }); });
  if (suppliers) {
    const sh = data_().getSheetByName('T_ap_suppliers');
    if (sh) values_(sh).forEach(function (r) { try { const e = String(JSON.parse(r[4] || '{}').email || '').trim().toLowerCase(); if (e) ok[e] = true; } catch (x) { /* none */ } });
  }
  return (list || []).filter(function (a) { return ok[String(a).toLowerCase()]; });
}
function allPrefTypes_() {
  const sh = data_().getSheetByName('T_shared_settings');
  if (!sh) return [];
  let out = [];
  values_(sh).forEach(function (r) {
    if (String(r[0]) !== 'email_prefs') return;
    try { const row = JSON.parse(r[4] || '{}'); const prefs = typeof row.value === 'string' ? JSON.parse(row.value) : row.value; out = Object.keys(prefs || {}); } catch (e) { /* none */ }
  });
  return out;
}
// A device's words in an email: shown as text (no links, no HTML).
function asText_(s) {
  return '<pre style="font-family:inherit;white-space:pre-wrap">' + String(s || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|tr|li|h\d)>/gi, '\n').replace(/<[^>]*>/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').slice(0, 20000) + '</pre>';
}

function staffEmail_(addr) {
  const a = String(addr).trim().toLowerCase();
  const sh = data_().getSheetByName('T_onboarding');
  if (!sh) return false;
  return values_(sh).some(function (r) { try { return String(JSON.parse(r[4] || '{}').email || '').trim().toLowerCase() === a; } catch (e) { return false; } });
}

// v9: which channel – '_S1' for a shop that has its own channel, '' for the main one.
function viberSuffix_(shopCode) {
  const c = String(shopCode || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 10);
  return c ? '_' + c : '';
}
function viberKey_(shopCode) {
  const sfx = viberSuffix_(shopCode);
  return sfx && PropertiesService.getScriptProperties().getProperty('VIBER_TOKEN' + sfx) ? sfx : '';
}
// Send a text to the Viber Channel (Channels Post API). The sender must be the channel's super admin.
// v9: a shop's message goes to that shop's channel when it has one, otherwise to the main channel.
function viber_(text, shopCode) {
  const props = PropertiesService.getScriptProperties();
  const sfx = viberKey_(shopCode);
  const token = props.getProperty('VIBER_TOKEN' + sfx);
  if (!token) return { ok: true, off: true };
  const call = function (path, payload) {
    const r = UrlFetchApp.fetch('https://chatapi.viber.com/pa/' + path, { method: 'post', contentType: 'application/json', headers: { 'X-Viber-Auth-Token': token },
      payload: JSON.stringify(payload || {}), muteHttpExceptions: true });
    try { return JSON.parse(r.getContentText()); } catch (e) { return { status: -1, status_message: r.getContentText().slice(0, 200) }; }
  };
  let from = props.getProperty('VIBER_FROM' + sfx);
  if (!from) {
    const info = call('get_account_info', {});
    if (info.status !== 0) return { ok: false, error: 'Viber: ' + (info.status_message || 'token not accepted') };
    const admin = (info.members || []).filter(function (m) { return m.role === 'superadmin'; })[0] || (info.members || [])[0];
    if (!admin) return { ok: false, error: 'Viber: the channel has no admin member.' };
    from = admin.id;
    props.setProperty('VIBER_FROM' + sfx, from);
  }
  let r = call('post', { from: from, type: 'text', text: text });
  if (r.status !== 0 && /webhook/i.test(String(r.status_message || ''))) {
    const h = viberWebhook_(call);
    if (!h.ok) return h;
    r = call('post', { from: from, type: 'text', text: text });
  }
  if (r.status !== 0) { props.deleteProperty('VIBER_FROM' + sfx); return { ok: false, error: 'Viber: ' + (r.status_message || 'not sent') }; }
  return { ok: true };
}
// A Viber Channel only takes posts once it has a webhook: this cloud's own web address (Viber checks it answers).
function viberWebhook_(call) {
  const props = PropertiesService.getScriptProperties();
  let url = props.getProperty('VIBER_HOOK_URL') || '';
  if (!url) { try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; } }
  url = url.replace(/\/dev$/, '/exec');
  if (!/^https:\/\//.test(url)) return { ok: false, error: 'Viber: the cloud does not know its own web address yet – deploy it as a web app (Deploy › Manage deployments), then try again.' };
  const w = call('set_webhook', { url: url, event_types: [], send_name: false, send_photo: false });
  if (w.status !== 0) return { ok: false, error: 'Viber: could not set the channel webhook – ' + (w.status_message || 'refused') };
  props.setProperty('VIBER_HOOK_URL', url);
  return { ok: true };
}

// v7: photos and PDFs in the Viber Channel. Viber can only take a file from a web link, so each file is saved in
// Drive › Viber posts, made viewable by link just long enough for Viber to copy it, then made private again
// (viberUnshare_ runs on every request and closes links older than 10 minutes). The text goes first.
function viberFiles_(text, files, shopCode) {
  const t = viber_(text, shopCode);
  if (!t.ok || t.off) return t;
  const props = PropertiesService.getScriptProperties();
  const sfx = viberKey_(shopCode);
  const token = props.getProperty('VIBER_TOKEN' + sfx);
  const from = props.getProperty('VIBER_FROM' + sfx);
  const folder = path_(['Viber posts', Utilities.formatDate(new Date(), 'Pacific/Fiji', 'yyyy-MM')]);
  let shared = [];
  try { shared = JSON.parse(props.getProperty('VIBER_SHARED') || '[]'); } catch (e) { shared = []; }
  const errors = [];
  files.forEach(function (f, i) {
    try {
      const mime = String(f.mime || 'image/jpeg');
      const name = String(f.name || ('file-' + i)).replace(/[^A-Za-z0-9 ._-]+/g, '-').slice(0, 80);
      const file = saveFile_(folder, name, String(f.data || ''), mime);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      shared.push({ id: file.getId(), at: Date.now() });
      const isPic = /^image\/(jpeg|png)$/.test(mime);
      const media = 'https://drive.google.com/uc?export=' + (isPic ? 'view' : 'download') + '&id=' + file.getId();
      const msg = isPic ? { type: 'picture', text: String(f.caption || '').slice(0, 700), media: media }
        : { type: 'file', media: media, size: file.getSize(), file_name: name };
      const r = UrlFetchApp.fetch('https://chatapi.viber.com/pa/post', { method: 'post', contentType: 'application/json', headers: { 'X-Viber-Auth-Token': token },
        payload: JSON.stringify(Object.assign({ from: from }, msg)), muteHttpExceptions: true });
      let j = {};
      try { j = JSON.parse(r.getContentText()); } catch (e) { j = { status: -1, status_message: r.getContentText().slice(0, 200) }; }
      if (j.status !== 0) errors.push(name + ': ' + (j.status_message || 'not sent'));
    } catch (e) { errors.push(String(e && e.message || e)); }
  });
  props.setProperty('VIBER_SHARED', JSON.stringify(shared.slice(-200)));
  return errors.length ? { ok: true, partial: true, error: 'Viber: ' + errors.join('; ') } : { ok: true };
}
function viberUnshare_() {
  const props = PropertiesService.getScriptProperties();
  let shared = [];
  try { shared = JSON.parse(props.getProperty('VIBER_SHARED') || '[]'); } catch (e) { return; }
  if (!shared.length) return;
  const keep = [];
  shared.forEach(function (x) {
    if (Date.now() - Number(x.at || 0) < 10 * 60000) { keep.push(x); return; }
    try { DriveApp.getFileById(x.id).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { /* gone already */ }
  });
  props.setProperty('VIBER_SHARED', JSON.stringify(keep));
  viberCleanup_();
}
// v8: the copies made for Viber (Drive › Viber posts) go to the Drive bin after 30 days (once a day, at most 150 files
// a time). The Viber Channel keeps its own copy; the originals stay with the invoice / cheque in the app and in Drive.
function viberCleanup_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Pacific/Fiji', 'yyyy-MM-dd');
  if (props.getProperty('VIBER_CLEANED') === today) return;
  props.setProperty('VIBER_CLEANED', today);
  try {
    const roots = root_().getFoldersByName('Viber posts');
    if (!roots.hasNext()) return;
    const cutoff = Date.now() - 30 * 86400000;
    let n = 0;
    const months = roots.next().getFolders();
    while (months.hasNext() && n < 150) {
      const files = months.next().getFiles();
      while (files.hasNext() && n < 150) {
        const f = files.next();
        if (f.getDateCreated().getTime() < cutoff) { try { f.setTrashed(true); n++; } catch (e) { /* next time */ } }
      }
    }
    if (n) log_('Viber copies binned', [n]);
  } catch (e) { /* never stops a request */ }
}

function ocr_(blob) {
  const file = Drive.Files.create({ name: 'eb-ocr-temp', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'en' });
  try {
    return DocumentApp.openById(file.id).getBody().getText();
  } catch (e) {
    return '';
  } finally {
    try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e) { /* ignore */ }
  }
}

// ---------------------------------------------------------------- online dashboard
function doGetRun_(e) {
  if (e && e.parameter && e.parameter.link) return linkPage_(String(e.parameter.link));
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('Easy Buys Business Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // 1.0.0 build 1: the dashboard cannot be shown inside another website (no framing).
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// The page an invite link opens on the new phone or tablet (email and WhatsApp do not open app links directly).
const APP_PACKAGE = 'com.easybuys.bakerymanager';
function linkPage_(code) {
  if (!/^EB1-[A-Za-z0-9_\-=]{10,600}$/.test(code)) return HtmlService.createHtmlOutput('<p style="font:18px Arial">This invite link is not complete. Ask for a new one.</p>');
  const intent = 'intent://link?c=' + encodeURIComponent(code) + '#Intent;scheme=easybuys;package=' + APP_PACKAGE + ';end';
  const html = '<div style="font-family:Arial,sans-serif;max-width:520px;margin:24px auto;padding:0 16px;color:#2b2320">'
    + '<h2>Link this device to Easy Buys</h2>'
    + '<p>1. Install the Easy Buys app first.</p>'
    + '<p>2. Tap: <a target="_top" href="' + intent + '" style="display:block;margin:12px 0;padding:16px;background:#1e7d3c;color:#fff;border-radius:12px;text-align:center;font-size:20px;font-weight:bold;text-decoration:none">Open in the Easy Buys app</a></p>'
    + '<p>3. If nothing happens: open the app, tap <b>Link to Easy Buys</b> and paste this code:</p>'
    + '<textarea id="c" readonly style="width:100%;height:110px;font-size:13px">' + code + '</textarea>'
    + '<button onclick="var t=document.getElementById(\'c\');t.select();try{navigator.clipboard.writeText(t.value)}catch(e){document.execCommand(\'copy\')}this.textContent=\'Copied\'" style="margin-top:8px;padding:14px;width:100%;font-size:18px;border-radius:12px;border:2px solid #ccc;background:#fff">Copy the code</button>'
    + '<p style="color:#777">The invite works once and expires 48 hours after it was made.</p></div>';
  return HtmlService.createHtmlOutput(html).setTitle('Link to Easy Buys').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function linkUrl_(url, linkCode) {
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'link=' + encodeURIComponent(linkCode);
}

// The dashboard password: the sample one is refused (anyone can read it in the public code), and after 10 wrong
// passwords the dashboard stays shut for an hour. Returns null when the password is right, or the answer to send back.
const DASH_MAX_FAILS = 10;
// 1.0.0 build 1: the wrong tries are counted per visitor (Google's anonymous key for the browser), so someone guessing
// cannot lock the owner out; the password is preferably kept in Script properties › DASHBOARD_PASSWORD.
function dashPass_() { return PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD') || DASHBOARD_PASSWORD; }
function dashCheck_(password) {
  const pass = dashPass_();
  if (!pass || pass === 'choose-a-dashboard-password' || String(pass).length < 6) {
    return { ok: false, error: 'The dashboard password has not been set in the cloud program yet (step 2 at the top).' };
  }
  let caller = '';
  try { caller = Session.getTemporaryActiveUserKey() || ''; } catch (e) { caller = ''; }
  const cache = CacheService.getScriptCache();
  // 1.0.0 build 1 (hardening, abuse C19/C20): a browser Google can tell apart is locked after 10 wrong passwords (only
  // that browser). Callers Google cannot tell apart (no user key) share one counter, so that counter never locks: after
  // 20 wrong passwords in an hour only one password is checked every 5 seconds – guessing is slowed to a crawl, and the
  // owner still gets in (at most a few seconds' wait).
  const key = 'dash_fails_' + sha_(caller || 'anyone').slice(0, 16);
  const fails = Number(cache.get(key) || '0');
  if (caller && fails >= DASH_MAX_FAILS) return { ok: false, error: 'Too many wrong passwords from this browser. Try again in an hour.' };
  if (!caller && fails >= 2 * DASH_MAX_FAILS) {
    const last = Number(cache.get('dash_slow_at') || '0');
    if (Date.now() - last < 5000) return { ok: false, slow: true, error: 'Too many wrong passwords – wait 5 seconds and try again.' };
    cache.put('dash_slow_at', String(Date.now()), 3600);
  }
  if (String(password || '') !== String(pass)) {
    cache.put(key, String(fails + 1), 3600);
    failCount_('dash', caller, { what: '5 wrong online dashboard passwords in 10 minutes from one browser' });
    return { ok: false, error: 'Wrong password' };
  }
  return null;
}

// Called by the dashboard page every 30 seconds.
function getDashboardRun_(password) {
  const bad = dashCheck_(password);
  if (bad) return bad;
  const props = PropertiesService.getScriptProperties();
  let snap = null;
  const fileId = props.getProperty('SNAPSHOT_FILE');
  if (fileId) snap = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8'));
  else if (props.getProperty('SNAPSHOT')) snap = JSON.parse(props.getProperty('SNAPSHOT'));
  const receivedAt = props.getProperty('SNAPSHOT_AT');
  const applied = snap && snap.planSeq ? Number(snap.planSeq) : 0;
  const pending = planEdits_().filter(function (x) { return x.seq > applied; });
  return { ok: true, snapshot: snap, receivedAt: receivedAt, files: recentFiles_(), pendingPlans: pending };
}

// ---------------------------------------------------------------- production plan from the dashboard
function planEdits_() {
  const raw = PropertiesService.getScriptProperties().getProperty('PLAN_EDITS');
  try { return raw ? JSON.parse(raw) : []; } catch (e) { return []; }
}

// Called by the dashboard when "Save plan" is pressed.
// items: { productId: quantity }  (0 or empty = not planned)
function savePlanOnlineRun_(password, date, items, byName) {
  const bad = dashCheck_(password);
  if (bad) return bad;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) return { ok: false, error: 'Invalid date' };
  const clean = {};
  Object.keys(items || {}).forEach(function (k) {
    const v = items[k];
    if (v === null || v === undefined || v === '') return;
    const n = Number(v);
    if (!isFinite(n) || n < 0 || n > 100000) throw new Error('Quantities must be between 0 and 100000.');
    if (/^\d+$/.test(k)) clean[k] = Math.round(n * 1000) / 1000;
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    let list = planEdits_();
    const seq = Number(props.getProperty('PLAN_SEQ') || '0') + 1;
    // Only the newest change for each day is kept waiting.
    list = list.filter(function (x) { return x.date !== date; });
    list.push({ seq: seq, date: date, items: clean, by: String(byName || 'Online dashboard').slice(0, 40), at: new Date().toISOString() });
    while (list.length > 30) list.shift();
    props.setProperty('PLAN_EDITS', JSON.stringify(list));
    props.setProperty('PLAN_SEQ', String(seq));
    log_('Plan changes', [date, JSON.stringify(clean), byName || '']);
    return { ok: true, seq: seq };
  } finally {
    lock.releaseLock();
  }
}

function recentFiles_() {
  const out = [];
  [['Shift reports', 'Report'], ['Purchase orders', 'Purchase order']].forEach(function (pair) {
    const top = root_().getFoldersByName(pair[0]);
    if (!top.hasNext()) return;
    const months = top.next().getFolders();
    while (months.hasNext()) {
      const files = months.next().getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (f.getMimeType() !== MimeType.PDF) continue;
        out.push({ type: pair[1], name: f.getName().replace(/\.pdf$/, ''), url: f.getUrl(), date: f.getDateCreated().toISOString() });
      }
    }
  });
  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  return out.slice(0, 40);
}

// Run this once from the editor to give the script permission to use Drive and Gmail.
function authorize() {
  root_();
  Logger.log('Cloud folder: ' + root_().getUrl());
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}


// ================================================================ EASY BUYS v3: shared records
// 1.0.0 build 10: per request memory (reset at the start of every request): the data spreadsheet is opened once, the
// people and the shared settings are read once, and logs / security events wait until the write lock is let go.
let REQ_ = {};
function data_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DATA_ID');
  if (id && REQ_.ss && REQ_.ssId === id) return REQ_.ss;
  if (id) { try { const ss0 = SpreadsheetApp.openById(id); REQ_.ss = ss0; REQ_.ssId = id; return ss0; } catch (e) { /* deleted: make a new one */ } }
  const ss = SpreadsheetApp.create(DATA_NAME);
  try { DriveApp.getFileById(ss.getId()).moveTo(root_()); } catch (e) { /* stays in My Drive */ }
  props.setProperty('DATA_ID', ss.getId());
  return ss;
}
function sheet_(name, header) {
  const ss = data_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    // Plain text everywhere: otherwise Sheets turns "2026-10-08 10:00:00" into a date and "007" into 7.
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), header.length)).setNumberFormat('@');
    sh.appendRow(header);
    sh.setFrozenRows(1);
    const first = ss.getSheetByName('Sheet1');
    if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  return sh;
}
// 1.0.0 build 10: the people and the shared settings are read once per request (every record of a push used to read
// the whole people sheet again); forgetValues_ after writing them.
const MEMO_SHEETS = { T_users: true, T_shared_settings: true };
function values_(sh) {
  const name = sh.getName ? String(sh.getName()) : '';
  const memo = MEMO_SHEETS[name] === true;
  if (memo && REQ_.vals && REQ_.vals[name]) return REQ_.vals[name];
  const last = sh.getLastRow();
  const out = last < 2 ? [] : sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().map(function (r) { return r.map(text_); });
  if (memo) (REQ_.vals = REQ_.vals || {})[name] = out;
  return out;
}
function forgetValues_(name) { if (!REQ_.vals) return; if (name) delete REQ_.vals[name]; else REQ_.vals = {}; }
// Safety net if a cell was ever typed or formatted by hand.
function text_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  return v === null || v === undefined ? '' : unesc_(String(v));
}
const TABLE_HEADER = ['key', 'updated_at', 'device', 'seq', 'json'];
const CHANGES_HEADER = ['seq', 'table', 'key', 'row', 'device', 'received_at'];
const DEVICE_HEADER = ['id', 'name', 'kind', 'status', 'token_hash', 'first_seen', 'last_seen', 'app', 'invited_by', 'note'];
// 1.0.0 build 10: uses (how many devices may join with it, 1–5), used_n (how many have), cancelled (by the owner)
const INVITE_HEADER = ['code', 'created_at', 'expires_at', 'created_by', 'email', 'used_at', 'used_by', 'uses', 'used_n', 'cancelled'];

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function now_() { return new Date().toISOString(); }
// 1.0.0 build 10: the write lock is waited for a few seconds only (tryLock); if another request still holds it the
// request is answered "busy" at once (see doPost) and the app tries again. Only writes take it: pulls, the devices list,
// ping and the till PCs' quick checks never wait for it.
const LOCK_WAIT_MS = 8000;
function withLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs || LOCK_WAIT_MS)) { const e = new Error('The cloud is busy – trying again.'); e.busy = true; throw e; }
  REQ_.locked = (REQ_.locked || 0) + 1;
  try { return fn(); } finally { lock.releaseLock(); REQ_.locked -= 1; }
}
// Work that does not need the lock (logs, security events): done when the lock is let go, or at the end of the request.
function later_(fn) { if (REQ_.locked) (REQ_.later = REQ_.later || []).push(fn); else fn(); }
// 1.0.0 build 10 (review): a run that is not a request (a trigger, or Run from the editor: compactChanges, the daily
// security summary, tidyRecords …, and the dashboard page's calls) starts with fresh request memory and writes the logs
// and security events it held back while it had the lock when it ends (they used to be lost).
function entry_(fn) { REQ_ = {}; try { return fn(); } finally { flushLater_(); } }
function doGet() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return doGetRun_.apply(null, a); }); }
function getDashboard() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return getDashboardRun_.apply(null, a); }); }
function savePlanOnline() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return savePlanOnlineRun_.apply(null, a); }); }
function setup() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return setupRun_.apply(null, a); }); }
function hashPins() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return hashPinsRun_.apply(null, a); }); }
function securityDigest() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return securityDigestRun_.apply(null, a); }); }
function checkHistory() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return checkHistoryRun_.apply(null, a); }); }
function compactChanges() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return compactChangesRun_.apply(null, a); }); }
function tidyRecords() { const a = Array.prototype.slice.call(arguments); return entry_(function () { return tidyRecordsRun_.apply(null, a); }); }
function flushLater_() {
  const list = REQ_.later || [];
  REQ_.later = [];
  list.forEach(function (fn) { try { fn(); } catch (e) { /* a log never stops a request */ } });
}

// Store records from a device (or the main software). The newest change wins.
// 1.0.0 build 1 (script 12): only known tables; times in the future refused; the change history only grows (with a
// running fingerprint); locked records and settings need an owner / admin session (lockedChange_); PINs are hashed and
// locked with the pepper (hashUserRow_); each accepted record keeps who the cloud checked (_vby).
// who: the device (and its person) – none for the main software.
function push_(records, deviceId, who) {
  return withLock_(function () {
    who = who || { deviceId: deviceId, main: true };
    // The main software's key (MAIN_KEY) is not a person either: it may not make owners, change roles / PINs or locked
    // settings (audit C3 – until every till PC has its own key, MAIN_KEY is on the till PCs).
    const boss = isBoss_(who);
    const bootstrap = !!who.viaKey && !anyBossProof_();      // a new business: the main tablet sets up the first owner
    const oldAdmin = !boss && transition_() && !!who.admin && !who.signed; // an old app on an admin device, for 7 days
    const props = props_();
    let seq = Number(props.getProperty('SEQ') || '0');
    let fresh = null;
    try { fresh = JSON.parse(props.getProperty('FRESH') || 'null'); } catch (e) { fresh = null; }
    const changes = sheet_('Changes', CHANGES_HEADER);
    const ss = data_();
    const tz = Session.getScriptTimeZone();
    const future = Utilities.formatDate(new Date(Date.now() + 26 * 3600000), tz, 'yyyy-MM-dd HH:mm:ss');
    // 1.0.0 build 1 (hardening): the closed-month lock, read once per push (money records dated in a closed month)
    let lockDay = null;
    try { const gl = sharedSetting_('gl_lock'); lockDay = gl && gl.date ? String(gl.date).slice(0, 10) : null; } catch (e) { lockDay = null; }
    const byTable = {};
    const rejected = [];
    records.forEach(function (r) {
      if (!r || !r.t || !/^[a-z_0-9]+$/.test(r.t) || r.k === undefined || r.k === null) return;
      if (KNOWN_TABLES.indexOf(r.t) < 0 && !ss.getSheetByName('T_' + r.t)) { rejected.push({ t: String(r.t), k: String(r.k), why: 'unknown table' }); return; }
      (byTable[r.t] = byTable[r.t] || []).push(r);
    });
    const accepted = [];
    const newChanges = [];
    const history = [];
    Object.keys(byTable).forEach(function (table) {
      const sh = sheet_('T_' + table, TABLE_HEADER);
      const rows = keysOf_(sh); // only key and time (the records themselves can be big: till figures)
      const index = {};
      rows.forEach(function (row, i) { index[String(row[0])] = { i: i + 2, updated: String(row[1]) }; });
      const appends = [];
      // 1.0.0 build 10: the rows already in the cloud that this push changes are read together (a few blocks), not one
      // read per record; rows that sit next to each other are written together too.
      const haveNos = [];
      byTable[table].forEach(function (r) { const h = index[String(r.k)]; if (h) haveNos.push(h.i); });
      const haveRows = rowsAt_(sh, haveNos);
      const updates = {};
      byTable[table].forEach(function (r) {
        const key = String(r.k);
        const updated = String(r.u || (r.row && r.row.updated_at) || '');
        const have = index[key];
        if (updated.replace('T', ' ').slice(0, 19) > future) { rejected.push({ t: table, k: key, why: 'time in the future' }); log_('Refused', [deviceId, table + ' ' + key, 'time in the future ' + updated]); return; }
        if (table === 'audit_log' && have) {
          // 1.0.0 build 10 (review): a push sent again after our time limit ran out (by design) carries the same history
          // lines: the very same line from the same device is taken quietly (nothing written, no security alert).
          if (sameHistoryRow_(haveRows[have.i], r.row, updated, deviceId)) { accepted.push({ t: table, k: key, again: true }); return; }
          rejected.push({ t: table, k: key, why: 'the change history cannot be changed' }); log_('Refused', [deviceId, 'audit_log ' + key, 'history change']); return;
        }
        if (have && have.updated > updated) { rejected.push({ t: table, k: key, why: 'newer change already in the cloud' }); return; }
        if (fresh && fresh.tables.indexOf(table) >= 0 && updated && updated < fresh.at) { rejected.push({ t: table, k: key, why: 'from before the fresh start' }); return; }
        let row = {};
        Object.keys(r.row || {}).forEach(function (c) { row[c] = r.row[c]; });
        delete row._vby;
        const haveRow = have ? (haveRows[have.i] || null) : null;
        // Locked records and settings: an owner / admin session (the device gets the cloud's copy back to show what is in force).
        const why = bootstrap ? null : (lockedChange_(table, key, table === 'users' ? hashUserRow_(row, key, false) : row, haveRow, boss, oldAdmin)
          || periodChange_(table, row, haveRow, boss, lockDay) || shopChange_(table, row, haveRow, boss, who));
        if (why) {
          let cur = null;
          if (haveRow) { try { cur = forDevice_(table, JSON.parse(haveRow[4] || 'null'), key, who); } catch (e) { cur = null; } }
          rejected.push({ t: table, k: key, why: why, code: 'LOGIN', restore: cur, gone: !have });
          log_('Refused', [deviceId, table + ' ' + key, why]);
          return;
        }
        // Security events for the owner: a locked setting / record changed (who, what, before → after), Link warnings.
        if (boss || bootstrap || oldAdmin) {
          try {
            if (lockedChange_(table, key, table === 'users' ? hashUserRow_(row, key, false) : row, haveRow, false, false)) {
              let bf = null; try { bf = haveRow ? JSON.parse(haveRow[4] || 'null') : null; } catch (e) { bf = null; }
              const show = function (o) { if (!o) return 'nothing'; const c = {}; Object.keys(o).forEach(function (x) { if (!/pin|password|proof|_vby|face|updated/.test(x)) c[x] = o[x]; }); return JSON.stringify(c).slice(0, 400); };
              secEvent_('locked_change', { device: who.deviceId, person: who.user ? who.user.name + ' (' + who.user.role + ')' : (bootstrap ? 'first set-up' : 'old app on an admin device'), what: table + ' ' + key + ': ' + show(bf) + ' → ' + show(row), refused: false });
            }
          } catch (e) { /* never stops the push */ }
        }
        if (table === 'aronium_feed' && row && String(row.kind || '') === 'link_event') secEvent_('link_event', { device: 'till PC ' + String(row.source || key.split('|')[0]), shop: String(row.shop_code || ''), what: String(row.data || row.text || 'Aronium Link event').slice(0, 500), refused: false });
        if (table === 'users') {
          // A plain PIN (old apps) never stays; a record without a PIN (the main software, a name change) keeps the PIN in force.
          row = hashUserRow_(row, key, boss || bootstrap || oldAdmin || !have);
          if (!row.pin_hash && haveRow) { try { const cur = JSON.parse(haveRow[4] || '{}'); if (cur.pin_hash) row.pin_hash = cur.pin_hash; } catch (e) { /* none */ } }
        }
        // 1.0.0 build 10: a device that is not an admin device never received the bank / account / TIN fields (forDevice_),
        // so its copy has them empty: a change it sends (a supplier's phone number, or the Aronium supplier round on a
        // phone holding the link job) keeps the cloud's values instead of wiping them.
        if (haveRow && !who.admin && !who.forMain && !who.main && MASKED_TABLES.indexOf(table) >= 0) {
          try {
            const cur = JSON.parse(haveRow[4] || '{}') || {};
            Object.keys(cur).forEach(function (c) { if (MASKED_FIELD.test(c) && c !== 'on_account' && (row[c] === null || row[c] === undefined || row[c] === '')) row[c] = cur[c]; });
          } catch (e) { /* the row as sent */ }
        }
        // Who the cloud vouches for: the person logged in (session); the first owner set up on the main tablet; an old app on
        // an admin device during the 7 days. Devices only apply locked records the cloud vouches for (services/locked.js).
        if (who.user) row._vby = { uid: who.user.uid, role: who.user.role };
        else if (bootstrap) row._vby = { uid: 'first-setup', role: 'owner' };
        else if (oldAdmin) row._vby = { uid: who.deviceId, role: 'admin', old: true };
        seq += 1;
        const json = JSON.stringify(row);
        const line = [esc_(key), esc_(updated), esc_(deviceId), seq, json];
        let rowNo;
        if (have) {
          rowNo = have.i;
          updates[rowNo] = line;
          haveRows[rowNo] = line.map(String);
          have.updated = updated;
        } else {
          rowNo = sh.getLastRow() + 1 + appends.length;
          appends.push(line);
          index[key] = { i: rowNo, updated: updated };
        }
        if (table === 'audit_log') history.push({ seq: seq, key: key, json: json });
        newChanges.push([seq, table, esc_(key), rowNo, esc_(deviceId), now_()]);
        accepted.push({ t: table, k: key });
      });
      writeRuns_(sh, updates);
      if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, TABLE_HEADER.length).setValues(appends);
      if (MEMO_SHEETS['T_' + table]) forgetValues_('T_' + table);
    });
    if (newChanges.length) changes.getRange(changes.getLastRow() + 1, 1, newChanges.length, CHANGES_HEADER.length).setValues(newChanges);
    historyChain_(history);
    props.setProperty('SEQ', String(seq));
    return { ok: true, accepted: accepted.length, rejected: rejected, seq: seq };
  });
}

// The history line already in the cloud is the very same as the one sent (same time, same device, same content).
function stableJson_(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v === undefined ? null : v);
  if (Array.isArray(v)) return '[' + v.map(stableJson_).join(',') + ']';
  return '{' + Object.keys(v).filter(function (k) { return k !== '_vby' && v[k] !== undefined; }).sort().map(function (k) { return JSON.stringify(k) + ':' + stableJson_(v[k]); }).join(',') + '}';
}
function sameHistoryRow_(haveRow, row, updated, deviceId) {
  if (!haveRow) return false;
  if (text_(haveRow[1]) !== String(updated) || text_(haveRow[2]) !== String(deviceId)) return false;
  let cur = null;
  try { cur = JSON.parse(String(haveRow[4] || 'null')); } catch (e) { return false; }
  return !!cur && stableJson_(cur) === stableJson_(row || {});
}
// Key and time of every row (columns A–B only).
function keysOf_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues().map(function (r) { return r.map(text_); });
}
// Some rows of a table. 1.0.0 build 10: rows close to each other are read as one block (a gap of up to RUN_GAP rows is
// read along), so 800 rows of a pull are a handful of reads instead of 800 (the first sync of a phone took 6 minutes);
// the whole table in one read when most of it is wanted.
const RUN_GAP = 20;
function rowsAt_(sh, rowNos) {
  const out = {};
  const last = sh.getLastRow();
  if (last < 2 || !rowNos.length) return out;
  if (rowNos.length * 4 >= last) {
    const all = values_(sh);
    rowNos.forEach(function (n) { if (all[n - 2]) out[n] = all[n - 2]; });
    return out;
  }
  runsOf_(rowNos.filter(function (n) { return n >= 2 && n <= last; }), RUN_GAP).forEach(function (run) {
    const vals = sh.getRange(run[0], 1, run[1] - run[0] + 1, TABLE_HEADER.length).getValues();
    for (let n = run[0]; n <= run[1]; n++) out[n] = vals[n - run[0]].map(text_);
  });
  // only the rows asked for
  const want = {}; rowNos.forEach(function (n) { want[n] = true; });
  Object.keys(out).forEach(function (n) { if (!want[n]) delete out[n]; });
  return out;
}
// [first, last] row runs of the sorted row numbers, joining rows at most `gap` apart.
function runsOf_(rowNos, gap) {
  const sorted = rowNos.slice().sort(function (a, b) { return a - b; });
  const runs = [];
  sorted.forEach(function (n) {
    const cur = runs[runs.length - 1];
    if (cur && n - cur[1] <= gap) cur[1] = Math.max(cur[1], n); else runs.push([n, n]);
  });
  return runs;
}
// Write changed rows: rows next to each other (no gap) in one write.
function writeRuns_(sh, updates) {
  const nos = Object.keys(updates).map(Number);
  if (!nos.length) return;
  runsOf_(nos, 1).forEach(function (run) {
    const lines = [];
    for (let n = run[0]; n <= run[1]; n++) lines.push(updates[n]);
    sh.getRange(run[0], 1, lines.length, TABLE_HEADER.length).setValues(lines);
  });
}
const PULL_MAX_CHARS = 3500000; // one answer stays small enough for a tablet on a slow connection

// The first row of the Changes sheet with a change number after `since` (the sheet is in change-number order, also after
// compactChanges). Before compacting, row n+1 holds change n, so older apps' positions mean the same.
function firstAfter_(sh, since, last) {
  let lo = 2; let hi = last + 1;
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    const v = Number(sh.getRange(mid, 1, 1, 1).getValues()[0][0]) || 0;
    if (v > since) hi = mid; else lo = mid + 1;
  }
  return lo;
}
// Changes after `since`, made by other devices (the main software gets everything). who: the device asking – what it may
// receive is trimmed by forDevice_; the change history only goes to admin devices.
// 1.0.0 build 10: nothing new (the device is at the newest change number) → answered without reading the sheets;
// limit: a smaller page asked by the app after a page took too long; left: change rows still to come after this page
// (the app shows "Receiving 800 of 2853…").
function pull_(since, deviceId, forMain, who, limit) {
  who = who || { deviceId: deviceId, forMain: !!forMain, admin: !!forMain };
  if (forMain) who.forMain = true;
  const seqNow = Number(props_().getProperty('SEQ') || '0');
  if (seqNow > 0 && since >= seqNow) return { ok: true, records: [], next: since, more: false, left: 0 };
  const changes = sheet_('Changes', CHANGES_HEADER);
  const last = changes.getLastRow();
  const start = last >= 2 ? firstAfter_(changes, since, last) : 2;
  if (start > last) return { ok: true, records: [], next: since, more: false, left: 0 };
  const count = Math.min(limit > 0 ? Math.max(50, Math.min(PULL_LIMIT, limit)) : PULL_LIMIT, last - start + 1);
  const list = changes.getRange(start, 1, count, CHANGES_HEADER.length).getValues().map(function (r) { return r.map(text_); });
  const want = [];
  const seen = {};
  list.forEach(function (c, i) {
    const table = String(c[1]);
    if (!forMain && String(c[4]) === String(deviceId)) return;
    if (!forMain && table === 'audit_log' && !who.admin) return;
    const id = table + '|' + c[2];
    if (seen[id]) return;
    seen[id] = true;
    want.push({ i: i, table: table, rowNo: Number(c[3]) });
  });
  const byTable = {};
  want.forEach(function (w) { (byTable[w.table] = byTable[w.table] || []).push(w.rowNo); });
  const rows = {};
  Object.keys(byTable).forEach(function (t) { const sh = data_().getSheetByName('T_' + t); rows[t] = sh ? rowsAt_(sh, byTable[t]) : {}; });
  const out = [];
  let size = 0;
  let used = count;
  for (let k = 0; k < want.length; k++) {
    const w = want[k];
    const row = rows[w.table][w.rowNo];
    if (!row || String(row[0]) !== String(list[w.i][2])) continue; // compacted / moved: the newest copy comes with its own change
    if (size > PULL_MAX_CHARS && out.length) { used = w.i; break; } // the rest comes in the next page
    size += String(row[4] || '').length;
    let parsed = {};
    try { parsed = JSON.parse(row[4]); } catch (e) { parsed = {}; }
    const vby = parsed && parsed._vby ? parsed._vby : null;
    if (parsed) delete parsed._vby;
    const rec = { t: w.table, k: String(row[0]), u: String(row[1]), d: String(row[2]), row: forDevice_(w.table, parsed, String(row[0]), who) };
    if (vby) rec.vby = vby;
    out.push(rec);
  }
  const next = used > 0 ? Number(list[used - 1][0]) || since : since;
  return { ok: true, records: out, next: next, more: start + used <= last, left: Math.max(0, last - (start + used) + 1), page: used };
}

// ================================================================ devices & invites
// 1.0.0 build 10: the Devices sheet is read on every request (who is asking): its rows are kept in the script cache,
// stamped with a version number in Script properties (DEVROWS_V) that every write to the sheet moves on – a copy made
// before a write is never used after it. A request reads the sheet at most once.
function deviceRows_() {
  if (REQ_.devRows) return REQ_.devRows;
  const p = props_();
  const v = p.getProperty('DEVROWS_V') || '0';
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  if (cache) {
    try { const c = JSON.parse(cache.get('DEVROWS') || 'null'); if (c && c.v === v && Array.isArray(c.rows)) { REQ_.devRows = c.rows; REQ_.devCached = true; return c.rows; } } catch (e) { /* read the sheet */ }
  }
  const rows = values_(sheet_('Devices', DEVICE_HEADER));
  REQ_.devRows = rows;
  if (cache) { try { const txt = JSON.stringify({ v: v, rows: rows }); if (txt.length < 90000) cache.put('DEVROWS', txt, 600); } catch (e) { /* too big: read each time */ } }
  return rows;
}
function devicesChanged_() {
  const p = props_();
  p.setProperty('DEVROWS_V', String((Number(p.getProperty('DEVROWS_V') || '0') || 0) + 1));
  REQ_.devRows = null;
}
function devices_() {
  const all = ctlAll_();
  return deviceRows_().map(function (r) {
    const c = all[String(r[0])] || {};
    return { id: r[0], name: r[1], kind: r[2], status: r[3], first_seen: r[5], last_seen: r[6], app: r[7], invited_by: r[8],
      until: c.until || null, endDate: c.endDate || null, shop: c.shop || null, roles: c.roles || [], stat: c.stat || null, hist: c.hist || [],
      attempts: c.attempts || [], waitingCmds: (c.cmds || []).length, wipe: c.wipe || null, lastLog: c.lastLog || null,
      admin: isAdminDevice_(String(r[0]), /cloud key/.test(String(r[9] || ''))), adminMarked: !!c.admin, signs: c.signs || null };
  });
}
// found.sh: the sheet, opened only when something is written (the caller then calls devicesChanged_()).
function findDevice_(match) {
  const rows = deviceRows_();
  for (let i = 0; i < rows.length; i++) if (match(rows[i])) return { get sh() { return sheet_('Devices', DEVICE_HEADER); }, row: i + 2, v: rows[i] };
  return null;
}
// 1.0.0 build 10: last seen is written at most every 2 minutes per device (and when its name, kind or app changes) –
// it used to be three writes on every request from every device.
const TOUCH_EVERY_MS = 2 * 60000;
function touchDevice_(body, deviceId, extra) {
  const d = body.device || {};
  const found = findDevice_(function (r) { return String(r[0]) === String(deviceId); });
  const now = now_();
  if (found) {
    const name = d.name || found.v[1]; const kind = d.kind || found.v[2]; const app = body.app || '';
    const seen = Date.parse(String(found.v[6] || '').replace(' ', 'T'));
    const fresh = isFinite(seen) && Date.now() - seen < TOUCH_EVERY_MS && Date.now() - seen >= 0;
    if (fresh && String(name) === String(found.v[1]) && String(kind) === String(found.v[2]) && String(app) === String(found.v[7])) return found;
    found.sh.getRange(found.row, 2, 1, 2).setValues([[name, kind]]);
    found.sh.getRange(found.row, 7, 1, 2).setValues([[now, app]]);
    devicesChanged_();
    return found;
  }
  sheet_('Devices', DEVICE_HEADER).appendRow([deviceId, d.name || 'Device', d.kind || '', (extra && extra.status) || 'active',
    (extra && extra.tokenHash) || '', now, now, body.app || '', (extra && extra.invitedBy) || '', (extra && extra.note) || '']);
  devicesChanged_();
  return null;
}

// Every request (except join and the main software) must come from a linked, unblocked device.
// 1.0.0 build 1 (script 12): a device proves who it is by SIGNING the request with its device pass (sig, see sign_);
// the pass itself no longer travels once the device knows the cloud is version 12. Old apps (no sig) still work until
// they update – unless the owner switched on "every device must sign" (STRICT_SIGN) – and a device that signed once may
// never send an unsigned request again. Unknown devices, removed devices and wrong passes are refused. A new device
// holding the cloud key waits for the owner when another device already uses the key.
function auth_(body) {
  const deviceId = String((body.device && body.device.id) || '');
  const strict = PropertiesService.getScriptProperties().getProperty('STRICT_SIGN') === '1';
  if (body.token || (body.sig && !body.key)) {
    let d = null;
    if (body.token) {
      const h = sha_(body.token);
      d = findDevice_(function (r) { return String(r[4]) === h; });
      if (!d) { noteAttemptByToken_(h); return { ok: false, removed: true, error: 'This device is not linked any more. Link it again from another device (Admin › Devices).' }; }
    } else {
      d = deviceId ? findDevice_(function (r) { return String(r[0]) === deviceId; }) : null;
      if (!d) { log_('Refused', [deviceId, 'unknown device']); return { ok: false, removed: true, error: 'This device is not known to the cloud. Link it again from another device (Admin › Devices).' }; }
      if (d.v[3] === 'removed' || !String(d.v[4] || '')) { noteAttempt_(String(d.v[0]), 'removed'); return { ok: false, removed: true, error: 'This device was removed by the owner. Link it again.' }; }
    }
    const id = String(d.v[0]);
    if (d.v[3] === 'blocked') { noteAttempt_(id, 'blocked'); return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' }; }
    if (d.v[3] === 'waiting') return { ok: false, waiting: true, error: 'Waiting for the owner to approve this device.' };
    const s = checkSig_(body, id, String(d.v[4]), strict);
    if (s) { noteAttempt_(id, s.why); return s; }
    const g = gate_(id, d);
    if (g) return g;
    touchDevice_(body, id);
    return { ok: true, deviceId: id, viaKey: false, admin: isAdminDevice_(id, false), signed: !!body.sig };
  }
  const KEY = apiKey_();
  if (!KEY || KEY.indexOf('PASTE') === 0) return { ok: false, error: 'The cloud key has not been set in the cloud program yet.' };
  if (body.key !== KEY) return { ok: false, error: 'Wrong cloud key. Copy the key from the app into the cloud program.' };
  const id = deviceId || 'first-device';
  const d = findDevice_(function (r) { return String(r[0]) === id; });
  if (d && d.v[3] === 'blocked') { noteAttempt_(id, 'blocked'); return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' }; }
  if (d && d.v[3] === 'removed') { noteAttempt_(id, 'removed'); return { ok: false, removed: true, error: 'This device was removed by the owner. Link it again.' }; }
  if (d && d.v[3] === 'waiting') return { ok: false, waiting: true, error: 'Waiting for the owner to approve this device (it used the cloud key).' };
  if (d) { const g = gate_(id, d); if (g) return g; }
  const s = checkSig_(body, id, sha_(KEY), strict);
  if (s) { noteAttempt_(id, s.why); return s; }
  if (!d) {
    // The cloud key is not a pass for any number of new devices: the first device holding it is the main tablet;
    // another new device with the key waits until the owner approves it (Admin › Devices).
    const others = deviceRows_().filter(function (r) { return /cloud key/.test(String(r[9] || '')) && String(r[3]) === 'active'; });
    if (others.length) {
      touchDevice_(body, id, { status: 'waiting', note: 'asked with the cloud key' });
      log_('Refused', [id, 'new device with the cloud key – waiting for the owner']);
      return { ok: false, waiting: true, error: 'A new device used the cloud key. The owner must approve it first (Admin › Devices on the main tablet).' };
    }
  }
  touchDevice_(body, id, { status: 'active', note: 'linked with the cloud key' });
  return { ok: true, deviceId: id, viaKey: true, admin: true, signed: !!body.sig };
}

// sig = sha256( sha256(pass) | device id | kind | sentAt | sha256(JSON of the records, or "null") ) – pass = the
// device's own pass, or the cloud key on the main tablet. The cloud only keeps sha256(pass), all it needs to check.
function sign_(passHash, deviceId, body) {
  return sha_([passHash, deviceId, String(body.kind || ''), String(body.sentAt || ''), sha_(JSON.stringify(body.records || null))].join('|'));
}
function checkSig_(body, id, passHash, strict) {
  const c = ctlAll_()[id] || {};
  if (!body.sig) {
    if (strict || !transition_()) return { ok: false, why: 'unsigned', error: 'Update the Easy Buys app on this device – the cloud now only accepts signed requests.' };
    if (c.signs) return { ok: false, why: 'unsigned', error: 'This device signs its requests, but this one was not signed. Refused (it may not have come from the device).' };
    return null;
  }
  if (!passHash || String(body.sig) !== sign_(passHash, id, body)) return { ok: false, why: 'bad signature', error: 'The request was not signed with this device\'s pass. Refused. Link the device again if this keeps happening.' };
  const t = Date.parse(String(body.sentAt || '').replace(' ', 'T'));
  if (!isFinite(t) || Math.abs(Date.now() - t) > 7 * 86400000) return { ok: false, why: 'old request', error: 'The request is too old, or the date on this device is wrong. Check the date and time on this device.' };
  if (!c.signs) { const all = ctlAll_(); ctl_(all, id).signs = now_(); ctlSaveSafe_(all); }
  return null;
}

// An "admin device": the main tablet (cloud key), a device the owner marked as an admin device in Admin › Devices, or a
// device only owners / admins may log in on. Role and PIN changes, locked settings and device changes come only from these.
function isAdminDevice_(id, viaKey) {
  if (viaKey) return true;
  const c = ctlAll_()[String(id)] || {};
  if (c.admin) return true;
  const roles = c.roles || [];
  return roles.length > 0 && roles.every(function (r) { return r === 'owner' || r === 'admin'; });
}

function code_(n) {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < n; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}
function linkCode_(url, code) {
  return 'EB1-' + Utilities.base64EncodeWebSafe(JSON.stringify({ u: url, i: code }));
}

// A single-use invite for a new device. The cloud key is never inside it.
function inviteCreate_(body, who) {
  const code = code_(8);
  const created = new Date();
  // 1.0.0 build 10 ("Add a phone" on Home): one QR code for several staff phones in one go – up to 5 devices, for 1 hour
  const uses = Math.max(1, Math.min(5, Math.floor(Number(body.uses) || 1)));
  const hours = Math.max(0.25, Math.min(INVITE_HOURS, Number(body.hours) || INVITE_HOURS));
  const expires = new Date(created.getTime() + hours * 3600000);
  const email = String(body.email || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'That email address does not look right.' };
  const url = String(body.url || '');
  sheet_('Invites', INVITE_HEADER).appendRow([code, created.toISOString(), expires.toISOString(), esc_((body.by || '') + ' @ ' + who.deviceId), email, '', '', uses, 0, '']);
  // 1.0.0 build 1: an invite for the owner's / an admin's own phone makes it an admin device when it joins.
  if (body.adminDevice) { const p = props_(); let l = []; try { l = JSON.parse(p.getProperty('INVITE_ADMIN') || '[]'); } catch (e) { l = []; } p.setProperty('INVITE_ADMIN', JSON.stringify(l.concat([code]).slice(-20))); }
  const link = linkCode_(url, code);
  let emailed = false;
  if (email) {
    const tap = linkUrl_(url, link);
    mail_({
      to: email, name: 'Easy Buys',
      subject: 'Link your device to Easy Buys',
      htmlBody: '<p>Bula! ' + (body.by || 'The owner') + ' invited this device to the Easy Buys app.</p>'
        + '<ol><li>Install the Easy Buys app on the tablet or phone.</li>'
        + '<li>On that device, open this email and tap: <a href="' + tap + '"><b>Link this device to Easy Buys</b></a> and then <b>Open in the Easy Buys app</b>.</li>'
        + '<li>If the link does not open the app: open the app, tap <b>Link to Easy Buys</b>, and paste this code:<br><code style="font-size:13px">' + link + '</code></li></ol>'
        + '<p>This invite works once and expires in ' + INVITE_HOURS + ' hours.</p>',
    });
    emailed = true;
  }
  return { ok: true, code: code, linkCode: link, linkUrl: url ? linkUrl_(url, link) : '', expires: expires.toISOString(), emailed: emailed, uses: uses };
}
// The owner cancels an invite (the QR code stops working at once).
function inviteCancel_(body, who) {
  const code = String(body.code || '').trim().toUpperCase();
  return withLock_(function () {
    const sh = sheet_('Invites', INVITE_HEADER);
    const rows = values_(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) !== code) continue;
      sh.getRange(i + 2, 10, 1, 1).setValues([[now_() + ' by ' + esc_(String(body.by || '') + ' @ ' + who.deviceId)]]);
      return { ok: true, cancelled: code };
    }
    return { ok: false, error: 'Invite not found.' };
  }, 5000);
}

// New device: swap the invite for its own access token.
// 1.0.0 build 1: wrong codes are limited (30 an hour for everyone, then joining waits an hour); a device id already
// linked cannot be taken over; new devices wait for the owner's approval unless the owner switched approval off.
function join_(body) {
  const code = String(body.invite || '').trim().toUpperCase();
  if (!/^[A-Z0-9]{6,12}$/.test(code)) return { ok: false, error: 'Invite not found. Ask for a new one.' };
  const hour = new Date().toISOString().slice(0, 13);
  const cache = typeof CacheService !== 'undefined' ? CacheService.getScriptCache() : null;
  const bad = cache ? Number(cache.get('join_bad_' + hour) || '0') : 0;
  if (bad >= 30) return { ok: false, error: 'Too many wrong invite codes – try again in an hour.' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return busyReply_(); // 1.0.0 build 10: the app tries again by itself
  REQ_.locked = (REQ_.locked || 0) + 1; // logs and security events wait until the lock is let go
  try {
    const sh = sheet_('Invites', INVITE_HEADER);
    const rows = values_(sh);
    let row = -1;
    for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === code) row = i;
    if (row < 0) {
      if (cache) cache.put('join_bad_' + hour, String(bad + 1), 3700);
      if (bad + 1 === 10) secEvent_('invite_guess', { device: String((body.device && body.device.id) || ''), what: '10 wrong invite codes within the hour' });
      return { ok: false, error: 'Invite not found. Ask for a new one.' };
    }
    const inv = rows[row];
    const d = body.device || {};
    if (!d.id || !/^[A-Za-z0-9._:-]{3,80}$/.test(String(d.id))) return { ok: false, error: 'Device id missing.' };
    const token = Utilities.getUuid() + Utilities.getUuid();
    const existing = findDevice_(function (r) { return String(r[0]) === String(d.id); });
    // 1.0.0 build 10: the phone asked again after its first try ran out of time (the join had gone through, the answer was
    // lost): the same join (its nonce, with this invite) within 10 minutes gets a fresh pass instead of "already linked"
    // or "already used" – checked before the invite's uses, since that try already counted.
    const nonceH = body.nonce ? sha_(String(body.nonce)) : '';
    if (existing && nonceH && String(existing.v[3]) !== 'removed') {
      const cj = ctlAll_()[String(d.id)] || {};
      if (cj.joinNonce === nonceH && cj.joinCode === code && Date.now() - Number(cj.joinAt || 0) < 10 * 60000) {
        existing.sh.getRange(existing.row, 5, 1, 1).setValues([[sha_(token)]]);
        devicesChanged_();
        return { ok: true, token: token, approved: String(existing.v[3]) === 'active', admin: !!cj.admin, again: true };
      }
    }
    const uses = Math.max(1, Number(inv[7]) || 1);
    const usedN = Number(inv[8]) || (inv[5] ? 1 : 0);
    if (inv[9]) return { ok: false, error: 'The owner cancelled this invite. Ask for a new QR code.' };
    if (usedN >= uses) return { ok: false, error: uses > 1 ? 'This QR code has been used by ' + uses + ' phones already. Ask for a new one.' : 'This invite has already been used. Ask for a new one.' };
    if (new Date(inv[2]).getTime() < Date.now()) return { ok: false, error: 'This invite has expired. Ask for a new one.' };
    const approval = PropertiesService.getScriptProperties().getProperty('REQUIRE_APPROVAL') !== '0';
    if (existing && String(existing.v[3]) !== 'removed' && String(existing.v[4] || '')) { log_('Refused', [String(d.id), 'join with a device id already linked']); return { ok: false, error: 'A device with this id is already linked. Remove the old one first (Admin › Devices).' }; }
    if (existing) { existing.sh.getRange(existing.row, 4, 1, 2).setValues([[approval ? 'waiting' : 'active', sha_(token)]]); devicesChanged_(); }
    else touchDevice_(body, String(d.id), { status: approval ? 'waiting' : 'active', tokenHash: sha_(token), invitedBy: inv[3], note: 'invite ' + code });
    // 1.0.0 build 10: an invite for several phones counts them (used_by: the devices, newest last)
    const by = usedN ? String(inv[6] || '') + ',' + String(d.id) : String(d.id);
    sh.getRange(row + 2, 6, 1, 4).setValues([[inv[5] || now_(), esc_(by.slice(-400)), uses, usedN + 1]]);
    let adminList = [];
    try { adminList = JSON.parse(props_().getProperty('INVITE_ADMIN') || '[]'); } catch (e) { adminList = []; }
    // one save for the device's control note (never stops the join: the device is linked already)
    const all0 = ctlAll_();
    if (adminList.indexOf(code) >= 0 || nonceH || (all0[String(d.id)] && all0[String(d.id)].removedAt)) {
      const all = all0; const c = ctl_(all, String(d.id));
      delete c.removedAt;
      if (adminList.indexOf(code) >= 0) { c.admin = true; hist_(c, 'Joined as an admin device (owner\'s invite)', inv[3]); }
      if (nonceH) { c.joinNonce = nonceH; c.joinCode = code; c.joinAt = Date.now(); }
      ctlSaveSafe_(all);
    }
    secEvent_('device', { device: String(d.name || d.id), what: 'A new device joined with an invite' + (approval ? ' – it waits for your approval' : '') + (adminList.indexOf(code) >= 0 ? ' (as an admin device)' : ''), refused: false });
    return { ok: true, token: token, approved: !approval, admin: adminList.indexOf(code) >= 0 };
  } finally { lock.releaseLock(); REQ_.locked -= 1; }
}

function deviceStatus_(id, status, who) {
  if (['active', 'blocked'].indexOf(status) < 0) return { ok: false, error: 'Choose active or blocked.' };
  if (String(id) === String(who.deviceId) && status === 'blocked') return { ok: false, error: 'A device cannot block itself.' };
  const d = findDevice_(function (r) { return String(r[0]) === String(id); });
  if (!d) return { ok: false, error: 'Device not found.' };
  d.sh.getRange(d.row, 4).setValue(status);
  devicesChanged_();
  return { ok: true, devices: devices_() };
}

// ---------------------------------------------------------------- v3.6 device control (Admin › Devices)
// Kept in a script property: { [deviceId]: { until, endDate, shop, roles, name, cmds: [], hist: [], attempts: [], stat, wipe, removedTokens } }.
function ctlAll_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DEVCTL') || '{}'); } catch (e) { return {}; } }
// 1.0.0 build 10 (review): DEVCTL is ONE Script property, and a property value over 9 KB cannot be saved (setProperty
// throws – every pull and join then failed). Before each save it is kept small: devices removed more than 30 days ago
// are forgotten, 8 history lines and 5 attempts per device, the join fields (the lost-answer check) only for 10 minutes;
// if it is still too big, removed devices go first, then the older history.
const CTL_MAX_BYTES = 8500;
const CTL_HIST = 8;
const CTL_ATTEMPTS = 5;
const bytes_ = function (s) { try { return unescape(encodeURIComponent(s)).length; } catch (e) { return String(s).length * 3; } };
function ctlPrune_(all) {
  const now = Date.now();
  Object.keys(all).forEach(function (id) {
    const c = all[id];
    if (!c || typeof c !== 'object') { delete all[id]; return; }
    const gone = c.removedAt ? Date.parse(String(c.removedAt)) : NaN;
    if (Number.isFinite(gone) && now - gone > 30 * 86400000) { delete all[id]; return; }
    if (c.joinAt !== undefined && !(now - Number(c.joinAt || 0) < 10 * 60000)) { delete c.joinNonce; delete c.joinCode; delete c.joinAt; }
    c.hist = (c.hist || []).slice(0, CTL_HIST).map(function (h) { return { at: h.at, text: String(h.text || '').slice(0, 140), by: String(h.by || '').slice(0, 80) }; });
    c.attempts = (c.attempts || []).slice(0, CTL_ATTEMPTS);
    c.cmds = (c.cmds || []).slice(-10);
    if (c.removedTokens) c.removedTokens = c.removedTokens.slice(-3);
  });
  let s = JSON.stringify(all);
  if (bytes_(s) <= CTL_MAX_BYTES) return s;
  // still too big: removed devices go (oldest first), then shorter history and attempts everywhere
  const removed = Object.keys(all).filter(function (id) { return all[id].removedAt; }).sort(function (a, b) { return String(all[a].removedAt).localeCompare(String(all[b].removedAt)); });
  for (let i = 0; i < removed.length && bytes_(s) > CTL_MAX_BYTES; i++) { delete all[removed[i]]; s = JSON.stringify(all); }
  [3, 1, 0].forEach(function (n) {
    if (bytes_(s) <= CTL_MAX_BYTES) return;
    Object.keys(all).forEach(function (id) { all[id].hist = (all[id].hist || []).slice(0, n); all[id].attempts = (all[id].attempts || []).slice(0, Math.min(n, 2)); });
    s = JSON.stringify(all);
  });
  return s;
}
function ctlSave_(all) { PropertiesService.getScriptProperties().setProperty('DEVCTL', ctlPrune_(all)); }
// A save that never stops the request (a pull, a join, a refused request noting an attempt).
function ctlSaveSafe_(all) { try { ctlSave_(all); return true; } catch (e) { log_('Problems', ['DEVCTL not saved', String(e && e.message || e).slice(0, 200)]); return false; } }
function ctl_(all, id) { all[id] = all[id] || { cmds: [], hist: [], attempts: [] }; return all[id]; }
function hist_(c, text, by) { c.hist = [{ at: now_(), text: String(text || '').slice(0, 140), by: String(by || '').slice(0, 80) }].concat(c.hist || []).slice(0, CTL_HIST); }
function noteAttempt_(id, why) {
  const all = ctlAll_(); const c = ctl_(all, id);
  c.attempts = [{ at: now_(), why: why }].concat(c.attempts || []).slice(0, CTL_ATTEMPTS);
  ctlSaveSafe_(all);
}
function noteAttemptByToken_(h) {
  const all = ctlAll_();
  Object.keys(all).forEach(function (id) { if ((all[id].removedTokens || []).indexOf(h) >= 0) { const c = all[id]; c.attempts = [{ at: now_(), why: 'removed' }].concat(c.attempts || []).slice(0, CTL_ATTEMPTS); } });
  ctlSaveSafe_(all);
}
// Refuse a suspended, ended or wiped device (enforced here, whatever the device does).
function gate_(id, d) {
  const all = ctlAll_();
  const c = all[id];
  if (!c) return null;
  if (c.wipe === 'pending') {
    c.wipe = 'done'; c.removedAt = now_(); hist_(c, 'Wiped (the device cleared its data)', 'cloud');
    if (d && d.sh) { c.removedTokens = (c.removedTokens || []).concat([String(d.v[4])]).slice(-5); d.sh.getRange(d.row, 4, 1, 2).setValues([['removed', '']]); devicesChanged_(); }
    ctlSaveSafe_(all);
    return { ok: false, wipe: true, removed: true, error: 'The owner wiped this device.' };
  }
  const nowIso = now_();
  if (c.until) {
    if (c.until > nowIso) { c.attempts = [{ at: nowIso, why: 'suspended' }].concat(c.attempts || []).slice(0, CTL_ATTEMPTS); ctlSaveSafe_(all); return { ok: false, lock: { kind: 'suspended', until: c.until }, error: 'Access paused by the owner until ' + c.until.slice(0, 16).replace('T', ' ') + '.' }; }
    delete c.until; hist_(c, 'Pause ended', 'cloud'); ctlSaveSafe_(all);
  }
  if (c.endDate && c.endDate < nowIso.slice(0, 10)) {
    c.attempts = [{ at: nowIso, why: 'ended' }].concat(c.attempts || []).slice(0, CTL_ATTEMPTS); ctlSaveSafe_(all);
    return { ok: false, lock: { kind: 'ended', until: c.endDate }, error: 'Access for this device ended on ' + c.endDate + '.' };
  }
  return null;
}
// What the device reports on each pull (waiting items, last health check / backup / login), and what it is told back.
// 1.0.0 build 10: written only when there is something new (commands handed over, or the device's figures changed /
// are more than 5 minutes old) – every pull from every device used to rewrite the whole device-control property.
function devicePulled_(id, body, out) {
  const all = ctlAll_(); const c = ctl_(all, id);
  let dirty = (c.cmds || []).length > 0;
  if (body.stat) {
    const st = { at: now_(), pending: body.stat.pending, health: body.stat.health || null, backup: body.stat.backup || null, login: body.stat.login || null };
    const was = c.stat || {};
    const same = JSON.stringify([was.pending, was.health, was.backup, was.login]) === JSON.stringify([st.pending, st.health, st.backup, st.login]);
    const age = Date.now() - Date.parse(String(was.at || ''));
    if (!same || !(age < 5 * 60000)) { c.stat = st; dirty = true; }
  }
  // 1.0.0 build 1: a device clock an hour or more out, or a time zone that is not Fiji's, may be tampering.
  // 1.0.0 hardening: more than 5 minutes out (the app stops money work then), or a time zone that is not Fiji
  if (body.stat && ((Math.abs(Number(body.stat.clockSkewMin) || 0) > 5) || (body.stat.tz && !/Fiji|^\+1[23]|^UTC\+12/.test(String(body.stat.tz)))) && !onceIn_('clock_' + id, 3600)) {
    secEvent_('clock', { device: String(c.name || id), what: 'Clock ' + (Number(body.stat.clockSkewMin) || 0) + ' minutes from Google\'s time; time zone ' + String(body.stat.tz || 'not said'), refused: false });
  }
  out.control = { shop: c.shop || null, roles: c.roles || [], name: c.name || null, endDate: c.endDate || null };
  out.commands = c.cmds || [];
  c.cmds = [];
  // 1.0.0 build 10 (review): a pull never fails because the device-control note could not be saved
  if (dirty) ctlSaveSafe_(all);
}
// True when the same thing was already noted within `secs` (script cache); notes it otherwise.
function onceIn_(key, secs) {
  try { const cache = CacheService.getScriptCache(); if (cache.get('once_' + key)) return true; cache.put('once_' + key, '1', secs); } catch (e) { /* no cache: not limited */ }
  return false;
}
const DEVICE_ACTIONS = ['suspend', 'block', 'allow', 'remove', 'rename', 'shop', 'roles', 'end', 'sync', 'log', 'signout', 'wipe', 'admin'];
function deviceControl_(body, who) {
  return withLock_(function () {
    const id = String(body.id || '');
    const a = String(body.action || '');
    if (DEVICE_ACTIONS.indexOf(a) < 0) return { ok: false, error: 'Unknown action.' };
    const d = findDevice_(function (r) { return String(r[0]) === id; });
    if (!d) return { ok: false, error: 'Device not found.' };
    if (id === String(who.deviceId) && ['suspend', 'block', 'remove', 'wipe', 'end'].indexOf(a) >= 0) return { ok: false, error: 'A device cannot do that to itself. Use another device.' };
    const all = ctlAll_(); const c = ctl_(all, id);
    const by = String(body.by || '') + ' @ ' + who.deviceId;
    if (a === 'suspend') {
      const until = String(body.until || '');
      if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(until)) return { ok: false, error: 'Until when? Choose a date and time.' };
      c.until = until; hist_(c, 'Paused until ' + until.slice(0, 16).replace('T', ' '), by);
    } else if (a === 'block') { d.sh.getRange(d.row, 4).setValue('blocked'); hist_(c, 'Blocked', by); }
    else if (a === 'allow') { d.sh.getRange(d.row, 4).setValue('active'); delete c.until; delete c.removedAt; if (c.endDate && c.endDate < now_().slice(0, 10)) delete c.endDate; hist_(c, 'Allowed again', by); }
    else if (a === 'remove') { c.removedAt = now_(); c.removedTokens = (c.removedTokens || []).concat([String(d.v[4])]).slice(-5); d.sh.getRange(d.row, 4, 1, 2).setValues([['removed', '']]); hist_(c, 'Removed (must be linked again with a new invite)', by); }
    else if (a === 'rename') { const n = String(body.name || '').trim(); if (!n) return { ok: false, error: 'Type the new name.' }; c.name = n; d.sh.getRange(d.row, 2).setValue(n); hist_(c, 'Renamed to ' + n, by); }
    else if (a === 'shop') { c.shop = String(body.shop || '') || null; hist_(c, 'Shop set to ' + (c.shop || 'none'), by); }
    else if (a === 'roles') { c.roles = (body.roles || []).map(String); hist_(c, c.roles.length ? 'Only these roles may log in: ' + c.roles.join(', ') : 'Every role may log in', by); }
    else if (a === 'end') { const e = String(body.endDate || ''); if (e && !/^\d{4}-\d{2}-\d{2}$/.test(e)) return { ok: false, error: 'Choose the last day.' }; c.endDate = e || null; hist_(c, e ? 'Access ends after ' + e : 'No end date', by); }
    else if (a === 'wipe') { c.wipe = 'pending'; hist_(c, 'Wipe asked (on its next contact)', by); }
    else if (a === 'admin') { c.admin = !!body.on; hist_(c, c.admin ? 'Marked as an admin device (roles, PINs and locked settings may be changed on it)' : 'No longer an admin device', by); }
    else { c.cmds = (c.cmds || []).concat([{ kind: a, at: now_(), by: String(body.by || '') }]).slice(-10); hist_(c, { sync: 'Asked to sync now', log: 'Asked for its problem log', signout: 'Asked to sign everyone out' }[a], by); }
    ctlSave_(all);
    if (['block', 'allow', 'remove', 'rename'].indexOf(a) >= 0) devicesChanged_();
    if (['wipe', 'remove', 'block', 'allow', 'admin'].indexOf(a) >= 0) secEvent_('device', { device: String(d.v[1] || id), person: who.user ? who.user.name : by, what: { wipe: 'Device to be wiped', remove: 'Device removed', block: 'Device blocked', allow: 'Device allowed again', admin: body.on ? 'Marked as an admin device' : 'No longer an admin device' }[a] + ' (by ' + by + ')', refused: false });
    return { ok: true, devices: devices_() };
  });
}
function problemLogSave_(body, who) {
  const text = String(body.text || '').slice(0, 45000);
  sheet_('Problem logs', ['at', 'device', 'text']).appendRow([now_(), who.deviceId, text]);
  const all = ctlAll_(); const c = ctl_(all, who.deviceId); c.lastLog = now_(); hist_(c, 'Sent its problem log', 'device'); ctlSaveSafe_(all);
  return { ok: true };
}
function problemLogGet_(id) {
  const rows = values_(sheet_('Problem logs', ['at', 'device', 'text'])).filter(function (r) { return String(r[1]) === String(id); });
  const last = rows[rows.length - 1];
  return { ok: true, at: last ? last[0] : null, text: last ? String(last[2]) : '' };
}

// Settings changed from the app (Admin › Devices), so nobody has to open the Google editor again.
function cloudSettings_(body) {
  const props = PropertiesService.getScriptProperties();
  if (body.requireApproval === true || body.requireApproval === false) props.setProperty('REQUIRE_APPROVAL', body.requireApproval ? '1' : '0');
  // 1.0.0 build 1: once every device runs the new app, the owner switches on "every device must sign".
  if (body.strictSign === true || body.strictSign === false) props.setProperty('STRICT_SIGN', body.strictSign ? '1' : '0');
  if (body.linkCutover === true || body.linkCutover === false) props.setProperty('LINK_CUTOVER', body.linkCutover ? '1' : '0');
  return { ok: true, requireApproval: props.getProperty('REQUIRE_APPROVAL') !== '0', strictSign: props.getProperty('STRICT_SIGN') === '1', linkCutover: props.getProperty('LINK_CUTOVER') === '1' };
}

// ================================================================ main Easy Buys software
function mainApi_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('MAIN_KEY');
  if (!key || body.key !== key) return { ok: false, error: 'Not allowed' };
  // 1.0.0 build 1: once every till PC has its own key (the owner pressed "Till keys done" – LINK_CUTOVER), the old Links'
  // main_push / main_link_check with the shared key are refused, and the owner is told.
  if (props_().getProperty('LINK_CUTOVER') === '1' && (body.kind === 'main_push' || body.kind === 'main_link_check')) {
    secEvent_('link_key', { device: 'till PC ' + String(body.source || '?'), what: 'An old Link used the shared main key after the cut-over' });
    return { ok: false, error: 'Not allowed – this till PC needs its own till key (Aronium › Till PCs › Till key).' };
  }
  PropertiesService.getScriptProperties().setProperty('LAST_MAIN_SYNC', now_());
  if (body.kind === 'main_pull') return pull_(Number(body.since) || 0, 'main', MAIN_TABLES_PULL_ALL);
  if (body.kind === 'main_push') return push_(body.records || [], 'main');
  if (body.kind === 'main_link_check') return linkCheck_(body);
  return { ok: true, devices: devices_() };
}

// v3.9.4 the till program's test buttons: did the cloud keep its test, and has the main tablet answered it?
// Only the two small test rows of that till are read (found by their key, not by reading the whole sheet).
function linkCheck_(body) {
  const src = String(body.source || '').replace(/[^A-Za-z0-9-]/g, '');
  const out = { ok: true, version: SCRIPT_VERSION, ping: null, pong: null, devices: [] };
  const sh = data_().getSheetByName('T_aronium_feed');
  if (sh && src) {
    ['ping', 'pong'].forEach(function (k) {
      const key = src + '|' + k + '|ping|0';
      let rowNo = 0;
      try { const cell = sh.getRange('A:A').createTextFinder(key).matchEntireCell(true).findNext(); rowNo = cell ? cell.getRow() : 0; } catch (e) {
        const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
        for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) { rowNo = i + 2; break; }
      }
      if (!rowNo) return;
      try { out[k] = JSON.parse(JSON.parse(String(sh.getRange(rowNo, 5, 1, 1).getValues()[0][0] || '{}')).data || 'null'); } catch (e) { out[k] = null; }
    });
  }
  out.devices = deviceRows_().filter(function (r) { return String(r[3]) === 'active' && !/^(bridge-|first-device)/.test(String(r[0])); })
    .map(function (r) { return { name: String(r[1]), last_seen: text_(r[6]) }; });
  return out;
}

// Run once from the editor: creates the data spreadsheet and the main-software key.
function setupRun_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MAIN_KEY')) props.setProperty('MAIN_KEY', Utilities.getUuid() + Utilities.getUuid());
  data_();
  ['Changes', 'Devices', 'Invites'].forEach(function (n) { sheet_(n, n === 'Changes' ? CHANGES_HEADER : n === 'Devices' ? DEVICE_HEADER : INVITE_HEADER); });
  securityDigestTrigger_(); // 1.0.0 build 1: the daily security summary at 7:00 am Fiji time
  Logger.log('Cloud folder: ' + root_().getUrl());
  Logger.log('Data spreadsheet: ' + data_().getUrl());
  Logger.log('Main-software key (keep private): ' + props.getProperty('MAIN_KEY'));
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}
// Optional: new devices must be approved by the owner before they sync. Run once to turn on/off.
function requireApprovalOn() { PropertiesService.getScriptProperties().setProperty('REQUIRE_APPROVAL', '1'); }
function requireApprovalOff() { PropertiesService.getScriptProperties().setProperty('REQUIRE_APPROVAL', '0'); }

// ================================================================ v3.6 Start fresh (go-live)
// Only the main tablet (signed in with the cloud key) may do it. The listed tables are emptied (the owner logins stay),
// the change list keeps working (old entries just point at nothing), and records from before the fresh start that a
// device still had waiting are refused when they arrive. Test photos and documents move to a "Test data" folder.
function freshStart_(body) {
  // 1.0.0 build 9: the owner's session is checked before this (needOwner_) – from any device; devices no longer hold the cloud key
  const tables = (body.tables || []).filter(function (t) { return /^[a-z_0-9]+$/.test(t); });
  const at = String(body.at || '');
  if (!tables.length || !at) return { ok: false, error: 'Nothing to clear.' };
  return withLock_(function () {
    const ss = data_();
    let cleared = 0;
    tables.forEach(function (t) {
      const sh = ss.getSheetByName('T_' + t);
      if (!sh || sh.getLastRow() < 2) return;
      const rows = values_(sh);
      const keep = rows.filter(function (r) {
        let row = {};
        try { row = JSON.parse(r[4] || '{}'); } catch (e) { row = {}; }
        if (String(r[1]) >= at) return true;                       // made after the fresh start
        if (t === 'users' && row.role === 'owner') return true;     // owner logins stay
        if (t === 'shared_settings' && String(r[0]) === 'fresh_start') return true;
        return false;
      });
      cleared += rows.length - keep.length;
      sh.getRange(2, 1, rows.length, sh.getLastColumn()).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
    });
    forgetValues_();
    PropertiesService.getScriptProperties().setProperty('FRESH', JSON.stringify({ id: String(body.id || ''), at: at, tables: tables }));
    let moved = 0;
    if (body.moveFiles) {
      const root = root_();
      const test = root.createFolder('Test data (before go-live) ' + at.slice(0, 10));
      ['Photos', 'Documents', 'Reports', 'Purchase orders'].forEach(function (name) {
        const it = root.getFoldersByName(name);
        while (it.hasNext()) { it.next().moveTo(test); moved += 1; }
      });
    }
    log_('Fresh start', [String(body.id || ''), at, tables.length, cleared]);
    return { ok: true, cleared: cleared, foldersMoved: moved };
  });
}

// ---------------------------------------------------------------- v3.5 My day (Google Calendar)
const MYDAY_CAL = 'Easy Buys – My day';
function mydayDay_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Bad date');
  return new Date(date + 'T00:00:00');
}
// Meetings on one day from the default (main) calendar of the Easy Buys account. The app's own calendars are left out.
function mydayEvents_(date) {
  const day = mydayDay_(date);
  const tz = Session.getScriptTimeZone();
  const cal = CalendarApp.getDefaultCalendar();
  if (!cal || cal.getName().indexOf('Easy Buys Roster – ') === 0 || cal.getName() === MYDAY_CAL) return { ok: true, events: [] };
  const events = cal.getEventsForDay(day).filter(function (ev) { return ev.getTag('ebmd') !== '1' && ev.getTag('eb') !== '1'; }).map(function (ev) {
    const vis = ev.getVisibility ? String(ev.getVisibility()) : '';
    return {
      id: ev.getId(), title: ev.getTitle(), allDay: ev.isAllDayEvent(),
      start: Utilities.formatDate(ev.getStartTime(), tz, 'HH:mm'), end: Utilities.formatDate(ev.getEndTime(), tz, 'HH:mm'),
      location: ev.getLocation ? ev.getLocation() : '', private: /PRIVATE|CONFIDENTIAL/i.test(vis),
    };
  });
  return { ok: true, date: date, events: events };
}
// The day's app tasks replace whatever the app put in the "Easy Buys – My day" calendar for that day before.
function mydayOut_(date, items, name) {
  const day = mydayDay_(date);
  const tz = Session.getScriptTimeZone();
  const calName = MYDAY_CAL; // always the app's own calendar, whatever a device sends
  const cal = CalendarApp.getCalendarsByName(calName)[0] || CalendarApp.createCalendar(calName, { timeZone: tz });
  const end = new Date(day.getTime() + 86400000);
  cal.getEvents(day, end).forEach(function (ev) { if (ev.getTag('ebmd') === '1') ev.deleteEvent(); });
  let n = 0;
  (items || []).slice(0, 200).forEach(function (x) {
    if (!/^\d{2}:\d{2}$/.test(String(x.start)) || !/^\d{2}:\d{2}$/.test(String(x.end))) return;
    const st = new Date(date + 'T' + x.start + ':00');
    let en = new Date(date + 'T' + x.end + ':00');
    if (en <= st) en = new Date(st.getTime() + 15 * 60000);
    const ev = cal.createEvent(String(x.title || 'Task').slice(0, 200), st, en, { description: String(x.note || '') });
    ev.setTag('ebmd', '1');
    if (x.key) ev.setTag('key', String(x.key).slice(0, 200));
    n += 1;
  });
  return { ok: true, events: n, calendar: calName };
}

// ================================================================ v3.5 documents in Drive, visibility, backups
// Every finished document: Easy Buys › Documents › <type folder> [› year › month]. One row per upload in the
// "Files" sheet: a repeated upload id returns the file already there (no duplicates on retry), and the size and
// SHA-256 of what arrived are returned so the device can check the upload.
const FILES_HEADER = ['upload_id', 'file_id', 'name', 'path', 'size', 'sha256', 'doc_type', 'access', 'device', 'uploaded_at', 'ref_uid', 'mime'];
const RESTRICTED_ROLES = ['payroll', 'admin', 'owner'];
function filesSheet_() { return sheet_('Files', FILES_HEADER); }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function bytesSha_(bytes) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)); }
function cleanName_(s, max) { return String(s || '').replace(/[\\/:*?"<>|]+/g, '-').replace(/^\.+/, '').trim().slice(0, max || 60); }
function fileRow_(fileId) {
  if (!fileId) return null;
  const sh = data_().getSheetByName('Files');
  if (!sh) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) if (String(rows[i][1]) === String(fileId)) return rows[i];
  return null;
}
function insideFolder_(file, folderId, depth) {
  let level = [];
  const ps = file.getParents();
  while (ps.hasNext()) level.push(ps.next());
  for (let d = 0; d < (depth || 1) && level.length; d++) {
    const next = [];
    for (let i = 0; i < level.length; i++) {
      if (level[i].getId() === folderId) return true;
      const up = level[i].getParents();
      while (up.hasNext()) next.push(up.next());
    }
    level = next;
  }
  return false;
}

function filePut_(body, who) {
  return withLock_(function () {
    const uploadId = String(body.uploadId).slice(0, 80);
    const sh = filesSheet_();
    const rows = values_(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) !== uploadId) continue;
      try {
        const old = DriveApp.getFileById(String(rows[i][1]));
        if (!old.isTrashed()) return { ok: true, id: old.getId(), url: old.getUrl(), size: Number(rows[i][4]), sha256: String(rows[i][5]), existing: true };
      } catch (e) { /* gone: upload again below */ }
    }
    const bytes = Utilities.base64Decode(String(body.data || ''));
    const sha = bytesSha_(bytes);
    if (body.sha256 && sha !== String(body.sha256).toLowerCase()) return { ok: false, error: 'The file arrived damaged (its fingerprint is different). It will be sent again.' };
    if (body.size !== undefined && Number(body.size) !== bytes.length) return { ok: false, error: 'The file arrived incomplete. It will be sent again.' };
    const path = (Array.isArray(body.path) ? body.path : [body.folder || 'other']).slice(0, 4).map(function (n) { return cleanName_(n, 50); }).filter(String);
    const name = cleanName_(body.fileName || 'document', 120);
    const mime = String(body.mime || 'application/octet-stream');
    const f = path_(['Documents'].concat(path.length ? path : ['other'])).createFile(Utilities.newBlob(bytes, mime, name));
    sh.appendRow([uploadId, f.getId(), name, path.join(' / '), String(bytes.length), sha, String(body.docType || ''), String(body.access || 'normal'), who.deviceId, now_(), String(body.refUid || ''), mime]);
    return { ok: true, id: f.getId(), url: f.getUrl(), size: bytes.length, sha256: sha };
  });
}

// The role of a person, from the staff records synced to the cloud (not what the device says).
function userRole_(uid) {
  if (!uid) return null;
  const sh = data_().getSheetByName('T_users');
  if (!sh) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(uid)) continue;
    let u = {};
    try { u = JSON.parse(rows[i][4] || '{}'); } catch (e) { u = {}; }
    if (String(u.active) === '0') return { role: '', active: false };
    return { role: String(u.role || ''), active: true };
  }
  return null;
}
function sharedSetting_(key) {
  const sh = data_().getSheetByName('T_shared_settings');
  if (!sh) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) !== key) continue;
    try { return JSON.parse(JSON.parse(rows[i][4] || '{}').value || 'null'); } catch (e) { return null; }
  }
  return null;
}
// Documents uploaded by v3.4 apps: type and access class from the synced document list.
function legacyDocType_(fileId) {
  const sh = data_().getSheetByName('T_documents');
  if (!sh) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) {
    let d = {};
    try { d = JSON.parse(rows[i][4] || '{}'); } catch (e) { continue; }
    if (String(d.cloud_id) !== String(fileId)) continue;
    const ref = String(d.ref_type || '');
    const type = d.doc_type || ({ contract: 'contract', applicant: 'contract', licence: 'licence', invoice: 'invoice', payment: 'payment' }[ref] || 'report');
    return { type: type, access: d.access_class || (type === 'contract' || type === 'payslip' ? 'restricted' : 'normal') };
  }
  return null;
}
// May this person see this type of document? Same rule as the app (Admin › Documents › Who can see documents).
function docAllowed_(user, type, access, legacyFile) {
  if (!user || !user.uid) {
    // Older apps (v3.4) do not say who is logged in: they keep opening what they could before, never restricted files.
    if (access === 'restricted' && !legacyFile) return { ok: false, error: 'Update the Easy Buys app to open this document.' };
    return { ok: true };
  }
  const who = userRole_(user.uid);
  if (!who) return { ok: false, error: 'Your login has not reached the cloud yet. Tap Sync now, then try again.' };
  if (!who.active) return { ok: false, error: 'This login is not active.' };
  if (who.role === 'owner' || who.role === 'admin') return { ok: true };
  const vis = sharedSetting_('doc_visibility') || {};
  const roles = Array.isArray(vis[type]) ? vis[type] : (access === 'restricted' ? RESTRICTED_ROLES : null);
  if (roles && roles.indexOf(who.role) < 0) return { ok: false, error: 'Your role may not open this type of document. Ask the owner.' };
  return { ok: true };
}
function fileGet_(body, who) {
  const id = String(body.id || '');
  const row = fileRow_(id);
  let file;
  try { file = DriveApp.getFileById(id); file.getName(); } catch (e) { return { ok: false, error: 'This file is no longer in Google Drive.' }; }
  // 1.0.0 (stream C) photos (invoices, cheques, deposit slips…) live in Photos › month › type: Today fetches them when the phone's copy was cleared
  // script 16: and in Records › month › [shop ›] kind – but never a payslip copy (only the owner opens those, in Drive).
  const inRecords = insideFolder_(file, path_(['Records']).getId(), 4) && !inPayslips_(file);
  if (!row && !insideFolder_(file, path_(['Documents']).getId(), 5) && !insideFolder_(file, path_(['Photos']).getId(), 4) && !inRecords) return { ok: false, error: 'Not an Easy Buys document.' };
  const legacy = row ? null : legacyDocType_(id);
  const type = row ? String(row[6] || '') : legacy ? legacy.type : '';
  const access = row ? String(row[7] || 'normal') : legacy ? legacy.access : 'normal';
  // 1.0.0 build 1: the person is the one the cloud checked (their session), not whoever the device says is logged in.
  // Without a session (old apps) a restricted file is never given.
  const allowed = docAllowed_(who.user ? { uid: who.user.uid } : null, type, access, !row);
  if (!allowed.ok) { log_('Refused', [who.deviceId, body.user ? body.user.uid : '(old app)', id, type, allowed.error]); return allowed; }
  const bytes = file.getBlob().getBytes();
  return { ok: true, data: Utilities.base64Encode(bytes), mime: file.getMimeType(), name: file.getName(), size: bytes.length, sha256: bytesSha_(bytes) };
}

// Nightly backups from the main tablet: Backups › Daily (kept 30 days) and the first one of each month also in
// Backups › Monthly (kept 12 months), each with a file manifest (every document and photo and where its copy is).
const DAILY_KEEP_DAYS = 30;
const MONTHLY_KEEP = 12;
function backupTiered_(body, who) {
  const root = path_(['Backups']);
  const tier = body.tier === 'manual' ? 'Manual' : 'Daily';
  const folder = folder_(root, tier);
  const name = cleanName_(body.fileName || ('backup-' + now_().slice(0, 10) + '.json'), 80);
  const f = saveFile_(folder, name, Utilities.base64Encode(String(body.data || ''), Utilities.Charset.UTF_8), 'application/json');
  const manifestName = name.replace(/\.json$/, '') + '-manifest.json';
  const m = body.manifest ? saveFile_(folder, manifestName, Utilities.base64Encode(String(body.manifest), Utilities.Charset.UTF_8), 'application/json') : null;
  let monthly = false;
  if (tier === 'Daily') {
    const month = String(body.date || now_()).slice(0, 7);
    const mf = folder_(root, 'Monthly');
    const it = mf.getFiles();
    let have = false;
    while (it.hasNext()) if (it.next().getName().indexOf('backup-' + month) === 0) have = true;
    if (!have) { f.makeCopy(name, mf); if (m) m.makeCopy(manifestName, mf); monthly = true; }
    pruneBackups_(folder, function (date) { return date < daysAgo_(DAILY_KEEP_DAYS); });
    pruneBackups_(mf, function (date) { return date.slice(0, 7) < monthsAgo_(MONTHLY_KEEP - 1); });
  }
  log_('Backups', [who.deviceId, tier, name, monthly ? 'also monthly' : '']);
  return { ok: true, id: f.getId(), manifestId: m ? m.getId() : null, monthly: monthly };
}
function daysAgo_(n) { return new Date(Date.now() - n * 86400000).toISOString().slice(0, 10); }
function monthsAgo_(n) { const d = new Date(); d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() - n); return d.toISOString().slice(0, 7); }
// Backups older than the limit go to the Drive bin (Google keeps them there 30 more days).
function pruneBackups_(folder, tooOld) {
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const m = /^backup-(\d{4}-\d{2}-\d{2})/.exec(f.getName());
    if (m && tooOld(m[1])) f.setTrashed(true);
  }
  // 1.0.0: the parts of a backup made in parts go with it (an index kept elsewhere – the monthly copy – still points
  // at them, so a folder is kept while any index made from it is younger than the monthly limit: parts are only
  // binned here for the daily ones when no monthly copy was made from them)
  const fs = folder.getFolders();
  while (fs.hasNext()) {
    const d = fs.next();
    const m = /^backup-(\d{4}-\d{2}-\d{2}).*\.parts-/.exec(d.getName());
    if (m && tooOld(m[1]) && !partsInMonthly_(d.getId())) { try { d.setTrashed(true); } catch (e) { /* kept */ } }
  }
}
function partsInMonthly_(folderId) {
  const it = path_(['Backups']).getFoldersByName('Monthly');
  if (!it.hasNext()) return false;
  const files = it.next().getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getSize() > 200000 || /-manifest\.json$/.test(f.getName())) continue; // an index is small
    try { const x = JSON.parse(f.getBlob().getDataAsString('UTF-8')); if (x && x._upload && x._upload.folder === folderId) return true; } catch (e) { /* not an index */ }
  }
  return false;
}
// 1.0.0 speed: a backup in parts. The parts of one upload go into "<name>.parts-<upload>" next to where the backup is
// kept (Backups, › Daily or › Manual); backup_done then keeps the index file <name> ({ _meta, _parts: [file ids] }).
// While an upload is not finished a part sent again (the app went on after being cut off) replaces the one before;
// once the index is kept the parts can no longer change (a kept backup is never swapped).
function backupPartsHome_(body) {
  const root = path_(['Backups']);
  const tier = body.tier === 'plain' ? null : (body.tier === 'manual' ? 'Manual' : 'Daily');
  const folder = tier ? folder_(root, tier) : root;
  const name = cleanName_(body.fileName || ('backup-' + now_().slice(0, 10) + '.json'), 80);
  const upload = String(body.upload || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60);
  if (!upload) throw new Error('No upload id.');
  return { root: root, folder: folder, tier: tier, name: name, upload: upload, partsName: name.replace(/\.json$/, '') + '.parts-' + upload };
}
function backupPart_(body, who) {
  const h = backupPartsHome_(body);
  const parts = folder_(h.folder, h.partsName);
  if (parts.getFilesByName('done').hasNext()) return { ok: false, error: 'This backup is already kept; its parts cannot change.' };
  const n = Math.floor(Number(body.n));
  if (!(n >= 0 && n < 10000)) return { ok: false, error: 'Bad part number.' };
  const partName = 'part-' + ('0000' + n).slice(-4) + '.json';
  const old = parts.getFilesByName(partName);
  while (old.hasNext()) old.next().setTrashed(true);
  parts.createFile(Utilities.newBlob(Utilities.base64Decode(Utilities.base64Encode(String(body.data || ''), Utilities.Charset.UTF_8)), 'application/json', partName));
  return { ok: true, n: n };
}
function backupDone_(body, who) {
  const h = backupPartsHome_(body);
  const it = h.folder.getFoldersByName(h.partsName);
  if (!it.hasNext()) return { ok: false, error: 'No parts of this backup came in.' };
  const parts = it.next();
  const count = Math.floor(Number(body.parts));
  const ids = [];
  for (let n = 0; n < count; n++) {
    const f = parts.getFilesByName('part-' + ('0000' + n).slice(-4) + '.json');
    if (!f.hasNext()) return { ok: false, error: 'Part ' + (n + 1) + ' of the backup is missing – it is sent again on the next try.' };
    ids.push(f.next().getId());
  }
  if (!parts.getFilesByName('done').hasNext()) parts.createFile(Utilities.newBlob(Utilities.base64Decode(Utilities.base64Encode(now_())), 'text/plain', 'done'));
  const index = {};
  if (body.head && typeof body.head === 'object') Object.keys(body.head).forEach(function (k) { index[k] = body.head[k]; });
  index._meta = body.meta || { app: 'Easy Buys', made: now_() };
  index._parts = ids;
  index._upload = { id: h.upload, folder: parts.getId() };
  if (!h.tier) {
    const f0 = saveFile_(h.folder, h.name, Utilities.base64Encode(JSON.stringify(index), Utilities.Charset.UTF_8), 'application/json');
    log_('Backups', [who.deviceId, 'Older', h.name, 'in ' + count + ' parts']);
    return { ok: true, id: f0.getId(), parts: count };
  }
  // the same keeping as a backup in one file (daily 30 days, the first of each month 12 months, manual for ever)
  const r = backupTiered_({ tier: body.tier, date: body.date, fileName: h.name, data: JSON.stringify(index), manifest: body.manifest }, who);
  return { ok: r.ok, id: r.id, manifestId: r.manifestId, monthly: r.monthly, parts: count };
}
function backupGetPart_(body, who) {
  const file = DriveApp.getFileById(String(body.id || ''));
  if (!insideFolder_(file, path_(['Backups']).getId(), 2)) return { ok: false, error: 'Not an Easy Buys backup.' };
  const index = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  const n = Math.floor(Number(body.n));
  if (!index || !index._parts || !(n >= 0 && n < index._parts.length)) return { ok: false, error: 'No such part.' };
  const part = DriveApp.getFileById(index._parts[n]);
  if (!insideFolder_(part, path_(['Backups']).getId(), 3)) return { ok: false, error: 'Not an Easy Buys backup.' };
  return { ok: true, n: n, text: part.getBlob().getDataAsString('UTF-8') };
}
function backupFiles_() {
  const root = path_(['Backups']);
  const list = [];
  const add = function (folder, tier) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/-manifest\.json$/.test(f.getName())) continue;
      list.push({ id: f.getId(), name: f.getName(), size: f.getSize(), at: f.getDateCreated().toISOString(), tier: tier });
    }
  };
  add(root, 'Older');
  ['Daily', 'Monthly', 'Manual'].forEach(function (t) { const it = root.getFoldersByName(t); if (it.hasNext()) add(it.next(), t); });
  list.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
  return list;
}

// ================================================================ v3.5 Windows software: full file mirror
// main_files: every document filed since a time (and the backups); main_file_get: one file with its fingerprint.
// Protected by the main-software key (Script properties › MAIN_KEY, made by setup()).
function mainFiles_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('MAIN_KEY');
  if (!key || body.key !== key) return { ok: false, error: 'Not allowed' };
  if (body.kind === 'main_file_get') {
    const id = String(body.id || '');
    let file;
    try { file = DriveApp.getFileById(id); file.getName(); } catch (e) { return { ok: false, error: 'Not found' }; }
    if (!fileRow_(id) && !insideFolder_(file, path_(['Documents']).getId(), 5) && !insideFolder_(file, path_(['Backups']).getId(), 2)) return { ok: false, error: 'Not an Easy Buys file' };
    const bytes = file.getBlob().getBytes();
    return { ok: true, id: id, name: file.getName(), mime: file.getMimeType(), size: bytes.length, sha256: bytesSha_(bytes), data: Utilities.base64Encode(bytes) };
  }
  const since = String(body.since || '');
  const sh = data_().getSheetByName('Files');
  const files = (sh ? values_(sh) : []).filter(function (r) { return String(r[9]) > since; }).map(function (r) {
    return { kind: 'document', uploadId: r[0], id: r[1], name: r[2], path: r[3], size: Number(r[4]), sha256: r[5], docType: r[6], access: r[7], device: r[8], at: r[9], refUid: r[10], mime: r[11] };
  });
  backupFiles_().forEach(function (b) { if (b.at > since) files.push({ kind: 'backup', id: b.id, name: b.name, path: 'Backups / ' + b.tier, size: b.size, at: b.at }); });
  files.sort(function (a, b) { return a.at < b.at ? -1 : 1; });
  const limit = Math.min(1000, Number(body.limit) || 500);
  const page = files.slice(0, limit);
  PropertiesService.getScriptProperties().setProperty('LAST_MAIN_SYNC', now_());
  return { ok: true, files: page, next: page.length ? page[page.length - 1].at : since, more: files.length > limit };
}

// ================================================================ 1.0.0 invoices read by AI (Google Gemini)
// The owner's own free key from aistudio.google.com is kept in Project settings › Script properties › GEMINI_KEY
// (never in this file – the code is public). The app sends the pages (photos as JPEG, or a PDF) and hints (the
// suppliers it knows); the answer is JSON in the shape below. Every figure is still checked by a person in the app.
// Models tried in turn: GEMINI_MODEL in Script properties (if set), then these. When Google retires a model, or one
// model's free limit is used up, the next one is tried – nothing to change in this file.
const AI_MODELS = ['gemini-3.8-flash', 'gemini-3.1-flash-lite', 'gemini-2.5-flash'];
const AI_MODEL = AI_MODELS[0];
function aiModels_() {
  const own = String(PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || '').trim().replace(/^models\//, '');
  return (own ? [own] : []).concat(AI_MODELS.filter(function (m) { return m !== own; }));
}
const AI_MAX_BYTES = 19 * 1024 * 1024;   // Gemini takes about 20 MB of pages in one request
function aiInvoiceSchema_(hints) {
  const n = function (d) { return { type: 'NUMBER', nullable: true, description: d }; };
  const s = function (d) { return { type: 'STRING', nullable: true, description: d }; };
  const schema = {
    type: 'OBJECT',
    properties: {
      supplier: s('The company that SOLD the goods and issued the invoice (never Easy Buys, the buyer)'),
      supplier_tin: s("The supplier's TIN / VAT number, digits only (never 50-14315-0-7, Easy Buys' own)"),
      invoice_no: s("The supplier's invoice number (not the order number, account number or the VMS / SDC fiscal number)"),
      invoice_date: s('The invoice date as YYYY-MM-DD (Fiji writes dates day/month/year)'),
      prices_incl_vat: { type: 'BOOLEAN', nullable: true, description: 'true if the unit prices and line amounts printed include VAT (VIP), false if they are before VAT (VEP)' },
      subtotal_excl: n('Total before VAT (VEP)'),
      vat: n('Total VAT (also called GST or Tax) in dollars – not a percentage'),
      total_incl: n('The total to pay including VAT (VIP / Invoice Total / Total Incl. VAT)'),
      lines: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            code: s("The supplier's item code"),
            description: { type: 'STRING', description: 'The product name as printed' },
            qty: n('Quantity'),
            unit: s('Unit: Each, Carton, Bag, Outer, Pkt, Kg …'),
            unit_price: n('Price of one unit as printed'),
            disc_pct: n('Discount percent on the line, if any'),
            vat_pct: n('VAT percent of the line: 12.5 standard, 0 zero rated or exempt'),
            line_total: n('Line amount BEFORE VAT if the invoice shows it, otherwise as printed'),
          },
          required: ['description'],
        },
      },
      notes: s('Anything unclear: handwriting, a page cut off, totals that do not add up, more than one invoice'),
      confidence: n('How sure you are of the figures, 0 to 1'),
    },
    required: ['lines'],
  };
  // 1.0.0 build 54: a power or water bill (EFL / WAF) also gives its account number and due date
  if (hints && hints.bill) {
    schema.properties.account_no = s('The customer ACCOUNT number on the bill, digits only (not the invoice, meter or bill number)');
    schema.properties.due_date = s('The date the bill must be paid by (Due Date / Pay By / Due on) as YYYY-MM-DD');
  }
  return schema;
}
function aiPrompt_(hints) {
  const h = hints || {};
  const known = (Array.isArray(h.suppliers) ? h.suppliers : []).map(function (x) { return String(x || '').replace(/[\r\n]+/g, ' ').slice(0, 80); }).filter(String).slice(0, 200);
  const vat = Number(h.vatPct) > 0 && Number(h.vatPct) < 50 ? Number(h.vatPct) : 12.5;
  return [
    'You read supplier invoices for a grocery business in Fiji (currency FJD). The pages are one invoice (sometimes a delivery note or receipt with it).',
    'The BUYER is Easy Buys Shopping Centre (also written Easy Buys Shopping Center, Easy Buy, EBSC), TIN 50-14315-0-7, 29 Drasa Avenue / Tukani Street, Lautoka. The buyer is NEVER the supplier: the supplier is the company whose name, logo, TIN and phone are at the top of the invoice.',
    'Fiji VAT is ' + vat + '% (standard); some lines are zero rated (VAT 0). VEP = value excluding VAT, VIP = value including VAT. The VAT is an amount in dollars, never the percentage.',
    'Dates are day/month/year. Handwritten amounts may be written 18-00 meaning 18.00. Read every item line, in order, across all pages; do not invent lines. Do not take the SDC / VMS fiscal invoice number as the invoice number.',
    'Give numbers as plain numbers (no $ or commas). Use null for anything not on the invoice.',
    known.length ? 'Known suppliers (use the exact name when it is one of these): ' + known.join('; ') + '.' : '',
    h.bill ? aiBillPrompt_(h.bill) : '',
  ].filter(String).join('\n');
}
function aiScrub_(msg, key) {
  let m = String(msg || '');
  if (key) m = m.split(key).join('***');
  return m.replace(/key=[A-Za-z0-9_\-]+/g, 'key=***').slice(0, 300);
}
// files: [{ mimeType, data (base64) }] (at most 6), hints: { suppliers: [names], vatPct }
// → { ok: true, invoice, model } | { ok: false, code: 'NO_AI_KEY' | 'AI_ERROR', error }
// opts (1.0.0 build 54): { prompt, schema } to read another paper with the same models (the FCCC gas prices).
function aiReadInvoice_(files, hints, opts) {
  const key = String(PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') || '').trim();
  if (!key) return { ok: false, code: 'NO_AI_KEY', error: 'NO_AI_KEY: the AI reader is not switched on – add GEMINI_KEY in Project settings › Script properties.' };
  const started = Date.now();
  try {
    const parts = [{ text: (opts && opts.prompt) || aiPrompt_(hints) }];
    let bytes = 0;
    (files || []).slice(0, 6).forEach(function (f) {
      const mime = /pdf/i.test(String(f && f.mimeType)) ? 'application/pdf' : 'image/jpeg';
      const data = String((f && f.data) || '').replace(/^data:[^,]*,/, '');
      if (!data) return;
      bytes += data.length;
      parts.push({ inline_data: { mime_type: mime, data: data } });
    });
    if (parts.length < 2) return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: no pages were sent.' };
    if (bytes > AI_MAX_BYTES) return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: the pages are too big for one reading (about 13 MB at most).' };
    const body = function (model, think) {
      const gc = { temperature: 0, responseMimeType: 'application/json', responseSchema: (opts && opts.schema) || aiInvoiceSchema_(hints) };
      // quick answer: a fetch from Apps Script may take about 60 seconds at most
      if (think && /^gemini-2\./.test(model)) gc.thinkingConfig = { thinkingBudget: 0 };
      // the newer models "think" before answering unless told not to – much slower for reading an invoice
      else if (think) gc.thinkingConfig = { thinkingLevel: 'low' };
      return JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: gc });
    };
    let res = null; let code = 0; let json = null; let used = ''; let lastWhy = '';
    const models = aiModels_();
    for (let i = 0; i < models.length; i++) {
      const model = models[i];
      const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key);
      let think = true;
      for (let attempt = 0; attempt < 3; attempt++) {
        res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: body(model, think), muteHttpExceptions: true });
        code = res.getResponseCode();
        try { json = JSON.parse(res.getContentText()); } catch (e) { json = null; }
        const why0 = json && json.error && json.error.message ? String(json.error.message) : '';
        // a model that does not take the "quick answer" setting: once more without it
        if (code === 400 && think && /thinking/i.test(why0)) { think = false; continue; }
        // busy: one more try, only when there is time left (Apps Script stops a request at 6 minutes)
        // (a busy model – 503 – is not tried again: the next model is asked straight away)
        if (code === 500 && attempt === 0 && Date.now() - started < 60000) { Utilities.sleep(3000); continue; }
        break;
      }
      used = model;
      if (code === 200) break;
      lastWhy = json && json.error && json.error.message ? json.error.message : ('HTTP ' + code);
      if (code === 400 && /api key/i.test(lastWhy)) return { ok: false, code: 'NO_AI_KEY', error: 'NO_AI_KEY: the GEMINI_KEY in Script properties was not accepted – make a new key at aistudio.google.com.' };
      // retired / unknown model, or this model's free limit used up: the next model
      // (also when Google is busy – "high demand" / overloaded – another model often has room)
      const nextOk = code === 404 || code === 429 || code === 500 || code === 503 || /no longer available|not found|not supported|deprecated|retired|high demand|overloaded|unavailable|try again later/i.test(lastWhy);
      if (!nextOk || Date.now() - started > 50000) break; // the phone is waiting: at most about a minute in all
    }
    if (code !== 200) {
      if (code === 429) return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: the free AI limit is used up for now – try again later.' };
      if (code === 503 || /high demand|overloaded/i.test(lastWhy)) return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: Google\'s AI is busy right now – tap "Read again" in a few minutes.' };
      return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: ' + aiScrub_(lastWhy, key) };
    }
    const cand = json && json.candidates && json.candidates[0];
    if (!cand) return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: no answer' + (json && json.promptFeedback && json.promptFeedback.blockReason ? ' (' + json.promptFeedback.blockReason + ')' : '') + '.' };
    const out = ((cand.content && cand.content.parts) || []).map(function (p) { return p.text || ''; }).join('').trim();
    let invoice;
    try { invoice = JSON.parse(out.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '')); } catch (e) {
      return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: the answer was not complete' + (cand.finishReason ? ' (' + cand.finishReason + ')' : '') + '.' };
    }
    if (!invoice || typeof invoice !== 'object') return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: the answer was empty.' };
    return { ok: true, invoice: invoice, model: used, ms: Date.now() - started };
  } catch (e) {
    // a timeout or no connection from Google's side: the app reads it on the phone instead
    return { ok: false, code: 'AI_ERROR', error: 'AI_ERROR: ' + aiScrub_(e && e.message ? e.message : e, key) };
  }
}

// ================================================================ 1.0.0 build 54: power / water bills read by AI
// hints.bill = 'power' (EFL) or 'water' (WAF): the same invoice reading, told it is a utility bill, with account_no and
// due_date added to the answer (aiInvoiceSchema_).
function aiBillPrompt_(kind) {
  const who = kind === 'water' ? 'Water Authority of Fiji (WAF)' : kind === 'power' ? 'Energy Fiji Limited (EFL)' : 'Energy Fiji Limited (EFL) or the Water Authority of Fiji (WAF)';
  return [
    'This is a utility bill (or a screenshot of the bill checker) from ' + who + ', not a goods invoice. EFL = Energy Fiji Limited (electricity), WAF = Water Authority of Fiji (water).',
    'What matters: account_no = the customer account number (digits only; not the meter number, bill number or invoice number), total_incl = the total amount due now (Total Amount Due / Amount Due / Total Payable, 0 when nothing is owing), due_date = the date to pay by as YYYY-MM-DD.',
    'supplier = Energy Fiji Limited or Water Authority of Fiji. invoice_no = the bill or invoice number if printed, else null. invoice_date = the bill or issue date. Put any overdue / arrears amount in notes. A bill has no item lines: give lines as an empty list unless charges are listed.',
  ].join('\n');
}

// ================================================================ 1.0.0 build 54: FCCC LPG (gas) prices read in the cloud
// The phone could not read the FCCC page or PDF (no text in the PDF, the website refusing it, the web test build):
// Apps Script fetches fccc.gov.fj/master-price-list/gas/, finds the newest "LPG Authorization (Effective dd/mm/yyyy)",
// and reads Schedule 1 (Viti Levu & Vanua Levu) of its PDF with Gemini (GEMINI_KEY), or with Google Drive's text
// reading when there is no AI key. → { ok, effective, url, prices: { kg: { w, r }, … } | text, engine }
//   | { ok: false, code: 'FETCH' | 'NO_LIST' | 'NOT_READ', error }
const FCCC_GAS_PAGE = 'https://fccc.gov.fj/master-price-list/gas/';
const FCCC_MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function fcccIso_(y, m, d) { return y + '-' + ('0' + Number(m)).slice(-2) + '-' + ('0' + Number(d)).slice(-2); }
function fcccFullUrl_(href) {
  const u = String(href || '').replace(/&amp;/g, '&').trim();
  if (/^https?:\/\//i.test(u)) return u;
  if (u.indexOf('//') === 0) return 'https:' + u;
  if (u.indexOf('/') === 0) return 'https://fccc.gov.fj' + u;
  return FCCC_GAS_PAGE + u;
}
function fcccGasList_(html) {
  const out = [];
  const re = /<a\b[^>]*?href\s*=\s*["']([^"']+?\.pdf(?:[?#][^"']*)?)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(String(html || '')))) {
    const label = m[2].replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
    const url = fcccFullUrl_(m[1]);
    if (!/LPG/i.test(label) && !/LPG/i.test(url)) continue;
    let effective = null;
    const d = /Authori[sz]ation\s*\(?\s*Effective\s*:?\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/i.exec(label);
    if (d) effective = fcccIso_(d[3], d[2], d[1]);
    else if (/Authori[sz]ation/i.test(label + url)) {
      const f = /Effective[-_ ]+(\d{1,2})[-_ ]+([A-Za-z]{3,9})[-_ ]+(\d{2,4})/i.exec(url.split('/').pop() || '');
      const mo = f && FCCC_MONTHS[f[2].slice(0, 3).toLowerCase()];
      if (mo) effective = fcccIso_(f[3].length === 2 ? '20' + f[3] : f[3], mo, f[1]);
    }
    if (!effective || out.some(function (x) { return x.url === url; })) continue;
    out.push({ url: url, effective: effective });
  }
  return out.sort(function (a, b) { return a.effective < b.effective ? 1 : -1; });
}
function fcccLpgSchema_() {
  const n = function (d) { return { type: 'NUMBER', nullable: true, description: d }; };
  const line = function (d) { return { type: 'OBJECT', nullable: true, description: d, properties: { w: n('Maximum WHOLESALE price in dollars'), r: n('Maximum RETAIL price in dollars') } }; };
  return {
    type: 'OBJECT',
    properties: {
      effective_date: { type: 'STRING', nullable: true, description: 'The date the prices start, YYYY-MM-DD' },
      kg: line('Cylinder (residential) – Per KG'),
      c4_5: line('4.5 KG Cylinder'),
      c12: line('12 KG Cylinder'),
      c40: line('40 KG Cylinder'),
      c50: line('50 KG Cylinder'),
      auto: line('Automotive gas (autogas) – Per Litre'),
      notes: { type: 'STRING', nullable: true, description: 'Anything unclear' },
    },
  };
}
function fcccLpgPrompt_() {
  return [
    'This PDF is the Fijian Competition and Consumer Commission (FCCC) "Final Authorisation of Prices for the Supply of LPG Products in Fiji".',
    'Read ONLY Schedule 1 (Viti Levu and Vanua Levu) – never Schedule 2 or other islands. For each product give the maximum wholesale price (w) and retail price (r) in Fiji dollars as plain numbers.',
    'kg = the residential cylinder "Per KG" line (NOT the "Bulk – Ex Factory" per kg line). c4_5, c12, c40, c50 = the 4.5, 12, 40 and 50 KG cylinders. auto = Automotive Gas per litre. Use null for anything not printed (NA).',
  ].join('\n');
}
function fcccLpg_() {
  const get = function (url) { return UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (Linux; Android 14) EasyBuysApp', Accept: 'text/html,application/pdf,*/*' } }); };
  let page;
  try { page = get(FCCC_GAS_PAGE); } catch (e) { return { ok: false, code: 'FETCH', error: 'FETCH: could not open the FCCC gas page (' + String(e && e.message ? e.message : e).slice(0, 120) + ').' }; }
  if (page.getResponseCode() !== 200) return { ok: false, code: 'FETCH', error: 'FETCH: the FCCC gas page answered ' + page.getResponseCode() + '.' };
  const list = fcccGasList_(page.getContentText());
  if (!list.length) return { ok: false, code: 'NO_LIST', error: 'NO_LIST: no LPG authorisation was found on the FCCC gas page.' };
  const n = list[0];
  let pdf;
  try { pdf = get(n.url); } catch (e) { return { ok: false, code: 'FETCH', error: 'FETCH: could not download the ' + n.effective + ' authorisation.' }; }
  if (pdf.getResponseCode() !== 200) return { ok: false, code: 'FETCH', error: 'FETCH: the ' + n.effective + ' authorisation answered ' + pdf.getResponseCode() + '.' };
  const bytes = pdf.getBlob().getBytes();
  const ai = aiReadInvoice_([{ mimeType: 'application/pdf', data: Utilities.base64Encode(bytes) }], {}, { prompt: fcccLpgPrompt_(), schema: fcccLpgSchema_() });
  if (ai.ok && ai.invoice && (ai.invoice.c12 || ai.invoice.c4_5)) return { ok: true, effective: n.effective, url: n.url, prices: ai.invoice, engine: 'ai', model: ai.model };
  // no AI key, or the AI could not read it: Google Drive's text reading of the PDF (the app finds the prices in it)
  let text = '';
  try { text = ocr_(Utilities.newBlob(bytes, 'application/pdf', 'fccc-lpg.pdf')); } catch (e) { text = ''; }
  if (text) return { ok: true, effective: n.effective, url: n.url, text: String(text).slice(0, 60000), engine: 'ocr' };
  return { ok: false, code: 'NOT_READ', error: 'NOT_READ: the ' + n.effective + ' authorisation could not be read' + (ai.error ? ' (' + String(ai.error).slice(0, 150) + ')' : '') + '.' };
}

// ================================================================ 1.0.0 build 1 (script 12): security
// The cloud no longer trusts a device blindly (security audit C1–C4, H3, H5, H6, M1–M3):
//  • every request is signed with the device's pass (auth_ / checkSig_ above);
//  • the PERSON is checked too: logging in with a PIN on a device asks the cloud for a session (user_login). Owner /
//    admin work (locked records and settings, device control, invites, backups, calendar, Viber, Start clean…) needs an
//    owner / admin session from that device – a device alone is not enough;
//  • PINs: PBKDF2 with a salt per person (p2:…); in the data spreadsheet the hash is also locked with a secret only the
//    cloud holds (PIN_PEPPER); the "proof" a device sends at login is checked against a value kept only in Script
//    properties; owner / admin PIN hashes, bank details and face data are not sent to devices that do not need them;
//  • the change history (audit_log) can only be added to, with a running fingerprint, and goes to admin devices;
//  • times in the future are refused; only known tables are accepted (devices cannot make new sheets);
//  • old apps keep working for 7 days after this script is installed, or until the owner presses "All phones updated".
const PIN_ITER = 2000;
const PIN_TOP_ROLES = ['owner', 'admin', 'manager', 'payroll'];
const BOSS_ROLES = ['owner', 'admin'];
const USER_LOCKED_COLS = ['role', 'pin_hash', 'perms', 'permissions', 'works_shifts', 'own_till', 'active', 'email'];
// Locked settings (the same list as src/services/locked.js LOCKED_SHARED – tests/cloud12.test.js keeps the two the same).
const LOCKED_SHARED = ['role_perms', 'locked_admin_ok', 'modules_on', 'update_repo', 'go_live', 'gl_lock', 'gl_map', 'acct_settings', 'ap_rules',
  'pay_rules', 'pay_cycle', 'pay_signoff_rules', 'tax_calendar_rules', 'employer_details', 'till_rules', 'cig_tally', 'owed_rules',
  'viber_channels', 'doc_settings', 'main_tablet_device', 'fresh_start', 'clean_start', 'link_settings', 'ai_settings', 'backup_settings',
  'email_prefs', 'doc_visibility', 'watchdog', 'email_logins', 'recharge_setup'];
// Settings where only some fields are locked (src/services/locked.js LOCKED_FIELDS).
const LOCKED_FIELDS = { go_live: ['liveFrom', 'freshFrom'] };
const LOCKED_TABLES = { users: 'cols', shops: '*', tills: '*', gl_accounts: '*', bin_records: '*', bin_parts: '*', pay_profiles: '*',
  ap_suppliers: ['bank_name', 'bank_account'], gl_journals: 'opening' };
// The tables the app syncs (src/services/records.js SYNC_TABLES – a test keeps this list complete). Anything else is refused,
// so a device cannot fill the spreadsheet with sheets of its own.
const KNOWN_TABLES = ['aronium_refunds', 'shift_ends', 'handover_notes', 'verify_trail', 'cheque_steps', 'deposit_plans', 'users', 'shops', 'shared_settings', 'punches', 'timesheet_checks', 'leave_types', 'leave_entitlements', 'leave_requests', 'rosters',
  'ap_suppliers', 'ap_invoices', 'ap_payments', 'ap_allocations', 'bank_accounts', 'cheque_books', 'cheques', 'bank_entries', 'ap_recurring', 'loan_entries',
  'tills', 'till_counts', 'cash_moves', 'deposits', 'bank_statements', 'bank_lines', 'pay_profiles', 'staff_debts', 'staff_charges', 'pay_runs', 'pay_lines',
  'inv_items', 'inv_item_shops', 'inv_moves', 'inv_cost_changes', 'po_orders', 'po_lines', 'inv_receipts', 'inv_counts', 'inv_transfers', 'inv_count_lines',
  'inv_count_areas', 'verifications', 'money_snaps', 'alert_log', 'safe_counts', 'supplier_statements', 'gl_accounts', 'gl_journals', 'fixed_assets',
  'gl_budgets', 'aronium_batches', 'gl_recurring', 'gl_year_close', 'stat_payments', 'roster_shifts', 'job_openings', 'applicants', 'onboarding',
  'signatures', 'contracts', 'pay_signoffs', 'documents', 'licences', 'licence_events', 'bank_checks', 'wallets', 'wallet_moves', 'wallet_imports',
  'wallet_lines', 'wallet_shift_notes', 'bank_closings', 'po_versions', 'supplier_schedules', 'aronium_suppliers', 'supplier_layouts', 'invoice_reads',
  'till_shifts', 'aronium_eods', 'till_shortages', 'staff_credit_imports', 'staff_credit_balances', 'credit_settlements', 'regular_terms',
  'regular_receipts', 'loan_statements', 'loan_recons', 'owner_moves', 'personal_regulars', 'pay_slip_status', 'tax_proofs', 'staff_details',
  'roster_changes', 'myday_items', 'myday_marks', 'issued_docs', 've_items', 've_due', 've_work', 've_faults', 've_fuel', 've_checks',
  'accountant_queries', 'tax_returns', 'tax_return_versions', 'till_surplus', 'cheque_cover', 'handovers', 'fccc_lists', 'inv_count_entries',
  'inv_item_suppliers', 'aronium_feed', 'duty_lists', 'duty_runs', 'roster_marks', 'verify_queries', 'shop_accounts', 'bin_records', 'bin_parts',
  'utility_checks', 'product_changes', 'supplier_owed', 'inv_supplier_catalog', 'cig_tallies', 'audit_log', 'recharge_moves', 'recharge_sales',
  // tables the main software or older apps may still send
  'suppliers', 'materials', 'products', 'recipes', 'equipment', 'checklists', 'checklist_items', 'shift_names', 'steps', 'product_steps', 'product_decks',
  'safe_counts', 'shop_profiles', 'notifications', 'messages', 'stock_takes', 'handover_notes'];

function props_() { return PropertiesService.getScriptProperties(); }
function secret_(name) {
  const p = props_();
  let s = p.getProperty(name);
  if (!s) { s = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); p.setProperty(name, s); }
  return s;
}
// The day this script (12) first ran: old apps are accepted for 7 days after it, or until "All phones updated".
function transition_() {
  const p = props_();
  if (p.getProperty('STRICT_SIGN') === '1') return false;
  let since = Number(p.getProperty('V12_SINCE') || '0');
  if (!since) { since = Date.now(); p.setProperty('V12_SINCE', String(since)); }
  return Date.now() - since < 7 * 86400000;
}
function bytesOf_(s) { return Utilities.newBlob(String(s)).getBytes ? Utilities.newBlob(String(s)).getBytes() : Array.from(Utilities.base64Decode(Utilities.base64Encode(String(s)))); }
function hexOf_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
// PBKDF2-HMAC-SHA256, one 32-byte block (the same as src/services/pins.js).
function pbkdf2_(password, salt, iter) {
  const key = bytesOf_(password);
  let u = Utilities.computeHmacSha256Signature(bytesOf_(salt).concat([0, 0, 0, 1]), key);
  const t = u.slice();
  for (let i = 1; i < iter; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (let j = 0; j < t.length; j++) t[j] ^= u[j];
  }
  return hexOf_(t);
}
function pinSalt_(offered) {
  const p = props_();
  let s = p.getProperty('PIN_SALT');
  if (!s) {
    const o = String(offered || '');
    s = /^[A-Za-z0-9-]{16,80}$/.test(o) ? o : Utilities.getUuid() + Utilities.getUuid();
    p.setProperty('PIN_SALT', s);
  }
  return s;
}
// p2:<length>:<tag>:<salt>:<iterations>:<hash>. tag = one hex letter (so a PIN pad finds the few people it may be).
function pinHash_(pin) {
  const p = String(pin === undefined || pin === null ? '' : pin).trim();
  if (!/^\d{4,6}$/.test(p)) return null;
  const tag = sha_(pinSalt_() + ':' + p).charAt(0);
  const salt = sha_(Utilities.getUuid()).slice(0, 16);
  return 'p2:' + p.length + ':' + tag + ':' + salt + ':' + PIN_ITER + ':' + pbkdf2_(p, salt, PIN_ITER);
}
// What a device sends at login (it never keeps it): PBKDF2 of the PIN with a different salt. The cloud keeps only
// sha256(pepper + proof), in Script properties (not the spreadsheet, not on any device).
function pinProof_(pin, uid) { return pbkdf2_(String(pin), 'proof|' + uid + '|' + pinSalt_(), PIN_ITER); }
function proofKeep_(uid, proof) { if (/^[0-9a-f]{64}$/.test(String(proof || ''))) props_().setProperty('PV_' + uid, sha_(secret_('PIN_PEPPER') + ':' + proof)); }
function proofOk_(uid, proof) { const v = props_().getProperty('PV_' + uid); return !!v && /^[0-9a-f]{64}$/.test(String(proof || '')) && v === sha_(secret_('PIN_PEPPER') + ':' + proof); }
// In the spreadsheet a PIN hash is also locked with the pepper (x1:…), so a copy of the spreadsheet alone is not enough to
// try PINs. Devices get it back unlocked (only those that may have it).
function keystream_(uid, n) { let out = []; let i = 0; while (out.length < n) { const h = sha_(secret_('PIN_PEPPER') + '|' + uid + '|' + i); for (let j = 0; j < h.length; j += 2) out.push(parseInt(h.substr(j, 2), 16)); i++; } return out.slice(0, n); }
function lockHash_(uid, h) { if (!h || /^x1:/.test(h)) return h || null; const b = String(h).split('').map(function (c) { return c.charCodeAt(0); }); const k = keystream_(uid, b.length); return 'x1:' + hexOf_(b.map(function (x, i) { return x ^ k[i]; })); }
function unlockHash_(uid, x) {
  if (!x || !/^x1:[0-9a-f]+$/.test(x)) return x || null;
  const hex = x.slice(3); const b = []; for (let i = 0; i < hex.length; i += 2) b.push(parseInt(hex.substr(i, 2), 16));
  const k = keystream_(uid, b.length);
  return b.map(function (v, i) { return String.fromCharCode(v ^ k[i]); }).join('');
}
// A users record on its way into the spreadsheet: a plain PIN (old apps) becomes its hash and proof, then is dropped; a
// proof sent with a new PIN (_proof) is kept in Script properties only when allowed; the hash is locked with the pepper.
function hashUserRow_(row, uid, mayProof) {
  const out = {};
  Object.keys(row || {}).forEach(function (k) { out[k] = row[k]; });
  if (out.pin !== undefined && out.pin !== null && String(out.pin) !== '') {
    if (!out.pin_hash) out.pin_hash = pinHash_(out.pin);
    if (mayProof) proofKeep_(uid, pinProof_(out.pin, uid));
  }
  if (out._proof && mayProof) proofKeep_(uid, out._proof);
  delete out.pin; delete out._proof; delete out._vby;
  if (out.pin_hash) out.pin_hash = lockHash_(uid, out.pin_hash);
  return out;
}
// Once: every plain PIN already in the cloud copy is hashed and removed (old apps sent them plain).
function migratePins_() {
  const p = props_();
  if (p.getProperty('PINS_HASHED') === '2') return 0;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return 0;
  try {
    const sh = data_().getSheetByName('T_users');
    let n = 0;
    if (sh && sh.getLastRow() >= 2) {
      values_(sh).forEach(function (r, i) {
        let row = null;
        try { row = JSON.parse(r[4] || '{}'); } catch (e) { return; }
        if (!row || (row.pin === undefined && (!row.pin_hash || /^x1:/.test(row.pin_hash)))) return;
        sh.getRange(i + 2, 5, 1, 1).setValues([[JSON.stringify(hashUserRow_(row, String(r[0]), true))]]);
        n += 1;
      });
    }
    forgetValues_('T_users');
    p.setProperty('PINS_HASHED', '2');
    if (n) log_('PINs hashed', [n]);
    // Records of the locked kinds already in the cloud before script 12 count as the owner's (they were made by the
    // business before the change): marked once so a device linked later still takes them.
    if (p.getProperty('VBY_MARKED') !== '1') {
      ['users', 'shared_settings'].concat(Object.keys(LOCKED_TABLES)).forEach(function (t) {
        const s2 = data_().getSheetByName('T_' + t);
        if (!s2 || s2.getLastRow() < 2) return;
        const vals = s2.getRange(2, 5, s2.getLastRow() - 1, 1).getValues();
        let changed = false;
        vals.forEach(function (v) { try { const row = JSON.parse(text_(v[0]) || '{}'); if (row && !row._vby) { row._vby = { uid: 'before-1.0', role: 'owner' }; v[0] = JSON.stringify(row); changed = true; } } catch (e) { /* leave it */ } });
        if (changed) s2.getRange(2, 5, vals.length, 1).setValues(vals);
      });
      p.setProperty('VBY_MARKED', '1');
    }
    return n;
  } finally { lock.releaseLock(); }
}
function hashPinsRun_() { props_().deleteProperty('PINS_HASHED'); Logger.log('PINs hashed: ' + migratePins_()); }

function cloudUser_(uid) {
  const sh = data_().getSheetByName('T_users');
  if (!sh || !uid) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(uid)) continue;
    try { const u = JSON.parse(rows[i][4] || '{}'); u.uid = String(rows[i][0]); return u; } catch (e) { return null; }
  }
  return null;
}
function anyBossProof_() {
  const sh = data_().getSheetByName('T_users');
  if (!sh) return false;
  return values_(sh).some(function (r) { try { const u = JSON.parse(r[4] || '{}'); return BOSS_ROLES.indexOf(String(u.role)) >= 0 && String(u.active) !== '0' && !!props_().getProperty('PV_' + r[0]); } catch (e) { return false; } });
}

// ---------------------------------------------------------------- sessions: the person, checked by the cloud
function sessionsAll_() { try { return JSON.parse(props_().getProperty('SESSIONS') || '{}'); } catch (e) { return {}; } }
function sessionsSave_(all) {
  const now = Date.now();
  Object.keys(all).forEach(function (k) { if (!(all[k].exp > now)) delete all[k]; });
  props_().setProperty('SESSIONS', JSON.stringify(all));
}
// user_login: { uid, proof } → { ok, session, role, until }. 5 wrong tries for a person (or 20 from a device) lock for 15 minutes.
function userLogin_(body, who) {
  const uid = String(body.uid || '');
  const p = props_();
  let fails = {};
  try { fails = JSON.parse(p.getProperty('LOGIN_FAILS') || '{}'); } catch (e) { fails = {}; }
  const now = Date.now();
  const fu = fails['u:' + uid] || { n: 0, until: 0 };
  const fd = fails['d:' + who.deviceId] || { n: 0, until: 0 };
  if (fu.until > now || fd.until > now) return { ok: false, locked: true, error: 'Too many wrong PINs – try again in 15 minutes.' };
  const u = cloudUser_(uid);
  const good = u && String(u.active) !== '0' && proofOk_(uid, body.proof);
  if (!good) {
    fu.n += 1; fd.n += 1;
    if (fu.n >= 5) { fu.n = 0; fu.until = now + 15 * 60000; }
    if (fd.n >= 20) { fd.n = 0; fd.until = now + 15 * 60000; }
    fails['u:' + uid] = fu; fails['d:' + who.deviceId] = fd;
    p.setProperty('LOGIN_FAILS', JSON.stringify(fails));
    log_('Refused', [who.deviceId, 'user_login', uid, u ? 'wrong PIN' : 'unknown person']);
    failCount_('pin', uid + '|' + who.deviceId, { device: who.deviceId, person: u ? String(u.name || uid) : uid, what: '5 wrong PINs in 10 minutes for this person on this device' });
    if (u && !p.getProperty('PV_' + uid)) return { ok: false, noProof: true, error: 'This person\'s PIN has not reached the cloud yet.' };
    return { ok: false, error: 'The cloud did not accept this PIN.' };
  }
  delete fails['u:' + uid]; delete fails['d:' + who.deviceId];
  p.setProperty('LOGIN_FAILS', JSON.stringify(fails));
  const token = Utilities.getUuid() + Utilities.getUuid();
  const all = sessionsAll_();
  const until = now + 12 * 3600000;
  all[sha_(token)] = { uid: uid, device: who.deviceId, exp: until };
  sessionsSave_(all);
  return { ok: true, session: token, role: String(u.role || ''), until: new Date(until).toISOString() };
}
// The person behind this request (from its session), with the role they have in the cloud NOW.
function sessionUser_(body, who) {
  if (!body || !body.session || !who) return null;
  const s = sessionsAll_()[sha_(String(body.session))];
  if (!s || !(s.exp > Date.now()) || s.device !== who.deviceId) return null;
  const u = cloudUser_(s.uid);
  if (!u || String(u.active) === '0') return null;
  return { uid: s.uid, role: String(u.role || ''), name: String(u.name || '') };
}
function isBoss_(who) { return !!(who && who.user && BOSS_ROLES.indexOf(who.user.role) >= 0); }
// Owner / admin work: an owner / admin session on this device. During the 7 days for old apps, an old app on an admin
// device (one that does not sign yet) is still let through.
function needOwner_(who, body, what) {
  if (isBoss_(who)) return null;
  if (transition_() && who && who.admin && !body.sig) return null;
  log_('Refused', [who ? who.deviceId : '', what, who && who.user ? who.user.role : 'no owner login']);
  return { ok: false, notAdmin: true, needLogin: true, error: what + ' needs the owner or admin to be logged in on this device (with the internet on).' };
}
function needAdmin_(who, what) { return needOwner_(who, {}, what); }

// Why a pushed change is refused (null = allowed). boss: an owner / admin session; old: an old app during the 7 days.
function lockedChange_(table, key, row, haveRow, boss, oldAdmin) {
  if (boss) return null;
  let before = null;
  if (haveRow) { try { before = JSON.parse(haveRow[4] || 'null'); } catch (e) { before = null; } }
  const same = function (a, b) { return String(a === undefined || a === null ? '' : a) === String(b === undefined || b === null ? '' : b); };
  const changed = function (cols) {
    if (!before) {
      if (table === 'gl_accounts' && String(row.system) === '1') return false; // the app's own standard accounts
      return cols === '*' ? true : cols.some(function (c) { return row[c] !== undefined && row[c] !== null && String(row[c]) !== '' && String(row[c]) !== '0'; });
    }
    const list = cols === '*' ? Object.keys(row).filter(function (c) { return ['updated_at', 'updated_device', '_vby'].indexOf(c) < 0; }) : cols;
    return list.some(function (c) { return row[c] !== undefined && !same(row[c], before[c]); });
  };
  const LOGIN = 'needs the owner or admin to be logged in on the device that made it';
  if (table === 'users') {
    if (!before) return PIN_TOP_ROLES.indexOf(String(row.role || '')) >= 0 && !oldAdmin ? 'Adding an owner, admin, manager or payroll login ' + LOGIN + '.' : null;
    const cols = USER_LOCKED_COLS.filter(function (c) {
      if (row[c] === undefined) return false;
      if (c === 'pin_hash') return !!row[c] && !same(unlockHash_(key, row[c]), unlockHash_(key, before[c]));
      return !same(row[c], before[c]);
    });
    if (!cols.length) return null;
    return oldAdmin && cols.every(function (c) { return c !== 'role' || BOSS_ROLES.indexOf(String(row.role)) < 0; }) ? null : 'Changing ' + cols.join(', ') + ' ' + LOGIN + '.';
  }
  if (oldAdmin) return null;
  if (table === 'shared_settings') {
    if (LOCKED_SHARED.indexOf(String(key)) < 0) return null;
    if (before && same(row.value, before.value)) return null;
    const lf = LOCKED_FIELDS[String(key)];
    if (lf) {
      const P = function (v) { try { const x = typeof v === 'string' ? JSON.parse(v) : v; return x && typeof x === 'object' ? x : {}; } catch (e) { return {}; } };
      const a = P(before ? before.value : null); const b = P(row.value);
      if (lf.every(function (k) { return same(a[k], b[k]); })) return null;
    }
    return 'This setting is locked – it ' + LOGIN + '.';
  }
  const spec = LOCKED_TABLES[table];
  if (!spec) return null;
  if (spec === 'opening') return String(row.kind || (before && before.kind) || '') === 'opening' && changed(['lines', 'jdate', 'status', 'kind']) ? 'Opening balances are locked – they ' + LOGIN + '.' : null;
  return changed(spec) ? 'This is locked (' + table + ') – it ' + LOGIN + '.' : null;
}

// 1.0.0 build 1 (hardening, audit H2 / abuse C5): a money record dated in a closed month (on or before gl_lock) can
// only be changed by an owner / admin session – the app refuses it too, this stops a phone that skips the app.
const PERIOD_TABLES = /^(cash_moves|ap_invoices|ap_payments|ap_allocations|bank_entries|bank_lines|bank_closings|gl_journals|till_counts|till_shifts|till_shortages|till_surplus|aronium_eods|deposits|safe_counts|pay_runs|pay_lines|staff_charges|staff_debts|credit_settlements|wallet_moves|wallet_lines|loan_entries|loan_recons|owner_moves|cig_tallies|stat_payments|tax_returns|fixed_assets|cheques|regular_receipts|inv_moves|inv_receipts)$/;
function rowDay_(r) {
  if (!r) return '';
  const k = ['move_date', 'invoice_date', 'pay_date', 'entry_date', 'jdate', 'shift_date', 'count_date', 'eod_date', 'z_date', 'line_date', 'run_date', 'deposit_date', 'charge_date', 'debt_date', 'settle_date', 'receipt_date', 'cheque_date', 'disposed_date', 'period_end', 'date'];
  for (let i = 0; i < k.length; i++) if (r[k[i]] && /^\d{4}-\d{2}-\d{2}/.test(String(r[k[i]]))) return String(r[k[i]]).slice(0, 10);
  return '';
}
function periodChange_(table, row, haveRow, boss, lockDay) {
  if (boss || !lockDay || !PERIOD_TABLES.test(table)) return null;
  let before = null;
  if (haveRow) { try { before = JSON.parse(haveRow[4] || 'null'); } catch (e) { before = null; } }
  const d1 = rowDay_(row); const d0 = rowDay_(before);
  if ((d1 && d1 <= lockDay) || (d0 && d0 <= lockDay)) return 'The books are closed up to ' + lockDay + ' – a record dated in a closed month needs the owner or admin to be logged in.';
  return null;
}
// 1.0.0 build 1 (hardening, abuse C10): a person tied to one shop (staff, supervisor, baker, shop manager) only
// changes that shop's records. Owner / admin, and people with no shop (payroll …), are not limited here.
function shopChange_(table, row, haveRow, boss, who) {
  if (boss || !who || !who.user || table === 'users' || table === 'shared_settings' || table === 'audit_log') return null;
  const u = cloudUser_(who.user.uid);
  const mine = u && u.shop_code ? String(u.shop_code) : '';
  if (!mine) return null;
  let before = null;
  if (haveRow) { try { before = JSON.parse(haveRow[4] || 'null'); } catch (e) { before = null; } }
  const s1 = row && row.shop_code ? String(row.shop_code) : '';
  const s0 = before && before.shop_code ? String(before.shop_code) : '';
  if ((s1 && s1 !== mine) || (s0 && s0 !== mine)) return 'This record belongs to the other shop – only that shop, the owner or admin can change it.';
  return null;
}

// ---------------------------------------------------------------- email logins (for the online version later)
// Off until the owner switches "Email logins" on (locked setting email_logins). A login needs the person's email, the
// password the owner set for them, and a 6-digit code sent to that email (5 minutes, 5 tries). Every answer is the same
// whether the email exists or not; tries are limited per email and per caller (kept in the cache, not the settings).
function emailOn_() { const v = sharedSetting_('email_logins'); return v === true || (v && v.on === true); }
function caller_(body) { let k = ''; try { k = Session.getTemporaryActiveUserKey() || ''; } catch (e) { k = ''; } return sha_(k || String((body.device && body.device.id) || 'anyone')).slice(0, 20); }
function userByEmail_(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) return null;
  const sh = data_().getSheetByName('T_users');
  if (!sh) return null;
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) { try { const u = JSON.parse(rows[i][4] || '{}'); if (String(u.email || '').trim().toLowerCase() === e) { u.uid = String(rows[i][0]); return u; } } catch (x) { /* next */ } }
  return null;
}
function limited_(cache, key, max, secs) { const n = Number(cache.get(key) || '0'); if (n >= max) return true; cache.put(key, String(n + 1), secs); return false; }
const EMAIL_SAME = 'If this email may log in, a code has been sent to it.';
const EMAIL_FAIL = 'The email, password or code is not right (or the code is too old). Try again.';
function emailLoginStart_(body) {
  const cache = CacheService.getScriptCache();
  const e = String(body.email || '').trim().toLowerCase();
  if (limited_(cache, 'el_c_' + caller_(body), 10, 3600) || limited_(cache, 'el_e_' + sha_(e).slice(0, 20), 3, 900)) return { ok: true, message: EMAIL_SAME };
  if (!emailOn_()) return { ok: true, message: EMAIL_SAME };
  const u = userByEmail_(e);
  if (!u || String(u.active) === '0' || !props_().getProperty('PW_' + u.uid)) return { ok: true, message: EMAIL_SAME };
  const code = String(100000 + Math.floor(Math.random() * 900000));
  cache.put('el_code_' + sha_(e).slice(0, 20), JSON.stringify({ h: sha_(secret_('PIN_PEPPER') + ':' + code), exp: Date.now() + 5 * 60000, tries: 0 }), 300);
  mailEssential_({ to: e, subject: 'Easy Buys login code', body: 'Your Easy Buys login code is ' + code + '. It works for 5 minutes. If you did not ask for it, tell the owner.', name: 'Easy Buys' });
  return { ok: true, message: EMAIL_SAME };
}
function emailLoginFinish_(body) {
  const cache = CacheService.getScriptCache();
  const e = String(body.email || '').trim().toLowerCase();
  const ck = 'el_code_' + sha_(e).slice(0, 20);
  if (limited_(cache, 'el_f_' + caller_(body), 20, 3600)) return { ok: false, error: EMAIL_FAIL };
  if (!emailOn_()) return { ok: false, error: EMAIL_FAIL };
  let c = null;
  try { c = JSON.parse(cache.get(ck) || 'null'); } catch (x) { c = null; }
  if (!c || !(c.exp > Date.now()) || c.tries >= 5) { failCount_('email', e, { person: e, what: '5 wrong email logins in 10 minutes' }); return { ok: false, error: EMAIL_FAIL }; }
  const u = userByEmail_(e);
  const pw = u ? props_().getProperty('PW_' + u.uid) : null;
  const codeOk = c.h === sha_(secret_('PIN_PEPPER') + ':' + String(body.code || '').trim());
  let passOk = false;
  if (pw) { const parts = pw.split(':'); passOk = parts.length === 3 && pbkdf2_(secret_('PIN_PEPPER') + ':' + String(body.password || ''), parts[1], Number(parts[0]) || PIN_ITER) === parts[2]; }
  if (!u || String(u.active) === '0' || !codeOk || !passOk) {
    failCount_('email', e, { person: e, what: '5 wrong email logins in 10 minutes' });
    c.tries += 1;
    if (c.tries >= 5) cache.put(ck, 'null', 1); else cache.put(ck, JSON.stringify(c), Math.max(1, Math.round((c.exp - Date.now()) / 1000)));
    return { ok: false, error: EMAIL_FAIL };
  }
  cache.put(ck, 'null', 1);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const all = sessionsAll_();
  const until = Date.now() + 12 * 3600000;
  all[sha_(token)] = { uid: u.uid, device: 'web:' + caller_(body), exp: until };
  sessionsSave_(all);
  log_('Email login', [u.uid, caller_(body)]);
  return { ok: true, session: token, role: String(u.role || ''), until: new Date(until).toISOString() };
}
// The owner sets (or resets) a person's email-login password: the owner logged in AND their PIN typed again (proof).
function emailPasswordSet_(body, who) {
  if (!who.user || who.user.role !== 'owner' || !proofOk_(who.user.uid, body.ownerProof)) return { ok: false, error: 'The owner must be logged in and type their PIN again.' };
  const u = cloudUser_(String(body.uid || ''));
  if (!u) return { ok: false, error: 'This person has not reached the cloud yet.' };
  const pw = String(body.password || '');
  if (pw.length < 10) return { ok: false, error: 'The password must be at least 10 characters.' };
  const salt = sha_(Utilities.getUuid()).slice(0, 16);
  const iter = PIN_ITER * 5;
  props_().setProperty('PW_' + u.uid, iter + ':' + salt + ':' + pbkdf2_(secret_('PIN_PEPPER') + ':' + pw, salt, iter));
  log_('Email login', [u.uid, 'password set by the owner', who.deviceId]);
  return { ok: true };
}

// ---------------------------------------------------------------- security alerts to the owner's email
// Every security event goes into the SecurityLog sheet (add only, safe cells) and – at once, at most one email per kind
// every 10 minutes (with how many more happened) – to the owner's address: Script properties › SECURITY_EMAIL, else the
// support address (the owner's). A daily summary goes at 7:00 am Fiji time (securityDigest, its trigger is made by setup()).
// A free Gmail account sends about 100 emails a day: 20 are always kept for the daily summary, crash reports and login
// codes; when fewer are left, alerts (and reports, orders, packs) wait and are told in the next summary instead.
const SEC_RESERVE = 20;
const SEC_KINDS = {
  refused: 'A request was refused', login_fail: 'Many wrong logins', invite_guess: 'Someone is guessing invite codes', link_key: 'A till PC key was misused',
  link_event: 'Aronium Link warning', locked_change: 'A locked setting was changed', device: 'A device was added, removed, blocked or wiped',
  clock: 'A device clock or time zone is wrong (not Fiji time, or more than 5 minutes out) – money work is stopped on it', backup: 'A backup was downloaded', clean_start: 'Start clean was used',
};
const SEC_TODO = {
  refused: 'Nothing was changed. If you do not know this device, block it in Admin › Devices.',
  login_fail: 'If it was not you or your staff, check who is near the tablet and change the PINs (Staff › the person › PIN).',
  invite_guess: 'Joining is paused for an hour by itself. If it goes on, tell support.',
  link_key: 'Nothing was accepted. Make a new key for that till PC (Admin › Aronium link) and check who used the PC.',
  link_event: 'Check the till PC: the Aronium Link must keep running and read the real Aronium database.',
  locked_change: 'If you did not make or allow this change, change your PIN and look at Admin › Change history.',
  device: 'If you did not do this, block the device in Admin › Devices at once.',
  clock: 'Check the date, time and time zone on that device (Settings › Date & time › automatic).',
  backup: 'If you did not download it, change your PIN and check who can log in as owner or admin.',
  clean_start: 'If you did not do this, call support now – the copy of everything is in Drive › Backups › Clean start.',
};
function fijiNow_() { return Utilities.formatDate(new Date(), 'Pacific/Fiji', 'yyyy-MM-dd HH:mm'); }
function secEmail_() { return props_().getProperty('SECURITY_EMAIL') || SUPPORT_EMAIL; }
function quotaLeft_() { try { return MailApp.getRemainingDailyQuota(); } catch (e) { return 0; } }
// Every email of the cloud goes through here. essential: login codes, crash reports and the daily summary (they may use
// the reserve). Returns true when sent; false when it had to wait (low quota) – noted for the summary.
function mailEssential_(opts) { return mail_(opts, true); }
function mail_(opts, essential) {
  const left = quotaLeft_();
  if (left <= 0 || (!essential && left <= SEC_RESERVE)) {
    const p = props_();
    const n = Number(p.getProperty('MAIL_HELD') || '0') + 1;
    p.setProperty('MAIL_HELD', String(n));
    try { sheet_('SecurityLog', SECLOG_HEADER).appendRow([fijiNow_(), 'mail_held', '', '', '', esc_('Not sent (daily email limit): ' + String(opts.subject || '').slice(0, 120)), '']); } catch (e) { /* next time */ }
    return false;
  }
  MailApp.sendEmail(opts);
  return true;
}
const SECLOG_HEADER = ['at_fiji', 'kind', 'device', 'person', 'shop', 'what', 'todo'];
function secEvent_(kind, info) {
  info = info || {};
  if (REQ_.locked) { later_(function () { secEvent_(kind, info); }); return false; } // 1.0.0 build 10: never inside the lock
  try {
    const row = [fijiNow_(), kind, esc_(String(info.device || '')), esc_(String(info.person || '')), esc_(String(info.shop || '')), esc_(String(info.what || '').slice(0, 1500)), esc_(SEC_TODO[kind] || '')];
    sheet_('SecurityLog', SECLOG_HEADER).appendRow(row);
  } catch (e) { /* the log never stops a request */ }
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('sec_last_' + kind)) { cache.put('sec_cnt_' + kind, String(Number(cache.get('sec_cnt_' + kind) || '0') + 1), 3600); return false; }
    const more = Number(cache.get('sec_cnt_' + kind) || '0');
    const body = [
      'Easy Buys security alert – ' + (SEC_KINDS[kind] || kind),
      '',
      'When (Fiji time): ' + fijiNow_(),
      info.device ? 'Device: ' + info.device : '',
      info.person ? 'Person: ' + info.person : '',
      info.shop ? 'Shop: ' + info.shop : '',
      'What happened: ' + String(info.what || ''),
      info.refused === false ? '' : 'What the cloud did: refused it – nothing was changed.',
      more ? 'The same happened ' + more + ' more time(s) since the last email.' : '',
      '',
      'What to do: ' + (SEC_TODO[kind] || 'Look at Admin › Change history.'),
    ].filter(function (x) { return x !== ''; }).join('\n');
    const sent = mail_({ to: secEmail_(), subject: 'Easy Buys security: ' + (SEC_KINDS[kind] || kind), body: body, name: 'Easy Buys' }, false);
    if (sent) { cache.put('sec_last_' + kind, '1', 600); cache.put('sec_cnt_' + kind, '0', 3600); }
    return sent;
  } catch (e) { return false; }
}
// Counted problems that become an alert at 5 in 10 minutes (wrong PINs, email logins, dashboard passwords).
function failCount_(kind, key, info) {
  const cache = CacheService.getScriptCache();
  const k = 'secf_' + kind + '_' + sha_(String(key)).slice(0, 12);
  const n = Number(cache.get(k) || '0') + 1;
  cache.put(k, String(n), 600);
  if (n === 5) secEvent_('login_fail', info);
}
// The daily summary (7:00 am Fiji): every security event and crash report of the last 24 hours, or "all quiet".
function securityDigestRun_() {
  const since = Utilities.formatDate(new Date(Date.now() - 86400000), 'Pacific/Fiji', 'yyyy-MM-dd HH:mm');
  const ss = data_();
  const sec = ss.getSheetByName('SecurityLog') ? values_(ss.getSheetByName('SecurityLog')).filter(function (r) { return String(r[0]) >= since; }) : [];
  const crashes = ss.getSheetByName('Problem logs') ? values_(ss.getSheetByName('Problem logs')).filter(function (r) { return String(r[0]).replace('T', ' ').slice(0, 16) >= since; }) : [];
  const held = Number(props_().getProperty('MAIL_HELD') || '0');
  const lines = ['Easy Buys – security and problems in the last 24 hours (to ' + fijiNow_() + ' Fiji time)', ''];
  if (!sec.length && !crashes.length && !held) lines.push('All quiet: no security events and no crash reports.');
  if (sec.length) {
    lines.push('Security events (' + sec.length + '):');
    sec.slice(-60).forEach(function (r) { lines.push('• ' + r[0] + ' – ' + (SEC_KINDS[r[1]] || r[1]) + (r[2] ? ' – device ' + r[2] : '') + (r[3] ? ' – ' + r[3] : '') + ': ' + String(r[5]).slice(0, 200)); });
  }
  if (crashes.length) { lines.push(''); lines.push('Problem reports from devices (' + crashes.length + '):'); crashes.slice(-20).forEach(function (r) { lines.push('• ' + String(r[0]).slice(0, 16) + ' – device ' + r[1] + ': ' + String(r[2]).split('\n')[0].slice(0, 160)); }); }
  if (held) lines.push('', held + ' email(s) were held back because the daily email limit was nearly used up – they are listed above or in the SecurityLog sheet.');
  const sent = mail_({ to: secEmail_(), subject: 'Easy Buys daily security summary', body: lines.join('\n'), name: 'Easy Buys' }, true);
  if (sent) props_().deleteProperty('MAIL_HELD');
  return { ok: sent, events: sec.length, crashes: crashes.length };
}
function securityDigestTrigger_() {
  try {
    if (typeof ScriptApp === 'undefined') return false;
    const have = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'securityDigest'; });
    if (!have) ScriptApp.newTrigger('securityDigest').timeBased().atHour(7).everyDays(1).inTimezone('Pacific/Fiji').create();
    return true;
  } catch (e) { return false; }
}

// ---------------------------------------------------------------- the change history: add only, with a fingerprint
function historyChain_(entries) {
  if (!entries.length) return;
  const p = props_();
  let fp = p.getProperty('HIST_FP') || '';
  const sh = sheet_('History chain', ['seq', 'key', 'fp', 'at']);
  const rows = entries.map(function (e) { fp = sha_(fp + '|' + e.key + '|' + e.json); return [String(e.seq), esc_(e.key), fp, now_()]; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  p.setProperty('HIST_FP', fp);
}
// Run from the editor: is the history unbroken? (every entry's fingerprint follows from the one before)
function checkHistoryRun_() {
  const ss = data_();
  const t = ss.getSheetByName('T_audit_log'); const c = ss.getSheetByName('History chain');
  if (!t || !c) return { ok: true, entries: 0 };
  const byKey = {}; values_(t).forEach(function (r) { byKey[String(r[0])] = String(r[4]); });
  let fp = ''; let bad = 0;
  values_(c).forEach(function (r) { fp = sha_(fp + '|' + r[1] + '|' + (byKey[String(r[1])] || '')); if (fp !== String(r[2])) bad++; });
  Logger.log(bad ? 'History CHANGED in ' + bad + ' place(s)' : 'History unbroken');
  return { ok: !bad, bad: bad };
}

// ---------------------------------------------------------------- safe cells (no formulas from devices)
function esc_(v) { const s = v === null || v === undefined ? '' : String(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function unesc_(s) { return /^'[=+\-@]/.test(s) ? s.slice(1) : s; }

// What a device may receive of a record (owner / admin PIN hashes, bank details and face data stay with admin devices).
// The fields a device that is not an admin device does not get (and whose empty copy never wipes the cloud's: push_).
const MASKED_TABLES = ['pay_profiles', 'ap_suppliers', 'staff_details'];
const MASKED_FIELD = /bank|account_no|tin$/;
function forDevice_(table, row, key, who) {
  if (!row || who.forMain) return row;
  const admin = !!who.admin;
  if (table === 'users') {
    const out = {}; Object.keys(row).forEach(function (k) { out[k] = row[k]; });
    delete out.pin;
    const top = BOSS_ROLES.indexOf(String(out.role || '')) >= 0;
    // 1.0.0 build 1 (hardening, abuse C1): a non-admin device set to a shop only gets the PIN hashes of that shop's
    // people (and people with no shop); a device with no shop yet gets the staff hashes so they can log in offline.
    const devShop = admin ? '' : String((ctlAll_()[who.deviceId] || {}).shop || '');
    const otherShop = !!devShop && !!out.shop_code && String(out.shop_code) !== devShop;
    out.pin_hash = (admin || (!top && !otherShop)) && out.pin_hash ? unlockHash_(key, out.pin_hash) : null;
    if (!admin) { if (!devShop || String(out.shop_code || '') !== devShop) { out.face = null; } }
    return out;
  }
  if (!admin && MASKED_TABLES.indexOf(table) >= 0) {
    const out = {}; Object.keys(row).forEach(function (k) { out[k] = MASKED_FIELD.test(k) && k !== 'on_account' ? null : row[k]; });
    return out;
  }
  return row;
}

// ---------------------------------------------------------------- per-till keys for the Aronium Link (C3)
// The owner makes one key per till PC (Admin › Aronium link): EBL1.<till>.<secret>. It may only send that till's
// aronium_feed records and run that till's link check – nothing else, nothing to read.
function linkKeys_() { try { return JSON.parse(props_().getProperty('LINK_KEYS') || '{}'); } catch (e) { return {}; } }
function linkKeyCreate_(body, who) {
  const src = String(body.source || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 20);
  if (!/^S\d+-T\d+$/.test(src)) return { ok: false, error: 'Which till? (e.g. S1-T1)' };
  const key = 'EBL1.' + src + '.' + sha_(Utilities.getUuid() + Utilities.getUuid()).slice(0, 48);
  const all = linkKeys_();
  Object.keys(all).forEach(function (h) { if (all[h].source === src) delete all[h]; }); // a new key replaces the old one
  all[sha_(key)] = { source: src, at: now_(), by: who.user ? who.user.name : who.deviceId };
  props_().setProperty('LINK_KEYS', JSON.stringify(all));
  log_('Link keys', [src, 'new key', who.deviceId]);
  return { ok: true, key: key, source: src };
}
function linkApi_(body) {
  // the Link sends { kind, source, till_key, records } (stream D); linkKey is the same key under its first name
  const k = linkKeys_()[sha_(String(body.till_key || body.linkKey || ''))];
  if (!k || (body.source && String(body.source) !== k.source)) { secEvent_('link_key', { device: 'till PC ' + String(body.source || '?'), what: 'A till PC request with a wrong or old key' }); return { ok: false, error: 'Not allowed' }; }
  if (body.kind === 'link_check') {
    // 1.0.0 build 5: the Link asks every ~30 s (pull_only: the cheap answer) and says when it has answered a pull
    if (body.pulled) pullDone_(k.source);
    const pull = pullOpen_(k.source);
    if (body.pull_only) return { ok: true, version: SCRIPT_VERSION, pull: pull };
    const chk = linkCheck_({ source: k.source });
    chk.pull = pull;
    return chk;
  }
  const own = function (r) {
    if (!r || r.t !== 'aronium_feed' || String(r.k).indexOf(k.source + '|') !== 0 || !r.row) return false;
    if (r.row.source !== undefined && String(r.row.source) !== k.source) return false;
    if (r.row.uid !== undefined && String(r.row.uid).indexOf(k.source + '|') !== 0) return false;
    return JSON.stringify(r.row).length < 60000; // a receipts part is up to ~36 KB
  };
  const recs = (body.records || []).filter(own);
  if (recs.length !== (body.records || []).length) { secEvent_('link_key', { device: 'till PC ' + k.source, what: 'The till PC key tried to send records that are not its own Aronium figures' }); return { ok: false, error: 'This till may only send its own Aronium records.' }; }
  const res = push_(recs, 'link-' + k.source, { deviceId: 'link-' + k.source, link: true });
  if (res && res.ok) { if (body.full) pullDone_(k.source); res.pull = pullOpen_(k.source); }
  return res;
}

// ---------------------------------------------------------------- 1.0.0 build 5: "Get the End of day now" (owner, 2 Nov)
// The owner / admin (any shop) or a manager / supervisor (own shop) asks that shop's till PCs to send now instead of
// waiting for their next round. The request is kept in Script properties (and the script cache, so the till PCs' checks
// every 30 s cost almost nothing): { S1: { ms, at, by, device, done: { 'S1-T1': time } } }. Each till PC is told pull: true
// until it has sent after the request (a full send, or its "pulled" answer). One request per shop every 30 s; a request
// is forgotten after 10 minutes. Nothing from the request goes into a cell.
const PULL_ROLES = ['owner', 'admin', 'manager', 'supervisor'];
const PULL_GAP_MS = 30000;
const PULL_OPEN_MS = 10 * 60000;
function pullCache_() { try { return typeof CacheService !== 'undefined' ? CacheService.getScriptCache() : null; } catch (e) { return null; } }
function pullAll_() {
  const c = pullCache_();
  let txt = null;
  try { txt = c ? c.get('LINK_PULLS') : null; } catch (e) { txt = null; }
  if (txt === null || txt === undefined) txt = props_().getProperty('LINK_PULLS') || '{}';
  try { const o = JSON.parse(txt); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; }
}
function pullSave_(all) {
  const txt = JSON.stringify(all);
  props_().setProperty('LINK_PULLS', txt);
  const c = pullCache_();
  try { if (c) c.put('LINK_PULLS', txt, 600); } catch (e) { /* the properties copy is read instead */ }
}
function pullTills_(shop) {
  const ks = linkKeys_();
  const out = [];
  Object.keys(ks).forEach(function (h) { const s = String(ks[h].source || ''); if (s.indexOf(shop + '-') === 0 && out.indexOf(s) < 0) out.push(s); });
  return out.sort();
}
function pullView_(shop, p) {
  return { ok: true, shopCode: shop, at: p ? p.at : null, by: p ? p.by : '', tills: pullTills_(shop), answered: p ? Object.keys(p.done || {}).sort() : [] };
}
function linkPullRequest_(body, who) {
  const shop = String(body.shopCode || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 8);
  if (!/^S\d+$/.test(shop)) return { ok: false, error: 'Which shop?' };
  const role = who && who.user ? who.user.role : '';
  const u = role ? cloudUser_(who.user.uid) : null;
  if (!u || PULL_ROLES.indexOf(role) < 0) {
    log_('Refused', [who ? who.deviceId : '', 'link pull ' + shop, role || 'no login']);
    return { ok: false, needLogin: !role, error: 'Only the owner, an admin, a manager or a supervisor can ask the till PCs (logged in, with the internet on).' };
  }
  if (BOSS_ROLES.indexOf(role) < 0 && String(u.shop_code || '') !== shop) return { ok: false, error: 'You can only ask your own shop\'s till PCs.' };
  return withLock_(function () {
    const all = pullAll_();
    const t = Date.now();
    Object.keys(all).forEach(function (s) { if (!(t - Number(all[s] && all[s].ms) < PULL_OPEN_MS)) delete all[s]; });
    const p = all[shop];
    if (body.check) return pullView_(shop, p);
    if (p && t - Number(p.ms) < PULL_GAP_MS) {
      const v = pullView_(shop, p);
      v.already = true; v.wait = Math.ceil((PULL_GAP_MS - (t - Number(p.ms))) / 1000);
      return v;
    }
    all[shop] = { ms: t, at: now_(), by: esc_(String(u.name || who.user.name || '').slice(0, 60)), device: esc_(String(who.deviceId || '').slice(0, 60)), done: {} };
    pullSave_(all);
    log_('Link pull', [shop, all[shop].by, all[shop].device]);
    return pullView_(shop, all[shop]);
  });
}
// A till PC: is there a request for its shop it has not answered yet?
function pullOpen_(source) {
  const p = pullAll_()[String(source).split('-')[0]];
  return !!p && Date.now() - Number(p.ms) < PULL_OPEN_MS && !(p.done || {})[source];
}
function pullDone_(source) {
  if (!pullOpen_(source)) return;
  withLock_(function () {
    const all = pullAll_();
    const p = all[String(source).split('-')[0]];
    if (!p) return;
    p.done = p.done || {};
    p.done[source] = now_();
    pullSave_(all);
  });
}
// Run from the editor after every till PC has its own key: a new main-software key (the old one stops working).
function rotateMainKey() { props_().setProperty('MAIN_KEY', Utilities.getUuid() + Utilities.getUuid()); Logger.log('New main-software key (keep private): ' + props_().getProperty('MAIN_KEY')); }

// ---------------------------------------------------------------- Changes sheet: compact (keep the newest entry per record)
// Run from the editor or a monthly trigger. Pull positions are change numbers (seq), so devices carry on where they were.
function compactChangesRun_() {
  return withLock_(function () {
    const sh = sheet_('Changes', CHANGES_HEADER);
    const last = sh.getLastRow();
    if (last < 3) return 0;
    const rows = sh.getRange(2, 1, last - 1, CHANGES_HEADER.length).getValues().map(function (r) { return r.map(text_); });
    const newest = {};
    rows.forEach(function (r, i) { newest[r[1] + '|' + r[2]] = i; });
    const keep = rows.filter(function (r, i) { return newest[r[1] + '|' + r[2]] === i; });
    sh.getRange(2, 1, rows.length, CHANGES_HEADER.length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, CHANGES_HEADER.length).setValues(keep);
    log_('Changes compacted', [rows.length, keep.length]);
    return rows.length - keep.length;
  });
}

// The supplier set-up pack (an Excel file the owner keeps in Drive – never in the app's code): the file whose id is in
// Script properties › SETUP_PACK_FILE_ID, or else the file with this name. → { ok, name, mime, size, b64 }
const SETUP_PACK_NAME = 'Easy Buys set-up pack – suppliers from invoices.xlsx';
function setupPack_() {
  const props = PropertiesService.getScriptProperties();
  let file = null;
  const id = props.getProperty('SETUP_PACK_FILE_ID');
  if (id) { try { file = DriveApp.getFileById(id); if (!file || !file.getName || (file.isTrashed && file.isTrashed())) file = null; } catch (e) { file = null; } }
  if (!file) {
    const it = DriveApp.getFilesByName(SETUP_PACK_NAME);
    while (it.hasNext()) { const f = it.next(); if (!(f.isTrashed && f.isTrashed())) { file = f; break; } }
  }
  if (!file) return { ok: false, code: 'NO_PACK', error: 'The supplier set-up pack is not in Google Drive. Upload "' + SETUP_PACK_NAME + '" to the Easy Buys Google Drive (or set SETUP_PACK_FILE_ID in Script properties).' };
  const mime = String(file.getMimeType ? file.getMimeType() || '' : '');
  if (mime === 'application/vnd.google-apps.spreadsheet') return { ok: false, code: 'NOT_XLSX', error: 'The set-up pack was turned into a Google Sheet. Upload it again as the .xlsx file (Drive › Settings › untick "Convert uploads").' };
  const bytes = file.getBlob().getBytes();
  if (!bytes || bytes.length < 100) return { ok: false, code: 'EMPTY', error: 'The set-up pack in Drive is empty.' };
  return { ok: true, name: file.getName(), mime: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', size: bytes.length, b64: Utilities.base64Encode(bytes) };
}

// Start clean (owner): the owner's PIN is checked HERE against the cloud copy (pinHash = p1:… of the PIN typed), a copy of
// the whole data spreadsheet is kept in Drive › Backups › Clean start, then every table is cleared (the owner logins stay
// so the owner can still sign in) and every device clears itself on its next sync (the fresh start marker).
function cleanStart_(body, who) {
  // The owner, logged in (session), types the PIN again: the proof of that PIN is checked here (it is kept on no device).
  const uid = String(body.ownerUid || '');
  if (!uid || !/^[0-9a-f]{64}$/.test(String(body.proof || ''))) return { ok: false, error: 'The owner\'s PIN is needed again.' };
  const owner = cloudUser_(uid);
  if (!owner || owner.role !== 'owner' || String(owner.active) === '0' || (who.user && who.user.uid !== uid)) return { ok: false, error: 'Only the owner can start clean.' };
  if (!props_().getProperty('PV_' + uid)) return { ok: false, error: 'The owner\'s PIN has not reached the cloud yet. Sync first, then try again.' };
  if (!proofOk_(uid, body.proof)) { log_('Refused', [who.deviceId, 'clean start', 'wrong owner PIN']); return { ok: false, error: 'Wrong PIN – nothing was cleared.' }; }
  const at = String(body.at || now_());
  const id = String(body.id || ('CLEAN-' + Utilities.getUuid()));
  const ss = data_();
  let copyId = null;
  try {
    const folder = folder_(path_(['Backups']), 'Clean start');
    copyId = DriveApp.getFileById(ss.getId()).makeCopy('Easy Buys data before Start clean ' + at.slice(0, 19).replace(/[: ]/g, '-'), folder).getId();
  } catch (e) { return { ok: false, error: 'The cloud copy could not be made – nothing was cleared. ' + String(e && e.message || e) }; }
  // every table except the change history (it only ever grows; the copy above holds it too)
  const tables = ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return /^T_[a-z_0-9]+$/.test(n) && n !== 'T_audit_log'; }).map(function (n) { return n.slice(2); });
  const r = freshStart_({ key: apiKey_(), id: id, at: at, tables: tables, moveFiles: false });
  if (!r.ok) return r;
  const old = oldFilesAway_(at);
  PropertiesService.getScriptProperties().setProperty('CLEAN_START', JSON.stringify({ id: id, at: at, by: String(body.by || ''), device: who.deviceId, copyId: copyId, oldFolder: old.name }));
  log_('Start clean', [id, at, who.deviceId, String(body.by || ''), copyId, r.cleared, old.moved]);
  secEvent_('clean_start', { device: who.deviceId, person: who.user ? who.user.name : String(body.by || ''), what: 'Start clean: ' + r.cleared + ' records cleared, ' + old.moved + ' Drive folders moved to "' + old.name + '"', refused: false });
  return { ok: true, id: id, at: at, cleared: r.cleared, copyId: copyId, tables: tables.length, foldersMoved: old.moved, oldFolder: old.name };
}
// Start clean never deletes a Drive file: every Easy Buys folder in the cloud folder (Photos, Documents, Shift reports,
// Purchase orders, Viber posts…) except Backups moves into ONE folder "Before 1.0.0 – old and test data <date>". Nothing is
// shared or unshared (the files stay private); the owner can open or delete them later. The app starts with new, empty
// folders (they are made again when first needed).
function oldFilesAway_(at) {
  const root = root_();
  const name = 'Before 1.0.0 – old and test data ' + String(at).slice(0, 10);
  const keep = ['Backups', name];
  const names = [];
  const it0 = root.getFolders ? root.getFolders() : null;
  if (it0) while (it0.hasNext()) { const f = it0.next(); if (keep.indexOf(f.getName()) < 0 && !/^Before 1\.0\.0 – old and test data /.test(f.getName())) names.push(f); }
  if (!names.length) return { moved: 0, name: name };
  const dest = folder_(root, name);
  names.forEach(function (f) { f.moveTo(dest); });
  return { moved: names.length, name: name };
}

// ---------------------------------------------------------------- 1.0.0 script 16: readable record photos in Drive
// Easy Buys › Records › "2026-10 October" › "Shop 1 - Drasa Avenue" (or "Business") › Invoices | Cheques, and in the
// month folder itself: Deposits, Z reports, Payslips. The same names as the app's src/services/photoNames.js (tested
// against each other in tests/photoNames.test.js).
const RECORD_FOLDERS = { invoice: 'Invoices', payment: 'Cheques', stub: 'Cheques', deposit: 'Deposits', till: 'Z reports' };
const RECORD_MONTH_LEVEL = { deposit: 1, till: 1, Deposits: 1, 'Z reports': 1, Payslips: 1, 'Purchase orders': 1, 'Shift reports': 1 };
const RECORD_TYPE_FOLDERS = ['Invoices', 'Cheques', 'Deposits', 'Z reports', 'Payslips', 'Purchase orders', 'Shift reports'];
const NAME_LIMIT = 115;
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
function cleanPart_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[<>:"/\\|?*\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '');
}
function monthFolderOf_(date) {
  const m = /^(\d{4})-(\d{2})-\d{2}/.exec(String(date || ''));
  if (!m || +m[2] < 1 || +m[2] > 12) return null;
  return m[1] + '-' + m[2] + ' ' + MONTH_NAMES[+m[2] - 1];
}
function shopFolderOf_(shop) {
  if (!shop) return 'Business';
  const name = cleanPart_(shop.name);
  const n = /^S(\d+)$/i.exec(String(shop.code || ''));
  if (/^shop\s*\d+/i.test(name)) return name;
  if (n) return name ? 'Shop ' + n[1] + ' - ' + name : 'Shop ' + n[1];
  return name || cleanPart_(shop.code) || 'Business';
}
function shopShortOf_(shop) {
  if (!shop) return '';
  const n = /^S(\d+)$/i.exec(String(shop.code || ''));
  if (n) return 'Shop ' + n[1];
  const m = /^shop\s*(\d+)/i.exec(cleanPart_(shop.name));
  return m ? 'Shop ' + m[1] : (cleanPart_(shop.name) || cleanPart_(shop.code));
}
function money_(v) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0) return '';
  const p = n.toFixed(2).split('.');
  return '$' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
}
function chequeNo_(v) { const s = cleanPart_(v); return /^\d{1,5}$/.test(s) ? ('000000' + s).slice(-6) : s; }
function invoiceNo_(v) { const s = cleanPart_(v); return !s ? '' : (/^inv/i.test(s) ? s : 'INV ' + s); }
// { refType, record, shop, chequeNo, page, ext } → { fileName, monthFolder, shopFolder, typeFolder } or null (old place).
function photoPlace_(o) {
  const refType = o.refType; const record = o.record; const shop = o.shop || null; const page = Number(o.page) || 1;
  const typeFolder = RECORD_FOLDERS[refType];
  if (!typeFolder || !record) return null;
  let date; let head = null; let party = ''; let tail = '';
  if (refType === 'invoice') {
    date = record.invoice_date || record.received_date; party = record.supplier_name;
    tail = record.kind === 'credit' ? 'Credit note' + (record.invoice_no ? ' ' + cleanPart_(record.invoice_no) : '') : (invoiceNo_(record.invoice_no) || 'Invoice');
  } else if (refType === 'payment') {
    date = record.pay_date; const c = chequeNo_(o.chequeNo || record.cheque_no);
    head = c ? 'Cheque ' + c : 'Payment'; party = record.supplier_name; tail = money_(record.amount);
  } else if (refType === 'deposit') {
    date = record.deposit_date || String(record.created_at || '').slice(0, 10); head = 'Deposit'; party = shopShortOf_(shop);
    tail = money_(record.total || (Number(record.cash_amount || 0) + Number(record.cheque_amount || 0)));
  } else if (refType === 'till') {
    date = record.shift_date || String(record.counted_at || '').slice(0, 10);
    const z = cleanPart_(record.z_number); head = z ? 'Z ' + z : 'Z report';
    const till = cleanPart_(record.till_name); const where = shopShortOf_(shop);
    party = /shop/i.test(till) || !where ? till : (where + ' ' + till).trim();
  } else {
    date = record.cheque_date; const c2 = chequeNo_(o.chequeNo || record.cheque_no);
    head = c2 ? 'Cheque ' + c2 : 'Cheque'; party = record.payee; tail = money_(record.amount);
  }
  const monthFolder = monthFolderOf_(date);
  if (!monthFolder) return null;
  const day = String(date).slice(0, 10);
  const suffix = page > 1 ? ' (page ' + page + ')' : '';
  let who = cleanPart_(party);
  const build = function () {
    const bits = [head ? day + ' ' + head : day];
    if (who) bits.push(who);
    if (tail) bits.push(cleanPart_(tail));
    const s = head ? bits.join(' - ') : (day + ' ' + bits.slice(1).join(' - ')).trim();
    return s + suffix;
  };
  let stem = build();
  if (stem.length > NAME_LIMIT && who) { who = who.slice(0, Math.max(1, who.length - (stem.length - NAME_LIMIT))).trim(); stem = build(); }
  if (stem.length > NAME_LIMIT) stem = stem.slice(0, NAME_LIMIT).trim();
  const e = /^(pdf|jpg|jpeg|png)$/i.test(String(o.ext || '')) ? String(o.ext).toLowerCase() : 'jpg';
  return { fileName: stem + '.' + e, monthFolder: monthFolder, shopFolder: RECORD_MONTH_LEVEL[refType] ? '' : shopFolderOf_(shop), typeFolder: typeFolder };
}
// The folder and the cleaned name for a request that says where (app 1.0.0 with script 16), or null (the old place).
function recordsFolder_(b, allowed) {
  const type = String(b.typeFolder || '');
  if (allowed && allowed.indexOf(type) < 0) return null;
  const month = cleanPart_(b.monthFolder);
  const name = cleanPart_(b.niceName).slice(0, 120).replace(/ +(\.[A-Za-z0-9]+)$/, '$1');
  if (RECORD_TYPE_FOLDERS.indexOf(type) < 0 || !/^\d{4}-\d{2} [A-Za-z]+$/.test(month) || !name) return null;
  const shop = RECORD_MONTH_LEVEL[type] ? '' : (cleanPart_(b.shopFolder).slice(0, 60) || 'Business');
  return { folder: path_(['Records', month, shop, type].filter(String)), name: name };
}
function inPayslips_(file) {
  const ps = file.getParents();
  while (ps.hasNext()) if (ps.next().getName() === 'Payslips') return true;
  return false;
}
// A name not used yet in the folder: "x.jpg", "x (2).jpg", …
function freeName_(folder, name) {
  if (!folder.getFilesByName(name).hasNext()) return name;
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? name.slice(0, dot) : name; const ext = dot > 0 ? name.slice(dot) : '';
  let out = name;
  for (let i = 2; i < 500 && folder.getFilesByName(out).hasNext(); i++) out = stem + ' (' + i + ')' + ext;
  return out;
}
const TIDY_MS = 270000; // ~4½ minutes: Apps Script stops a run at 6
// Owner / admin (Admin › Backup) or the script editor: the photos already in Drive › Easy Buys › Photos › month › kind
// named "<kind>-<record id>-<n>.jpg" are renamed and moved into Records (moving keeps the file id, so the links saved on
// the records still open them). Unknown ones stay where they are; nothing is deleted. Time-boxed: { moved, left, more }.
function tidyRecords_(maxMs) {
  const started = Date.now();
  const budget = Math.max(0, Math.min(TIDY_MS, maxMs === undefined || maxMs === null || maxMs === '' ? TIDY_MS : Number(maxMs) || 0));
  const top = root_().getFoldersByName('Photos');
  if (!top.hasNext()) return { ok: true, moved: 0, left: 0, more: false };
  const photos = top.next();
  const cache = {};
  const table = function (t) {
    if (cache[t]) return cache[t];
    const out = {}; const sh = data_().getSheetByName('T_' + t);
    if (sh) values_(sh).forEach(function (r) { try { const o = JSON.parse(r[4] || 'null'); if (o) out[String(r[0])] = o; } catch (e) { /* skip */ } });
    cache[t] = out; return out;
  };
  const byPayment = function (payUid) {
    if (!cache._alloc) { cache._alloc = {}; const a = table('ap_allocations'); Object.keys(a).forEach(function (k) { (cache._alloc[a[k].payment_uid] = cache._alloc[a[k].payment_uid] || []).push(a[k]); }); }
    const inv = table('ap_invoices');
    const list = (cache._alloc[payUid] || []).slice().sort(function (x, y) { return String(x.uid) < String(y.uid) ? -1 : 1; });
    for (let i = 0; i < list.length; i++) { const v = inv[list[i].invoice_uid]; if (v && v.shop_code) return v.shop_code; }
    return null;
  };
  const shopOf = function (code) { return code ? (table('shops')[code] || { code: code, name: '' }) : null; };
  const lookup = function (refType, refId) {
    let record = null; let shop = null; let cq = null;
    if (refType === 'invoice') { record = table('ap_invoices')[refId] || null; if (record) shop = shopOf(record.shop_code); }
    else if (refType === 'payment') {
      record = table('ap_payments')[refId] || null;
      if (record) { if (!record.cheque_no && record.cheque_uid) { const c = table('cheques')[record.cheque_uid]; cq = c ? c.cheque_no : null; } shop = shopOf(record.shop_code || byPayment(record.uid)); }
    } else if (refType === 'deposit' || refType === 'till') { record = table(refType === 'deposit' ? 'deposits' : 'till_counts')[refId] || null; if (record) shop = shopOf(record.shop_code); }
    else if (refType === 'stub') {
      record = table('cheques')[refId] || null;
      if (record) { const pay = record.payment_uid ? table('ap_payments')[record.payment_uid] : null; shop = shopOf(record.shop_code || (pay && pay.shop_code) || byPayment(record.payment_uid)); }
    }
    return record ? { record: record, shop: shop, chequeNo: cq } : null;
  };
  // Every old-named photo of the kinds that move, grouped by record (pages in photo order).
  const groups = {}; const order = [];
  let left = 0;
  const months = photos.getFolders();
  while (months.hasNext()) {
    const m = months.next();
    Object.keys(RECORD_FOLDERS).forEach(function (kind) {
      const it = m.getFoldersByName(kind);
      while (it.hasNext()) {
        const files = it.next().getFiles();
        while (files.hasNext()) {
          const f = files.next();
          const p = /^([a-z_]+)-(.+)-(\d+)\.(jpe?g|pdf|png)$/i.exec(String(f.getName()));
          if (!p || p[1] !== kind) { left++; continue; }
          const key = kind + '|' + p[2];
          if (!groups[key]) { groups[key] = { kind: kind, refId: p[2], files: [] }; order.push(key); }
          groups[key].files.push({ f: f, n: Number(p[3]), ext: p[4] });
        }
      }
    });
  }
  let moved = 0; let more = false;
  for (let gi = 0; gi < order.length; gi++) {
    if (moved > 0 && Date.now() - started >= budget) { more = true; break; } // at least one record moved each run
    const g = groups[order[gi]];
    const found = lookup(g.kind, g.refId);
    g.files.sort(function (a, b) { return a.n - b.n; });
    g.files.forEach(function (x, i) {
      const place = found ? photoPlace_({ refType: g.kind, record: found.record, shop: found.shop, chequeNo: found.chequeNo, page: i + 1, ext: x.ext }) : null;
      const where = place ? recordsFolder_({ niceName: place.fileName, monthFolder: place.monthFolder, shopFolder: place.shopFolder, typeFolder: place.typeFolder }) : null;
      if (!where) { left++; return; }
      try { x.f.moveTo(where.folder); x.f.setName(freeName_(where.folder, where.name)); moved++; } catch (e) { left++; }
    });
  }
  log_('Tidy photos', [moved + ' moved', left + ' left', more ? 'more to do' : 'done']);
  return { ok: true, moved: moved, left: left, more: more };
}
// Run from the script editor (Run › tidyRecords) – again until the log says "done".
function tidyRecordsRun_() {
  const r = tidyRecords_();
  Logger.log('Moved ' + r.moved + ', left where they are ' + r.left + (r.more ? ' – run again for the rest.' : ' – done.'));
  return r;
}


// ================================================================ 1.0.0 build 10: job leases (owner, 2 Nov: "It needs all the
// devices to allow sync instead of the main tablet")
// The jobs only one device should do at a time (bringing in the till PCs' products / suppliers / staff credit / FCCC,
// the automatic Viber posts, the web checks, the accountant pack, the nightly backup) go to whichever linked device asks
// first: it holds the job for ttlSec and keeps it by asking again on each sync; when it stops asking, the job is free
// after ttlSec and the next device that asks takes it. The main device chosen in Admin (shared setting
// main_tablet_device) comes first: it takes a job from an ordinary device. An admin device comes before a staff phone.
// The backup and the accountant pack need the whole data, so only admin devices (and the main device) are given those.
// mark: a short note the holder keeps with the job (the day of the last backup, the month of the last pack), so the next
// holder does not do it again. Kept in Script properties (LEASES).
const LEASE_ADMIN_JOBS = ['backup', 'pack'];
// 1.0.0 build 10 (review): only the app's own jobs (anything else is refused, never kept)
const LEASE_JOBS = ['link', 'viber', 'web', 'pack', 'backup'];
function leasesAll_() { try { const o = JSON.parse(props_().getProperty('LEASES') || '{}'); return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; } }
// A job held by another device is NEVER taken while its time runs (two devices doing the Aronium products at once made
// the same products twice). The main device / an admin device only comes first for a free job: asking for a job a
// lower device holds puts it next in line (want) – the holder is not renewed any more (it stops at its next check) and
// the job goes to the waiting device once the holder's time is up. A want older than twice the job's time is forgotten.
function jobLease_(body, who) {
  const list = Array.isArray(body.jobs) ? body.jobs : [{ job: body.job, ttlSec: body.ttlSec, mark: body.mark, release: body.release }];
  const jobs = list.filter(function (x) { return x && LEASE_JOBS.indexOf(String(x.job || '')) >= 0; }).slice(0, 12);
  if (!jobs.length || jobs.length !== list.filter(function (x) { return x; }).slice(0, 12).length) return { ok: false, error: 'Which job? Only ' + LEASE_JOBS.join(', ') + '.' };
  let main = '';
  try { main = String(sharedSetting_('main_tablet_device') || ''); } catch (e) { main = ''; }
  const me = String(who.deviceId);
  const prio = main && main === me ? 2 : (who.admin ? 1 : 0);
  let name = me;
  try { const r = deviceRows_().filter(function (x) { return String(x[0]) === me; })[0]; if (r && r[1]) name = String(r[1]); } catch (e) { /* the id */ }
  return withLock_(function () {
    const all = leasesAll_();
    const now = Date.now();
    const out = {};
    jobs.forEach(function (x) {
      const job = String(x.job);
      const ttl = Math.max(30, Math.min(3600, Number(x.ttlSec) || 600)) * (prio === 2 ? 2 : 1);
      const cur = all[job] || null;
      const live = cur && Number(cur.until) > now;
      const mine = live && String(cur.holder) === me;
      const markOf = function () { return cur ? cur.mark || '' : ''; };
      if (x.release) {
        if (mine) { cur.until = now; }
        out[job] = { granted: false, holder: live && !mine ? cur.holder : '', holderName: live && !mine ? cur.name : '', until: null, mark: markOf() };
        return;
      }
      // the backup and the accountant pack: an admin device only (being the main device is not enough – they need the
      // whole data). Not asked again for the job's time (staff phones used to ask under the write lock on every sync).
      if (LEASE_ADMIN_JOBS.indexOf(job) >= 0 && !who.admin) {
        out[job] = { granted: false, holder: live ? cur.holder : '', holderName: live ? cur.name : '', until: new Date(now + ttl * 1000).toISOString(), untilMs: now + ttl * 1000, mark: markOf(), why: 'admin device only' };
        return;
      }
      const want = cur && cur.want && now - Number(cur.want.at || 0) < 2 * ttl * 1000 ? cur.want : null;
      const wantOther = want && String(want.holder) !== me && Number(want.prio) > prio ? want : null;
      let take;
      if (mine) take = !(want && String(want.holder) !== me && Number(want.prio) > (Number(cur.prio) || 0)); // handing over: not renewed
      else if (live) take = false;                                                                           // never taken while it runs
      else take = !wantOther;                                                                                // free: the device waiting first
      if (take) {
        const mark = x.mark !== undefined && x.mark !== null && String(x.mark) !== '' ? String(x.mark).slice(0, 60) : markOf();
        all[job] = { holder: me, name: String(name).slice(0, 60), prio: prio, until: now + ttl * 1000, since: mine ? cur.since : now, mark: mark, at: now };
        if (want && !wantOther && String(want.holder) !== me) all[job].want = want; // a lower / equal device waiting: kept, it gets nothing earlier
        out[job] = { granted: true, holder: me, holderName: all[job].name, until: new Date(all[job].until).toISOString(), untilMs: all[job].until, mark: mark, took: !!(cur && String(cur.holder) !== me) };
        return;
      }
      if (mine) {
        // handing over to a device that comes first: this one stops now; the time it was given still runs out
        if (x.mark !== undefined && x.mark !== null && String(x.mark) !== '') cur.mark = String(x.mark).slice(0, 60);
        out[job] = { granted: false, handover: true, holder: String(want.holder), holderName: want.name || want.holder, until: new Date(Number(cur.until)).toISOString(), untilMs: Number(cur.until), mark: markOf() };
        return;
      }
      if (cur && prio > (Number(cur.prio) || 0) && !wantOther && (!want || String(want.holder) === me || Number(want.prio) < prio)) cur.want = { holder: me, name: String(name).slice(0, 60), prio: prio, at: now };
      else if (cur && want && String(want.holder) === me) cur.want.at = now;
      const busyUntil = live ? Number(cur.until) : now + 60000; // free but kept for the device waiting first
      out[job] = { granted: false, holder: live ? cur.holder : String((wantOther || {}).holder || ''), holderName: live ? (cur.name || cur.holder) : String((wantOther || {}).name || ''), until: new Date(busyUntil).toISOString(), untilMs: busyUntil, mark: markOf(), next: !!(cur && cur.want && String(cur.want.holder) === me) };
    });
    // forget jobs nobody asked for in a week
    Object.keys(all).forEach(function (j) { if (!(Number(all[j].until) > now - 7 * 86400000)) delete all[j]; });
    props_().setProperty('LEASES', JSON.stringify(all));
    const first = out[String(jobs[0].job)];
    return { ok: true, leases: out, granted: first.granted, holder: first.holder, until: first.until, mark: first.mark, prio: prio, version: SCRIPT_VERSION };
  }, 5000);
}

// Viber "once" keys: kept in the Viber once sheet (and the cache, for the quick answer) for 15 days.
function viberOnceSeen_(key) {
  const h = sha_(key).slice(0, 24);
  try { if (CacheService.getScriptCache().get('vonce_' + h)) return true; } catch (e) { /* the sheet */ }
  const sh = data_().getSheetByName('Viber once');
  if (!sh || sh.getLastRow() < 2) return false;
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === h) return true;
  return false;
}
// A post with this key is on its way (claimed inside the lock; for 2 minutes at most).
function viberOnceClaimed_(key) { try { return !!CacheService.getScriptCache().get('vclaim_' + sha_(key).slice(0, 24)); } catch (e) { return false; } }
function viberOnceClaim_(key, on) {
  try { const c = CacheService.getScriptCache(); const k = 'vclaim_' + sha_(key).slice(0, 24); if (on) c.put(k, '1', 120); else c.remove(k); } catch (e) { /* the sheet still knows once posted */ }
}
// The alert was sorted: its key goes (it is posted again if the same alert comes back).
function viberOnceForget_(key) {
  const h = sha_(key).slice(0, 24);
  try { CacheService.getScriptCache().remove('vonce_' + h); } catch (e) { /* the sheet */ }
  try {
    const sh = data_().getSheetByName('Viber once');
    if (!sh || sh.getLastRow() < 2) return;
    const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === h) sh.getRange(i + 2, 1, 1, 1).setValues([['forgotten']]);
  } catch (e) { /* kept 15 days at most */ }
}
function viberOnceKeep_(key) {
  const h = sha_(key).slice(0, 24);
  try { CacheService.getScriptCache().put('vonce_' + h, '1', 21600); } catch (e) { /* the sheet */ }
  try {
    const sh = sheet_('Viber once', ['key', 'at']);
    sh.appendRow([h, now_()]);
    // keep it short: the oldest rows go once it passes 3000 rows
    if (sh.getLastRow() > 3000) {
      const keep = sh.getRange(1002, 1, sh.getLastRow() - 1001, 2).getValues();
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
      sh.getRange(2, 1, keep.length, 2).setValues(keep);
    }
  } catch (e) { /* the cache still knows for 6 hours */ }
}
