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

const ROOT_NAME = 'Easy Buys Bakery Cloud';   // existing folder name kept so nothing moves
const DATA_NAME = 'Easy Buys Data (do not edit by hand)';
const INVITE_HOURS = 48;
const SCRIPT_VERSION = 6;                 // v3.7 / 1.0: documents (checked uploads, visibility, nightly backups, Windows file API); device control (pause, end date, remove, wipe, commands); v3.6: Start fresh
const SUPPORT_EMAIL = 'prasaddiptesh@gmail.com'; // "Send to support" problem reports (Admin & support)
const PULL_LIMIT = 800;
const MAIN_TABLES_PULL_ALL = true;   // the main software receives every table, including the change history

// ---------------------------------------------------------------- helpers
function json_(obj) {
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
function saveFile_(folder, fileName, base64, mime) {
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, fileName);
  return folder.createFile(blob);
}
function log_(sheetName, row) {
  const ssIt = root_().getFilesByName('Bakery Cloud Log');
  const ss = ssIt.hasNext() ? SpreadsheetApp.open(ssIt.next()) : (function () {
    const created = SpreadsheetApp.create('Bakery Cloud Log');
    DriveApp.getFileById(created.getId()).moveTo(root_());
    return created;
  })();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.appendRow([new Date()].concat(row));
}
function cleanEmails_(list) {
  return (list || []).map(function (e) { return String(e).trim(); }).filter(function (e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); });
}

// ---------------------------------------------------------------- tablet → cloud
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // --- the main Easy Buys software (office PC) ---
    if (body.kind === 'main_pull' || body.kind === 'main_push' || body.kind === 'main_devices') return json_(mainApi_(body));
    if (body.kind === 'main_files' || body.kind === 'main_file_get') return json_(mainFiles_(body)); // v3.5 Windows software: full file mirror
    // --- a new device joining with an invite (no key yet) ---
    if (body.kind === 'join') return json_(join_(body));
    const who = auth_(body);
    if (!who.ok) return json_({ ok: false, error: who.error, blocked: !!who.blocked, waiting: !!who.waiting, lock: who.lock || null, wipe: !!who.wipe, removed: !!who.removed });
    if (body.kind === 'push') return json_(push_(body.records || [], who.deviceId));
    if (body.kind === 'pull') { const out = pull_(Number(body.since) || 0, who.deviceId, false); devicePulled_(who.deviceId, body, out); return json_(out); }
    if (body.kind === 'invite_create') return json_(inviteCreate_(body, who));
    if (body.kind === 'devices') return json_({ ok: true, devices: devices_() });
    if (body.kind === 'device_status') return json_(deviceStatus_(body.id, body.status, who));
    if (body.kind === 'device_control') return json_(deviceControl_(body, who));
    if (body.kind === 'problem_log') return json_(problemLogSave_(body, who));
    if (body.kind === 'problem_log_get') return json_(problemLogGet_(body.id));
    if (body.kind === 'cloud_settings') return json_(cloudSettings_(body));
    if (body.kind === 'fresh_start') return json_(freshStart_(body));
    switch (body.kind) {
      case 'ping':
        return json_({ ok: true, account: Session.getEffectiveUser().getEmail(), folder: root_().getUrl(), device: who.deviceId, version: SCRIPT_VERSION });

      case 'snapshot': {
        const props = PropertiesService.getScriptProperties();
        const text = JSON.stringify(body.snapshot);
        // Script properties hold up to ~9 KB per value, so larger snapshots go to a Drive file.
        if (text.length < 8500) { props.setProperty('SNAPSHOT', text); props.deleteProperty('SNAPSHOT_FILE'); }
        else {
          const f = saveFile_(root_(), 'dashboard-status.json', Utilities.base64Encode(text, Utilities.Charset.UTF_8), 'application/json');
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
        const folder = path_(['Shift reports', body.month || 'Other']);
        const pdf = saveFile_(folder, body.fileName + '.pdf', body.pdf, MimeType.PDF);
        const xlsx = body.xlsx ? saveFile_(folder, body.fileName + '.xlsx', body.xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') : null;
        const to = cleanEmails_(body.emails);
        if (to.length) {
          MailApp.sendEmail({
            to: to.join(','), subject: body.subject,
            htmlBody: '<p>The shift has been balanced and approved. The full report is attached (PDF and Excel) and saved in Google Drive.</p>' + body.html,
            attachments: [pdf.getBlob()].concat(xlsx ? [xlsx.getBlob()] : []),
            name: 'Easy Buys Bakery',
          });
        }
        log_('Reports', [body.shiftCode, body.fileName, to.join(', '), pdf.getUrl()]);
        return json_({ ok: true, url: pdf.getUrl(), emailed: to.length });
      }

      case 'po': {
        const folder = path_(['Purchase orders', body.month || 'Other']);
        const pdf = saveFile_(folder, body.fileName + '.pdf', body.pdf, MimeType.PDF);
        const to = cleanEmails_(body.emails);
        if (to.length) {
          MailApp.sendEmail({
            to: to.join(','), subject: body.subject,
            htmlBody: '<p>Stock has reached the minimum level. Purchase order attached for reordering'
              + (body.supplierEmail ? ' (supplier email: ' + body.supplierEmail + ')' : '') + '.</p>' + body.html,
            attachments: [pdf.getBlob()], name: 'Easy Buys Bakery',
          });
        }
        log_('Purchase orders', [body.fileName, to.join(', '), pdf.getUrl()]);
        return json_({ ok: true, url: pdf.getUrl(), emailed: to.length });
      }

      case 'photo': {
        const folder = path_(['Photos', body.month || 'Other', body.refType || 'other']);
        const f = saveFile_(folder, body.fileName, body.data, MimeType.JPEG);
        return json_({ ok: true, url: f.getUrl() });
      }

      case 'ocr': {
        // Read the text of invoice photos with Google's free OCR (needs Services › Drive API added once).
        if (typeof Drive === 'undefined' || !Drive.Files) return json_({ ok: true, off: true });
        const texts = (body.files || []).slice(0, 4).map(function (f, i) {
          return ocr_(Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType || 'image/jpeg', 'page-' + (i + 1) + '.jpg'));
        });
        return json_({ ok: true, text: texts.join('\n') });
      }

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

      case 'backup_list': {
        // v3.5: the backups kept in Drive › Backups, › Daily, › Monthly and › Manual (newest first) for Admin › Backup & restore.
        return json_({ ok: true, files: backupFiles_().slice(0, Number(body.limit) || 60) });
      }

      case 'backup_get': {
        // Only files kept in Easy Buys › Backups can be fetched this way.
        const file = DriveApp.getFileById(String(body.id || ''));
        if (!insideFolder_(file, path_(['Backups']).getId(), 2)) return json_({ ok: false, error: 'Not an Easy Buys backup.' });
        return json_({ ok: true, name: file.getName(), text: file.getBlob().getDataAsString('UTF-8') });
      }

      case 'viber_setup': {
        // Viber Channel "Post API": the token comes from the channel's Settings › Developer tools. Kept only here.
        const token = String(body.token || '').trim();
        const current = PropertiesService.getScriptProperties().getProperty('VIBER_TOKEN');
        // Replacing or removing a token needs the current one (so a lost tablet cannot take the alerts over).
        if (current && String(body.oldToken || '').trim() !== current) return json_({ ok: false, error: 'A Viber Channel is already connected. Enter its current token to replace or remove it.' });
        if (!token) { PropertiesService.getScriptProperties().deleteProperty('VIBER_TOKEN'); return json_({ ok: true, message: 'Viber Channel disconnected.' }); }
        PropertiesService.getScriptProperties().setProperty('VIBER_TOKEN', token);
        PropertiesService.getScriptProperties().deleteProperty('VIBER_FROM');
        const t = viber_('Easy Buys alerts are connected to this channel.');
        if (!t.ok) PropertiesService.getScriptProperties().deleteProperty('VIBER_TOKEN');
        return json_(t.ok ? { ok: true, message: 'Connected – a test message was sent to the channel.' } : t);
      }

      case 'email': {
        // Emails to staff (welcome, signed contract). Sent from the Google account the cloud runs on.
        const to = cleanEmails_([body.to]);
        if (!to.length) return json_({ ok: false, error: 'No valid email address.' });
        // Only to people being onboarded (their email is in the synced onboarding records) – a device cannot email anyone else.
        if (!staffEmail_(to[0])) return json_({ ok: false, error: 'This address is not a staff member being onboarded.' });
        const atts = (body.attachments || []).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/pdf', a.name || 'document.pdf'); });
        MailApp.sendEmail({ to: to.join(','), subject: String(body.subject || 'Easy Buys'), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to.join(', '), body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true });
      }

      case 'payslip_email': {
        // v3.5 payslips: only to an address saved in a person's pay details (or onboarding) – a device cannot email anyone else.
        const to = cleanEmails_([body.to]);
        if (!to.length) return json_({ ok: false, error: 'No valid email address.' });
        if (!payEmail_(to[0]) && !staffEmail_(to[0])) return json_({ ok: false, error: 'This address is not in anyone\'s pay details.' });
        const atts = (body.attachments || []).slice(0, 2).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/pdf', a.name || 'payslip.pdf'); });
        MailApp.sendEmail({ to: to[0], subject: String(body.subject || 'Easy Buys payslip').slice(0, 200), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to[0], body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true });
      }

      case 'pack_email': {
        // 1.0 monthly accountant pack: only to the addresses the owner saved in Admin › Emails (synced setting).
        const to = prefEmails_('accountant_pack');
        if (!to.length) return json_({ ok: false, error: 'No accountant email saved in Admin › Emails.' });
        const atts = (body.attachments || []).slice(0, 3).map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'application/octet-stream', a.name || 'pack.xlsx'); });
        MailApp.sendEmail({ to: to.join(','), subject: String(body.subject || 'Easy Buys accountant pack').slice(0, 200), body: String(body.text || ''), attachments: atts, name: 'Easy Buys' });
        log_('Emails', [to.join(', '), body.subject, atts.length + ' attachment(s)']);
        return json_({ ok: true, emailed: to.length });
      }

      case 'support_email': {
        // v3.5 "Send to support": problem reports go only to the app's support address (fixed here, not chosen by the device).
        MailApp.sendEmail({ to: SUPPORT_EMAIL, subject: String(body.subject || 'Easy Buys problem report').slice(0, 200),
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

      case 'viber':
        return json_(viber_(String(body.text || '').slice(0, 7000)));

      default:
        return json_({ ok: false, error: 'Unknown request: ' + body.kind });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

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

function staffEmail_(addr) {
  const a = String(addr).trim().toLowerCase();
  const sh = data_().getSheetByName('T_onboarding');
  if (!sh) return false;
  return values_(sh).some(function (r) { try { return String(JSON.parse(r[4] || '{}').email || '').trim().toLowerCase() === a; } catch (e) { return false; } });
}

// Send a text to the Viber Channel (Channels Post API). The sender must be the channel's super admin.
function viber_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('VIBER_TOKEN');
  if (!token) return { ok: true, off: true };
  const call = function (path, payload) {
    const r = UrlFetchApp.fetch('https://chatapi.viber.com/pa/' + path, { method: 'post', contentType: 'application/json', headers: { 'X-Viber-Auth-Token': token },
      payload: JSON.stringify(payload || {}), muteHttpExceptions: true });
    try { return JSON.parse(r.getContentText()); } catch (e) { return { status: -1, status_message: r.getContentText().slice(0, 200) }; }
  };
  let from = props.getProperty('VIBER_FROM');
  if (!from) {
    const info = call('get_account_info', {});
    if (info.status !== 0) return { ok: false, error: 'Viber: ' + (info.status_message || 'token not accepted') };
    const admin = (info.members || []).filter(function (m) { return m.role === 'superadmin'; })[0] || (info.members || [])[0];
    if (!admin) return { ok: false, error: 'Viber: the channel has no admin member.' };
    from = admin.id;
    props.setProperty('VIBER_FROM', from);
  }
  const r = call('post', { from: from, type: 'text', text: text });
  if (r.status !== 0) { props.deleteProperty('VIBER_FROM'); return { ok: false, error: 'Viber: ' + (r.status_message || 'not sent') }; }
  return { ok: true };
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
function doGet(e) {
  if (e && e.parameter && e.parameter.link) return linkPage_(String(e.parameter.link));
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('Easy Buys Bakery – Live Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

// Called by the dashboard page every 30 seconds.
function getDashboard(password) {
  if (!DASHBOARD_PASSWORD || password !== DASHBOARD_PASSWORD) return { ok: false, error: 'Wrong password' };
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
function savePlanOnline(password, date, items, byName) {
  if (!DASHBOARD_PASSWORD || password !== DASHBOARD_PASSWORD) return { ok: false, error: 'Wrong password' };
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
function data_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DATA_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* deleted: make a new one */ } }
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
function values_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues().map(function (r) { return r.map(text_); });
}
// Safety net if a cell was ever typed or formatted by hand.
function text_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  return v === null || v === undefined ? '' : String(v);
}
const TABLE_HEADER = ['key', 'updated_at', 'device', 'seq', 'json'];
const CHANGES_HEADER = ['seq', 'table', 'key', 'row', 'device', 'received_at'];
const DEVICE_HEADER = ['id', 'name', 'kind', 'status', 'token_hash', 'first_seen', 'last_seen', 'app', 'invited_by', 'note'];
const INVITE_HEADER = ['code', 'created_at', 'expires_at', 'created_by', 'email', 'used_at', 'used_by'];

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function now_() { return new Date().toISOString(); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Store records from a device (or the main software). The newest change wins.
function push_(records, deviceId) {
  return withLock_(function () {
    const props = PropertiesService.getScriptProperties();
    let seq = Number(props.getProperty('SEQ') || '0');
    let fresh = null;
    try { fresh = JSON.parse(props.getProperty('FRESH') || 'null'); } catch (e) { fresh = null; }
    const changes = sheet_('Changes', CHANGES_HEADER);
    const byTable = {};
    records.forEach(function (r) {
      if (!r || !r.t || !/^[a-z_]+$/.test(r.t) || r.k === undefined || r.k === null) return;
      (byTable[r.t] = byTable[r.t] || []).push(r);
    });
    const accepted = [];
    const rejected = [];
    const newChanges = [];
    Object.keys(byTable).forEach(function (table) {
      const sh = sheet_('T_' + table, TABLE_HEADER);
      const rows = values_(sh);
      const index = {};
      rows.forEach(function (row, i) { index[String(row[0])] = { i: i + 2, updated: String(row[1]) }; });
      const appends = [];
      byTable[table].forEach(function (r) {
        const key = String(r.k);
        const updated = String(r.u || (r.row && r.row.updated_at) || '');
        const have = index[key];
        if (have && have.updated > updated) { rejected.push({ t: table, k: key, why: 'newer change already in the cloud' }); return; }
        if (fresh && fresh.tables.indexOf(table) >= 0 && updated && updated < fresh.at) { rejected.push({ t: table, k: key, why: 'from before the fresh start' }); return; }
        seq += 1;
        const line = [key, updated, deviceId, seq, JSON.stringify(r.row || {})];
        let rowNo;
        if (have) {
          rowNo = have.i;
          sh.getRange(rowNo, 1, 1, line.length).setValues([line]);
        } else {
          rowNo = sh.getLastRow() + 1 + appends.length;
          appends.push(line);
          index[key] = { i: rowNo, updated: updated };
        }
        newChanges.push([seq, table, key, rowNo, deviceId, now_()]);
        accepted.push({ t: table, k: key });
      });
      if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, TABLE_HEADER.length).setValues(appends);
    });
    if (newChanges.length) changes.getRange(changes.getLastRow() + 1, 1, newChanges.length, CHANGES_HEADER.length).setValues(newChanges);
    props.setProperty('SEQ', String(seq));
    return { ok: true, accepted: accepted.length, rejected: rejected, seq: seq };
  });
}

// Changes after `since`, made by other devices (the main software gets everything).
function pull_(since, deviceId, forMain) {
  const changes = sheet_('Changes', CHANGES_HEADER);
  const last = changes.getLastRow();
  const total = Math.max(0, last - 1);
  if (since >= total) return { ok: true, records: [], next: total, more: false };
  const count = Math.min(PULL_LIMIT, total - since);
  const list = changes.getRange(since + 2, 1, count, CHANGES_HEADER.length).getValues().map(function (r) { return r.map(text_); });
  const tables = {};
  const out = [];
  const seen = {};
  list.forEach(function (c) {
    const table = String(c[1]);
    if (!forMain && (table === 'audit_log' || String(c[4]) === String(deviceId))) return;
    const id = table + '|' + c[2];
    if (seen[id]) return;   // only the latest version of a record in this batch
    seen[id] = true;
    if (!tables[table]) tables[table] = values_(sheet_('T_' + table, TABLE_HEADER));
    const row = tables[table][Number(c[3]) - 2];
    if (!row) return;
    let parsed = {};
    try { parsed = JSON.parse(row[4]); } catch (e) { parsed = {}; }
    out.push({ t: table, k: String(row[0]), u: String(row[1]), d: String(row[2]), row: parsed });
  });
  const next = since + count;
  return { ok: true, records: out, next: next, more: next < total };
}

// ================================================================ devices & invites
function devices_() {
  const all = ctlAll_();
  return values_(sheet_('Devices', DEVICE_HEADER)).map(function (r) {
    const c = all[String(r[0])] || {};
    return { id: r[0], name: r[1], kind: r[2], status: r[3], first_seen: r[5], last_seen: r[6], app: r[7], invited_by: r[8],
      until: c.until || null, endDate: c.endDate || null, shop: c.shop || null, roles: c.roles || [], stat: c.stat || null, hist: c.hist || [],
      attempts: c.attempts || [], waitingCmds: (c.cmds || []).length, wipe: c.wipe || null, lastLog: c.lastLog || null };
  });
}
function findDevice_(match) {
  const sh = sheet_('Devices', DEVICE_HEADER);
  const rows = values_(sh);
  for (let i = 0; i < rows.length; i++) if (match(rows[i])) return { sh: sh, row: i + 2, v: rows[i] };
  return null;
}
function touchDevice_(body, deviceId, extra) {
  const d = body.device || {};
  const found = findDevice_(function (r) { return String(r[0]) === String(deviceId); });
  const now = now_();
  if (found) {
    found.sh.getRange(found.row, 2, 1, 1).setValue(d.name || found.v[1]);
    found.sh.getRange(found.row, 3, 1, 1).setValue(d.kind || found.v[2]);
    found.sh.getRange(found.row, 7, 1, 2).setValues([[now, body.app || '']]);
    return found;
  }
  sheet_('Devices', DEVICE_HEADER).appendRow([deviceId, d.name || 'Device', d.kind || '', (extra && extra.status) || 'active',
    (extra && extra.tokenHash) || '', now, now, body.app || '', (extra && extra.invitedBy) || '', (extra && extra.note) || '']);
  return null;
}

// Every request (except join and the main software) must come from a linked, unblocked device.
function auth_(body) {
  const deviceId = String((body.device && body.device.id) || '');
  if (body.token) {
    const h = sha_(body.token);
    const d = findDevice_(function (r) { return String(r[4]) === h; });
    if (!d) { noteAttemptByToken_(h); return { ok: false, removed: true, error: 'This device is not linked any more. Link it again from another device (Admin › Devices).' }; }
    if (d.v[3] === 'blocked') noteAttempt_(String(d.v[0]), 'blocked');
    if (d.v[3] === 'blocked') return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' };
    if (d.v[3] === 'waiting') return { ok: false, waiting: true, error: 'Waiting for the owner to approve this device.' };
    const g = gate_(String(d.v[0]), d);
    if (g) return g;
    touchDevice_(body, String(d.v[0]));
    return { ok: true, deviceId: String(d.v[0]) };
  }
  if (!API_KEY || API_KEY.indexOf('PASTE') === 0) return { ok: false, error: 'The cloud key has not been set in the cloud program yet.' };
  if (body.key !== API_KEY) return { ok: false, error: 'Wrong cloud key. Copy the key from the app into the cloud program.' };
  const id = deviceId || 'first-device';
  const d = findDevice_(function (r) { return String(r[0]) === id; });
  if (d && d.v[3] === 'blocked') { noteAttempt_(id, 'blocked'); return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' }; }
  if (d && d.v[3] === 'removed') { noteAttempt_(id, 'removed'); return { ok: false, removed: true, error: 'This device was removed by the owner. Link it again.' }; }
  if (d) { const g = gate_(id, d); if (g) return g; }
  touchDevice_(body, id, { status: 'active', note: 'linked with the cloud key' });
  return { ok: true, deviceId: id };
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
  const expires = new Date(created.getTime() + INVITE_HOURS * 3600000);
  const email = String(body.email || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'That email address does not look right.' };
  const url = String(body.url || '');
  sheet_('Invites', INVITE_HEADER).appendRow([code, created.toISOString(), expires.toISOString(), (body.by || '') + ' @ ' + who.deviceId, email, '', '']);
  const link = linkCode_(url, code);
  let emailed = false;
  if (email) {
    const tap = linkUrl_(url, link);
    MailApp.sendEmail({
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
  return { ok: true, code: code, linkCode: link, linkUrl: url ? linkUrl_(url, link) : '', expires: expires.toISOString(), emailed: emailed };
}

// New device: swap the invite for its own access token.
function join_(body) {
  return withLock_(function () {
    const code = String(body.invite || '').trim().toUpperCase();
    const sh = sheet_('Invites', INVITE_HEADER);
    const rows = values_(sh);
    let row = -1;
    for (let i = 0; i < rows.length; i++) if (String(rows[i][0]) === code) row = i;
    if (row < 0) return { ok: false, error: 'Invite not found. Ask for a new one.' };
    const inv = rows[row];
    if (inv[5]) return { ok: false, error: 'This invite has already been used. Ask for a new one.' };
    if (new Date(inv[2]).getTime() < Date.now()) return { ok: false, error: 'This invite has expired. Ask for a new one.' };
    const d = body.device || {};
    if (!d.id) return { ok: false, error: 'Device id missing.' };
    const token = Utilities.getUuid() + Utilities.getUuid();
    const approval = PropertiesService.getScriptProperties().getProperty('REQUIRE_APPROVAL') === '1';
    const existing = findDevice_(function (r) { return String(r[0]) === String(d.id); });
    if (existing) existing.sh.getRange(existing.row, 4, 1, 2).setValues([[approval ? 'waiting' : 'active', sha_(token)]]);
    else touchDevice_(body, String(d.id), { status: approval ? 'waiting' : 'active', tokenHash: sha_(token), invitedBy: inv[3], note: 'invite ' + code });
    sh.getRange(row + 2, 6, 1, 2).setValues([[now_(), String(d.id)]]);
    return { ok: true, token: token, approved: !approval };
  });
}

function deviceStatus_(id, status, who) {
  if (['active', 'blocked'].indexOf(status) < 0) return { ok: false, error: 'Choose active or blocked.' };
  if (String(id) === String(who.deviceId) && status === 'blocked') return { ok: false, error: 'A device cannot block itself.' };
  const d = findDevice_(function (r) { return String(r[0]) === String(id); });
  if (!d) return { ok: false, error: 'Device not found.' };
  d.sh.getRange(d.row, 4).setValue(status);
  return { ok: true, devices: devices_() };
}

// ---------------------------------------------------------------- v3.6 device control (Admin › Devices)
// Kept in a script property: { [deviceId]: { until, endDate, shop, roles, name, cmds: [], hist: [], attempts: [], stat, wipe, removedTokens } }.
function ctlAll_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DEVCTL') || '{}'); } catch (e) { return {}; } }
function ctlSave_(all) { PropertiesService.getScriptProperties().setProperty('DEVCTL', JSON.stringify(all)); }
function ctl_(all, id) { all[id] = all[id] || { cmds: [], hist: [], attempts: [] }; return all[id]; }
function hist_(c, text, by) { c.hist = [{ at: now_(), text: text, by: by || '' }].concat(c.hist || []).slice(0, 25); }
function noteAttempt_(id, why) {
  const all = ctlAll_(); const c = ctl_(all, id);
  c.attempts = [{ at: now_(), why: why }].concat(c.attempts || []).slice(0, 10);
  ctlSave_(all);
}
function noteAttemptByToken_(h) {
  const all = ctlAll_();
  Object.keys(all).forEach(function (id) { if ((all[id].removedTokens || []).indexOf(h) >= 0) { const c = all[id]; c.attempts = [{ at: now_(), why: 'removed' }].concat(c.attempts || []).slice(0, 10); } });
  ctlSave_(all);
}
// Refuse a suspended, ended or wiped device (enforced here, whatever the device does).
function gate_(id, d) {
  const all = ctlAll_();
  const c = all[id];
  if (!c) return null;
  if (c.wipe === 'pending') {
    c.wipe = 'done'; hist_(c, 'Wiped (the device cleared its data)', 'cloud');
    if (d && d.sh) { c.removedTokens = (c.removedTokens || []).concat([String(d.v[4])]).slice(-5); d.sh.getRange(d.row, 4, 1, 2).setValues([['removed', '']]); }
    ctlSave_(all);
    return { ok: false, wipe: true, removed: true, error: 'The owner wiped this device.' };
  }
  const nowIso = now_();
  if (c.until) {
    if (c.until > nowIso) { c.attempts = [{ at: nowIso, why: 'suspended' }].concat(c.attempts || []).slice(0, 10); ctlSave_(all); return { ok: false, lock: { kind: 'suspended', until: c.until }, error: 'Access paused by the owner until ' + c.until.slice(0, 16).replace('T', ' ') + '.' }; }
    delete c.until; hist_(c, 'Pause ended', 'cloud'); ctlSave_(all);
  }
  if (c.endDate && c.endDate < nowIso.slice(0, 10)) {
    c.attempts = [{ at: nowIso, why: 'ended' }].concat(c.attempts || []).slice(0, 10); ctlSave_(all);
    return { ok: false, lock: { kind: 'ended', until: c.endDate }, error: 'Access for this device ended on ' + c.endDate + '.' };
  }
  return null;
}
// What the device reports on each pull (waiting items, last health check / backup / login), and what it is told back.
function devicePulled_(id, body, out) {
  const all = ctlAll_(); const c = ctl_(all, id);
  if (body.stat) c.stat = { at: now_(), pending: body.stat.pending, health: body.stat.health || null, backup: body.stat.backup || null, login: body.stat.login || null };
  out.control = { shop: c.shop || null, roles: c.roles || [], name: c.name || null, endDate: c.endDate || null };
  out.commands = c.cmds || [];
  c.cmds = [];
  ctlSave_(all);
}
const DEVICE_ACTIONS = ['suspend', 'block', 'allow', 'remove', 'rename', 'shop', 'roles', 'end', 'sync', 'log', 'signout', 'wipe'];
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
    else if (a === 'allow') { d.sh.getRange(d.row, 4).setValue('active'); delete c.until; if (c.endDate && c.endDate < now_().slice(0, 10)) delete c.endDate; hist_(c, 'Allowed again', by); }
    else if (a === 'remove') { c.removedTokens = (c.removedTokens || []).concat([String(d.v[4])]).slice(-5); d.sh.getRange(d.row, 4, 1, 2).setValues([['removed', '']]); hist_(c, 'Removed (must be linked again with a new invite)', by); }
    else if (a === 'rename') { const n = String(body.name || '').trim(); if (!n) return { ok: false, error: 'Type the new name.' }; c.name = n; d.sh.getRange(d.row, 2).setValue(n); hist_(c, 'Renamed to ' + n, by); }
    else if (a === 'shop') { c.shop = String(body.shop || '') || null; hist_(c, 'Shop set to ' + (c.shop || 'none'), by); }
    else if (a === 'roles') { c.roles = (body.roles || []).map(String); hist_(c, c.roles.length ? 'Only these roles may log in: ' + c.roles.join(', ') : 'Every role may log in', by); }
    else if (a === 'end') { const e = String(body.endDate || ''); if (e && !/^\d{4}-\d{2}-\d{2}$/.test(e)) return { ok: false, error: 'Choose the last day.' }; c.endDate = e || null; hist_(c, e ? 'Access ends after ' + e : 'No end date', by); }
    else if (a === 'wipe') { c.wipe = 'pending'; hist_(c, 'Wipe asked (on its next contact)', by); }
    else { c.cmds = (c.cmds || []).concat([{ kind: a, at: now_(), by: String(body.by || '') }]).slice(-10); hist_(c, { sync: 'Asked to sync now', log: 'Asked for its problem log', signout: 'Asked to sign everyone out' }[a], by); }
    ctlSave_(all);
    return { ok: true, devices: devices_() };
  });
}
function problemLogSave_(body, who) {
  const text = String(body.text || '').slice(0, 45000);
  sheet_('Problem logs', ['at', 'device', 'text']).appendRow([now_(), who.deviceId, text]);
  const all = ctlAll_(); const c = ctl_(all, who.deviceId); c.lastLog = now_(); hist_(c, 'Sent its problem log', 'device'); ctlSave_(all);
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
  return { ok: true, requireApproval: props.getProperty('REQUIRE_APPROVAL') === '1' };
}

// ================================================================ main Easy Buys software
function mainApi_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('MAIN_KEY');
  if (!key || body.key !== key) return { ok: false, error: 'Not allowed' };
  PropertiesService.getScriptProperties().setProperty('LAST_MAIN_SYNC', now_());
  if (body.kind === 'main_pull') return pull_(Number(body.since) || 0, 'main', MAIN_TABLES_PULL_ALL);
  if (body.kind === 'main_push') return push_(body.records || [], 'main');
  return { ok: true, devices: devices_() };
}

// Run once from the editor: creates the data spreadsheet and the main-software key.
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MAIN_KEY')) props.setProperty('MAIN_KEY', Utilities.getUuid() + Utilities.getUuid());
  data_();
  ['Changes', 'Devices', 'Invites'].forEach(function (n) { sheet_(n, n === 'Changes' ? CHANGES_HEADER : n === 'Devices' ? DEVICE_HEADER : INVITE_HEADER); });
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
  if (!body.key || body.key !== API_KEY) return { ok: false, error: 'Start fresh only from the main tablet (the one holding the cloud key).' };
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
  if (!row && !insideFolder_(file, path_(['Documents']).getId(), 5)) return { ok: false, error: 'Not an Easy Buys document.' };
  const legacy = row ? null : legacyDocType_(id);
  const type = row ? String(row[6] || '') : legacy ? legacy.type : '';
  const access = row ? String(row[7] || 'normal') : legacy ? legacy.access : 'normal';
  const allowed = docAllowed_(body.user, type, access, !row);
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
