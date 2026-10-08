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
    // --- a new device joining with an invite (no key yet) ---
    if (body.kind === 'join') return json_(join_(body));
    const who = auth_(body);
    if (!who.ok) return json_({ ok: false, error: who.error, blocked: !!who.blocked, waiting: !!who.waiting });
    if (body.kind === 'push') return json_(push_(body.records || [], who.deviceId));
    if (body.kind === 'pull') return json_(pull_(Number(body.since) || 0, who.deviceId, false));
    if (body.kind === 'invite_create') return json_(inviteCreate_(body, who));
    if (body.kind === 'devices') return json_({ ok: true, devices: devices_() });
    if (body.kind === 'device_status') return json_(deviceStatus_(body.id, body.status, who));
    if (body.kind === 'cloud_settings') return json_(cloudSettings_(body));
    switch (body.kind) {
      case 'ping':
        return json_({ ok: true, account: Session.getEffectiveUser().getEmail(), folder: root_().getUrl(), device: who.deviceId, version: 3 });

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
        return json_({ ok: true, planEdits: edits });
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

      case 'backup': {
        const folder = path_(['Backups']);
        saveFile_(folder, body.fileName, Utilities.base64Encode(body.data, Utilities.Charset.UTF_8), 'application/json');
        return json_({ ok: true });
      }

      default:
        return json_({ ok: false, error: 'Unknown request: ' + body.kind });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
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
  return values_(sheet_('Devices', DEVICE_HEADER)).map(function (r) {
    return { id: r[0], name: r[1], kind: r[2], status: r[3], first_seen: r[5], last_seen: r[6], app: r[7], invited_by: r[8] };
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
    if (!d) return { ok: false, error: 'This device is not linked any more. Link it again from another device (Admin › Devices).' };
    if (d.v[3] === 'blocked') return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' };
    if (d.v[3] === 'waiting') return { ok: false, waiting: true, error: 'Waiting for the owner to approve this device.' };
    touchDevice_(body, String(d.v[0]));
    return { ok: true, deviceId: String(d.v[0]) };
  }
  if (!API_KEY || API_KEY.indexOf('PASTE') === 0) return { ok: false, error: 'The cloud key has not been set in the cloud program yet.' };
  if (body.key !== API_KEY) return { ok: false, error: 'Wrong cloud key. Copy the key from the app into the cloud program.' };
  const id = deviceId || 'first-device';
  const d = findDevice_(function (r) { return String(r[0]) === id; });
  if (d && d.v[3] === 'blocked') return { ok: false, blocked: true, error: 'This device has been blocked by the owner.' };
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
