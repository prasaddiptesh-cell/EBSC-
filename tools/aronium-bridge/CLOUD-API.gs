/**
 * Easy Buys POS bridge API scaffold.
 *
 * Install this file into the same Apps Script project as Code.gs and add the
 * two dispatch lines documented in CLOUD-API-CONTRACT.md before deployment.
 * Set Script Properties:
 *   POS_TOKEN_S1-TILL1 = unique long random token
 *   POS_TOKEN_S1-TILL2 = unique long random token
 *   POS_TOKEN_S2-TILL1 = unique long random token
 *   POS_QUERY_TOKEN = separate long random token for authorized central queries
 *
 * Do not commit real tokens. Google Sheets is suitable for a controlled pilot;
 * assess storage volume/latency before production-scale POS ingestion.
 */
const POS_EVENTS_HEADER_ = [
  'event_id', 'source_id', 'event_type', 'event_time',
  'payload_json', 'payload_sha256', 'received_at', 'schema_version'
];
const POS_MAX_BATCH_ = 50;
const POS_MAX_EVENT_CHARS_ = 80000;

function posHash_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function posEventsSheet_() {
  return sheet_('POS_Events', POS_EVENTS_HEADER_);
}
function posTokenValid_(sourceId, token) {
  if (!/^[A-Z0-9_-]{3,40}$/.test(String(sourceId || ''))) return false;
  const expected = PropertiesService.getScriptProperties().getProperty('POS_TOKEN_' + sourceId);
  return !!expected && !!token && posHash_(String(token)) === posHash_(expected);
}
function posIngest_(body) {
  const sourceId = String(body.source_id || '');
  if (!posTokenValid_(sourceId, body.token)) return { ok: false, error: 'POS bridge authentication failed.' };
  if (!Array.isArray(body.events) || body.events.length < 1 || body.events.length > POS_MAX_BATCH_) {
    return { ok: false, error: 'Batch must contain between 1 and ' + POS_MAX_BATCH_ + ' events.' };
  }
  return withLock_(function () {
    const sh = posEventsSheet_();
    const outcomes = [];
    let accepted = 0, duplicates = 0, rejected = 0;
    body.events.forEach(function (ev) {
      const eventId = String((ev && ev.event_id) || '');
      const eventType = String((ev && ev.event_type) || '');
      const eventTime = String((ev && ev.event_time) || '');
      const schemaVersion = String((ev && ev.schema_version) || '1');
      if (!/^[A-Za-z0-9:._-]{1,180}$/.test(eventId) ||
          !/^[a-z][a-z0-9_]{1,50}$/.test(eventType) ||
          !eventTime || !ev || typeof ev.payload !== 'object' || ev.payload === null) {
        rejected++;
        outcomes.push({ event_id: eventId, status: 'rejected', reason: 'Invalid event fields.' });
        return;
      }
      const payload = JSON.stringify(ev.payload);
      if (payload.length > POS_MAX_EVENT_CHARS_) {
        rejected++;
        outcomes.push({ event_id: eventId, status: 'rejected', reason: 'Payload exceeds size limit.' });
        return;
      }
      const hash = posHash_(payload);
      // Search only the event_id column. Stable event IDs must be unique across all sources.
      const last = sh.getLastRow();
      let existingRow = 0;
      if (last >= 2) {
        const found = sh.getRange(2, 1, last - 1, 1).createTextFinder(eventId).matchEntireCell(true).findNext();
        if (found) existingRow = found.getRow();
      }
      if (existingRow) {
        const prior = sh.getRange(existingRow, 1, 1, POS_EVENTS_HEADER_.length).getDisplayValues()[0];
        if (String(prior[1]) === sourceId && String(prior[5]) === hash) {
          duplicates++;
          outcomes.push({ event_id: eventId, status: 'already_processed' });
        } else {
          rejected++;
          outcomes.push({ event_id: eventId, status: 'rejected', reason: 'Event ID already exists with a different source or payload.' });
        }
        return;
      }
      sh.appendRow([eventId, sourceId, eventType, eventTime, payload, hash, now_(), schemaVersion]);
      accepted++;
      outcomes.push({ event_id: eventId, status: 'accepted' });
    });
    return { ok: true, source_id: sourceId, accepted: accepted, already_processed: duplicates, rejected: rejected, outcomes: outcomes };
  });
}

function posQuery_(body) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('POS_QUERY_TOKEN');
  if (!expected || !body.token || posHash_(String(body.token)) !== posHash_(expected)) {
    return { ok: false, error: 'POS query authentication failed.' };
  }
  const sourceId = String(body.source_id || '');
  const eventType = String(body.event_type || '');
  const from = String(body.from || '');
  const to = String(body.to || '');
  const limit = Math.max(1, Math.min(200, Number(body.limit) || 100));
  const offset = Math.max(0, Number(body.offset) || 0);
  const sh = posEventsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, events: [], next_offset: null };
  // Bounded result page; query scans newest-first. For large production volumes,
  // replace this sheet-backed store with a database indexed by source/time.
  const start = Math.max(2, last - Math.min(5000, last - 1) + 1);
  const rows = sh.getRange(start, 1, last - start + 1, POS_EVENTS_HEADER_.length).getDisplayValues();
  const matches = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (sourceId && r[1] !== sourceId) continue;
    if (eventType && r[2] !== eventType) continue;
    if (from && r[3] < from) continue;
    if (to && r[3] >= to) continue;
    matches.push({ event_id: r[0], source_id: r[1], event_type: r[2], event_time: r[3],
      payload: JSON.parse(r[4]), received_at: r[6], schema_version: r[7] });
  }
  const page = matches.slice(offset, offset + limit);
  const next = offset + page.length < matches.length ? offset + page.length : null;
  return { ok: true, events: page, next_offset: next, scanned_rows: rows.length };
}
