/**
 * Job Application Tracker — Gmail sync
 * ------------------------------------
 * Scans a Gmail label for recruiting emails, matches each to an application
 * in your private Gist CSV, auto-updates its status, and sets an alert flag
 * so the tracker shows a ❗ badge.
 *
 * SETUP (one-time):
 *  1. Project Settings ▸ Script Properties ▸ add:
 *       GITHUB_TOKEN = your GitHub token (classic, with `gist` scope)
 *       GIST_ID      = the Gist ID shown in the tracker's sync card
 *  2. (optional) GMAIL_LABEL — defaults to "Jobs"
 *  3. Run `syncGmail` once to authorise.
 *  4. Deploy ▸ New deployment ▸ Web app ▸ Execute as "Me", Access "Anyone".
 *     Copy the /exec URL and paste it into the tracker's "Check email" button.
 *
 * Non-destructive: processed threads get a "Jobs/Processed" label so they're
 * never handled twice. Nothing is deleted or archived.
 */

var FILENAME = 'job-applications.csv';
var DEFAULT_LABEL = 'Jobs';
var PROCESSED_LABEL = 'Jobs/Processed';

// Company legal suffixes stripped before matching.
var SUFFIXES = /\b(inc|incorporated|ltd|limited|llc|plc|gmbh|co|corp|corporation|group|technologies|technology|labs|the)\b/g;

// Status ordering so we never move an application backwards (e.g. Interview -> Screening).
var RANK = { Applied: 0, Screening: 1, Interview: 2, Offer: 3 };

/** Called by the tracker's "Check email" button (JSONP). */
function doGet(e) {
  var summary;
  try { summary = syncGmail(); }
  catch (err) { summary = { ok: false, error: String(err) }; }
  var json = JSON.stringify(summary || {});
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + json + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function syncGmail() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var gistId = props.getProperty('GIST_ID');
  var labelName = props.getProperty('GMAIL_LABEL') || DEFAULT_LABEL;
  if (!token || !gistId) {
    Logger.log('Missing GITHUB_TOKEN or GIST_ID in Script Properties. Aborting.');
    return { ok: false, error: 'Missing GITHUB_TOKEN or GIST_ID' };
  }

  var processed = getOrCreateLabel(PROCESSED_LABEL);
  var query = 'label:"' + labelName + '" -label:"' + PROCESSED_LABEL + '" newer_than:14d';
  var threads = GmailApp.search(query, 0, 40);
  if (!threads.length) { Logger.log('No new job emails.'); return { ok: true, scanned: 0, matched: 0 }; }

  var data = readEntries(token, gistId);      // { header, rows }
  if (!data) { Logger.log('Could not read Gist. Aborting.'); return { ok: false, error: 'Could not read Gist' }; }
  var rows = data.rows;
  var changed = false;
  var matched = 0;

  threads.forEach(function (thread) {
    var msg = thread.getMessages()[thread.getMessageCount() - 1]; // latest message in thread
    var fromRaw = msg.getFrom();                                  // "Name <addr@domain>"
    var fromEmail = extractEmail(fromRaw);
    var fromName = fromRaw.replace(/<[^>]*>/, '').replace(/"/g, '').trim();
    var subject = msg.getSubject() || '';
    var body = (msg.getPlainBody() || '').slice(0, 2000);

    var entry = matchEntry(rows, fromEmail, fromName, subject, body);
    if (entry) {
      var newStatus = classify(subject + ' ' + body);
      var reason = describe(newStatus, entry.company);
      if (newStatus && shouldAdvance(entry.status, newStatus)) entry.status = newStatus;
      entry.alert = '1';
      entry.alertMsg = reason;
      entry.updatedAt = String(new Date().getTime());
      changed = true;
      matched++;
      Logger.log('Matched "' + subject + '" -> ' + entry.company + ' (' + reason + ')');
    } else {
      Logger.log('No match for "' + subject + '" from ' + fromEmail);
    }
    thread.addLabel(processed); // mark handled either way, so we don't re-scan it
  });

  if (changed) {
    writeEntries(token, gistId, data.header, rows);
    Logger.log('Gist updated.');
  }
  return { ok: true, scanned: threads.length, matched: matched };
}

/* ---------- Matching ---------- */

function matchEntry(rows, fromEmail, fromName, subject, body) {
  var domainRoot = domainOf(fromEmail);                 // e.g. "teya" from careers@teya.com
  var haystack = norm(fromName + ' ' + subject + ' ' + body + ' ' + domainRoot);
  var best = null, bestLen = 0;

  rows.forEach(function (e) {
    var company = norm(e.company);
    if (company.length < 3) return;
    var hit = false;
    // Strong: sender domain root matches the company (either direction).
    if (domainRoot && domainRoot.length >= 3 &&
        (company.indexOf(domainRoot) !== -1 || domainRoot.indexOf(company) !== -1)) {
      hit = true;
    }
    // Otherwise: company name appears as a whole token in name/subject/body.
    if (!hit && new RegExp('\\b' + escapeReg(company) + '\\b').test(haystack)) hit = true;
    if (hit && company.length > bestLen) { best = e; bestLen = company.length; }
  });
  return best;
}

function classify(text) {
  var t = ' ' + text.toLowerCase() + ' ';
  if (/(unfortunately|not moving forward|not be (moving|progress)|won'?t be progressing|decided (not|to not)|regret to inform|other candidates|not (be )?(success|selected)|will not be proceeding)/.test(t))
    return 'Rejected';
  if (/(pleased to offer|offer of employment|job offer|formally offer|we('| a)re delighted to offer|extend an offer)/.test(t))
    return 'Offer';
  if (/(interview|schedule a (call|chat)|book a time|meet the team|your availability|next (stage|round|step)|phone screen|video call|hiring manager|meet with)/.test(t))
    return 'Interview';
  if (/(assessment|take[- ]?home|case study|technical test|coding (test|challenge)|exercise|questionnaire|screening call|screening stage)/.test(t))
    return 'Screening';
  return null; // general update — flag but don't change status
}

function describe(status, company) {
  var map = {
    Rejected: 'Rejection from ' + company,
    Offer: 'Offer from ' + company,
    Interview: 'Interview / next-stage email from ' + company,
    Screening: 'Assessment or screening from ' + company
  };
  return map[status] || 'New email from ' + company;
}

function shouldAdvance(current, next) {
  if (next === 'Rejected' || next === 'Offer') return true;          // terminal signals always win
  if (current === 'Rejected' || current === 'Offer' || current === 'Withdrawn') return false;
  return (RANK[next] || 0) > (RANK[current] || 0);                   // only move forward
}

/* ---------- Gist read / write (CSV, matches the tracker exactly) ---------- */

function readEntries(token, gistId) {
  var res = ghFetch('https://api.github.com/gists/' + gistId, 'get', token, null);
  if (res.getResponseCode() !== 200) { Logger.log('Gist GET ' + res.getResponseCode()); return null; }
  var gist = JSON.parse(res.getContentText());
  var file = gist.files && gist.files[FILENAME];
  if (!file) return { header: baseHeader(), rows: [] };
  var text = file.truncated && file.raw_url
    ? UrlFetchApp.fetch(file.raw_url).getContentText()
    : (file.content || '');
  return parseCsv(text);
}

function writeEntries(token, gistId, header, rows) {
  var csv = toCsv(header, rows);
  var payload = { files: {} };
  payload.files[FILENAME] = { content: csv };
  var res = ghFetch('https://api.github.com/gists/' + gistId, 'patch', token, payload);
  if (res.getResponseCode() !== 200) Logger.log('Gist PATCH ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 150));
}

function ghFetch(url, method, token, payload) {
  return UrlFetchApp.fetch(url, {
    method: method,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    contentType: 'application/json',
    payload: payload ? JSON.stringify(payload) : null,
    muteHttpExceptions: true
  });
}

/* ---------- CSV (mirrors the tracker's format: CRLF, RFC-4180 quoting) ---------- */

function baseHeader() {
  return ['id','createdAt','updatedAt','company','role','category','date','cv','status','link','jd','notes','alert','alertMsg'];
}

function parseCsv(text) {
  if (!text || !text.trim()) return { header: baseHeader(), rows: [] };
  var grid = [], row = [], field = '', i = 0, inQ = false;
  while (i < text.length) {
    var ch = text[i];
    if (inQ) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } inQ = false; i++; continue; }
      field += ch; i++; continue;
    } else {
      if (ch === '"') { inQ = true; i++; continue; }
      if (ch === ',') { row.push(field); field = ''; i++; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(field); grid.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++; continue;
    }
  }
  if (field !== '' || row.length) { row.push(field); grid.push(row); }
  var header = grid[0];
  var rows = [];
  for (var r = 1; r < grid.length; r++) {
    if (grid[r].length === 1 && grid[r][0] === '') continue;
    var obj = {};
    header.forEach(function (h, idx) { obj[h] = grid[r][idx] != null ? grid[r][idx] : ''; });
    if (obj.alert == null) obj.alert = '';
    if (obj.alertMsg == null) obj.alertMsg = '';
    rows.push(obj);
  }
  // Ensure alert columns exist in the header we write back.
  if (header.indexOf('alert') === -1) header.push('alert');
  if (header.indexOf('alertMsg') === -1) header.push('alertMsg');
  return { header: header, rows: rows };
}

function toCsv(header, rows) {
  var out = [header.join(',')];
  rows.forEach(function (e) {
    out.push(header.map(function (c) { return csvCell(e[c]); }).join(','));
  });
  return out.join('\r\n');
}

function csvCell(v) {
  var s = v == null ? '' : String(v);
  if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}

/* ---------- Small helpers ---------- */

function norm(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(SUFFIXES, ' ').replace(/\s+/g, ' ').trim();
}
function domainOf(email) {
  var m = String(email || '').match(/@([^>\s]+)/);
  if (!m) return '';
  var parts = m[1].toLowerCase().split('.');
  // drop common ATS/mail hosts so they don't become the "company"
  var host = parts.length >= 2 ? parts[parts.length - 2] : parts[0];
  var ats = ['ashbyhq','greenhouse','lever','workday','gmail','googlemail','outlook','hotmail','myworkday','icims','smartrecruiters','teamtailor','notifications'];
  return ats.indexOf(host) !== -1 ? '' : host;
}
function extractEmail(from) { var m = String(from).match(/<([^>]+)>/); return m ? m[1] : from; }
function escapeReg(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function getOrCreateLabel(name) { return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }
