/**
 * Job Application Tracker — Gmail sync
 * ------------------------------------
 * Scans your recent inbox and matches each email against the companies already
 * in your tracker (the Gist CSV). No manual labelling needed — your application
 * list IS the filter. Matches auto-update status and set an alert flag (❗).
 * Recruiting emails from companies NOT in the tracker auto-create a flagged
 * entry: company + role are extracted from the email and the role is matched
 * against the CV rules defined in the tracker's ⚙ settings (tracker-settings.json).
 *
 * SETUP (one-time):
 *  1. Project Settings ▸ Script Properties ▸ add:
 *       GITHUB_TOKEN = your GitHub token (classic, with `gist` scope)
 *       GIST_ID      = the Gist ID shown in the tracker's sync card
 *  2. (optional) SCAN_DAYS — how far back to scan the inbox (default 5)
 *  3. Run `syncGmail` once to authorise.
 *  4. Deploy ▸ New deployment ▸ Web app ▸ Execute as "Me", Access "Anyone".
 *     Copy the /exec URL and paste it into the tracker's "Check email" button.
 *
 * Non-destructive: never labels, moves, reads-flags, or deletes any mail. It
 * only reads. Every run rescans the whole window, so a dismissed alert will
 * reappear while its email is still within the last SCAN_DAYS days.
 */

var FILENAME = 'job-applications.csv';
var SETTINGS_FILE = 'tracker-settings.json';   // { cvMap: [{ cv, keywords[], category }] }
var DEFAULT_SCAN_DAYS = 2;

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

/**
 * DIAGNOSTIC — run this directly in the Apps Script editor (select it in the
 * function dropdown, hit Run, then check the execution log). No redeploy needed.
 * Answers: which Gmail account the script scans, whether that account can see
 * the Uniphore/Workday email, and whether it is sitting in spam.
 */
function debugFindEmail() {
  Logger.log('Script runs as: ' + Session.getActiveUser().getEmail());
  var q1 = GmailApp.search('from:myworkday.com newer_than:7d', 0, 10);
  Logger.log('from:myworkday.com (last 7d): ' + q1.length + ' thread(s)');
  q1.forEach(function (t) {
    var m = t.getMessages()[t.getMessageCount() - 1];
    Logger.log('  - "' + m.getSubject() + '" from ' + m.getFrom() + ' at ' + m.getDate());
  });
  var q2 = GmailApp.search('uniphore newer_than:7d', 0, 10);
  Logger.log('uniphore (normal search): ' + q2.length + ' thread(s)');
  var q3 = GmailApp.search('in:anywhere uniphore newer_than:7d', 0, 10);
  Logger.log('uniphore in:anywhere (incl spam/trash): ' + q3.length + ' thread(s)');
  q3.forEach(function (t) {
    var m = t.getMessages()[t.getMessageCount() - 1];
    Logger.log('  - "' + m.getSubject() + '" | inSpam: ' + t.isInSpam() + ' | inTrash: ' + t.isInTrash() +
      ' | labels: ' + t.getLabels().map(function (l) { return l.getName(); }).join(', '));
  });
  // What the real scan query sees right now:
  var days = Number(PropertiesService.getScriptProperties().getProperty('SCAN_DAYS')) || DEFAULT_SCAN_DAYS;
  var q4 = GmailApp.search('newer_than:' + days + 'd -in:sent -in:chats -in:draft', 0, 100);
  Logger.log('Real scan query (newer_than:' + days + 'd): ' + q4.length + ' thread(s); newest 3:');
  q4.slice(0, 3).forEach(function (t) {
    var m = t.getMessages()[t.getMessageCount() - 1];
    Logger.log('  - "' + m.getSubject() + '" from ' + m.getFrom() + ' at ' + m.getDate());
  });
}

function syncGmail() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var gistId = props.getProperty('GIST_ID');
  var days = Number(props.getProperty('SCAN_DAYS')) || DEFAULT_SCAN_DAYS;
  if (!token || !gistId) {
    Logger.log('Missing GITHUB_TOKEN or GIST_ID in Script Properties. Aborting.');
    return { ok: false, error: 'Missing GITHUB_TOKEN or GIST_ID' };
  }

  var data = readEntries(token, gistId);      // { header, rows }
  if (!data) { Logger.log('Could not read Gist. Aborting.'); return { ok: false, error: 'Could not read Gist' }; }
  var rows = data.rows;
  if (!rows.length) { Logger.log('No applications in tracker yet.'); return { ok: true, scanned: 0, matched: 0 }; }

  // Scan all recent received mail (inbox + archived), every run, no labels/tags.
  var query = 'newer_than:' + days + 'd -in:sent -in:chats -in:draft';
  var threads = GmailApp.search(query, 0, 100);
  if (!threads.length) { Logger.log('No recent mail to scan.'); return { ok: true, scanned: 0, matched: 0 }; }

  var changed = false;
  var matched = 0;
  var total = threads.length;
  setProgress(token, gistId, 'sync:0/' + total);

  threads.forEach(function (thread, idx) {
    // Report progress every few threads so the tracker can show a live count.
    if (idx > 0 && idx % 5 === 0) setProgress(token, gistId, 'sync:' + idx + '/' + total);
    var msg = thread.getMessages()[thread.getMessageCount() - 1]; // latest message in thread
    var fromRaw = msg.getFrom();                                  // "Name <addr@domain>"
    var fromEmail = extractEmail(fromRaw);
    var fromName = fromRaw.replace(/<[^>]*>/, '').replace(/"/g, '').trim();
    var subject = msg.getSubject() || '';
    var body = (msg.getPlainBody() || '').slice(0, 2000);

    var entry = matchEntry(rows, fromEmail, fromName, subject, body);
    var cls = classify(subject + ' ' + body);

    // Unknown company + application CONFIRMATION => you applied but never logged it.
    // Auto-create the entry (flagged for review) and match the role to a CV rule.
    // Only confirmations qualify: outcome-shaped emails (a stray "unfortunately"
    // in a personal conversation) must never invent an application.
    if (!entry) {
      if (cls !== 'Ack') {
        // Job-outcome-shaped email from an unknown sender — never auto-create
        // from these, but LOG it so silent skips are debuggable.
        if (cls) Logger.log('Skipped (' + cls + ', unknown company) "' + subject + '" from ' + fromEmail);
        return; // leave everything else completely untouched
      }
      var created = autoCreate(rows, data.cvMap, fromEmail, fromName, subject, body, msg.getDate(), cls);
      if (created) {
        rows.push(created); // so later emails in this run match it instead of duplicating
        changed = true;
        matched++;
        Logger.log('Auto-added ' + created.company + (created.role ? ' — ' + created.role : '') +
          (created.cv ? ' [' + created.cv + ']' : '') + ' from "' + subject + '"');
      } else {
        Logger.log('Ack but NOT auto-added (company unclear or duplicate) "' + subject + '" from ' + fromEmail);
      }
      return;
    }
    if (cls !== 'Rejected' && cls !== 'Offer' && cls !== 'Interview' && cls !== 'Screening') {
      // Confirmation, OTP, or no clear signal — not worth an alert.
      Logger.log((cls === 'Ack' ? 'Ack' : 'No signal') + ' (no alert) "' + subject + '" -> ' + entry.company);
      return;
    }

    var newStatus = cls;
    // Only alert when the email tells us something NEW. If the entry is already
    // at (or past) this status — e.g. already Rejected — rescans stay silent,
    // so a dismissed alert never comes back for the same news.
    if (!shouldAdvance(entry.status, newStatus)) {
      Logger.log('Known (no alert) "' + subject + '" -> ' + entry.company + ' (already ' + entry.status + ')');
      return;
    }
    entry.status = newStatus;
    entry.alert = '1';
    entry.alertMsg = describe(newStatus, entry.company);
    entry.updatedAt = String(new Date().getTime());
    changed = true;
    matched++;
    Logger.log('Matched "' + subject + '" -> ' + entry.company + ' (' + entry.alertMsg + ')');
  });

  if (changed) {
    writeEntries(token, gistId, data.header, rows);
    Logger.log('Gist updated.');
  }
  // Restore the normal description so the tracker stops reading progress.
  setProgress(token, gistId, 'Job Application Tracker — data (private)');
  Logger.log('Done. Scanned ' + threads.length + ' thread(s), flagged ' + matched + '.');
  return { ok: true, scanned: threads.length, matched: matched };
}

/* ---------- Matching ---------- */

function matchEntry(rows, fromEmail, fromName, subject, body) {
  var domainRoot = domainOf(fromEmail);                 // e.g. "teya" from careers@teya.com
  // Match on SENDER + SUBJECT only, never the body — otherwise a jobs digest or
  // newsletter that merely lists company names would match every tracked company.
  var haystack = norm(fromName + ' ' + subject);
  var tokens = haystack.split(' ');                     // "SourceWhale" -> one token "sourcewhale"
  var best = null, bestLen = 0;

  rows.forEach(function (e) {
    var company = norm(e.company);
    if (company.length < 3) return;
    var companyNoSpace = company.replace(/ /g, '');
    var hit = false;
    // Strong: sender domain root equals the company (spaces ignored).
    if (domainRoot && domainRoot.length >= 3 && domainRoot === companyNoSpace) hit = true;
    // Multi-word company appears verbatim in the text.
    if (!hit && company.indexOf(' ') !== -1 && new RegExp('\\b' + escapeReg(company) + '\\b').test(haystack)) hit = true;
    // Spaceless company is a whole token (catches "SourceWhale", and single-word names).
    if (!hit && companyNoSpace.length >= 3 && tokens.indexOf(companyNoSpace) !== -1) hit = true;
    if (hit && company.length > bestLen) { best = e; bestLen = company.length; }
  });
  return best;
}

/* ---------- Auto-create (unknown company + recruiting email) ---------- */

function autoCreate(rows, cvMap, fromEmail, fromName, subject, body, msgDate, cls) {
  var company = extractCompany(fromEmail, fromName, subject, body);
  if (!company) return null;
  // Final dedup guard against near-matches ("Proactive Appointments" vs
  // "Proactive.IT Appointments Ltd."): substring either way, or one name's
  // words being a subset of the other's.
  var cn = norm(company).replace(/ /g, '');
  if (cn.length < 3) return null;
  var cTokens = norm(company).split(' ').filter(function (w) { return w.length >= 2; });
  for (var i = 0; i < rows.length; i++) {
    var rNorm = norm(rows[i].company);
    var rn = rNorm.replace(/ /g, '');
    if (!rn) continue;
    if (rn === cn || rn.indexOf(cn) !== -1 || cn.indexOf(rn) !== -1) return null;
    var rTokens = rNorm.split(' ').filter(function (w) { return w.length >= 2; });
    if (tokenSubset(cTokens, rTokens) || tokenSubset(rTokens, cTokens)) return null;
  }
  var role = extractRole(subject, body);
  var rule = matchCv(role, cvMap);
  var now = new Date();
  var status = (cls === 'Ack') ? 'Applied' : cls;
  return {
    id: now.getTime().toString(36) + Math.floor(Math.random() * 1e8).toString(36),
    createdAt: String(now.getTime()),
    updatedAt: String(now.getTime()),
    company: company,
    role: role || '',
    category: rule ? (rule.category || 'Other') : 'Other',
    date: isoDate(msgDate || now),
    cv: rule ? rule.cv : '',
    status: status,
    link: '', jd: '',
    notes: 'Auto-added from email',
    alert: '1',
    alertMsg: (cls === 'Ack' ? 'Auto-added from email' : describe(status, company) + ' (auto-added)') + ' — review'
  };
}

// Words that mean a phrase is a JOB TITLE, not a company name. Workday-style
// confirmations say "thank you for applying to the Product Management Intern",
// which must not be read as a company called "Product Management Intern".
var ROLE_HINT = /\b(manager|management|intern(ship)?|analyst|engineer(ing)?|associate|consultant|consulting|director|graduate|developer|designer|scientist|specialist|coordinator|executive|officer|assistant|apprentice|researcher)\b/i;

function extractCompany(fromEmail, fromName, subject, body) {
  var text = subject + '\n\n' + String(body || '').slice(0, 800);
  // Same soft-wrap unwrap + hypothetical-clause guard as classify()/extractRole().
  text = text.replace(/([^\n])\n(?!\n)/g, '$1 ');
  text = text.replace(/\bif\b[^.!?\n]*[.!?\n]/gi, ' ');
  var pats = [
    /application (?:was |has been )?sent to ([^!,.\n(]{2,50})/i,
    /thanks? (?:so much )?for applying (?:to|at) ([^!,.\n(]{2,50})/i,
    /thank you (?:for applying (?:to|at)|from) ([^!,.\n(]{2,50})/i,
    /your application (?:to|with|at) ([^!,.\n(]{2,50})/i,
    /your interest in (?:joining )?([^!,.\n(]{2,50})/i,
    /(?:position|role|opening|team) at ([^!,.\n(]{2,50})/i
  ];
  for (var i = 0; i < pats.length; i++) {
    var m = pats[i].exec(text);
    // Skip captures that look like a job title — fall through to the sender name.
    if (m) { var c = cleanCompany(m[1]); if (c && !ROLE_HINT.test(c)) return c; }
  }
  // Sender display name: "Megan Pickett - SourceWhale" -> "SourceWhale";
  // "Encord Hiring Team" -> "Encord". Skip if it just looks like a person.
  var seg = String(fromName || '').split(/[-–—|·]/).pop();
  var c2 = cleanCompany(seg.replace(/\b(hiring|talent|recruit\w*|careers?|people|team|hr|acquisition|no[- ]?reply|notifications?|jobs?)\b/gi, ' '));
  if (c2 && !looksLikePerson(c2)) return c2;
  // Company domain, incl. the subdomain trick ATS mailers use (x@company.teamtailor-mail.com).
  var d = companyFromDomain(fromEmail);
  if (d) return d.charAt(0).toUpperCase() + d.slice(1);
  return null;
}

function cleanCompany(s) {
  s = String(s || '').replace(/["'’]s\b/g, '').replace(/["'”“]/g, '')
    .replace(/^\s*(the|at|to)\s+/i, '').replace(/\s+/g, ' ').trim()
    .replace(/[.,;:!\s]+$/, '');
  if (s.length < 2 || s.length > 40) return null;
  if (/\b(linkedin|indeed|glassdoor|otta|builtin|built in|welcome|congratulations|your|our|this|position|role|job)\b/i.test(s)) return null;
  return s;
}

function looksLikePerson(s) {
  var words = s.trim().split(/\s+/);
  return words.length === 2 && /^[A-Z][a-z]+$/.test(words[0]) && /^[A-Z][a-z]+$/.test(words[1]);
}

function companyFromDomain(email) {
  var m = String(email || '').match(/@([^>\s]+)/);
  if (!m) return '';
  var parts = m[1].toLowerCase().split('.');
  var ats = ['ashbyhq', 'greenhouse', 'greenhouse-mail', 'lever', 'lever-mail', 'hire', 'workday', 'myworkday',
    'icims', 'smartrecruiters', 'teamtailor', 'teamtailor-mail', 'workablemail', 'workable', 'bamboohr',
    'gmail', 'googlemail', 'outlook', 'hotmail', 'notifications', 'linkedin', 'indeed', 'mail', 'email'];
  var root = parts.length >= 2 ? parts[parts.length - 2] : parts[0];
  if (ats.indexOf(root) === -1) return root;
  // ATS root: the real company is often the subdomain (sourcewhale.teamtailor-mail.com)
  if (parts.length >= 3 && ats.indexOf(parts[0]) === -1 && parts[0].length >= 3) return parts[0];
  return '';
}

function extractRole(subject, body) {
  var text = subject + '\n\n' + String(body || '').slice(0, 800);
  // Unwrap soft line breaks, then drop hypothetical clauses ("if you do not get
  // selected ... for this role at this time") — same guards as classify().
  text = text.replace(/([^\n])\n(?!\n)/g, '$1 ');
  text = text.replace(/\bif\b[^.!?\n]*[.!?\n]/gi, ' ');
  // "applying to/for the X" — Workday-style confirmations put the job title here
  // ("thank you for applying to the Product Management Intern."), but so do
  // filler sentences ("applying for a job can be stressful"), so only trust a
  // capture that actually reads like a job title. Runs first: it is the most
  // precise pattern, validated against ROLE_HINT.
  var re2 = /applying (?:to|for) (?:the )?([^.,\n!:;]{3,60}?)(?: position| role| opening| vacancy| at |[.,!\n;:])/gi;
  var m2;
  while ((m2 = re2.exec(text))) {   // try every occurrence — subject may name the company, body the role
    var r2 = m2[1].replace(/^\s*(the|a|an|our|your)\s+/i, '').replace(/\s+/g, ' ').trim();
    if (r2.length >= 3 && r2.length <= 60 && ROLE_HINT.test(r2) && !/@/.test(r2)) return r2;
  }
  var pats = [
    /for the ([^.,\n!:;]{3,60}?) (?:position|role|opening|vacancy)/i,
    /application (?:for|to) (?:the )?([^.,\n!:;]{3,60}?) (?:position|role)/i,
    /application for (?:the )?([^.,\n!:;]{3,60}?) was sent/i,
    /interest in (?:the )?([^.,\n!:;]{3,60}?) (?:position|role)/i,
    /(?:position|role) of ([^.,\n!:;]{3,60})/i,
    /([^.,\n!:;]{3,60}?) (?:position|role) at /i
  ];
  for (var i = 0; i < pats.length; i++) {
    var m = pats[i].exec(text);
    if (m) {
      var r = m[1].replace(/^\s*(the|a|an|our|your)\s+/i, '').replace(/\s+/g, ' ').trim();
      if (r.length >= 3 && r.length <= 60 && !/@/.test(r)) return r;
    }
  }
  return '';
}

// Pick the CV rule whose keyword best matches the role title ("similar", not exact):
// each rule lists keywords ("product manager, pm, apm"); any whole-token/phrase hit
// counts, longest keyword wins so specific rules beat generic ones.
function matchCv(role, cvMap) {
  if (!role || !cvMap || !cvMap.length) return null;
  var r = ' ' + norm(role) + ' ';
  var rSolid = r.replace(/ /g, '');
  var best = null, bestLen = 0;
  cvMap.forEach(function (rule) {
    (rule.keywords || []).forEach(function (k) {
      var kk = norm(k);
      if (!kk) return;
      var hit = r.indexOf(' ' + kk + ' ') !== -1 ||
        (kk.length >= 6 && rSolid.indexOf(kk.replace(/ /g, '')) !== -1);
      if (hit && kk.length > bestLen) { best = rule; bestLen = kk.length; }
    });
  });
  return best;
}

function tokenSubset(a, b) {
  if (!a.length) return false;
  return a.every(function (w) { return b.indexOf(w) !== -1; });
}

function isoDate(d) {
  return d.getFullYear() + '-' +
    ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
    ('0' + d.getDate()).slice(-2);
}

function classify(text) {
  var t = ' ' + text.toLowerCase() + ' ';
  // Unwrap soft line breaks first: plain-text emails hard-wrap at ~75 chars, so
  // a conditional sentence can span several lines. Single newlines become
  // spaces; blank lines (paragraph breaks) survive as sentence boundaries.
  t = t.replace(/([^\n])\n(?!\n)/g, '$1 ');
  // Drop hypothetical clauses ("if your profile is not selected, don't be
  // discouraged") so conditional language in confirmations can't read as an outcome.
  t = t.replace(/\bif\b[^.!?\n]*[.!?\n]/g, ' ');
  // Order matters: strong outcomes first, then acknowledgement, so a rejection
  // that opens with "thank you for applying" is still read as a rejection.
  if (/(unfortunately|not (be )?(moving|to move|move|progress|proceed)(ing)? forward|won'?t be progressing|(decided|decision) (not|to not|not to)|regret to inform|other candidates|not (be )?(success|selected)|will not be proceeding|not to progress|wish(ing)? you (success|luck|well|(all )?the best) (in|with|for) your)/.test(t))
    return 'Rejected';
  if (/(pleased to offer|offer of employment|job offer|formally offer|we('| a)re delighted to offer|extend an offer)/.test(t))
    return 'Offer';
  if (/(invite you|invitation to interview|like to (invite|arrange|set up)|schedule (a|an|your)? ?(call|chat|interview|meeting)|book a (time|slot|call)|your availability|phone screen|video (call|interview)|next (stage|round)|move (you|your application) (forward|to the next))/.test(t))
    return 'Interview';
  if (/(assessment|take[- ]?home|case study|technical test|coding (test|challenge)|complete (a|an|the) (exercise|task)|questionnaire|screening call|screening stage)/.test(t))
    return 'Screening';
  // Application confirmation / acknowledgement — no action needed, don't alert.
  if (/(thank(s| you)( so much)? for (applying|your application)|received your application|application (has been )?received|we('| ha)ve received your|thanks for your (interest|application)|glad you found us|successfully (applied|submitted)|application (was )?submitted|your application (was|has been) sent|got your application|application received)/.test(t))
    return 'Ack';
  // OTP / verification / login mail from an ATS — never worth an alert.
  if (/(otp|one[- ]?time (code|password|passcode)|verification code|verify your (email|account)|confirm your email|sign[- ]?in code|login code|magic link)/.test(t))
    return 'Ack';
  return null; // matched sender but no clear signal — stay silent, don't guess
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
  if (next === current) return false;                                // same status — nothing to do
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
  var out;
  if (!file) out = { header: baseHeader(), rows: [] };
  else {
    var text = file.truncated && file.raw_url
      ? UrlFetchApp.fetch(file.raw_url).getContentText()
      : (file.content || '');
    out = parseCsv(text);
  }
  // CV auto-match rules, maintained from the tracker's ⚙ settings panel.
  out.cvMap = [];
  var sf = gist.files && gist.files[SETTINGS_FILE];
  if (sf && sf.content) {
    try { out.cvMap = (JSON.parse(sf.content).cvMap || []); } catch (e) {}
  }
  return out;
}

/** Description-only patch — the tracker polls this for a live progress count. */
function setProgress(token, gistId, text) {
  ghFetch('https://api.github.com/gists/' + gistId, 'patch', token, { description: text });
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
