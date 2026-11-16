/**
 * Lead tracker for akshen.info  (#/apply form -> Google Sheet + email summary)
 *
 * SETUP (about 5 minutes, once):
 *  1. Create a new Google Sheet (e.g. "Akshen – Leads").
 *  2. In the sheet: Extensions > Apps Script. Delete what's there, paste this whole file, Save.
 *  3. In the toolbar pick the function "setup" and press Run. Approve the permissions
 *     (Google warns "unverified app" because it's your own script: Advanced > Go to project).
 *     This creates the "Leads" tab with headers, filters and a Status dropdown.
 *  4. Deploy > New deployment > type "Web app".
 *       Execute as: Me        Who has access: Anyone
 *     Copy the Web app URL (ends in /exec).
 *  5. In index.html, paste that URL into SITE.funnel.endpoint and push the site.
 *  6. Run "sendTestLead" once (or submit the live form) to check a row and an email arrive.
 *
 * If you edit this script later: Deploy > Manage deployments > edit (pencil) > Version: New version.
 * Otherwise the live URL keeps running the old code.
 */

const CONFIG = {
  SHEET_NAME: 'Leads',
  NOTIFY_EMAIL: '',          // empty = the Google account that owns the script
  EMAIL_ON: 'all',           // 'all' leads, or 'qualified' to only email qualified ones
};

// Column order in the sheet: [header, key in the form data]
const COLUMNS = [
  ['Submitted', 'submitted'],
  ['Qualified', 'qualified'],
  ['Status', '_status'],
  ['Service', 'service'],
  ['Business type', 'business'],
  ['Plans', 'plans'],
  ['Role', 'role'],
  ['Bottleneck', 'bottleneck'],
  ['Timing', 'timing'],
  ['Name', 'name'],
  ['Business name', 'company'],
  ['Email', 'email'],
  ['Phone', 'phone'],
  ['Website / IG', 'website'],
  ['Message', 'message'],
  ['Consent', 'consent'],
  ['UTM source', 'utm_source'],
  ['UTM medium', 'utm_medium'],
  ['UTM campaign', 'utm_campaign'],
  ['UTM content', 'utm_content'],
  ['Ref', 'ref'],
  ['Page', 'page'],
  ['Notes', '_notes'],
];
const STATUSES = ['New', 'Contacted', 'Call booked', 'Proposal sent', 'Won', 'Lost', 'Not a fit'];

// Short words for the email subject
const SHORT = {
  business: {
    'Med spa or aesthetics clinic': 'med spa',
    'Dental, vet, pharmacy, or other clinic': 'clinic',
    'Restaurant, café, or bar': 'restaurant',
    'Car detailing or auto': 'auto detailing',
    'DJ, musician, or visual artist': 'artist',
    'Other local business': 'local business',
    'Brand, event, or creator': 'brand/event',
  },
  plans: {
    'Monthly content': 'monthly content',
    'A one-off shoot day': 'a shoot day',
    'Event coverage or recap': 'event coverage',
    'Music visualizer': 'a visualizer',
    'Editing footage I already have': 'editing',
    'Ads and new leads': 'ads',
  },
  timing: {
    'Right away': 'right away',
    'In the next month': 'next month',
    'Just exploring': 'just exploring',
  },
};

/* ---------------- web app ---------------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = parseBody_(e);
    if (data._gotcha) return json_({ ok: true });            // bot: pretend success, store nothing
    if (!data.email && !data.name) return json_({ ok: false, error: 'empty' });

    lock.waitLock(20000);
    const sheet = getSheet_();
    const row = appendLead_(sheet, data);
    lock.releaseLock();

    try { notify_(data, sheet, row); } catch (err) { console.error('Email failed: ' + err); }
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// Visiting the /exec URL in a browser shows this: a quick "is it live?" check
function doGet() {
  return json_({ ok: true, service: 'akshen lead tracker' });
}

/* ---------------- sheet ---------------- */

function setup() {
  const sheet = getSheet_();
  const n = COLUMNS.length;
  sheet.getRange(1, 1, 1, n).setValues([COLUMNS.map(c => c[0])])
    .setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  if (!sheet.getFilter()) sheet.getRange(1, 1, sheet.getMaxRows(), n).createFilter();

  const statusCol = colOf_('_status');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(rule);

  // highlight qualified leads
  const qCol = columnLetter_(colOf_('qualified'));
  const rules = sheet.getConditionalFormatRules().filter(r => true);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + qCol + '2="yes"')
    .setBackground('#e8f5e9')
    .setRanges([sheet.getRange(2, 1, sheet.getMaxRows() - 1, n)]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.getRange(2, colOf_('submitted'), sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd h:mm am/pm');
  sheet.setColumnWidth(colOf_('message'), 320);
  SpreadsheetApp.getActive().setSpreadsheetTimeZone('America/Toronto');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
}

function appendLead_(sheet, data) {
  if (sheet.getLastRow() === 0) setup();
  // any new field the form sends later gets its own column automatically
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const known = new Set(COLUMNS.map(c => c[1]));
  Object.keys(data).forEach(k => {
    if (k[0] === '_' || known.has(k) || headers.includes(k)) return;
    sheet.getRange(1, headers.length + 1).setValue(k).setFontWeight('bold');
    headers.push(k);
  });

  const byHeader = {};
  COLUMNS.forEach(([h, k]) => { byHeader[h] = data[k]; });
  byHeader['Status'] = 'New';
  byHeader['Submitted'] = data.submitted ? new Date(data.submitted) : new Date();
  headers.forEach(h => { if (!(h in byHeader)) byHeader[h] = data[h]; });

  const values = headers.map(h => clean_(byHeader[h]));
  sheet.appendRow(values);
  return sheet.getLastRow();
}

// Stops "=HYPERLINK(...)"-style formula injection from form input
function clean_(v) {
  if (v instanceof Date) return v;
  if (v === undefined || v === null) return '';
  let s = String(v).slice(0, 5000);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

/* ---------------- email ---------------- */

function notify_(d, sheet, row) {
  const qualified = d.qualified === 'yes';
  if (CONFIG.EMAIL_ON === 'qualified' && !qualified) return;
  const to = CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();

  // Subject + body read like the client wrote to you. The site sends the finished note (_inquiry);
  // older form versions don't, so a simple one is built here instead.
  const who = d.company || d.name || 'a new client';
  const subject = (qualified ? '\u2b50 ' : '') + 'Inquiry from ' + who;
  const note = String(d._inquiry || fallbackNote_(d)).replace(/\n--\n[\s\S]*$/, '').trim();   // drop the site's sign-off line

  const url = sheet.getParent().getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + row;
  const source = [d.utm_source, d.utm_medium, d.utm_campaign, d.ref].filter(String).join(' / ');
  const footer = [qualified ? 'Qualified' : 'Not qualified', d.service ? 'from the ' + d.service + ' page' : '',
                  source ? 'source: ' + source : ''].filter(String).join(' \u00b7 ');
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#222">' +
      note.split(/\n{2,}/).map(p => '<p style="margin:0 0 12px">' + esc_(p).replace(/\n/g, '<br>') + '</p>').join('') +
      '<hr style="border:0;border-top:1px solid #eee;margin:20px 0 10px">' +
      '<p style="font-size:12px;color:#999;margin:0">' + esc_(footer) + ' &nbsp;\u00b7&nbsp; <a href="' + url + '" style="color:#999">Open in the lead tracker</a></p>' +
    '</div>';

  const opts = { htmlBody: html, name: d.name ? d.name + ' (via akshen.info)' : 'akshen.info' };
  if (d.email && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) opts.replyTo = d.email;   // hit Reply to answer them directly
  MailApp.sendEmail(to, subject, note + '\n\n--\n' + footer + '\n' + url, opts);
}

function fallbackNote_(d) {
  const facts = [[d.name, d.company].filter(String).join(', ')];
  if (d.plans) facts.push('Looking for: ' + d.plans.toLowerCase());
  if (d.timing) facts.push('Timing: ' + d.timing.toLowerCase());
  if (d.bottleneck && d.bottleneck !== 'Something else') facts.push('Main challenge: ' + d.bottleneck.toLowerCase());
  facts.push([d.email, d.phone, d.website].filter(String).join(' \u00b7 '));
  return 'Hi Akshen,\n\n' + (d.message || 'I filled out the form on your site. Here are my details.') + '\n\n' + facts.join('\n');
}

/* ---------------- helpers ---------------- */

function parseBody_(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (_) {}
  }
  return e.parameter || {};
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function colOf_(key) { return COLUMNS.findIndex(c => c[1] === key) + 1; }
function columnLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function esc_(v) { return String(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

// Run from the editor to check the sheet + email without touching the site
function sendTestLead() {
  doPost({ postData: { contents: JSON.stringify({
    submitted: new Date().toISOString(), qualified: 'yes', service: 'clinic',
    business: 'Med spa or aesthetics clinic', plans: 'Ads and new leads', role: 'Owner',
    bottleneck: 'Not enough new customers', timing: 'In the next month',
    name: 'Test Lead', company: 'Test Med Spa', email: 'test@example.com', consent: 'yes',
    utm_source: 'test', page: 'https://akshen.info/#/apply/clinic',
  }) } });
}
