/****************************************************************************************
 * TikTok uAttend → Google Sheets punch sync
 * - Pulls the uAttend Punch Report (raw IN/OUT pairs) through the uAttend API
 * - Saves it on the "TikTok Punches" tab of your uAttend Timecards sheet
 * - Serves it to the TikTok W.E. Report's  ⟳ PULL FROM UATTEND  button
 * - Optional: runs by itself every morning (installDailyTrigger)
 *
 * SETUP (once)
 * 1. script.google.com → New project → paste this whole file → name it "TikTok Punch Sync" → Save.
 * 2. Project Settings (gear) → Script Properties → Add:
 *      UATTEND_API_KEY   = your uAttend API key
 *      SYNC_PASSWORD     = (optional) a password the report must send with each pull
 * 3. Run  testPull  once (▶ Run) and approve Google's permission prompts.  Check the Log tab.
 * 4. Deploy → New deployment → gear → Web app
 *      Execute as: Me      Who has access: Anyone
 *    → Deploy → copy the Web app URL (ends in /exec). Paste it + SYNC_PASSWORD into the report.
 * 5. (optional) Run  installDailyTrigger  to pull yesterday + today every morning at 8 AM.
 ****************************************************************************************/

const CONFIG = {
  SHEET_ID: '1lmycnNNY-i6Mv6z29soLTG62RtzdKNgOucb890K3Noo', // "uAttend Timecards"
  TAB: 'TikTok Punches',
  LOG_TAB: 'Log',
  // Only keep TikTok punches: every dept code seen in TikTok's punch reports (100/200/300 + 101/201/301 + ONSITE).
  // The report decides what is billed. Leave the list empty [] to keep every department in the uAttend account.
  DEPT_CODES: ['100', '101', '200', '201', '300', '301', 'ONSITE'],
  API: 'https://api.workwelltech.com/reports/punch',
  PAGE_SIZE: 500,
  DAILY_HOUR: 8,
};

const HEADER = ['EMP L NAME','EMP F NAME','EMP##','DATE','IN','OUT','TOTAL','DEPT CODE',
                'IN LOCATION','IN PUNCH METHOD','OUT LOCATION','OUT PUNCH METHOD','DEPT NAME','PUNCH ID','PULLED'];

/* ---------------- uAttend ---------------- */
function apiKey_() {
  const p = PropertiesService.getScriptProperties();
  const k = p.getProperty('UATTEND_API_KEY');
  if (!k) throw new Error('Add UATTEND_API_KEY in Project Settings → Script Properties.');
  return k.trim();
}

function fetchPunches_(start, end) {           // start/end = 'yyyy-mm-dd'
  const out = [];
  for (let page = 0; page < 200; page++) {
    const res = UrlFetchApp.fetch(CONFIG.API, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey_() },
      payload: JSON.stringify({ StartDate: start, EndDate: end, UsePaging: true, PageSize: CONFIG.PAGE_SIZE, PageNumber: page }),
    });
    const code = res.getResponseCode(), body = res.getContentText();
    if (code === 401 || code === 403) throw new Error('uAttend refused the API key (' + code + '). Check UATTEND_API_KEY.');
    if (code === 429) { Utilities.sleep(5000); page--; continue; }
    if (code !== 200) throw new Error('uAttend error ' + code + ': ' + body.slice(0, 300));
    const j = JSON.parse(body);
    if (j.Response && j.Response.Success === false) throw new Error('uAttend: ' + JSON.stringify(j.Response.Messages || j.Response));
    const items = j.PunchReportLineItems || [];
    out.push.apply(out, items);
    if (items.length < CONFIG.PAGE_SIZE) break;
  }
  return out;
}

/* uAttend gives dates as yyyy-mm-dd (sometimes with a time part) and times as hh:mm or hh:mm AM */
function d_(v) { if (!v) return ''; const s = String(v); const m = s.match(/(\d{4})-(\d{2})-(\d{2})/); if (m) return m[2] + '/' + m[3] + '/' + m[1]; const m2 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m2 ? s : s; }
function t_(v) {
  if (!v) return '';
  const s = String(v).trim(); let m = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?/);
  if (!m) return s;
  let h = +m[1]; const mm = m[2]; let ap = m[3] ? m[3].toUpperCase() : (h >= 12 ? 'PM' : 'AM');
  if (!m[3]) h = h % 12 || 12;
  return (h < 10 ? '0' : '') + h + ':' + mm + ' ' + ap;
}

function toRow_(it, stamp) {
  return [it.LastName || '', it.FirstName || '', it.PayrollNumber || String(it.UserId || ''),
          d_(it.PunchDate || it.InDate || it.OutDate), t_(it.InTime), t_(it.OutTime), it.Tot || '',
          String(it.DepartmentCode == null ? '' : it.DepartmentCode).trim(),
          it.InPunchLocation || '', it.InPunchMethod || '', it.OutPunchLocation || '', it.OutPunchMethod || '',
          it.DepartmentName || '', String(it.Id || ''), stamp];
}

/* ---------------- Sheet ---------------- */
function tab_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sh = ss.getSheetByName(CONFIG.TAB);
  if (!sh) { sh = ss.insertSheet(CONFIG.TAB); }
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold'); sh.setFrozenRows(1); }
  sh.getRange('D:G').setNumberFormat('@'); // keep dates / times as typed
  return sh;
}
function log_(msg) {
  try { const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID); const lg = ss.getSheetByName(CONFIG.LOG_TAB) || ss.insertSheet(CONFIG.LOG_TAB);
        lg.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), 'TIKTOK PUNCHES ' + msg]); } catch (e) {}
}
const ymd_ = s => { const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2) : String(s).slice(0, 10); };

/** Pull start..end from uAttend and replace those days on the TikTok Punches tab. */
function pull(start, end) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try {
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    let items = fetchPunches_(start, end);
    const keep = CONFIG.DEPT_CODES.map(String);
    if (keep.length) items = items.filter(it => keep.indexOf(String(it.DepartmentCode == null ? '' : it.DepartmentCode).trim().toUpperCase()) >= 0);
    const rows = items.map(it => toRow_(it, stamp));
    const sh = tab_();
    const last = sh.getLastRow();
    let old = last > 1 ? sh.getRange(2, 1, last - 1, HEADER.length).getValues() : [];
    old = old.filter(r => { const k = ymd_(r[3]); return !(k >= start && k <= end); });  // drop the days being re-pulled
    const all = old.concat(rows).sort((a, b) => ymd_(a[3]).localeCompare(ymd_(b[3])) || String(a[0]).localeCompare(String(b[0])) || String(a[4]).localeCompare(String(b[4])));
    if (last > 1) sh.getRange(2, 1, last - 1, HEADER.length).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, HEADER.length).setValues(all);
    log_(start + ' to ' + end + ': ' + rows.length + ' punches saved');
    return rows.length;
  } finally { lock.releaseLock(); }
}

/** Read start..end back from the sheet (no uAttend call). */
function read(start, end) {
  const sh = tab_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADER.length).getDisplayValues().filter(r => { const k = ymd_(r[3]); return k >= start && k <= end; });
}

/* ---------------- Web app for the report's button ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    const pw = PropertiesService.getScriptProperties().getProperty('SYNC_PASSWORD') || '';
    if (pw && p.k !== pw) throw new Error('Wrong or missing sync password.');  // optional: only checked if SYNC_PASSWORD is set
    const start = p.start, end = p.end || p.start;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start || '') || !/^\d{4}-\d{2}-\d{2}$/.test(end || '')) throw new Error('start and end must be yyyy-mm-dd.');
    let pulled = null;
    if (p.action === 'pull') pulled = pull(start, end);
    const rows = read(start, end);
    out = { ok: true, pulled: pulled, count: rows.length, title: 'uAttend ' + start + ' to ' + end, header: HEADER, rows: rows };
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  const json = JSON.stringify(out);
  if (p.callback && /^[\w.$]+$/.test(p.callback))
    return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Daily run + tests ---------------- */
function dailyPull() {               // yesterday + today (catches late OUT punches from overnight shifts)
  const tz = Session.getScriptTimeZone(), now = new Date(), y = new Date(now.getTime() - 86400000);
  pull(Utilities.formatDate(y, tz, 'yyyy-MM-dd'), Utilities.formatDate(now, tz, 'yyyy-MM-dd'));
}
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyPull').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyPull').timeBased().everyDays(1).atHour(CONFIG.DAILY_HOUR).create();
  Logger.log('Daily pull set for ' + CONFIG.DAILY_HOUR + ':00 (' + Session.getScriptTimeZone() + ').');
}
function testPull() {                // last 3 days
  const tz = Session.getScriptTimeZone(), now = new Date(), s = new Date(now.getTime() - 3 * 86400000);
  const n = pull(Utilities.formatDate(s, tz, 'yyyy-MM-dd'), Utilities.formatDate(now, tz, 'yyyy-MM-dd'));
  Logger.log(n + ' punches saved to "' + CONFIG.TAB + '".');
}
