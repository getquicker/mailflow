/**
 * MailFlow backend for GitHub Pages.
 *
 * Set Script Properties:
 *   MAILFLOW_ACCESS_CODE = a long random team passphrase
 *   SPREADSHEET_ID = your Google Sheet ID
 *   ENABLE_LIVE_SENDING = false (keep false while testing)
 *
 * Deploy as Web App: execute as you, access "Anyone" because GitHub Pages calls this API.
 * API calls require a team access code at login and a short-lived server-side session token.
 */
const MF_CONFIG = {
  SHEET_NAME: 'Sheet1',
  HEADER_ROW: 4,
  FIRST_ROW: 5,
  STATUS_COL: 6,
  LOG_SHEET: '_MailFlowLog',
  JOB_SHEET: '_MailFlowJobs',
  MAX_PER_RUN: 90,
  TOKEN_TTL_SECONDS: 21600
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  let result;
  try { result = { ok: true, data: routeGet_(p) }; }
  catch (err) { result = { ok: false, error: String(err && err.message || err) }; }
  const json = JSON.stringify(result);
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) throw new Error('Missing request body.');
    const payload = JSON.parse(raw);
    if (payload.action !== 'send') throw new Error('Unsupported POST action.');
    requireSession_(payload.token);
    if (payload.confirmSend !== true || String(payload.typedConfirmation || '') !== 'SEND') throw new Error('Campaign confirmation is required.');
    const jobId = String(payload.jobId || '');
    if (!/^mf_[A-Za-z0-9_-]{8,80}$/.test(jobId)) throw new Error('Invalid job ID.');
    const jobs = getJobsSheet_();
    if (findJobRow_(jobs, jobId)) throw new Error('This job ID has already been submitted.');
    const jobRow = jobs.getLastRow() + 1;
    jobs.getRange(jobRow, 1, 1, 5).setValues([[jobId, 'processing', new Date(), '', '']]);
    SpreadsheetApp.flush();
    try {
      const result = runCampaign_(payload);
      jobs.getRange(jobRow, 2, 1, 4).setValues([['done', new Date(), JSON.stringify(result), '']]);
    } catch (err) {
      jobs.getRange(jobRow, 2, 1, 4).setValues([['error', new Date(), '', String(err && err.message || err)]]);
    }
  } catch (err) {
    console.error(String(err && err.stack || err));
  }
  return ContentService.createTextOutput('accepted').setMimeType(ContentService.MimeType.TEXT);
}

function routeGet_(p) {
  const action = String(p.action || '');
  if (action === 'login') return login_(String(p.accessCode || ''));
  if (action === 'jobStatus') { requireSession_(p.token); return getJobStatus_(String(p.jobId || '')); }
  requireSession_(p.token);
  if (action === 'logout') return { loggedOut: true };
  if (action === 'contacts') return getContacts_();
  if (action === 'preview') return previewSheet_(String(p.subject || ''), String(p.body || ''));
  throw new Error('Unknown action.');
}

function login_(code) {
  const expected = PropertiesService.getScriptProperties().getProperty('MAILFLOW_ACCESS_CODE');
  if (!expected || expected.length < 16) throw new Error('Backend access code is not configured. Set MAILFLOW_ACCESS_CODE in Script Properties (16+ characters).');
  if (!constantTimeEquals_(code, expected)) throw new Error('Invalid team access code.');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('mf_session_' + token, 'valid', MF_CONFIG.TOKEN_TTL_SECONDS);
  return { token: token, user: 'team-member', expiresInSeconds: MF_CONFIG.TOKEN_TTL_SECONDS };
}

function requireSession_(token) {
  token = String(token || '');
  if (!/^[a-f0-9]{64}$/i.test(token) || CacheService.getScriptCache().get('mf_session_' + token) !== 'valid') throw new Error('Session expired or invalid. Sign in again.');
}

function getContacts_() {
  const sheet = getContactSheet_();
  const rows = getRawContacts_(sheet);
  const contacts = rows.map(c => Object.assign({}, c, {validEmail: validEmail_(c.email), alreadySent: /^SENT\b/i.test(c.status)}));
  return {
    spreadsheetName: SpreadsheetApp.openById(getSpreadsheetId_()).getName(),
    sheetName: sheet.getName(),
    contacts: contacts,
    remainingQuota: MailApp.getRemainingDailyQuota(),
    liveSendingEnabled: isLiveEnabled_(),
    maxPerRun: MF_CONFIG.MAX_PER_RUN
  };
}

function previewSheet_(subject, body) {
  if (!subject.trim() || !body.trim()) throw new Error('Subject and message are required.');
  const contacts = getRawContacts_(getContactSheet_());
  const eligible = contacts.filter(c => c.email && validEmail_(c.email) && !c.status);
  const invalid = contacts.filter(c => c.email && !validEmail_(c.email));
  const sent = contacts.filter(c => /^SENT\b/i.test(c.status));
  return {
    previews: eligible.slice(0, 10).map(c => Object.assign({}, c, {subject: render_(subject, c), body: render_(body, c)})),
    eligibleCount: eligible.length, invalidCount: invalid.length, alreadySentCount: sent.length,
    remainingQuota: MailApp.getRemainingDailyQuota(), maxPerRun: MF_CONFIG.MAX_PER_RUN
  };
}

function runCampaign_(payload) {
  if (!isLiveEnabled_()) throw new Error('Live sending is disabled. Set ENABLE_LIVE_SENDING=true in Script Properties only after testing.');
  const subject = String(payload.subject || '').trim();
  const body = String(payload.body || '').trim();
  if (!subject || !body || subject.length > 250 || body.length > 20000) throw new Error('Subject or message is missing or too long.');
  const sheetSource = String(payload.source || '') === 'sheet';
  let recipients, sheet = null;
  if (sheetSource) {
    sheet = getContactSheet_();
    recipients = getRawContacts_(sheet).filter(c => c.email && !c.status);
  } else {
    recipients = (Array.isArray(payload.recipients) ? payload.recipients : []).map(c => ({
      rowNumber: 0, name: String(c.name || '').trim(), email: String(c.email || '').trim(),
      company: String(c.company || '').trim(), role: String(c.role || '').trim(), status: String(c.status || '')
    })).filter(c => c.email && !c.status);
  }
  const seen = {};
  recipients = recipients.filter(c => {
    const key = c.email.toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return validEmail_(c.email) && !/^SENT\b/i.test(c.status);
  });
  const limit = Math.min(MF_CONFIG.MAX_PER_RUN, MailApp.getRemainingDailyQuota(), recipients.length);
  if (limit < 1) throw new Error('No eligible recipients or no daily quota remains.');
  const log = getLogSheet_();
  let sent = 0, failed = 0, skipped = 0;
  const sentEmails = [];
  for (let i = 0; i < recipients.length; i++) {
    const c = recipients[i];
    if (sent >= limit) { skipped++; continue; }
    try {
      MailApp.sendEmail({to:c.email, subject:render_(subject,c), body:render_(body,c), name:'MailFlow'});
      sent++; sentEmails.push(c.email);
      if (sheetSource && c.rowNumber) sheet.getRange(c.rowNumber, MF_CONFIG.STATUS_COL).setValue('SENT - ' + new Date().toISOString());
      log.appendRow([new Date(), c.email, c.name, c.company, render_(subject,c), 'SENT', '']);
    } catch (err) {
      failed++;
      if (sheetSource && c.rowNumber) sheet.getRange(c.rowNumber, MF_CONFIG.STATUS_COL).setValue('ERROR: ' + String(err).slice(0,120));
      log.appendRow([new Date(), c.email, c.name, c.company, render_(subject,c), 'ERROR', String(err).slice(0,500)]);
    }
    SpreadsheetApp.flush(); Utilities.sleep(350);
  }
  return {sent:sent, failed:failed, skipped:skipped, sentEmails:sentEmails, limit:limit, remainingQuota:MailApp.getRemainingDailyQuota()};
}

function getJobStatus_(jobId) {
  const sheet = getJobsSheet_(), row = findJobRow_(sheet, jobId);
  if (!row) return {state:'not_found', error:'Job not found. Check backend setup and sign in again.'};
  const values = sheet.getRange(row,1,1,5).getValues()[0];
  let result = {};
  if (values[3]) { try { result = JSON.parse(values[3]); } catch (e) {} }
  return Object.assign({jobId:values[0],state:values[1],updatedAt:values[2],error:values[4]||''},result);
}

function getRawContacts_(sheet) {
  const last = sheet.getLastRow();
  if (last < MF_CONFIG.FIRST_ROW) return [];
  const vals = sheet.getRange(MF_CONFIG.FIRST_ROW,1,last-MF_CONFIG.FIRST_ROW+1,Math.max(MF_CONFIG.STATUS_COL,6)).getDisplayValues();
  return vals.map((r,i)=>({
    rowNumber:MF_CONFIG.FIRST_ROW+i, company:String(r[0]||'').trim(), name:String(r[1]||'').trim(),
    role:String(r[2]||'').trim(), email:String(r[3]||'').trim(), outreachStatus:String(r[4]||'').trim(), status:String(r[5]||'').trim()
  })).filter(c=>c.company||c.name||c.email||c.role);
}
function getContactSheet_() {
  const ss=SpreadsheetApp.openById(getSpreadsheetId_()), sheet=ss.getSheetByName(MF_CONFIG.SHEET_NAME);
  if(!sheet) throw new Error('Sheet tab "'+MF_CONFIG.SHEET_NAME+'" not found. Set it in Code.gs.');
  return sheet;
}
function getSpreadsheetId_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id) throw new Error('SPREADSHEET_ID is missing in Apps Script Script Properties.');
  return id;
}
function isLiveEnabled_() {
  return String(PropertiesService.getScriptProperties().getProperty('ENABLE_LIVE_SENDING')||'false').toLowerCase()==='true';
}
function getLogSheet_() {
  const ss=SpreadsheetApp.openById(getSpreadsheetId_()); let s=ss.getSheetByName(MF_CONFIG.LOG_SHEET);
  if(!s){s=ss.insertSheet(MF_CONFIG.LOG_SHEET);s.appendRow(['Timestamp','Email','Name','Company','Subject','Status','Error']);} return s;
}
function getJobsSheet_() {
  const ss=SpreadsheetApp.openById(getSpreadsheetId_()); let s=ss.getSheetByName(MF_CONFIG.JOB_SHEET);
  if(!s){s=ss.insertSheet(MF_CONFIG.JOB_SHEET);s.appendRow(['Job ID','State','Updated At','Result JSON','Error']);} return s;
}
function findJobRow_(sheet,jobId) {
  if(!jobId||sheet.getLastRow()<2)return 0;
  const found=sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(jobId).matchEntireCell(true).findNext();
  return found?found.getRow():0;
}
function render_(template,c) {
  return String(template||'').replace(/\{\{\s*(Name|Email|Company|Role)\s*\}\}/gi,(_,key)=>{
    const m={name:c.name||'there',email:c.email||'',company:c.company||'',role:c.role||''};return m[key.toLowerCase()]||'';
  });
}
function validEmail_(email){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email||'').trim());}
function constantTimeEquals_(a,b) {
  a=String(a);b=String(b);let diff=a.length^b.length,n=Math.max(a.length,b.length);
  for(let i=0;i<n;i++)diff|=(a.charCodeAt(i)||0)^(b.charCodeAt(i)||0);
  return diff===0;
}