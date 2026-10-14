// ============================================================
// City of San Jacinto — Planning Time Tracker
// Google Apps Script Backend — JSONP version (CORS-free)
// ------------------------------------------------------------
// SETUP:
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Select all existing code → Delete → Paste this file
// 3. Save (floppy disk)
// 4. Deploy → Manage Deployments → pencil → New version → Deploy
// ============================================================

const SHEET_NAME_PROJECTS = 'Projects';
const SHEET_NAME_SESSIONS = 'Sessions';
const SHEET_NAME_DEPOSITS = 'DepositHistory';
const SHEET_NAME_PLANNERS = 'Planners';
const SHEET_NAME_AILOG    = 'AILog';

const WRITE_ACTIONS = {
  saveSession: saveSession,
  addProject: addProject,
  updateProject: updateProject,
  deleteProject: deleteProject,
  addDeposit: addDeposit,
  updateDeposit: updateDeposit,
  deleteDeposit: deleteDeposit,
  deleteSession: deleteSession,
  savePlanners: savePlanners,
  saveAILog: saveAILog,
  deleteAILog: deleteAILog
};

// ------------------------------------------------------------
// doGet handles ALL requests (GET + POST-via-GET for JSONP)
// This completely avoids CORS preflight issues
// ------------------------------------------------------------

function doGet(e) {
  const p = e.parameter;
  const callback = p.callback || 'callback';
  const action = p.action || '';

  try {
    let result;

    if (action === 'ping') {
      result = { status: 'ok' };

    } else if (action === 'getAll') {
      result = getAllData();

    } else if (WRITE_ACTIONS[action]) {
      const data = JSON.parse(decodeURIComponent(p.data));
      // One write at a time, so two planners saving at once can't
      // overwrite each other's balance updates
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        result = WRITE_ACTIONS[action](data);
      } finally {
        lock.releaseLock();
      }

    } else {
      result = { error: 'Unknown action: ' + action };
    }

    const json = JSON.stringify({ success: true, data: result });
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);

  } catch (ex) {
    const json = JSON.stringify({ success: false, error: ex.message });
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

// Keep doPost as fallback
function doPost(e) {
  return doGet(e);
}

// ------------------------------------------------------------
// Sheet helpers
// ------------------------------------------------------------

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#C00')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// Returns the 1-based sheet row whose id column matches, or -1
function findRowById(sheet, id) {
  if (!id) return -1;
  const values = sheet.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

function setupSheets() {
  getOrCreateSheet(SHEET_NAME_PROJECTS, [
    'id','rectype','case','name','applicant','deposit','spent','type','parentId','createdDate'
  ]);
  getOrCreateSheet(SHEET_NAME_SESSIONS, [
    'id','recIdx','case','name','rectype','activity','note','planner','seconds','charge','date','source','createdAt'
  ]);
  getOrCreateSheet(SHEET_NAME_DEPOSITS, [
    'id','projectId','projectCase','amount','note','date','createdAt'
  ]);
  getOrCreateSheet(SHEET_NAME_PLANNERS, ['name']);
  getOrCreateSheet(SHEET_NAME_AILOG, [
    'id','date','projCase','projName','planCase','tool','task','tier','review','notes','planner','createdAt'
  ]);
}

// ------------------------------------------------------------
// Read all data
// ------------------------------------------------------------

function getAllData() {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projects = sheetToObjects(ss.getSheetByName(SHEET_NAME_PROJECTS));
  const sessions = sheetToObjects(ss.getSheetByName(SHEET_NAME_SESSIONS));
  const deposits = sheetToObjects(ss.getSheetByName(SHEET_NAME_DEPOSITS));
  const planSheet = ss.getSheetByName(SHEET_NAME_PLANNERS);
  const planners = planSheet
    ? sheetToObjects(planSheet).map(r => r.name).filter(Boolean)
    : ['Frank','Maria','David','Sarah'];
  const aiLogs = sheetToObjects(ss.getSheetByName(SHEET_NAME_AILOG));
  return { projects, sessions, deposits, planners, aiLogs };
}

// ------------------------------------------------------------
// Projects
// ------------------------------------------------------------

function addProject(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_PROJECTS, [
    'id','rectype','case','name','applicant','deposit','spent','type','parentId','createdDate'
  ]);
  // The app sends its own id; if it is already here this is a retry
  if (findRowById(sheet, data.id) > 0) return { id: data.id, duplicate: true };
  const id = data.id || ('p_' + Date.now());
  sheet.appendRow([
    id, data.rectype||'project', data.case||'', data.name||'',
    data.applicant||'', Number(data.deposit)||0, Number(data.spent)||0,
    data.type||'', data.parentId||'',
    new Date().toLocaleDateString('en-US')
  ]);
  return { id };
}

function updateProject(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROJECTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(data.id)) {
      Object.keys(data).forEach(key => {
        const col = headers.indexOf(key);
        if (col >= 0) sheet.getRange(i+1, col+1).setValue(data[key]);
      });
      return { updated: data.id };
    }
  }
  return { error: 'Project not found' };
}

function deleteProject(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROJECTS);
  const values = sheet.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return { deleted: data.id };
    }
  }
  return { error: 'Not found' };
}

// ------------------------------------------------------------
// Sessions
// ------------------------------------------------------------

function saveSession(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_SESSIONS, [
    'id','recIdx','case','name','rectype','activity','note','planner',
    'seconds','charge','date','source','createdAt'
  ]);
  // A retried save must not charge the plan twice
  if (findRowById(sheet, data.id) > 0) return { id: data.id, duplicate: true };
  const id = data.id || ('s_' + Date.now());
  sheet.appendRow([
    id, data.recIdx||'', data.case||'', data.name||'',
    data.rectype||'project', data.activity||'—', data.note||'',
    data.planner||'', Number(data.seconds)||0, Number(data.charge)||0,
    data.date||new Date().toLocaleDateString('en-US'),
    data.source||'timer', new Date().toLocaleDateString('en-US')
  ]);
  updateProjectSpent(data.case, Number(data.charge)||0);
  return { id };
}

function deleteSession(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SESSIONS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  const chargeCol = headers.indexOf('charge');
  const caseCol = headers.indexOf('case');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(data.id)) {
      const charge = Number(values[i][chargeCol])||0;
      const projCase = values[i][caseCol];
      sheet.deleteRow(i+1);
      updateProjectSpent(projCase, -charge);
      return { deleted: data.id };
    }
  }
  return { error: 'Session not found' };
}

function updateProjectSpent(caseNum, delta) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROJECTS);
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const caseCol = headers.indexOf('case');
  const spentCol = headers.indexOf('spent');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][caseCol]) === String(caseNum)) {
      const cur = Number(values[i][spentCol])||0;
      sheet.getRange(i+1, spentCol+1).setValue(
        Math.max(0, parseFloat((cur + delta).toFixed(2)))
      );
      break;
    }
  }
}

// ------------------------------------------------------------
// Deposits
// ------------------------------------------------------------

function getDepositSheet() {
  return getOrCreateSheet(SHEET_NAME_DEPOSITS, [
    'id','projectId','projectCase','amount','note','date','createdAt'
  ]);
}

// Adds delta to the deposit total on the plan/project row
function adjustProjectDeposit(projectId, delta) {
  const projSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_PROJECTS);
  const values = projSheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  const depCol = headers.indexOf('deposit');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(projectId)) {
      const cur = Number(values[i][depCol])||0;
      projSheet.getRange(i+1, depCol+1).setValue(Math.max(0, parseFloat((cur+delta).toFixed(2))));
      break;
    }
  }
}

function addDeposit(data) {
  const sheet = getDepositSheet();
  if (findRowById(sheet, data.id) > 0) return { id: data.id, duplicate: true };
  const id = data.id || ('d_' + Date.now());
  const amt = Number(data.amount)||0;
  sheet.appendRow([
    id, data.projectId||'', data.projectCase||'', amt,
    data.note||'Deposit added',
    data.date||new Date().toLocaleDateString('en-US'),
    new Date().toLocaleDateString('en-US')
  ]);
  // adjustTotal is false for a new plan's starting deposit, which the
  // plan row already includes
  if (data.adjustTotal !== false) adjustProjectDeposit(data.projectId, amt);
  return { id };
}

function updateDeposit(data) {
  const sheet = getDepositSheet();
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Deposit not found' };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cur = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const amtCol = headers.indexOf('amount');
  const oldAmt = Number(cur[amtCol])||0;
  const newAmt = Number(data.amount)||0;
  sheet.getRange(row, amtCol+1).setValue(newAmt);
  if (data.note) sheet.getRange(row, headers.indexOf('note')+1).setValue(data.note);
  if (data.date) sheet.getRange(row, headers.indexOf('date')+1).setValue(data.date);
  adjustProjectDeposit(cur[headers.indexOf('projectId')], newAmt - oldAmt);
  return { updated: data.id };
}

function deleteDeposit(data) {
  const sheet = getDepositSheet();
  const row = findRowById(sheet, data.id);
  if (row < 0) return { error: 'Deposit not found' };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cur = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const amt = Number(cur[headers.indexOf('amount')])||0;
  sheet.deleteRow(row);
  adjustProjectDeposit(cur[headers.indexOf('projectId')], -amt);
  return { deleted: data.id };
}

// ------------------------------------------------------------
// Planners
// ------------------------------------------------------------

function savePlanners(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_PLANNERS, ['name']);
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last-1);
  data.planners.forEach(name => sheet.appendRow([name]));
  return { saved: data.planners.length };
}

// ------------------------------------------------------------
// AI Log
// ------------------------------------------------------------

function saveAILog(data) {
  const sheet = getOrCreateSheet(SHEET_NAME_AILOG, [
    'id','date','projCase','projName','planCase','tool','task','tier','review','notes','planner','createdAt'
  ]);
  if (findRowById(sheet, data.id) > 0) return { id: data.id, duplicate: true };
  const id = data.id || ('ai_' + Date.now());
  sheet.appendRow([
    id,
    data.date || new Date().toLocaleDateString('en-US'),
    data.projCase || '',
    data.projName || '',
    data.planCase || '',
    data.tool || '',
    data.task || '',
    data.tier || '',
    data.review || '',
    data.notes || '',
    data.planner || '',
    new Date().toLocaleDateString('en-US')
  ]);
  return { id };
}

function deleteAILog(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_AILOG);
  if (!sheet) return { error: 'AILog sheet not found' };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return { deleted: data.id };
    }
  }
  return { error: 'AI log entry not found' };
}
