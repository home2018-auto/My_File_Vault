/****************************************************************
 * FILE VAULT — Google Apps Script backend (multi-user)
 * Deploy: Extensions > Apps Script (in a Google Sheet) > paste this
 * file as Code.gs > Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Copy the deployed Web App URL into the dashboard's Settings.
 * This ONE deployment + ONE spreadsheet is shared by every account
 * that signs up from the dashboard — each account automatically gets
 * its own set of Files/Automation/Email tabs, kept completely apart
 * from every other account's tabs.
 *
 * NOTE: id & color (per data row) and passHash & salt (per user) are
 * kept as hidden columns — internal bookkeeping, not meant to be
 * edited by hand.
 *
 * SECURITY NOTE: this is a lightweight, personal-use login layer, not
 * enterprise-grade auth. The dashboard hashes the password in the
 * browser before it's ever sent, and this script salts + re-hashes it
 * again before storing — so the raw password never travels or sits in
 * a sheet cell. But there's no session-expiry, no rate-limiting, and
 * anyone with the Web App URL, a valid userId and its passHash can
 * call the API directly. That's an acceptable trade-off for a private
 * hobby dashboard shared with people you trust — not for anything
 * that needs real access control.
 ****************************************************************/

const HEADER_ROW = 6;   // rows 1-5 stay reserved/blank
const DATA_START_ROW = 7;
const TIMEZONE = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

const SHEET_CONFIG = {
  files:      { name: 'Files',      headers: ['sl','date','name','type','size','location','desc','note','photo','id','color'],
                dateFields: ['date'] },
  automation: { name: 'Automation', headers: ['sl','date','name','type','maker','desc','location','note','photo','id','color'],
                dateFields: ['date'] },
  email:      { name: 'Email',      headers: ['sl','date','name','email','password','recovery','mobile','passkey','lastUpdate','note','id','color'],
                dateFields: ['date','lastUpdate'] },
  passwordNote: { name: 'PasswordNote', headers: ['sl','date','siteName','name','userName','email','password','note','photo','id','color'],
                dateFields: ['date'] }
};

const USERS_SHEET_NAME = 'Users';
const USERS_HEADERS = ['fullName','userId','email','passHash','salt','createdAt'];

function doGet(e){
  const cb = e.parameter.callback;
  const action = e.parameter.action;
  const sheetKey = e.parameter.sheet;
  let result;
  try{
    if(action === 'test'){
      result = { status: 'ok' };
    } else if(action === 'signup'){
      result = withLock(()=> handleSignup(e.parameter));
    } else if(action === 'login'){
      result = handleLogin(e.parameter);
    } else if(action === 'read'){
      result = withAuth(e.parameter, ()=> ({ status: 'ok', data: readSheet(sheetKey, e.parameter.user) }));
    } else if(action === 'add' || action === 'update'){
      result = withAuth(e.parameter, ()=>{
        withLock(()=> upsertRow(sheetKey, e.parameter.user, JSON.parse(e.parameter.data)));
        return { status: 'ok' };
      });
    } else if(action === 'delete'){
      result = withAuth(e.parameter, ()=>{
        withLock(()=> deleteRow(sheetKey, e.parameter.user, JSON.parse(e.parameter.data).id));
        return { status: 'ok' };
      });
    } else if(action === 'setDriveFolder'){
      result = withAuth(e.parameter, ()=> moveSpreadsheetToFolder(e.parameter.folderName));
    } else {
      result = { status: 'error', message: 'unknown action' };
    }
  } catch(err){
    result = { status: 'error', message: String(err) };
  }
  const body = cb ? (cb + '(' + JSON.stringify(result) + ')') : JSON.stringify(result);
  return ContentService.createTextOutput(body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e){
  let result;
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.action === 'uploadPhoto'){
      if(!verifyUser(body.user, body.passHash)){
        result = { status: 'error', message: 'auth_failed' };
      } else {
        result = { status: 'ok', url: uploadPhotoToDrive(body.image, body.id, body.folder) };
      }
    } else {
      result = { status: 'error', message: 'unknown action' };
    }
  } catch(err){
    result = { status: 'error', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- auth ---------- */
function withAuth(p, fn){
  const auth = verifyUser(p.user, p.passHash);
  if(!auth) return { status: 'error', message: 'auth_failed' };
  return fn();
}

function handleSignup(p){
  const fullName = (p.fullName || '').trim();
  const userId = sanitizeUserId(p.userId || '');
  const email = (p.email || '').trim();
  const passHash = p.passHash || '';
  if(!fullName || !userId || !email || !passHash) return { status: 'error', message: 'missing_fields' };
  if(findUser(userId)) return { status: 'error', message: 'userId_taken' };
  const wasFirstUser = countUsers() === 0;
  createUser(fullName, userId, email, passHash);
  if(wasFirstUser) migrateLegacyDataForFirstUser(userId);
  return { status: 'ok' };
}

function handleLogin(p){
  const userId = sanitizeUserId(p.userId || '');
  const passHash = p.passHash || '';
  const auth = verifyUser(userId, passHash);
  if(!auth) return { status: 'error', message: 'invalid_credentials' };
  return { status: 'ok', fullName: auth.fullName, email: auth.email };
}

function sanitizeUserId(id){
  return String(id || '').replace(/[^A-Za-z0-9_\-]/g, '').slice(0, 40);
}

function sha256Hex(str){
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(b=>{
    const v = b < 0 ? b + 256 : b;
    return (v < 16 ? '0' : '') + v.toString(16);
  }).join('');
}

/* ---------- Users sheet ---------- */
function getUsersSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if(!sheet) sheet = ss.insertSheet(USERS_SHEET_NAME);
  const headerCell = sheet.getRange(HEADER_ROW, 1).getValue();
  if(!headerCell){
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, USERS_HEADERS.length);
    headerRange.setValues([USERS_HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(HEADER_ROW);
  }
  const passCol = USERS_HEADERS.indexOf('passHash') + 1;
  const saltCol = USERS_HEADERS.indexOf('salt') + 1;
  if(passCol > 0) sheet.hideColumns(passCol);
  if(saltCol > 0) sheet.hideColumns(saltCol);
  return sheet;
}

function countUsers(){
  const sheet = getUsersSheet();
  const lastRow = sheet.getLastRow();
  return lastRow < DATA_START_ROW ? 0 : lastRow - DATA_START_ROW + 1;
}

function findUser(userId){
  const sheet = getUsersSheet();
  const lastRow = sheet.getLastRow();
  if(lastRow < DATA_START_ROW) return null;
  const values = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, USERS_HEADERS.length).getValues();
  for(let i = 0; i < values.length; i++){
    if(String(values[i][1]) === userId){
      const obj = {};
      USERS_HEADERS.forEach((h, idx)=>{ obj[h] = values[i][idx]; });
      obj._row = DATA_START_ROW + i;
      return obj;
    }
  }
  return null;
}

function createUser(fullName, userId, email, passHash){
  const sheet = getUsersSheet();
  const salt = Utilities.getUuid();
  const stored = sha256Hex(passHash + salt);
  const row = [fullName, userId, email, stored, salt, new Date()];
  const targetRow = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
  sheet.getRange(targetRow, 1, 1, USERS_HEADERS.length).setValues([row]);
  SpreadsheetApp.flush();
}

function verifyUser(userId, passHash){
  userId = sanitizeUserId(userId || '');
  passHash = passHash || '';
  if(!userId || !passHash) return null;
  const user = findUser(userId);
  if(!user) return null;
  const stored = sha256Hex(passHash + user.salt);
  if(stored !== user.passHash) return null;
  return { fullName: user.fullName, email: user.email };
}

/* First person ever to sign up inherits whatever data already sits in
 * the original un-suffixed Files/Automation/Email tabs (from before
 * multi-user existed), by renaming those tabs to that account's own
 * tabs. Anyone who signs up after that just starts empty, as normal. */
function migrateLegacyDataForFirstUser(userId){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_CONFIG).forEach(key=>{
    const base = SHEET_CONFIG[key].name;
    const legacy = ss.getSheetByName(base);
    if(!legacy) return;
    const target = userSheetName(base, userId);
    if(ss.getSheetByName(target)) return;
    legacy.setName(target);
  });
}

/* ---------- one-time cleanup ----------
 * An earlier version of this file used setWrap(true), which force-grows
 * row height to fit long notes. This is NOT called automatically (so it
 * never fights with a row height you resize by hand later). Run it once
 * yourself: open this file in the Apps Script editor, pick
 * "resetRowHeightsOnce" from the function dropdown at the top, click Run,
 * then you're done. */
function resetRowHeightsOnce(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet=>{
    const rows = Math.max(sheet.getMaxRows() - HEADER_ROW + 1, 1);
    sheet.setRowHeightsForced(HEADER_ROW, rows, 21);
  });
}

/* ---------- concurrency guard ----------
 * Without this, several near-simultaneous requests (e.g. "send all
 * data") could all read the same lastRow at once and overwrite each
 * other — that was the cause of rows going missing/duplicated. */
function withLock(fn){
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{ return fn(); } finally{ lock.releaseLock(); }
}

/* ---------- per-user sheet helpers ---------- */
function userSheetName(base, userId){
  return base + '__' + sanitizeUserId(userId);
}

function getOrCreateSheet(sheetKey, userId){
  const cfg = SHEET_CONFIG[sheetKey];
  if(!cfg) throw new Error('unknown sheet: ' + sheetKey);
  const sheetName = userSheetName(cfg.name, userId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet) sheet = ss.insertSheet(sheetName);

  const headerCell = sheet.getRange(HEADER_ROW, 1).getValue();
  if(!headerCell){
    // First-time setup: rows 1-5 stay blank, header goes on row 6.
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, cfg.headers.length);
    headerRange.setValues([cfg.headers]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(HEADER_ROW);

    const bodyRangeForFilter = sheet.getRange(HEADER_ROW, 1, Math.max(sheet.getMaxRows() - HEADER_ROW + 1, 2), cfg.headers.length);
    const existingFilter = sheet.getFilter();
    if(existingFilter) existingFilter.remove();
    bodyRangeForFilter.createFilter();
  }

  // Runs every call (cheap, idempotent) so a sheet self-heals formatting
  // even if it was created before this logic existed.
  ensureBodyFormatting(sheet, cfg);

  // 'files' sheets created before the photo feature existed are missing
  // the 'photo' column — insert it in the right spot (just before id) so
  // old sheets catch up without disturbing any existing data/columns.
  if(sheetKey === 'files') ensureColumnExists(sheet, cfg, 'photo', 'id');

  return sheet;
}

/* Inserts headerName as a real column in the LIVE sheet if it's not
 * already there — positioned right before insertBeforeHeaderName (falls
 * back to appending at the end if that column can't be found). A no-op
 * once the column already exists, so safe to call on every request. */
function ensureColumnExists(sheet, cfg, headerName, insertBeforeHeaderName){
  const lastCol = Math.max(sheet.getLastColumn(), cfg.headers.length);
  const headerRow = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  if(headerRow.indexOf(headerName) > -1) return;
  const insertBeforeIdx = headerRow.indexOf(insertBeforeHeaderName); // 0-based
  if(insertBeforeIdx === -1){
    const col = lastCol + 1;
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(HEADER_ROW, col).setValue(headerName).setFontWeight('bold');
  } else {
    const col = insertBeforeIdx + 1; // 1-based
    sheet.insertColumnBefore(col);
    sheet.getRange(HEADER_ROW, col).setValue(headerName).setFontWeight('bold');
  }
}

/* Keep every cell exactly the size it already is — never let the row grow.
 * CLIP means text that doesn't fit is simply hidden inside the cell rather
 * than spilling over the border, and rather than forcing the row taller.
 * If you widen the column or make the row taller yourself later, the
 * hidden part shows automatically — the data itself is untouched, only
 * how much of it is visible changes. Safe to call repeatedly. */
function ensureBodyFormatting(sheet, cfg){
  const rows = Math.max(sheet.getMaxRows() - HEADER_ROW + 1, 2);
  const bodyRange = sheet.getRange(HEADER_ROW, 1, rows, cfg.headers.length);
  bodyRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  bodyRange.setVerticalAlignment('top');

  cfg.dateFields.forEach(f=>{
    const col = cfg.headers.indexOf(f) + 1;
    if(col > 0) sheet.getRange(HEADER_ROW + 1, col, Math.max(sheet.getMaxRows() - HEADER_ROW, 1), 1).setNumberFormat('@');
  });

  const idCol = cfg.headers.indexOf('id') + 1;
  const colorCol = cfg.headers.indexOf('color') + 1;
  if(idCol > 0) sheet.hideColumns(idCol);
  if(colorCol > 0) sheet.hideColumns(colorCol);
}

function idColumnOf(cfg){ return cfg.headers.indexOf('id') + 1; } // 1-based

function normalizeDateValue(v){
  if(Object.prototype.toString.call(v) === '[object Date]'){
    return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
  }
  return v;
}

function readSheet(sheetKey, userId){
  const sheet = getOrCreateSheet(sheetKey, userId);
  const cfg = SHEET_CONFIG[sheetKey];
  const lastRow = sheet.getLastRow();
  if(lastRow < DATA_START_ROW) return [];
  const idCol = idColumnOf(cfg);
  const values = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, cfg.headers.length).getValues();
  return values.filter(row => row[idCol - 1]).map(row => {
    const obj = {};
    cfg.headers.forEach((h, i) => {
      obj[h] = cfg.dateFields.indexOf(h) > -1 ? normalizeDateValue(row[i]) : row[i];
    });
    return obj;
  });
}

function upsertRow(sheetKey, userId, data){
  const sheet = getOrCreateSheet(sheetKey, userId);
  const cfg = SHEET_CONFIG[sheetKey];
  const idCol = idColumnOf(cfg);
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if(lastRow >= DATA_START_ROW){
    const ids = sheet.getRange(DATA_START_ROW, idCol, lastRow - DATA_START_ROW + 1, 1).getValues();
    for(let i = 0; i < ids.length; i++){
      if(ids[i][0] === data.id){ targetRow = DATA_START_ROW + i; break; }
    }
  }
  const rowValues = cfg.headers.map(h => data[h] !== undefined ? data[h] : '');
  if(targetRow === -1){
    targetRow = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
  }
  sheet.getRange(targetRow, 1, 1, cfg.headers.length).setValues([rowValues]);
  SpreadsheetApp.flush();
}

function deleteRow(sheetKey, userId, id){
  const sheet = getOrCreateSheet(sheetKey, userId);
  const cfg = SHEET_CONFIG[sheetKey];
  const idCol = idColumnOf(cfg);
  const lastRow = sheet.getLastRow();
  if(lastRow < DATA_START_ROW) return;
  const ids = sheet.getRange(DATA_START_ROW, idCol, lastRow - DATA_START_ROW + 1, 1).getValues();
  for(let i = 0; i < ids.length; i++){
    if(ids[i][0] === id){ sheet.deleteRow(DATA_START_ROW + i); break; }
  }
  SpreadsheetApp.flush();
}

/* ---------- photo upload to Drive ---------- */
function resolveDriveFolder(folderName){
  if(!folderName) return DriveApp.getRootFolder();
  const it = DriveApp.getFoldersByName(folderName);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}
/* ছবি মূল ফোল্ডারে সরাসরি না রেখে, তার ভেতরে "Photos" নামে একটা
 * সাব-ফোল্ডার (একবারই তৈরি হবে, পরে reuse হবে) বানিয়ে সেখানে রাখা হয় —
 * ফলে মূল ফোল্ডারে শুধু Sheet ফাইল আর একটা পরিষ্কার Photos সাব-ফোল্ডার থাকে। */
function resolvePhotosSubfolder(mainFolder){
  const it = mainFolder.getFoldersByName('Photos');
  if(it.hasNext()) return it.next();
  return mainFolder.createFolder('Photos');
}
const PHOTOS_SUBFOLDER_NAME = 'Photos';
/* মূল ফোল্ডারের ভেতরেই "Photos" নামে একটা সাব-ফোল্ডার — শিট ফাইলটা মূল
 * ফোল্ডারে থাকে, আর সব ছবি এই সাব-ফোল্ডারে গিয়ে জমা হয়, একসাথে
 * জগাখিচুড়ি না হয়ে। */
function resolvePhotosSubfolder(mainFolder){
  const it = mainFolder.getFoldersByName(PHOTOS_SUBFOLDER_NAME);
  if(it.hasNext()) return it.next();
  return mainFolder.createFolder(PHOTOS_SUBFOLDER_NAME);
}
/* সেটিংসে ফোল্ডারের নাম সেভ করলে এটা কল হয় — শুধু ছবি নয়, এই Google
 * Sheet ফাইলটা নিজেও ওই ফোল্ডারে সরিয়ে নেওয়া হয় (আগের সব লোকেশন থেকে
 * সরিয়ে), যাতে সিট আর ছবি একই জায়গায় থাকে। */
function moveSpreadsheetToFolder(folderName){
  if(!folderName) return { status: 'error', message: 'no folder name given' };
  const folder = resolveDriveFolder(folderName);
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const already = file.getParents();
  const oldParents = [];
  while(already.hasNext()){ oldParents.push(already.next()); }
  folder.addFile(file);
  oldParents.forEach(p=>{ if(p.getId() !== folder.getId()) p.removeFile(file); });
  return { status: 'ok', folderUrl: folder.getUrl() };
}
function uploadPhotoToDrive(dataUrl, id, folderName){
  const match = /^data:(.+);base64,(.*)$/.exec(dataUrl);
  const contentType = match ? match[1] : 'image/jpeg';
  const base64 = match ? match[2] : dataUrl;
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, contentType, (id || 'photo') + '_' + new Date().getTime() + '.jpg');
  const mainFolder = resolveDriveFolder(folderName);
  const folder = folderName ? resolvePhotosSubfolder(mainFolder) : mainFolder;
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}
