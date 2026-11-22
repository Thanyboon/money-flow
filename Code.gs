/** MONEY FLOW — Google Apps Script API (ปรับปรุงการรับส่งข้อมูล)
 * คงชื่อชีตและคอลัมน์เดิมทั้งหมด เพื่อรักษาข้อมูลเดิม
 * Script Properties: SPREADSHEET_ID
 *
 * รุ่นนี้ปิดระบบรหัสผ่านและ Session ตามที่เจ้าของแอปยืนยัน
 * ทุกคนที่เข้าถึง Web App URL สามารถอ่าน/แก้ไขข้อมูลในชีตได้
 * การสื่อสารจาก GitHub Pages ใช้ POST ผ่าน hidden iframe และ postMessage
 */
const MF = {
  VERSION: '1.1.0-public-no-password',
  TZ: 'Asia/Bangkok',
  SESSION_TTL: 21600,
  LOGIN_LIMIT: 8,
  LOGIN_WINDOW: 900,
  TABLES: {
    Settings:['id','key','value','updatedAt'],
    Transactions:['id','type','amount','categoryId','walletId','date','note','createdAt','updatedAt'],
    Wallets:['id','name','type','openingBalance','color','active','createdAt','updatedAt'],
    Categories:['id','name','type','icon','color','active'],
    Budgets:['id','month','categoryId','amount','createdAt','updatedAt'],
    SavingsGoals:['id','name','targetAmount','currentAmount','deadline','walletId','note','createdAt','updatedAt'],
    SavingsHistory:['id','goalId','type','amount','date','note','createdAt'],
    BankReconciliation:['id','walletId','date','actualBalance','systemBalance','difference','note','createdAt'],
    BillGroups:['id','name','createdAt'],
    BillMembers:['id','groupId','name','createdAt'],
    Bills:['id','groupId','title','date','payerMemberId','total','splitMethod','note','createdAt'],
    BillItems:['id','billId','description','amount','memberShares','createdAt'],
    BillPayments:['id','billId','fromMemberId','toMemberId','amount','paidAt','status','note','createdAt']
  }
};

/** GET เปิดให้ใช้เฉพาะ ping เท่านั้น ไม่รับ password/token ผ่าน URL */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if ((p.action || 'ping') === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({
      ok:true, app:'MONEY FLOW', version:MF.VERSION
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({
    ok:false, error:'POST_REQUIRED'
  })).setMimeType(ContentService.MimeType.JSON);
}

/** ทุกคำสั่งแอปส่งผ่าน POST; ผลลัพธ์ส่งกลับไปยัง iframe ด้วย nonce */
function doPost(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    const action = String(p.action || '');
    if (action === 'bootstrap') result = bootstrap_();
    else if (action === 'mutate') result = mutate_(p);
    else throw new Error('UNKNOWN_ACTION');
  } catch (err) {
    result = {ok:false, error:String(err && err.message || err)};
  }
  return postMessageHtml_(result, String(p.nonce || ''));
}

function postMessageHtml_(payload, nonce) {
  // JSON is escaped for safe embedding in a script element.
  const safePayload = JSON.stringify(payload)
    .replace(/</g, '\\u003c').replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026').replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  const safeNonce = nonce.replace(/[^a-zA-Z0-9]/g, '');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    'window.top.postMessage({mf:true,nonce:' + JSON.stringify(safeNonce) +
    ',payload:' + safePayload + '}, "*");' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function initialize_() {
  const ss = spreadsheet_();
  Object.keys(MF.TABLES).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = MF.TABLES[name];
    // Never overwrite existing headers or rows.
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  seedCategories_();
  return {ok:true, message:'ตรวจสอบโครงสร้างฐานข้อมูลเรียบร้อย', spreadsheetUrl:ss.getUrl(), sheets:Object.keys(MF.TABLES)};
}
function initializeDatabase() { return initialize_(); }

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('กรุณาตั้งค่า SPREADSHEET_ID ใน Script Properties');
  return SpreadsheetApp.openById(id);
}
function seedCategories_() {
  const sh = spreadsheet_().getSheetByName('Categories');
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['cat_income_salary','เงินเดือน','income','briefcase','#22C55E',true],
    ['cat_income_other','รายรับอื่น','income','wallet','#16A34A',true],
    ['cat_food','อาหาร','expense','utensils','#F97316',true],
    ['cat_transport','เดินทาง','expense','car','#3B82F6',true],
    ['cat_shopping','ช้อปปิ้ง','expense','shopping-bag','#A855F7',true],
    ['cat_bills','บิลและค่าสาธารณูปโภค','expense','receipt','#EAB308',true],
    ['cat_health','สุขภาพ','expense','heart-pulse','#EC4899',true],
    ['cat_entertainment','บันเทิง','expense','gamepad-2','#8B5CF6',true],
    ['cat_other','อื่น ๆ','expense','circle-ellipsis','#64748B',true]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}
function bootstrap_() {
  const ss = spreadsheet_(), data = {};
  Object.keys(MF.TABLES).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error('MISSING_SHEET:' + name);
    data[name] = readTable_(sh);
  });
  return {ok:true, data:data, serverTime:new Date().toISOString(), version:MF.VERSION};
}
function readTable_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '' && v !== null)).map(r => {
    const o = {};
    headers.forEach((h,i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, MF.TZ, "yyyy-MM-dd'T'HH:mm:ss");
      o[h] = v;
    });
    return o;
  });
}
function mutate_(p) {
  const table = String(p.table || '');
  if (!Object.prototype.hasOwnProperty.call(MF.TABLES, table)) throw new Error('INVALID_TABLE');
  const op = String(p.op || '');
  const sh = spreadsheet_().getSheetByName(table);
  if (!sh) throw new Error('MISSING_SHEET:' + table);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const headers = MF.TABLES[table], rows = readTable_(sh);
    if (op === 'upsert') {
      let row = typeof p.row === 'string' ? JSON.parse(p.row) : p.row;
      if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('INVALID_ROW');
      if (Object.keys(row).length > headers.length + 3) throw new Error('TOO_MANY_FIELDS');
      if (!row.id) row.id = Utilities.getUuid();
      row.id = String(row.id);
      validateRow_(table, row);
      const now = new Date().toISOString();
      if (headers.includes('updatedAt')) row.updatedAt = now;
      if (headers.includes('createdAt') && !row.createdAt) row.createdAt = now;
      const index = rows.findIndex(r => String(r.id) === row.id);
      const values = headers.map(h => safeCell_(row[h] === undefined ? '' : row[h]));
      if (index >= 0) sh.getRange(index+2,1,1,headers.length).setValues([values]);
      else sh.appendRow(values);
      return {ok:true, id:row.id, row:row};
    }
    if (op === 'delete') {
      const id = String(p.id || '');
      if (!id || id.length > 200) throw new Error('INVALID_ID');
      const index = rows.findIndex(r => String(r.id) === id);
      if (index < 0) return {ok:true, deleted:false};
      sh.deleteRow(index+2);
      return {ok:true, deleted:true, id:id};
    }
    throw new Error('INVALID_OPERATION');
  } finally {
    lock.releaseLock();
  }
}
function validateRow_(table, row) {
  const amountFields = {
    Transactions:['amount'], Wallets:['openingBalance'], Budgets:['amount'],
    SavingsGoals:['targetAmount','currentAmount'], SavingsHistory:['amount'],
    BankReconciliation:['actualBalance','systemBalance','difference'],
    Bills:['total'], BillItems:['amount'], BillPayments:['amount']
  };
  (amountFields[table] || []).forEach(k => {
    if (row[k] !== undefined && row[k] !== '' && !Number.isFinite(Number(row[k]))) throw new Error('INVALID_NUMBER:' + k);
  });
  if (table === 'Transactions') {
    if (!['income','expense'].includes(String(row.type))) throw new Error('INVALID_TRANSACTION_TYPE');
    if (!(Number(row.amount) > 0)) throw new Error('INVALID_AMOUNT');
    if (!row.walletId || !row.date) throw new Error('MISSING_TRANSACTION_FIELDS');
  }
  if (table === 'SavingsHistory' && !['deposit','withdraw'].includes(String(row.type))) throw new Error('INVALID_SAVINGS_TYPE');
  if (table === 'BillPayments' && row.amount !== undefined && Number(row.amount) <= 0) throw new Error('INVALID_AMOUNT');
  if (table === 'Bills' && Number(row.total) <= 0) throw new Error('INVALID_AMOUNT');
  Object.keys(row).forEach(k => {
    if (typeof row[k] === 'string' && row[k].length > 10000) throw new Error('FIELD_TOO_LONG:' + k);
  });
}
function safeCell_(value) {
  if (typeof value === 'string' && /^[=+\-@]/.test(value)) return "'" + value;
  return value;
}
