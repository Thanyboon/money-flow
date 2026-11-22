/** MONEY FLOW — Google Apps Script API
 * ติดตั้ง: ตั้งค่า Script Properties: SPREADSHEET_ID, APP_PASSWORD
 * Deploy เป็น Web app (Execute as: Me; Who has access: Anyone) แล้วใช้รหัสผ่านแอป
 */
const MF = {
  VERSION: '1.0.0', TZ: 'Asia/Bangkok',
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
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || 'ping';
    let result;
    if (action === 'login') result = login_(p.password || '');
    else if (action === 'ping') result = {ok:true,app:'MONEY FLOW',version:MF.VERSION};
    else {
      if (!validToken_(p.token)) throw new Error('SESSION_EXPIRED');
      if (action === 'bootstrap') result = bootstrap_();
      else if (action === 'mutate') result = mutate_(p);
      else if (action === 'init') result = initialize_();
      else throw new Error('UNKNOWN_ACTION');
    }
    return jsonp_(result, p.callback);
  } catch (err) { return jsonp_({ok:false,error:String(err.message || err)}, e && e.parameter && e.parameter.callback); }
}
function doPost(e) {
  // รองรับ POST form สำหรับการส่งข้อมูลที่ใหญ่กว่า URL; ตอบกลับเป็น JSON
  try {
    const p = (e && e.parameter) || {};
    if (!validToken_(p.token)) throw new Error('SESSION_EXPIRED');
    const result = mutate_({table:p.table, op:p.op, row:p.row, id:p.id});
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function login_(password) {
  const expected = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if (!expected) throw new Error('กรุณาตั้งค่า APP_PASSWORD ใน Script Properties ก่อน');
  if (!constantTimeEquals_(String(password), expected)) throw new Error('รหัสผ่านไม่ถูกต้อง');
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put('mf:'+token, '1', 21600);
  return {ok:true,token:token,expiresIn:21600};
}
function validToken_(token) { return !!token && CacheService.getScriptCache().get('mf:'+token) === '1'; }
function constantTimeEquals_(a,b) { if(a.length!==b.length)return false; let x=0; for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i); return x===0; }
function initialize_() {
  const ss = spreadsheet_();
  Object.keys(MF.TABLES).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = MF.TABLES[name];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  seedCategories_();
  return {ok:true,message:'สร้างโครงสร้างฐานข้อมูลเรียบร้อย',spreadsheetUrl:ss.getUrl(),sheets:Object.keys(MF.TABLES)};
}
function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('กรุณาตั้งค่า SPREADSHEET_ID ใน Script Properties');
  return SpreadsheetApp.openById(id);
}
function seedCategories_() {
  const sh=spreadsheet_().getSheetByName('Categories');
  if(sh.getLastRow()>1)return;
  const rows=[
    ['cat_income_salary','เงินเดือน','income','briefcase','#22C55E',true],['cat_income_other','รายรับอื่น','income','wallet','#16A34A',true],
    ['cat_food','อาหาร','expense','utensils','#F97316',true],['cat_transport','เดินทาง','expense','car','#3B82F6',true],['cat_shopping','ช้อปปิ้ง','expense','shopping-bag','#A855F7',true],
    ['cat_bills','บิลและค่าสาธารณูปโภค','expense','receipt','#EAB308',true],['cat_health','สุขภาพ','expense','heart-pulse','#EC4899',true],['cat_entertainment','บันเทิง','expense','gamepad-2','#8B5CF6',true],['cat_other','อื่น ๆ','expense','circle-ellipsis','#64748B',true]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}
function bootstrap_() {
  const ss=spreadsheet_(), data={};
  Object.keys(MF.TABLES).forEach(name=>data[name]=readTable_(ss.getSheetByName(name)));
  return {ok:true,data:data,serverTime:new Date().toISOString(),version:MF.VERSION};
}
function readTable_(sh) {
  const values=sh.getDataRange().getValues(); if(values.length<2)return [];
  const headers=values[0];
  return values.slice(1).filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>{
    const o={}; headers.forEach((h,i)=>{let v=r[i]; if(v instanceof Date)v=Utilities.formatDate(v,MF.TZ,"yyyy-MM-dd'T'HH:mm:ss"); o[h]=v;}); return o;
  });
}
function mutate_(p) {
  const table=String(p.table||''); if(!MF.TABLES[table])throw new Error('INVALID_TABLE');
  const op=String(p.op||''); const sh=spreadsheet_().getSheetByName(table);
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const headers=MF.TABLES[table], rows=readTable_(sh);
    if(op==='upsert') {
      let row=typeof p.row==='string'?JSON.parse(p.row):p.row; if(!row||typeof row!=='object')throw new Error('INVALID_ROW');
      if(!row.id)row.id=Utilities.getUuid();
      const now=new Date().toISOString(); if(headers.includes('updatedAt'))row.updatedAt=now; if(headers.includes('createdAt')&&!row.createdAt)row.createdAt=now;
      const index=rows.findIndex(r=>String(r.id)===String(row.id));
      const values=headers.map(h=>row[h]===undefined?'':row[h]);
      if(index>=0)sh.getRange(index+2,1,1,headers.length).setValues([values]); else sh.appendRow(values);
      return {ok:true,id:row.id,row:row};
    }
    if(op==='delete') {
      const id=String(p.id||''); const index=rows.findIndex(r=>String(r.id)===id); if(index<0)return {ok:true,deleted:false};
      sh.deleteRow(index+2); return {ok:true,deleted:true,id:id};
    }
    throw new Error('INVALID_OPERATION');
  } finally {lock.releaseLock();}
}
function jsonp_(obj, callback) {
  const json=JSON.stringify(obj).replace(/</g,'\\u003c').replace(/>/g,'\\u003e');
  if(callback && /^[A-Za-z_$][\w.$]*$/.test(callback)) return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันสาธารณะสำหรับกด Run ครั้งแรกจาก Apps Script Editor
function initializeDatabase() { return initialize_(); }
