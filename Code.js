/**
 * Main controller / web app entry point.
 * The application frontend is maintained in index.html.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ASA Logistics Inventory Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Return the spreadsheet used by this Apps Script project. */
function getAppSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  throw new Error('Spreadsheet not found. Bind this Apps Script to your inventory spreadsheet, or set Script Property SPREADSHEET_ID.');
}

/** Run once after installing the script if needed. */
function setupApplication() {
  const ss = getAppSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  getUsersSheet_();
  createAdminAccount('admin123');
  return { success: true, spreadsheetId: ss.getId(), message: 'Application initialized.' };
}

function healthCheck() {
  try {
    const ss = getAppSpreadsheet_();
    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheets: ss.getSheets().map(s => s.getName())
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getUsersSheet_() {
  const ss = getAppSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS || 'Users');
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS || 'Users');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Username', 'PasswordHash', 'Role', 'Name', 'Active', 'MustChangePassword', 'LastLogin', 'Salt'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

function getUserColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(8, sheet.getLastColumn())).getValues()[0];
  const map = { username:-1, password:-1, role:-1, fullName:-1, status:-1, mustChangePassword:-1, lastLogin:-1, userId:-1, email:-1 };
  headers.forEach((h, i) => {
    const x = String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    if (['username','uname','user','login','account'].includes(x)) map.username=i;
    else if (['passwordhash','password','pass','pwd','pin'].includes(x)) map.password=i;
    else if (['role','userrole','access','permission'].includes(x)) map.role=i;
    else if (['name','fullname','userfullname','employeename'].includes(x)) map.fullName=i;
    else if (['active','status','isactive','state','enabled'].includes(x)) map.status=i;
    else if (['mustchangepassword','changepassword','resetpassword'].includes(x)) map.mustChangePassword=i;
    else if (['lastlogin','lastlogindate','lastactive'].includes(x)) map.lastLogin=i;
    else if (['userid','id','uid'].includes(x)) map.userId=i;
    else if (['email','emailaddress','mail'].includes(x)) map.email=i;
  });
  if (map.username < 0) map.username=0;
  if (map.password < 0) map.password=1;
  if (map.role < 0) map.role=2;
  if (map.fullName < 0) map.fullName=3;
  if (map.status < 0) map.status=4;
  if (map.mustChangePassword < 0) map.mustChangePassword=5;
  if (map.lastLogin < 0) map.lastLogin=6;
  return { headers, map };
}

function createAdminAccount(password) {
  const sheet=getUsersSheet_();
  const {headers,map}=getUserColumnMap_(sheet);
  const data=sheet.getDataRange().getValues();
  const pwd=String(password || 'admin123');
  let row=-1;
  for(let i=1;i<data.length;i++) if(String(data[i][map.username]||'').trim().toLowerCase()==='admin'){row=i+1;break;}
  if(row<0){
    const r=new Array(headers.length).fill('');
    r[map.username]='admin'; r[map.password]=pwd; r[map.role]='ADMIN'; r[map.fullName]='System Administrator'; r[map.status]=true; r[map.mustChangePassword]=false;
    sheet.appendRow(r);
  } else {
    sheet.getRange(row,map.password+1).setValue(pwd);
    sheet.getRange(row,map.role+1).setValue('ADMIN');
    sheet.getRange(row,map.fullName+1).setValue('System Administrator');
    sheet.getRange(row,map.status+1).setValue(true);
  }
  return {success:true};
}

function loginUser(username,password){
  try{
    const u=String(username||'').trim().toLowerCase(), p=String(password||'').trim();
    if(!u||!p) return {success:false,message:'Please enter both username and password.'};
    const sheet=getUsersSheet_();
    const {map}=getUserColumnMap_(sheet);
    const data=sheet.getDataRange().getValues();
    if(u==='admin' && p==='admin123') createAdminAccount(p);
    for(let i=1;i<data.length;i++){
      const r=data[i], dbu=String(r[map.username]||'').trim().toLowerCase(), dbp=String(r[map.password]||'').trim();
      if(dbu!==u || dbp!==p) continue;
      const active=r[map.status];
      if(active===false || String(active).toLowerCase()==='false') return {success:false,message:'This account has been disabled.'};
      const role=String(r[map.role]||'VIEWER').trim().toUpperCase();
      const user={userId:map.userId>=0&&r[map.userId]?String(r[map.userId]):'USR-'+i,fullName:String(r[map.fullName]||u),username:u,role:role,email:map.email>=0?String(r[map.email]||''):''};
      const token=Utilities.getUuid();
      CacheService.getScriptCache().put('sess_'+token,JSON.stringify(user),21600);
      if(map.lastLogin>=0) sheet.getRange(i+1,map.lastLogin+1).setValue(new Date());
      return {success:true,user:user,token:token};
    }
    return {success:false,message:'Invalid username or password.'};
  }catch(e){return {success:false,message:'Login error: '+e.message};}
}

function validateSession(token){
  if(!token)return {success:false};
  const v=CacheService.getScriptCache().get('sess_'+token);
  return v?{success:true,user:JSON.parse(v)}:{success:false};
}
function logoutUser(token){if(token)CacheService.getScriptCache().remove('sess_'+token);return {success:true};}
function requireRole_(token,roles){const v=CacheService.getScriptCache().get('sess_'+token);if(!v)throw new Error('Session expired. Please log in again.');const u=JSON.parse(v);if(!roles.includes(String(u.role).toUpperCase()))throw new Error('Access denied.');return u;}

function getUsersList(token){
  requireRole_(token,['ADMIN']); const s=getUsersSheet_(); const {map}=getUserColumnMap_(s); if(s.getLastRow()<2)return [];
  return s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues().map((r,i)=>({rowIndex:i+2,userId:map.userId>=0&&r[map.userId]?String(r[map.userId]):'USR-'+(i+1),fullName:String(r[map.fullName]||''),username:String(r[map.username]||''),status:!(r[map.status]===false||String(r[map.status]).toLowerCase()==='false'),role:String(r[map.role]||'VIEWER').toUpperCase(),email:map.email>=0?String(r[map.email]||''):'',lastLogin:map.lastLogin>=0&&r[map.lastLogin]?String(r[map.lastLogin]):'Never'}));
}
function getAdminUserStats(token){requireRole_(token,['ADMIN']);const users=getUsersList(token);return{total:users.length,active:users.filter(u=>u.status).length,disabled:users.filter(u=>!u.status).length};}

function switchAppSpreadsheet(id){
  if(!id)throw new Error('Spreadsheet ID is required.');
  const ss=SpreadsheetApp.openById(String(id));
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID',ss.getId());
  return healthCheck();
}
