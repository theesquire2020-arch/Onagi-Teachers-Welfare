/* Coast Region Onagi Teachers Welfare — Google Apps Script backend.
 * Deploy as a web app (execute as owner). Configure Script Properties per README.
 */
const SUPER_ADMIN = { username: 'admin@gmail.com', password: 'chief001' };
const FUNDS = ['Saving','Hospital Bill','Benevolence','Education','Charity'];
const TABS = {
  Members: ['MemberID','FullName','NationalID','TSCNo','Phone','Email','PasswordHash','Salt','SchoolType','County','SubCounty','SchoolName','SchoolPhone','Role','Status','JoinedAt','Gender'],
  Beneficiaries: ['BeneficiaryID','MemberID','FullName','Relationship','Phone','CreatedAt','Category'],
  Contributions: ['ContributionID','MemberID','Period','Amount','Method','Reference','ProviderRequestID','Status','PaidAt','CreatedAt','Category','BeneficiaryID'],
  Saving: ['MemberID','FullName','Total'],
  'Hospital Bill': ['MemberID','FullName','Total'],
  Benevolence: ['MemberID','FullName','Total'],
  Education: ['MemberID','FullName','Total'],
  Charity: ['MemberID','FullName','Total'],
  Loans: ['LoanID','MemberID','Amount','Purpose','TermMonths','InterestRate','Status','ApprovedBy','ApprovedAt','CreatedAt'],
  LoanRepayments: ['RepaymentID','LoanID','MemberID','Amount','Method','Reference','ProviderRequestID','Status','PaidAt','CreatedAt'],
  Bonuses: ['BonusID','MemberID','Amount','Reason','Status','AwardedBy','CreatedAt'],
  Expenditures: ['ExpenditureID','Category','Amount','Purpose','BeneficiaryName','Reference','RecordedBy','SpentAt'],
  Settings: ['Key','Value','UpdatedAt'],
  Gallery: ['GalleryID','Title','ImageURL','Description','CreatedBy','CreatedAt'],
  Documents: ['DocumentID','Title','FileURL','Category','Visibility','CreatedBy','CreatedAt'],
  AuditLog: ['AuditID','Actor','Action','Details','CreatedAt']
};
const COUNTIES = {
  'Mombasa':['Mvita','Nyali','Likoni','Changamwe','Jomvu','Kisauni'],
  'Kwale':['Msambweni','Lunga Lunga','Matuga','Kinango'],
  'Kilifi':['Kilifi North','Kilifi South','Kaloleni','Rabai','Ganze','Malindi','Magarini'],
  'Lamu':['Lamu East','Lamu West'],
  'Tana River':['Bura','Galole','Garsen'],
  'Taita Taveta':['Voi','Mwatate','Wundanyi','Taveta']
};
const DEFAULT_SETTINGS = { associationName:'Coast Region Onagi Teachers Welfare', postalAddress:'P.O. Box, Coast Region, Kenya', printFooter:'Together, we grow stronger.', logoUrl:'', currency:'KES', monthlyContribution:'1000', savingCadence:'Monthly', deployedUrl:'https://script.google.com/macros/s/AKfycbyd4zRrxgvSHNmTBF3FMhMDPRE5RsUt6VVpZ3gzQBv-VOwVSj7iuggqsqFjkbK99jfHIw', databaseRegistry:'[]', loanInterestRate:'8', maximumLoanMultiplier:'3', contactEmail:'', contactPhone:'' };

function doGet(e) {
  try {
    ensureDatabase_();
    const page = HtmlService.createTemplateFromFile('Index');
    page.counties = JSON.stringify(COUNTIES);
    return page.evaluate().setTitle('Onagi Teachers Welfare').addMetaTag('viewport','width=device-width, initial-scale=1');
  } catch (error) { return HtmlService.createHtmlOutput('<div style="font:16px Arial,sans-serif;max-width:650px;margin:80px auto;padding:30px"><h1>Google Sheet not connected</h1><p>Ask the administrator to bind this Apps Script project to a Google Sheet, or set <code>SPREADSHEET_ID</code> in Script Properties to a real Google Sheet ID. Run <code>setupDatabase()</code>, authorize it, then redeploy the web app.</p><p>An Apps Script deployment URL is not a Google Sheet URL.</p></div>'); }
}
function include(file) { return HtmlService.createHtmlOutputFromFile(file).getContent(); }
function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (body.action === 'paymentCallback' || (body.Body && body.Body.stkCallback)) return jsonOutput(paymentCallback_(body, e.parameter || {}));
    const actions = { register, login, logout, getDashboard, getMembers, getContributions, getLoans, getBonuses, getBeneficiaries, getGallery, getDocuments, getSettings, getPerformance, getExpenditures, addExpenditure, testConnection, registerDatabase, requestLoan, addBeneficiary, initiatePayment, addMemberContribution, approveLoan, updateMember, awardBonus, saveSettings, switchDatabase, exportData, rebuildMatrices, addGalleryItem, addDocument, getAuditLog };
    if (!Object.prototype.hasOwnProperty.call(actions, body.action)) throw Error('Unknown action.');
    return jsonOutput(actions[body.action](body.payload || {}, body.token || ''));
  } catch (error) { return jsonOutput(fail_(error)); }
}
function jsonOutput(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function ok_(data, message) { return {success:true, message:message || 'Success', data:data == null ? null : data}; }
function fail_(error) { return {success:false, message:String(error && error.message || error || 'Request failed'), data:null}; }
function safe_(value, max) {
  const text = String(value == null ? '' : value).trim().slice(0, max || 250).replace(/[\u0000-\u001f\u007f]/g,'');
  return /^[=+@\-\t\r]/.test(text) ? "'" + text : text;
}
function required_(value, label, max) { const s = safe_(value,max); if (!s) throw Error(label + ' is required.'); return s; }
function number_(value, label) { const n = Number(value); if (!Number.isFinite(n) || n <= 0) throw Error(label + ' must be greater than zero.'); return Math.round(n * 100) / 100; }
function now_() { return new Date().toISOString(); }
function props_() { return PropertiesService.getScriptProperties(); }
function ensureDatabase_() {
  const id=props_().getProperty('SPREADSHEET_ID');
  if(!id) return locked_(()=>setupDatabase());
  const ss=SpreadsheetApp.openById(id);
  const needsSetup=Object.keys(TABS).some(name=>{
    const sh=ss.getSheetByName(name),head=TABS[name];
    if(!sh||!sh.getLastRow()) return true;
    if(FUNDS.includes(name)) return sh.getRange(1,1,1,2).getValues()[0].join('|')!==head.slice(0,2).join('|')||sh.getRange(sh.getLastRow(),1).getValue()!=='TOTAL';
    return sh.getRange(1,1,1,head.length).getValues()[0].join('|')!==head.join('|');
  });
  if(needsSetup) return locked_(()=>setupDatabase());
}
function db_() {
  const id = props_().getProperty('SPREADSHEET_ID');
  if (!id) throw Error('Database not configured. Run setupDatabase() in the spreadsheet-bound script first.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) { const sh = db_().getSheetByName(name); if (!sh) throw Error('Missing sheet: ' + name + '. Run setupDatabase().'); return sh; }
function rows_(name) {
  const sh = sheet_(name), data = sh.getDataRange().getValues(), headers = TABS[name];
  return data.slice(1).filter(r => r[0] !== '').map((r,i) => {
    const item = { _row:i+2 }; headers.forEach((h,j) => item[h] = r[j] instanceof Date ? r[j].toISOString() : r[j]); return item;
  });
}
function append_(name, values) { const headers = TABS[name], sh = sheet_(name); sh.appendRow(headers.map(h => values[h] == null ? '' : values[h])); }
function update_(name, row, patch) {
  const sh = sheet_(name), headers = TABS[name], range = sh.getRange(row,1,1,headers.length), values = range.getValues()[0];
  Object.keys(patch).forEach(key => { const i = headers.indexOf(key); if (i >= 0) values[i] = patch[key]; }); range.setValues([values]);
}
function nextId_(name, prefix) {
  const key = 'COUNTER_' + name;
  const existing = Number(props_().getProperty(key) || 0);
  const sheetMax = rows_(name).reduce((max,r) => Math.max(max,Number(String(r[TABS[name][0]]).split('-').pop()) || 0),0);
  const next = Math.max(existing,sheetMax) + 1; props_().setProperty(key,String(next));
  return prefix + '-' + String(next).padStart(4,'0');
}
function locked_(fn) { const lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) throw Error('System busy. Please try again.'); try { return fn(); } finally { lock.releaseLock(); } }
function audit_(actor,action,details) { append_('AuditLog',{AuditID:nextId_('AuditLog','AUD'),Actor:actor,Action:action,Details:safe_(details,500),CreatedAt:now_()}); }
function publicMember_(member) { if (!member) return null; const {PasswordHash,Salt,_row,...rest} = member; return rest; }
function setting_(key) { const row = rows_('Settings').find(r => r.Key === key); return row ? String(row.Value) : DEFAULT_SETTINGS[key] || ''; }
function registry_() {
  let saved=[];try{saved=JSON.parse(setting_('databaseRegistry')||'[]');if(!Array.isArray(saved)) saved=[];}catch(e){saved=[];}
  const ss=db_();if(!saved.some(x=>x.id===ss.getId())) saved.unshift({id:ss.getId(),name:ss.getName(),url:ss.getUrl()});
  return saved.filter(x=>x&&typeof x.id==='string'&&typeof x.name==='string').slice(0,20);
}
function saveRegistry_(list, ss) {
  const sh=(ss||db_()).getSheetByName('Settings'),data=sh.getDataRange().getValues();
  const index=data.findIndex((row,i)=>i>0&&row[0]==='databaseRegistry'),value=JSON.stringify(list);
  if(index<0) sh.appendRow(['databaseRegistry',value,now_()]);
  else sh.getRange(index+1,2,1,2).setValues([[value,now_()]]);
}
function validateDatabase_(ss) {
  return Object.keys(TABS).filter(name=>{
    const sh=ss.getSheetByName(name),headers=TABS[name];if(!sh||!sh.getLastRow())return true;
    const current=sh.getRange(1,1,1,Math.max(headers.length,sh.getLastColumn())).getValues()[0];
    return FUNDS.includes(name)?current.slice(0,2).join('|')!==headers.slice(0,2).join('|')||current[sh.getLastColumn()-1]!=='Total'||sh.getRange(sh.getLastRow(),1).getValue()!=='TOTAL':current.slice(0,headers.length).join('|')!==headers.join('|');
  });
}
function testConnection(payload,token) { try { const s=session_(token),ss=db_(),missing=validateDatabase_(ss);return ok_({connected:missing.length===0,databaseName:ss.getName(),spreadsheetUrl:s.role==='SuperAdmin'?ss.getUrl():'',tabCount:Object.keys(TABS).length-missing.length,missingTabs:missing,deployedUrl:setting_('deployedUrl')||DEFAULT_SETTINGS.deployedUrl},missing.length?'Database needs setup: '+missing.join(', '):'Connected to '+ss.getName()+'.'); }catch(e){return fail_(e);} }
function registerDatabase(payload,token) { try {const s=session_(token);requireRole_(s,['SuperAdmin']);return locked_(()=>{
  const url=required_(payload.spreadsheetUrl,'Google Sheet URL',500),name=required_(payload.databaseName,'Database name',80);
  if(!/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9_-]+(?:\/|$)/.test(url))throw Error('Use a Google Sheet URL, not an Apps Script deployment URL.');
  const ss=SpreadsheetApp.openByUrl(url),missing=validateDatabase_(ss);if(missing.length)throw Error('Target Sheet needs these tabs or headers: '+missing.join(', '));
  const list=registry_(),existing=list.find(x=>x.id===ss.getId());if(existing){existing.name=name;existing.url=ss.getUrl();}else{if(list.length>=20)throw Error('Maximum 20 saved databases.');list.push({id:ss.getId(),name,url:ss.getUrl()});}
  saveRegistry_(list);audit_(who_(s),'REGISTER_DATABASE',ss.getId());return ok_({databases:list},'Database saved in Settings. Select it below to switch.');
});}catch(e){return fail_(e);} }
function fund_(category) { const value = String(category || 'Saving'); if (!FUNDS.includes(value)) throw Error('Choose a valid contribution fund.'); return value; }
function period_(category, supplied) {
  if (category === 'Saving') {
    const today = new Date(), month = Utilities.formatDate(today,'Africa/Nairobi','yyyy-MM');
    if (setting_('savingCadence') !== 'Weekly') return month;
    const local = new Date(Utilities.formatDate(today,'Africa/Nairobi','yyyy-MM-dd')+'T12:00:00Z');
    const day = local.getUTCDay() || 7; local.setUTCDate(local.getUTCDate()+4-day);
    const year = local.getUTCFullYear(), jan1 = new Date(Date.UTC(year,0,1));
    return year+'-W'+String(Math.ceil((((local-jan1)/86400000)+1)/7)).padStart(2,'0');
  }
  const value = required_(supplied,'Period',7);
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(value)) throw Error('Period must be YYYY-MM.');
  return value;
}
function ensureMatrixFooter_(fund) {
  const sh=sheet_(fund), last=sh.getLastRow();
  if(last===1 || sh.getRange(last,1).getValue()!=='TOTAL') sh.appendRow(['TOTAL',...Array(Math.max(0,sh.getLastColumn()-1)).fill('')]);
}
function matrixMember_(member) {
  FUNDS.forEach(fund => { const sh=sheet_(fund); ensureMatrixFooter_(fund);
    const last=sh.getLastRow(), ids=last>2?sh.getRange(2,1,last-2,1).getValues().flat():[];
    if(!ids.includes(member.MemberID)){sh.insertRowsBefore(last,1);sh.getRange(last,1,1,2).setValues([[member.MemberID,member.FullName]]);}
  });
}
function matrixColumn_(fund, title) {
  const sh=sheet_(fund), headings=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let col=headings.indexOf(title)+1;
  if(!col){col=sh.getLastColumn();sh.insertColumnBefore(col);sh.getRange(1,col).setValue(title).setBackground('#0c716d').setFontColor('#ffffff').setFontWeight('bold');}
  return col;
}
function recalcMatrixTotals_(fund) {
  const sh=sheet_(fund);ensureMatrixFooter_(fund);
  const last=sh.getLastRow(), width=sh.getLastColumn(), sums=Array(Math.max(0,width-3)).fill(0);
  if(last>2){const values=sh.getRange(2,3,last-2,width-2).getValues();values.forEach((cells,i)=>{
    let total=0,hasValue=false;cells.slice(0,-1).forEach((value,j)=>{if(value!==''&&value!=null){hasValue=true;const n=Number(value)||0;sums[j]+=n;total+=n;}});
    sh.getRange(i+2,width).setValue(hasValue?total:'');
  });}
  sh.getRange(last,1,1,width).setValues([['TOTAL','',...sums,sums.reduce((a,b)=>a+b,0)]]);
  sh.getRange(last,1,1,width).setBackground('#e5f4ed').setFontWeight('bold');
}
function beneficiary_(category, beneficiaryId) {
  const found=rows_('Beneficiaries').find(r=>r.BeneficiaryID===beneficiaryId&&r.Category===category);
  if(!found) throw Error('Choose a beneficiary in the '+category+' fund.');
  return found;
}
function syncMatrix_(category, memberId, period, beneficiaryId) {
  const fund=fund_(category), sh=sheet_(fund), member=rows_('Members').find(m=>m.MemberID===memberId);
  if(!member) throw Error('Member not found for matrix update.');
  matrixMember_(member);
  const heading=fund==='Saving'||!beneficiaryId?period:beneficiary_(fund,beneficiaryId).FullName;
  const col=matrixColumn_(fund,heading), ids=sh.getRange(2,1,sh.getLastRow()-2,1).getValues().flat(), row=ids.indexOf(memberId)+2;
  const matchingIds=beneficiaryId?rows_('Beneficiaries').filter(r=>r.Category===fund&&r.FullName.toLowerCase()===heading.toLowerCase()).map(r=>r.BeneficiaryID):[];
  const paid=rows_('Contributions').filter(r=>r.MemberID===memberId&&fund_(r.Category)===fund&&r.Status==='Paid'&&(fund==='Saving'||!beneficiaryId?!r.BeneficiaryID&&r.Period===period:matchingIds.includes(r.BeneficiaryID)));
  sh.getRange(row,col).setValue(paid.reduce((sum,r)=>sum+Number(r.Amount||0),0));
  recalcMatrixTotals_(fund);
}
function matrix_(fund, role, memberId) {
  const sh=sheet_(fund), values=sh.getDataRange().getValues(), headings=values[0]||TABS[fund], columns=headings.slice(2,-1);
  const visible=values.slice(1,-1).filter(r=>r[0]&&(role!=='Member'||r[0]===memberId));
  const rows=visible.map(r=>({memberId:r[0],fullName:r[1],total:r[headings.length-1]===''?'':Number(r[headings.length-1]||0),amounts:columns.map((_,i)=>r[i+2]===''?'':Number(r[i+2]||0))}));
  return {name:fund,columns,rows,columnTotals:columns.map((_,i)=>rows.reduce((sum,r)=>sum+Number(r.amounts[i]||0),0)),grandTotal:rows.reduce((sum,r)=>sum+Number(r.total||0),0)};
}

function setupDatabase() {
  const configured=props_().getProperty('SPREADSHEET_ID');
  const ss = configured?SpreadsheetApp.openById(configured):SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw Error('Open a Google Sheet, then run setupDatabase() from its bound Apps Script project.');
  props_().setProperty('SPREADSHEET_ID',ss.getId());
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
    const headers = TABS[name];
    if (sh.getLastRow() === 0) { sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setBackground('#0c716d').setFontColor('#ffffff').setFontWeight('bold'); }
    else {
      const current=sh.getRange(1,1,1,Math.max(headers.length,sh.getLastColumn())).getValues()[0];
      if (FUNDS.includes(name)) { if (current.slice(0,2).join('|')!==headers.slice(0,2).join('|') || current[sh.getLastColumn()-1]!=='Total') throw Error('Headers differ in '+name+'.'); }
      else if (current.slice(0,headers.length).join('|')!==headers.join('|')) {
        const count=sh.getLastColumn();
        if (['Members','Contributions','Beneficiaries'].includes(name) && count<headers.length && current.slice(0,count).join('|')===headers.slice(0,count).join('|')) sh.getRange(1,count+1,1,headers.length-count).setValues([headers.slice(count)]);
        else throw Error('Headers differ in ' + name + '; resolve manually before continuing.');
      }
    }
  });
  const existing = rows_('Settings').map(r => r.Key);
  Object.keys(DEFAULT_SETTINGS).forEach(key => { if (!existing.includes(key)) append_('Settings',{Key:key,Value:DEFAULT_SETTINGS[key],UpdatedAt:now_()}); });
  if(!setting_('deployedUrl')){const row=rows_('Settings').find(r=>r.Key==='deployedUrl');update_('Settings',row._row,{Value:DEFAULT_SETTINGS.deployedUrl,UpdatedAt:now_()});}
  saveRegistry_(registry_(),ss);
  rebuildMatrices_();
  return ok_({spreadsheetId:ss.getId(),sheets:Object.keys(TABS)},'Database ready.');
}
function hashPassword(pw,salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(pw), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function token_() { return Utilities.getUuid() + Utilities.getUuid(); }
function sessionKey_(token) { return 'SESSION_' + hashPassword(token,'session'); }
function session_(token) {
  if (!token || typeof token !== 'string') throw Error('Please sign in.');
  const raw = PropertiesService.getUserProperties().getProperty(sessionKey_(token));
  if (!raw) throw Error('Session expired. Please sign in again.');
  const session = JSON.parse(raw);
  if (Date.now() > session.expires) { PropertiesService.getUserProperties().deleteProperty(sessionKey_(token)); throw Error('Session expired. Please sign in again.'); }
  if (session.role !== 'SuperAdmin') {
    const member = rows_('Members').find(r => r.MemberID === session.memberId);
    if (!member || member.Status !== 'Active') throw Error('Account is inactive.');
    session.role = member.Role; session.member = publicMember_(member);
  }
  return session;
}
function requireRole_(session, roles) { if (!roles.includes(session.role)) throw Error('You do not have permission to perform this action.'); }
function who_(s) { return s.memberId || 'SUPERADMIN'; }
function ownOrAdmin_(s,id) { if (s.role === 'Member' && id !== s.memberId) throw Error('Access denied.'); }
function issueSession_(memberId,role,member) {
  const token = token_(), session = {memberId,role,expires:Date.now()+8*60*60*1000};
  PropertiesService.getUserProperties().setProperty(sessionKey_(token),JSON.stringify(session));
  return ok_({token,role,member,expires:session.expires},'Signed in successfully.');
}
function register(payload) {
  try { return locked_(() => {
    const p = payload || {};
    const fullName = required_(p.fullName,'Full name',120), nationalId = required_(p.nationalId,'National ID',30);
    const tscNo = required_(p.tscNo,'TSC number',30), phone = required_(p.phone,'Phone',20);
    const email = required_(p.email,'Email',160).toLowerCase(), password = String(p.password || '');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw Error('Enter a valid email address.');
    if (email === SUPER_ADMIN.username) throw Error('This email address is reserved.');
    if (password.length < 10 || password.length > 128) throw Error('Password must be 10–128 characters.');
    if (!/^\+?\d{9,15}$/.test(phone.replace(/[\s-]/g,''))) throw Error('Enter a valid phone number.');
    const schoolType = required_(p.schoolType,'School type',20);
    if (!['Public','Private'].includes(schoolType)) throw Error('Invalid school type.');
    const county = required_(p.county,'County',40), subCounty = required_(p.subCounty,'Sub-county',50);
    if (!COUNTIES[county] || !COUNTIES[county].includes(subCounty)) throw Error('Choose a valid county and sub-county.');
    const members = rows_('Members');
    if (members.some(m => String(m.Email).toLowerCase() === email || String(m.NationalID) === nationalId || String(m.TSCNo).toLowerCase() === tscNo.toLowerCase())) throw Error('Email, National ID or TSC number is already registered.');
    const salt = token_(), id = nextId_('Members','CRT');
    const gender=required_(p.gender,'Title',20); if(!['Male','Female','Other'].includes(gender)) throw Error('Choose a valid title option.');
    append_('Members',{MemberID:id,FullName:fullName,NationalID:nationalId,TSCNo:tscNo,Phone:phone,Email:email,PasswordHash:hashPassword(password,salt),Salt:salt,SchoolType:schoolType,County:county,SubCounty:subCounty,SchoolName:required_(p.schoolName,'School name',120),SchoolPhone:required_(p.schoolPhone,'School phone',20),Role:'Member',Status:'Active',JoinedAt:now_(),Gender:gender});
    matrixMember_({MemberID:id,FullName:fullName});
    audit_(id,'REGISTER','New member registered'); return ok_({memberId:id},'Registration complete. You can now sign in.');
  }); } catch (error) { return fail_(error); }
}
function login(payload) {
  try {
    const identity = String((payload || {}).email || (payload || {}).username || '').trim().toLowerCase();
    const password = String((payload || {}).password || '');
    // Required embedded override; rotate credentials before production deployment.
    if (identity === SUPER_ADMIN.username && password === SUPER_ADMIN.password) return issueSession_('SUPERADMIN','SuperAdmin',{MemberID:'SUPERADMIN',FullName:'System Administrator',Role:'SuperAdmin',Email:SUPER_ADMIN.username});
    const member = rows_('Members').find(m => String(m.Email).toLowerCase() === identity);
    if (!member || member.PasswordHash !== hashPassword(password,member.Salt)) throw Error('Invalid email or password.');
    if (member.Status !== 'Active') throw Error('Account is inactive.');
    return issueSession_(member.MemberID,member.Role,publicMember_(member));
  } catch (error) { return fail_(error); }
}
function logout(payload,token) { try { PropertiesService.getUserProperties().deleteProperty(sessionKey_(token)); return ok_(null,'Signed out.'); } catch(e) { return fail_(e); } }
function getDashboard(payload,token) { try {
  const s = session_(token), members = rows_('Members'), contributions = rows_('Contributions'), loans = rows_('Loans'), bonuses = rows_('Bonuses');
  const filter = r => s.role !== 'Member' || r.MemberID === s.memberId;
  const paid = contributions.filter(r => filter(r) && r.Status === 'Paid');
  const mine = loans.filter(filter), visibleMembers = s.role === 'Member' ? 1 : members.length;
  const fundTotals=Object.fromEntries(FUNDS.map(f=>[f,paid.filter(r=>fund_(r.Category)===f).reduce((sum,r)=>sum+Number(r.Amount||0),0)]));
  return ok_({member:s.member || null,stats:{members:visibleMembers,totalContributions:paid.reduce((a,r) => a+Number(r.Amount||0),0),activeLoans:mine.filter(r => r.Status === 'Approved').length,pendingLoans:mine.filter(r => r.Status === 'Pending').length,totalBonuses:bonuses.filter(r => filter(r) && r.Status === 'Awarded').reduce((a,r) => a+Number(r.Amount||0),0)},contributions:paid.slice(-8).reverse(),loans:mine.slice(-5).reverse(),monthly:monthly_(paid),fundTotals});
} catch(e) { return fail_(e); } }
function monthly_(items) { const out = {}; items.forEach(r => { const key = String(r.Period || r.PaidAt || '').slice(0,7); if (/^\d{4}-\d{2}$/.test(key)) out[key] = (out[key]||0)+Number(r.Amount||0); }); return Object.keys(out).sort().slice(-12).map(month => ({month,amount:out[month]})); }
function getMembers(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return ok_(rows_('Members').map(publicMember_)); } catch(e){return fail_(e);} }
function getContributions(payload,token) { try { const s=session_(token); const items=rows_('Contributions').filter(r => s.role !== 'Member' || r.MemberID===s.memberId); return ok_({matrices:FUNDS.map(fund=>matrix_(fund,s.role,s.memberId)),records:items.reverse(),beneficiaries:rows_('Beneficiaries').filter(r=>r.Category&&r.Category!=='Saving').map(r=>({id:r.BeneficiaryID,name:r.FullName,category:r.Category})),savingCadence:setting_('savingCadence')}); } catch(e){return fail_(e);} }
function getLoans(payload,token) { try { const s=session_(token); return ok_({loans:rows_('Loans').filter(r => s.role!=='Member'||r.MemberID===s.memberId).reverse(),repayments:rows_('LoanRepayments').filter(r => s.role!=='Member'||r.MemberID===s.memberId).reverse()}); } catch(e){return fail_(e);} }
function getBonuses(payload,token) { try { const s=session_(token); return ok_(rows_('Bonuses').filter(r => s.role!=='Member'||r.MemberID===s.memberId).reverse()); } catch(e){return fail_(e);} }
function getExpenditures(payload,token) { try {const s=session_(token);requireRole_(s,['Admin','SuperAdmin']);const items=rows_('Expenditures').reverse();return ok_({items,total:items.reduce((sum,r)=>sum+Number(r.Amount||0),0)});}catch(e){return fail_(e);} }
function addExpenditure(payload,token) { try {const s=session_(token);requireRole_(s,['Admin','SuperAdmin']);return locked_(()=>{
  const category=required_(payload.category,'Category',40);if(!FUNDS.includes(category)&&category!=='Other')throw Error('Choose an expenditure category.');
  const expenditureId=nextId_('Expenditures','EXP'),item={ExpenditureID:expenditureId,Category:category,Amount:number_(payload.amount,'Amount'),Purpose:required_(payload.purpose,'Purpose',250),BeneficiaryName:safe_(payload.beneficiaryName,120),Reference:'EXP-RCPT-'+expenditureId,RecordedBy:who_(s),SpentAt:now_()};
  append_('Expenditures',item);audit_(who_(s),'ADD_EXPENDITURE',expenditureId);return ok_(item,'Expenditure recorded. Reference: '+item.Reference);
});}catch(e){return fail_(e);} }
function getBeneficiaries(payload,token) { try { const s=session_(token); const id=safe_((payload||{}).memberId,40)||(s.role==='Member'?s.memberId:''); if(id) ownOrAdmin_(s,id); return ok_(rows_('Beneficiaries').filter(r => !id||r.MemberID===id)); } catch(e){return fail_(e);} }
function addBeneficiary(payload,token) { try { const s=session_(token); return locked_(() => { const id=safe_(payload.memberId,40)||s.memberId; ownOrAdmin_(s,id); if (!rows_('Members').some(m => m.MemberID===id)) throw Error('Member not found.'); const category=fund_(payload.category); if(category==='Saving') throw Error('Saving does not use beneficiaries.'); const name=required_(payload.fullName,'Full name',120); if(['memberid','fullname','total'].includes(name.toLowerCase())||/^\d{4}-(?:\d{2}|W\d{2})$/.test(name)) throw Error('Use a beneficiary name, not a reserved table heading.'); if(rows_('Beneficiaries').some(r=>r.Category===category&&r.FullName.toLowerCase()===name.toLowerCase())) throw Error('A beneficiary with this name already has a column in '+category+'.'); const item={BeneficiaryID:nextId_('Beneficiaries','BEN'),MemberID:id,FullName:name,Relationship:required_(payload.relationship,'Relationship',40),Phone:required_(payload.phone,'Phone',20),CreatedAt:now_(),Category:category}; append_('Beneficiaries',item); matrixColumn_(category,name); recalcMatrixTotals_(category); audit_(who_(s),'ADD_BENEFICIARY',item.BeneficiaryID+' '+category); return ok_(item,'Added '+name+' as a column in '+category+'.'); }); } catch(e){return fail_(e);} }
function requestLoan(payload,token) { try { const s=session_(token); if(s.role==='SuperAdmin') throw Error('A member account is required to request a loan.'); return locked_(() => { const amount=number_(payload.amount,'Amount'), term=Number(payload.termMonths); if(!Number.isInteger(term)||term<1||term>60) throw Error('Term must be 1–60 months.'); const contributions=rows_('Contributions').filter(r => r.MemberID===s.memberId&&r.Status==='Paid'&&fund_(r.Category)==='Saving').reduce((a,r)=>a+Number(r.Amount),0); const ceiling=contributions*Number(setting_('maximumLoanMultiplier')); if(amount>ceiling) throw Error('Loan exceeds your eligible limit of KES '+ceiling.toLocaleString()+'.'); const item={LoanID:nextId_('Loans','LOAN'),MemberID:s.memberId,Amount:amount,Purpose:required_(payload.purpose,'Purpose',250),TermMonths:term,InterestRate:Number(setting_('loanInterestRate')),Status:'Pending',ApprovedBy:'',ApprovedAt:'',CreatedAt:now_()}; append_('Loans',item); audit_(who_(s),'REQUEST_LOAN',item.LoanID); return ok_(item,'Loan request submitted for review.'); }); } catch(e){return fail_(e);} }
function approveLoan(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const loan=rows_('Loans').find(r => r.LoanID===safe_(payload.loanId,40)); if(!loan) throw Error('Loan not found.'); if(loan.Status!=='Pending') throw Error('Loan already reviewed.'); const status=payload.approve===true?'Approved':'Declined'; update_('Loans',loan._row,{Status:status,ApprovedBy:who_(s),ApprovedAt:now_()}); audit_(who_(s),'REVIEW_LOAN',loan.LoanID+' '+status); return ok_({loanId:loan.LoanID,status},'Loan '+status.toLowerCase()+'.'); }); } catch(e){return fail_(e);} }
function addMemberContribution(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const id=required_(payload.memberId,'Member ID',40), member=rows_('Members').find(m=>m.MemberID===id); if(!member) throw Error('Member not found.'); const category=fund_(payload.category), period=period_(category,payload.period), beneficiaryId=category==='Saving'?'':required_(payload.beneficiaryId,'Beneficiary',40); if(beneficiaryId) beneficiary_(category,beneficiaryId); const contributionId=nextId_('Contributions','CON'); let ref='RCPT-'+contributionId; if(rows_('Contributions').some(r=>r.Reference===ref)) ref+='-'+Utilities.getUuid().slice(0,8); const item={ContributionID:contributionId,MemberID:id,Period:period,Amount:number_(payload.amount,'Amount'),Method:'Manual',Reference:ref,Status:'Paid',PaidAt:now_(),CreatedAt:now_(),Category:category,BeneficiaryID:beneficiaryId}; append_('Contributions',item); syncMatrix_(category,id,period,beneficiaryId); audit_(who_(s),'RECORD_CONTRIBUTION',item.ContributionID+' '+category); return ok_(item,category+' contribution receipted. Receipt: '+ref); }); } catch(e){return fail_(e);} }
function awardBonus(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const id=required_(payload.memberId,'Member ID',40); if(!rows_('Members').some(m=>m.MemberID===id)) throw Error('Member not found.'); const item={BonusID:nextId_('Bonuses','BON'),MemberID:id,Amount:number_(payload.amount,'Amount'),Reason:required_(payload.reason,'Reason',250),Status:'Awarded',AwardedBy:who_(s),CreatedAt:now_()}; append_('Bonuses',item); audit_(who_(s),'AWARD_BONUS',item.BonusID); return ok_(item,'Bonus awarded.'); }); } catch(e){return fail_(e);} }
function updateMember(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const member=rows_('Members').find(m=>m.MemberID===safe_(payload.memberId,40)); if(!member) throw Error('Member not found.'); if(member.Role==='SuperAdmin') throw Error('Super admin cannot be changed.'); const patch={}; if(payload.status != null) { if(!['Active','Inactive'].includes(payload.status)) throw Error('Invalid status.'); patch.Status=payload.status; } if(payload.role != null) { if(s.role!=='SuperAdmin') throw Error('Only SuperAdmin can assign roles.'); if(!['Member','Admin'].includes(payload.role)) throw Error('Invalid role.'); patch.Role=payload.role; } if(!Object.keys(patch).length) throw Error('No changes supplied.'); update_('Members',member._row,patch); audit_(who_(s),'UPDATE_MEMBER',member.MemberID+' '+JSON.stringify(patch)); return ok_(patch,'Member updated.'); }); } catch(e){return fail_(e);} }
function getGallery(payload,token) { try { session_(token); return ok_(rows_('Gallery').reverse()); } catch(e){return fail_(e);} }
function getDocuments(payload,token) { try { const s=session_(token); return ok_(rows_('Documents').filter(r=>s.role!=='Member'||r.Visibility==='All').reverse()); } catch(e){return fail_(e);} }
function validUrl_(url) { const s=required_(url,'URL',1000); if(!/^https:\/\/[^\s]+$/i.test(s)) throw Error('Use a secure https:// URL.'); return s; }
function addGalleryItem(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const item={GalleryID:nextId_('Gallery','GAL'),Title:required_(payload.title,'Title',120),ImageURL:validUrl_(payload.imageUrl),Description:safe_(payload.description,500),CreatedBy:who_(s),CreatedAt:now_()}; append_('Gallery',item); audit_(who_(s),'ADD_GALLERY',item.GalleryID); return ok_(item,'Gallery item added.'); }); } catch(e){return fail_(e);} }
function addDocument(payload,token) { try { const s=session_(token); requireRole_(s,['Admin','SuperAdmin']); return locked_(() => { const visibility=payload.visibility==='Staff'?'Staff':'All'; const item={DocumentID:nextId_('Documents','DOC'),Title:required_(payload.title,'Title',120),FileURL:validUrl_(payload.fileUrl),Category:required_(payload.category,'Category',60),Visibility:visibility,CreatedBy:who_(s),CreatedAt:now_()}; append_('Documents',item); audit_(who_(s),'ADD_DOCUMENT',item.DocumentID); return ok_(item,'Document added.'); }); } catch(e){return fail_(e);} }
function getSettings(payload,token) { try {session_(token);const ss=db_(),settings=Object.fromEntries(rows_('Settings').map(r=>[r.Key,r.Value]));delete settings.databaseRegistry;return ok_({...settings,deployedUrl:settings.deployedUrl||DEFAULT_SETTINGS.deployedUrl,spreadsheetUrl:ss.getUrl(),databaseStatus:'Live database',activeDatabaseId:ss.getId(),activeDatabaseName:ss.getName(),databases:registry_()});}catch(e){return fail_(e);} }
function saveSettings(payload,token) { try { const s=session_(token); requireRole_(s,['SuperAdmin']); return locked_(() => { const allowed=Object.keys(DEFAULT_SETTINGS).filter(key=>key!=='databaseRegistry'); const changed={}; allowed.forEach(key => { if(!Object.prototype.hasOwnProperty.call(payload,key)) return; let value=safe_(payload[key],500); if(['monthlyContribution','loanInterestRate','maximumLoanMultiplier'].includes(key)) { const n=Number(value); if(!Number.isFinite(n)||n<0||n>10000000) throw Error('Invalid '+key+'.'); value=String(n); } if(key==='savingCadence'&&!['Weekly','Monthly'].includes(value)) throw Error('Saving cadence must be Weekly or Monthly.'); if(key==='deployedUrl'&&value&&!/^https:\/\/script\.google\.com\/macros\/s\/[a-zA-Z0-9_-]+(?:\/(?:exec|dev))?(?:[?].*)?$/.test(value)) throw Error('Enter a Google Apps Script deployment URL.'); if(key==='logoUrl'&&value) value=validUrl_(value); const row=rows_('Settings').find(r=>r.Key===key); if(row) update_('Settings',row._row,{Value:value,UpdatedAt:now_()}); else append_('Settings',{Key:key,Value:value,UpdatedAt:now_()}); changed[key]=value; }); audit_(who_(s),'UPDATE_SETTINGS',JSON.stringify(changed)); return ok_(changed,'Settings saved.'); }); } catch(e){return fail_(e);} }
function exportData(payload,token) { try { const s=session_(token);requireRole_(s,['SuperAdmin']);const clean=name=>rows_(name).map(({_row,...record})=>record);return ok_({exportedAt:now_(),members:rows_('Members').map(publicMember_),beneficiaries:clean('Beneficiaries'),contributions:clean('Contributions'),loans:clean('Loans'),loanRepayments:clean('LoanRepayments'),bonuses:clean('Bonuses'),expenditures:clean('Expenditures'),settings:clean('Settings')},'Data export ready. Password hashes and salts are excluded.'); } catch(e){return fail_(e);} }
function rebuildMatrices(payload,token) { try { const s=session_(token);requireRole_(s,['SuperAdmin']);return locked_(()=>{rebuildMatrices_();audit_(who_(s),'REBUILD_MATRICES','Recalculated all contribution matrices');return ok_({funds:FUNDS},'All five matrices rebuilt from the paid contribution ledger.');}); } catch(e){return fail_(e);} }
function rebuildMatrices_() {
  FUNDS.forEach(fund=>{ensureMatrixFooter_(fund);const sh=sheet_(fund),last=sh.getLastRow(),width=sh.getLastColumn();if(last>2)sh.getRange(2,3,last-2,width-2).clearContent();});
  rows_('Members').forEach(matrixMember_);
  rows_('Beneficiaries').filter(r=>r.Category&&r.Category!=='Saving').forEach(r=>matrixColumn_(fund_(r.Category),r.FullName));
  const keys=new Set();rows_('Contributions').filter(r=>r.Status==='Paid').forEach(r=>{const key=[fund_(r.Category),r.MemberID,r.BeneficiaryID||r.Period].join('|');if(!keys.has(key)){keys.add(key);syncMatrix_(r.Category,r.MemberID,String(r.Period),r.BeneficiaryID);}});
  FUNDS.forEach(recalcMatrixTotals_);
}
function switchDatabase(payload,token) { try { const s=session_(token); requireRole_(s,['SuperAdmin']); return locked_(() => {
  const list=registry_(),entry=list.find(x=>x.id===safe_(payload.databaseId,120));
  if(!entry)throw Error('Choose a saved database in Settings first.');
  const target=SpreadsheetApp.openById(entry.id),missing=validateDatabase_(target);
  if(missing.length)throw Error('Target database has missing tabs or headers: '+missing.join(', '));
  saveRegistry_(list);
  props_().setProperty('SPREADSHEET_ID',target.getId());
  saveRegistry_(list,target);
  audit_(who_(s),'SWITCH_DATABASE','Switched to '+target.getId());
  return ok_({spreadsheetUrl:target.getUrl(),databaseStatus:'Live database',activeDatabaseName:entry.name},'Live database switched to '+entry.name+'.');
}); } catch(e){return fail_(e);} }
function getPerformance(payload,token) { try {
  const s=session_(token), allMembers=rows_('Members'), members=allMembers.filter(m=>s.role!=='Member'||m.MemberID===s.memberId);
  const contributions=rows_('Contributions').filter(r=>r.Status==='Paid'&&(s.role!=='Member'||r.MemberID===s.memberId));
  const loans=rows_('Loans'), currentSaving=period_('Saving'), multiplier=Number(setting_('maximumLoanMultiplier'));
  const byCounty={}, fundTotals=Object.fromEntries(FUNDS.map(f=>[f,0]));
  contributions.forEach(r=>{const m=allMembers.find(x=>x.MemberID===r.MemberID),amount=Number(r.Amount||0);if(m)byCounty[m.County]=(byCounty[m.County]||0)+amount;fundTotals[fund_(r.Category)]+=amount;});
  const memberPerformance=members.map(m=>{
    const paid=contributions.filter(r=>r.MemberID===m.MemberID), savings=paid.filter(r=>fund_(r.Category)==='Saving');
    const funds=Object.fromEntries(FUNDS.map(f=>[f,paid.filter(r=>fund_(r.Category)===f).reduce((sum,r)=>sum+Number(r.Amount||0),0)]));
    const ownLoans=loans.filter(l=>l.MemberID===m.MemberID), last=savings.map(r=>String(r.PaidAt||r.CreatedAt||'')).sort().pop()||'';
    return {memberId:m.MemberID,fullName:m.FullName,county:m.County,school:m.SchoolName,status:m.Status,funds,total:paid.reduce((sum,r)=>sum+Number(r.Amount||0),0),savingPeriods:new Set(savings.map(r=>String(r.Period))).size,lastSaved:last,savingIndicator:savings.some(r=>String(r.Period)===currentSaving)?'On track':savings.length?'Review saving':'No saving yet',loanLimit:funds.Saving*multiplier,activeLoans:ownLoans.filter(l=>l.Status==='Approved').length,pendingLoans:ownLoans.filter(l=>l.Status==='Pending').length};
  });
  return ok_({monthly:monthly_(contributions),byCounty:Object.keys(byCounty).map(county=>({county,amount:byCounty[county]})),memberCount:members.length,memberPerformance,fundTotals,savingPeriod:currentSaving});
} catch(e){return fail_(e);} }
function getAuditLog(payload,token) { try { const s=session_(token); requireRole_(s,['SuperAdmin']); return ok_(rows_('AuditLog').slice(-200).reverse()); } catch(e){return fail_(e);} }

/* Payment request only creates PENDING rows. Never mark paid until an authenticated
 * provider callback confirms success. No PIN is accepted or stored anywhere. */
function initiatePayment(payload,token) { try { const s=session_(token); if(s.role==='SuperAdmin') throw Error('Use a member account for payments.'); return locked_(() => {
  const provider=required_(payload.provider,'Provider',20), kind=payload.kind==='repayment'?'repayment':'contribution';
  if(!['MPESA','AIRTEL'].includes(provider)) throw Error('Unsupported provider.');
  const amount=number_(payload.amount,'Amount'), phone=required_(payload.phone,'Phone',20).replace(/[^\d]/g,'');
  if(!Number.isInteger(amount)) throw Error('Mobile money payments must use whole KES amounts.');
  if(!/^254\d{9}$/.test(phone)) throw Error('Enter phone in 2547XXXXXXXX format.');
  const id=kind==='repayment'?nextId_('LoanRepayments','REP'):nextId_('Contributions','CON');
  const created=now_(), reference='ONAGI-'+id;
  let loanId='',period='',category='',beneficiaryId='';
  if(kind==='repayment') { loanId=required_(payload.loanId,'Loan ID',40); const loan=rows_('Loans').find(r=>r.LoanID===loanId&&r.MemberID===s.memberId&&r.Status==='Approved'); if(!loan) throw Error('Approved loan not found.'); }
  else { category=fund_(payload.category); period=period_(category,payload.period);if(category!=='Saving'){beneficiaryId=required_(payload.beneficiaryId,'Beneficiary',40);beneficiary_(category,beneficiaryId);} }
  const name=kind==='repayment'?'LoanRepayments':'Contributions';
  append_(name,kind==='repayment'?{RepaymentID:id,LoanID:loanId,MemberID:s.memberId,Amount:amount,Method:provider,Reference:reference,Status:'Pending',PaidAt:'',CreatedAt:created}:{ContributionID:id,MemberID:s.memberId,Period:period,Amount:amount,Method:provider,Reference:reference,Status:'Pending',PaidAt:'',CreatedAt:created,Category:category,BeneficiaryID:beneficiaryId});
  try { const response=provider==='MPESA'?mpesaPush_(phone,amount,reference):airtelPush_(phone,amount,reference); const providerRef=response.CheckoutRequestID||response.transactionId||response.reference||reference; update_(name,rows_(name).find(r=>r[name==='Contributions'?'ContributionID':'RepaymentID']===id)._row,{ProviderRequestID:String(providerRef)}); audit_(who_(s),'PAYMENT_INITIATED',id+' '+provider); return ok_({id,reference:providerRef,status:'Pending'},'Payment prompt sent. Complete it on your phone.'); }
  catch(err) { update_(name,rows_(name).find(r=>r[name==='Contributions'?'ContributionID':'RepaymentID']===id)._row,{Status:'Failed'}); throw err; }
}); } catch(e){return fail_(e);} }
function mpesaPush_(phone,amount,reference) {
  const p=props_(), key=p.getProperty('MPESA_CONSUMER_KEY'), secret=p.getProperty('MPESA_CONSUMER_SECRET'), shortcode=p.getProperty('MPESA_SHORTCODE'), passkey=p.getProperty('MPESA_PASSKEY'), callback=p.getProperty('PAYMENT_CALLBACK_URL');
  if(!key||!secret||!shortcode||!passkey||!callback) throw Error('M-Pesa is not configured. Contact the administrator.');
  const base=p.getProperty('MPESA_BASE_URL')||'https://sandbox.safaricom.co.ke';
  const auth=UrlFetchApp.fetch(base+'/oauth/v1/generate?grant_type=client_credentials',{headers:{Authorization:'Basic '+Utilities.base64Encode(key+':'+secret)},muteHttpExceptions:true});
  if(auth.getResponseCode()!==200) throw Error('M-Pesa authentication failed.');
  const access=JSON.parse(auth.getContentText()).access_token, stamp=Utilities.formatDate(new Date(),'Africa/Nairobi','yyyyMMddHHmmss');
  const body={BusinessShortCode:shortcode,Password:Utilities.base64Encode(shortcode+passkey+stamp),Timestamp:stamp,TransactionType:'CustomerPayBillOnline',Amount:Math.round(amount),PartyA:phone,PartyB:shortcode,PhoneNumber:phone,CallBackURL:callback,AccountReference:reference,TransactionDesc:'Onagi welfare payment'};
  const res=UrlFetchApp.fetch(base+'/mpesa/stkpush/v1/processrequest',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+access},payload:JSON.stringify(body),muteHttpExceptions:true});
  const data=JSON.parse(res.getContentText()); if(res.getResponseCode()>=300||data.ResponseCode!=='0') throw Error('M-Pesa request was rejected: '+(data.errorMessage||data.ResponseDescription||'Please retry.')); return data;
}
function airtelPush_(phone,amount,reference) {
  // Airtel implementations vary by market/merchant: configure the approved gateway endpoint and bearer token.
  const p=props_(), endpoint=p.getProperty('AIRTEL_PAYMENT_URL'), bearer=p.getProperty('AIRTEL_BEARER_TOKEN');
  if(!endpoint||!bearer||!/^https:\/\//.test(endpoint)) throw Error('Airtel Money is not configured. Contact the administrator.');
  const res=UrlFetchApp.fetch(endpoint,{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+bearer},payload:JSON.stringify({reference,amount,currency:'KES',phone,callbackUrl:p.getProperty('PAYMENT_CALLBACK_URL')}),muteHttpExceptions:true});
  const data=JSON.parse(res.getContentText()); if(res.getResponseCode()>=300) throw Error('Airtel request was rejected.'); return data;
}
function paymentCallback_(body,params) {
  try { const expected=props_().getProperty('PAYMENT_CALLBACK_SECRET'); if(!expected||String(params.key||body.key||'')!==expected) throw Error('Unauthorized callback.'); return locked_(() => {
    const stk=body.Body&&body.Body.stkCallback;
    const reference=stk?String(stk.CheckoutRequestID):safe_(body.reference||body.transactionId,100);
    if(!reference) throw Error('Missing payment reference.');
    const name=['Contributions','LoanRepayments'].find(tab=>rows_(tab).some(r=>String(r.ProviderRequestID)===reference || String(r.Reference)===reference));
    if(!name) throw Error('Payment reference not found.');
    const row=rows_(name).find(r=>String(r.ProviderRequestID)===reference || String(r.Reference)===reference);
    if(row.Status!=='Pending') return ok_({status:row.Status},'Already processed.');
    const success=stk?Number(stk.ResultCode)===0:body.status==='SUCCESS';
    // Airtel callback must include verified provider reference; shared callback key must be kept secret.
    const metadata=stk&&stk.CallbackMetadata&&stk.CallbackMetadata.Item||[];
    const receipt=stk?String((metadata.find(x=>x.Name==='MpesaReceiptNumber')||{}).Value||''):safe_(body.receipt||body.providerReceipt,100);
    if(success&&!receipt) throw Error('Successful payment requires provider receipt.');
    const confirmedAmount=stk?(metadata.find(x=>x.Name==='Amount')||{}).Value:body.amount;
    if(success&&Number(confirmedAmount)!==Number(row.Amount)) throw Error('Payment amount mismatch. Reconcile with provider.');
    if(success&&rows_(name).some(r=>r.Reference===receipt&&r._row!==row._row)) throw Error('Receipt already recorded.');
    update_(name,row._row,{Status:success?'Paid':'Failed',PaidAt:success?now_():'',Reference:success?receipt:reference});
    if(success&&name==='Contributions') syncMatrix_(row.Category,row.MemberID,String(row.Period),row.BeneficiaryID);
    audit_('PAYMENT_CALLBACK',success?'PAYMENT_RECEIPTED':'PAYMENT_FAILED',(name==='Contributions'?row.ContributionID:row.RepaymentID)+' '+reference);
    return ok_({status:success?'Paid':'Failed',receipt:success?receipt:null},'Callback processed.');
  }); } catch(e){return fail_(e);} }
