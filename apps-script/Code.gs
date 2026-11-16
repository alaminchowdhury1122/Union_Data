/**
 * Union Data Backup + Passwords - Google Apps Script
 * Setup: Google Sheet > Extensions > Apps Script > paste this > Deploy > New deployment
 * Type: Web app | Execute as: Me | Who has access: Anyone > copy the /exec URL
 *
 * Roles:
 *   ADMIN_PASSWORD  -> can change everything (members, names, rates, entries)
 *   EDITOR_PASSWORD -> can add/edit contribution, donation and spending only
 *   everyone else   -> view only (no password needed)
 */
const ADMIN_PASSWORD  = 'CHANGE_ADMIN_PASSWORD';      // <-- put your own strong password
const EDITOR_PASSWORD = 'CHANGE_CONTRIBUTOR_PASSWORD'; // <-- shared with contributors

function doGet() {
  const json = rawSheet_().getRange('A1').getValue() || '{}';
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const role = roleOf_(body.password);
    if (role === 'blocked') return out_({ok:false, error:'too many tries'});
    if (!role) return out_({ok:false, error:'wrong password'});
    if (body.action === 'whoami') return out_({ok:true, role:role});

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      let db = body.data;
      if (!db || !db.members || !db.years) return out_({ok:false, error:'bad data'});
      if (role !== 'admin') {
        // contributors cannot change members (names, rates, add/remove): keep stored members
        const old = JSON.parse(rawSheet_().getRange('A1').getValue() || '{}');
        if (old.members) db.members = old.members;
      }
      rawSheet_().getRange('A1').setValue(JSON.stringify(db));
      writeReadable_(db);
      logBackup_(db, role);
    } finally { lock.releaseLock(); }
    return out_({ok:true, role:role});
  } catch (err) {
    return out_({ok:false, error:String(err)});
  }
}

// Returns 'admin', 'editor', 'blocked' (too many wrong tries) or ''
function roleOf_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = +(cache.get('fails') || 0);
  if (fails >= 15) return 'blocked';
  pw = String(pw || '');
  if (pw && pw === ADMIN_PASSWORD) return 'admin';
  if (pw && pw === EDITOR_PASSWORD) return 'editor';
  cache.put('fails', String(fails + 1), 600);   // after 15 wrong tries, wait 10 minutes
  return '';
}

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name){ const ss = SpreadsheetApp.getActiveSpreadsheet(); return ss.getSheetByName(name) || ss.insertSheet(name); }
function rawSheet_(){ return sheet_('RawData'); }

function writeReadable_(db){
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  Object.keys(db.years || {}).forEach(function(y){
    const d = db.years[y];
    const sh = sheet_('Year ' + y);
    sh.clear();
    const head = ['Name','Group','Monthly rate'].concat(months.map(function(m){return m+' paid';})).concat(months.map(function(m){return m+' donation';})).concat(['Total paid','Total donation']);
    const rows = [head];
    (db.members || []).forEach(function(m){
      const p = (d.pay && d.pay[m.id]) || [];
      let dn = (d.don && d.don[m.id]) || [];
      if (!Array.isArray(dn)) { const o = +dn || 0; dn = [o,0,0,0,0,0,0,0,0,0,0,0]; }
      const pa = [], da = [];
      for (let i=0;i<12;i++){ pa.push(+p[i]||0); da.push(+dn[i]||0); }
      const tp = pa.reduce(function(a,b){return a+b;},0), td = da.reduce(function(a,b){return a+b;},0);
      rows.push([m.n, m.g==='p'?'Probashi':'Elakabashi', m.r||0].concat(pa).concat(da).concat([tp,td]));
    });
    sh.getRange(1,1,rows.length,head.length).setValues(rows);
    sh.getRange(1,1,1,head.length).setFontWeight('bold').setBackground('#9d174d').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
    const start = rows.length + 3;
    sh.getRange(start,1).setValue('Spending').setFontWeight('bold');
    const sp = [['Event','Amount','Month']];
    (d.spend || []).forEach(function(s){ sp.push([s.r || '', +s.a || 0, months[s.m||0]]); });
    sh.getRange(start+1,1,sp.length,3).setValues(sp);
  });
}

function logBackup_(db, who){
  const sh = sheet_('BackupLog');
  if (sh.getLastRow() === 0) sh.appendRow(['Time','Role','Members','Size (chars)']);
  sh.appendRow([new Date(), who, (db.members||[]).length, JSON.stringify(db).length]);
}
