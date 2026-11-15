const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const storage=new Map(),server={};let cached=null,patches=0,failVerify=false;
function context(){
 const c=vm.createContext({console,setTimeout,clearTimeout,Date,CONFIG:{filePath:'test.xlsx'},msal:{PublicClientApplication:class{getActiveAccount(){return {homeAccountId:'a'};}}},document:{readyState:'loading',addEventListener(){}},window:{addEventListener(){}},localStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)}});
 vm.runInContext(fs.readFileSync('onedrive.js','utf8'),c);
 c.XLSX={read:JSON.parse,write:JSON.stringify,utils:{decode_cell:s=>({r:Number(s.slice(1))-1,c:s.charCodeAt(0)-65}),encode_cell:p=>String.fromCharCode(65+p.c)+(p.r+1),sheet_add_aoa:(sheet,rows,{origin})=>{const r=Number(origin.slice(1)),col=origin.charCodeAt(0);rows.forEach((row,i)=>row.forEach((v,j)=>sheet[String.fromCharCode(col+j)+(r+i)]={v}));}}};
 c.server=server;c.getToken=async()=>'';c.setExcelBusyMessage=()=>{};c.withExcelBusy=async(_,f)=>f();c.readExcelDownloadCache=async()=>cached;c.invalidateExcelDownloadCache=async()=>{cached=null;};c.storeExcelDownloadCache=async b=>{cached=b;};
 c.writeExcelRangeNow=async(sheet,range,values)=>{patches++;server[sheet+range]=values;return {values};};c.readExcelRange=async(sheet,range)=>({values:failVerify?[['WRONG']]:server[sheet+range]});
 return c;
}
(async()=>{
 let c=context(),old=JSON.stringify({Sheets:{Transaction:{A2:{v:'deleted purchase'}},'Budget Setup':{B2:{v:100}}}});
 await c.writeExcelRanges([{sheetName:'Transaction',rangeAddress:'A2:A2',values:[['']]},{sheetName:'Budget Setup',rangeAddress:'B2:B2',values:[[200]]}]);
 assert.equal(patches,2);assert.equal(c.readExcelRecentRanges().length,2);assert.equal(cached,null);
 c=context(); // Navigation/reload retains markers, not old JavaScript state.
 let fixed=JSON.parse(await c.reconcileExcelDownload(old,true));assert.equal(fixed.Sheets.Transaction.A2.v,'');assert.equal(fixed.Sheets['Budget Setup'].B2.v,200);assert.equal(c.readExcelRecentRanges().length,2,'Old downloadable snapshot cannot retire markers');
 await c.reconcileExcelDownload(JSON.stringify(fixed),false);assert.equal(c.readExcelRecentRanges().length,2,'A patched cache is not server persistence proof');
 server['Budget SetupB2:B2']=[[300]];fixed=JSON.parse(await c.reconcileExcelDownload(old,true));assert.equal(fixed.Sheets['Budget Setup'].B2.v,300,'Live external edits win, no old write replay');
 await c.reconcileExcelDownload(JSON.stringify(fixed),true);assert.equal(c.readExcelRecentRanges().length,0,'Retire only when server file and live ranges match');
 failVerify=true;await assert.rejects(c.writeExcelRange('Transaction','A3:A3',[['new']]),/could not be verified/);assert.equal(c.readExcelRecentRanges().length,1,'Uncertain saves stay marked for reconciliation');failVerify=false;
 assert.ok(c.excelCellEqual((Date.UTC(2026,5,28)-Date.UTC(1899,11,30))/86400000,'2026-06-28'));assert.ok(!c.excelCellEqual(1,'1'));assert.ok(c.excelCellEqual(null,''));
 c=context();c.getToken=async()=>'token';c.downloadExcelBytes=async()=>old;
 let lockQueue=Promise.resolve();c.navigator={locks:{request:(_,fn)=>{const work=lockQueue.then(fn);lockQueue=work.catch(()=>{});return work;}}};
 const save=c.writeExcelRange('Budget Setup','B2:B2',[[450]]),load=c.downloadExcelFile(true);
 const [,bytes]=await Promise.race([Promise.all([save,load]),new Promise((_,reject)=>setTimeout(()=>reject(new Error('Save/load lock deadlocked')),1000))]);
 assert.equal(JSON.parse(bytes).Sheets['Budget Setup'].B2.v,450,'Queued save completes before reload reconciliation');
 cached=null;let release;const gate=new Promise(resolve=>release=resolve);let downloads=0;
 c.downloadExcelBytes=async()=>{downloads++;if(downloads===1)await gate;return old;};
 const first=c.downloadExcelFile(false),forced=c.downloadExcelFile(true);release();await Promise.all([first,forced]);assert.equal(downloads,2,'Explicit reload does not reuse an older in-flight download');
 console.log('Read-after-write: deletion and budget recovery across reload, live conflict precedence, marker retirement and verification failures passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
