const assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const c=vm.createContext({console,setTimeout:(fn)=>{fn();return 0;},clearTimeout,CONFIG:{},msal:{PublicClientApplication:class{}},document:{readyState:'loading',addEventListener(){}},window:{addEventListener(){}}});
vm.runInContext(fs.readFileSync('onedrive.js','utf8'),c);
vm.runInContext('withExcelBusy=async(_,fn)=>fn();setExcelBusyMessage=()=>{};',c);
const ok=data=>({ok:true,json:async()=>data});const fail=(status,detail)=>({ok:false,status,text:async()=>detail,headers:{get:()=>null}});
(async()=>{
 let calls=[];c.fetch=async(url,options)=>{calls.push({url,options});return ok(url.endsWith('createSession')?{id:'session1'}:{values:[[1]]});};
 const url='https://graph.microsoft.com/v1.0/me/drive/root:/test.xlsx:/workbook/worksheets(1)/range(address=\'A1\')';
 await c.excelRangeRequest(url,'token');await c.excelRangeRequest(url,'token','PATCH',{values:[[2]]});
 assert.equal(calls.length,3);assert.equal(calls[2].options.headers['workbook-session-id'],'session1');
 let n=0;c.fetch=async()=>++n<3?fail(500,'GenericFileOpenError'):ok({});await c.excelRequest(url,{},true);assert.equal(n,3);
 n=0;c.fetch=async()=>{n++;return fail(500,'GenericFileOpenError');};await assert.rejects(c.excelRequest(url,{},true),/not confirmed saved/);assert.equal(n,3);
 n=0;await assert.rejects(c.excelRequest(url,{method:'POST'},false));assert.equal(n,1,'Non-idempotent POST never replayed');
 let seq=[fail(404,'invalidSessionNotFound'),ok({id:'new'}),ok({})];c.fetch=async()=>seq.shift();await c.excelRangeRequest(url,'token');assert.equal(seq.length,0);
 console.log('Excel session reuse, bounded transient retries, explicit failure, no POST replay and expired-session recovery passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
