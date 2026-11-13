const assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const c=vm.createContext({console,setTimeout,clearTimeout,CONFIG:{},msal:{PublicClientApplication:class{}},document:{readyState:'loading',addEventListener(){}},window:{addEventListener(){}}});
vm.runInContext(fs.readFileSync('onedrive.js','utf8'),c);
vm.runInContext('withExcelBusy=async(_,fn)=>fn();',c);
(async()=>{
 const calls=[],bytes=new ArrayBuffer(3);
 c.fetch=async(url,options)=>{calls.push({url,options});return {ok:true,json:async()=>({'@microsoft.graph.downloadUrl':'https://download.example/workbook?secret=temporary'}),arrayBuffer:async()=>bytes};};
 assert.equal(await c.downloadExcelBytes('Folder/test.xlsx','graph-token'),bytes);
 assert.ok(calls[0].url.endsWith('Folder/test.xlsx'));
 assert.equal(calls[0].options.headers.Authorization,'Bearer graph-token');
 assert.equal(calls[1].options.headers,undefined,'Bearer token must never reach download host');
 assert.equal(calls[1].options.credentials,'omit');
 assert.ok(calls.every(x=>!x.url.endsWith('/content')));
 c.fetch=async()=>({ok:true,json:async()=>({})});
 await assert.rejects(c.downloadExcelBytes('test.xlsx','token'),/did not provide/);
 let n=0;c.fetch=async()=>++n===1?{ok:true,json:async()=>({'@microsoft.graph.downloadUrl':'https://download.example/?secret=private'})}:Promise.reject(new TypeError('Failed to fetch'));
 await assert.rejects(c.downloadExcelBytes('test.xlsx','token'),e=>e.message.includes('saved data has not changed')&&!e.message.includes('private'));
 console.log('Browser workbook download: metadata resolution, token isolation, missing links and safe network failure passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
