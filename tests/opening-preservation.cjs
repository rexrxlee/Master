const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const fields={setupDate_0:{value:'2026-01-01'},setupAmt_0:{value:'200'},setupDate_1:{value:'2026-11-13'},setupAmt_1:{value:''}};
const original=[['Date','Transaction','Amount','Main Category','Sub Category','Account','Claimable'],['2026-01-01','Opening Balance',100,'Income','Opening Balance','UOB'],['2026-11-13','Transfer',5,'Transfer','Transfer Out','UOB'],['2026-01-01','Opening Balance',50,'Income','Opening Balance','Other']];
let rows=original,writes=[];
const c=vm.createContext({console,document:{getElementById:id=>fields[id]||null},log(){},alert(){},CONFIG:{sheetName:'Transaction'},getToken:async()=>'',getEncodedExcelPath:()=>'',graphGetJson:async()=>({values:rows,rowIndex:0,columnIndex:0}),writeExcelRange:async(sheet,range,values)=>writes.push({range,values})});
vm.runInContext(fs.readFileSync('setup.js','utf8'),c);vm.runInContext("setupAccounts=[{name:'UOB'},{name:'New'}]",c);
(async()=>{
 const options={allowEmpty:true,rethrow:true,silentSuccess:true};
 await c.saveOpeningBalances(options);
 assert.equal(writes.length,1);assert.equal(writes[0].range,'A2:F2');assert.equal(writes[0].values[0][2],200);assert.deepEqual(rows,original,'Existing transfer and other opening rows remain untouched');
 fields.setupAmt_0.value='';fields.setupAmt_1.value='25';writes=[];await c.saveOpeningBalances(options);
 assert.equal(writes[0].range,'A5:F5','New account balance appends after existing rows');
 rows=[...original,original[1]];fields.setupAmt_0.value='200';writes=[];
 await assert.rejects(c.saveOpeningBalances(options),/Multiple opening/);assert.equal(writes.length,0);
 console.log('Opening updates preserve transfers and unedited accounts, append new accounts, reject duplicates before writing.');
})().catch(e=>{console.error(e);process.exitCode=1;});
