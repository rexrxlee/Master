const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const elements=Object.fromEntries(['inDate','inTransaction','inSubCategory','inAmount','inAccount','incomeSuccess'].map(id=>[id,{value:'',style:{}}]));
let written=null,alerts=[];
const c=vm.createContext({console,document:{addEventListener(){},getElementById:id=>elements[id]},window:{addEventListener(){}},setTimeout:()=>0,alert:s=>alerts.push(s),log(){},CONFIG:{sheetName:'Transaction'},writeExcelRange:async(sheet,range,rows)=>{written=rows;}});
vm.runInContext(fs.readFileSync('add-transaction.js','utf8'),c);
vm.runInContext(`accounts=[{name:'Savings',type:'Savings'},{name:'UOB ONE',type:'Credit Card'}];currentMode='cashback';formatDateForExcel=s=>s;addSavedRowsToPage=()=>{};`,c);
Object.assign(elements.inDate,{value:'2026-11-13'});elements.inAmount.value='50';elements.inAccount.value='UOB ONE';elements.inSubCategory.value='Salary';
(async()=>{
 await c.saveIncome();assert.equal(alerts.length,0);assert.equal(written[0][3],'Income');assert.equal(written[0][4],'Cashback');assert.equal(written[0][5],'UOB ONE');assert.equal(written[0][2],50);
 const rows=[{date:'2026-11-01',transaction:'Opening Balance',amount:100,account:'UOB ONE',mainCat:'Income',subCat:'Opening Balance'},{date:'2026-11-13',transaction:'Cashback',amount:50,account:'UOB ONE',mainCat:'Income',subCat:'Cashback'}];
 const result=c.computeTransactionAccountBalances([{name:'UOB ONE',type:'Credit Card'},{name:'Savings',type:'Savings'}],rows);
 assert.equal(result[0].balance,50,'Cashback reduces owed amount');assert.equal(result[1].balance,0,'Savings unchanged');
 elements.inAmount.value='0';written=null;await c.saveIncome();assert.equal(written,null);
 elements.inAmount.value='50';elements.inAccount.value='Savings';await c.saveIncome();assert.equal(written,null,'Cashback mode requires card');
 console.log('Cashback writes positive card credit, reduces card debt without adding savings, rejects zero and non-card destination.');
})().catch(e=>{console.error(e);process.exitCode=1;});
