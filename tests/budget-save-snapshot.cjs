const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
let release,updates;const gate=new Promise(resolve=>release=resolve);const status={};
const c=vm.createContext({console,clearTimeout,setTimeout,document:{addEventListener(){},getElementById:()=>status},window:{addEventListener(){}},log(){},alert(){},writeExcelRanges:async rows=>{updates=JSON.parse(JSON.stringify(rows));await gate;}});
vm.runInContext(fs.readFileSync('budget.js','utf8'),c);
vm.runInContext("billsBudget=[{category:'Bill',allocated:100}];monthlyBudget=[{category:'Food',allocated:200}];buildSaveValues=rows=>rows.map(r=>[r.category,r.allocated]);buildBudgetFundingMap=()=>({});",c);
(async()=>{
 const saving=c.saveBudgetSetupToExcel({silent:true});
 vm.runInContext('monthlyBudget[0].allocated=300;markBudgetSetupDirty();',c);release();await saving;
 assert.equal(updates[1].values[0][1],200,'Save uses one captured snapshot');assert.equal(vm.runInContext('budgetSetupDirty',c),true);assert.match(status.textContent,/Newer changes/);
 await c.saveBudgetSetupToExcel({silent:true});assert.equal(updates[1].values[0][1],300);assert.equal(vm.runInContext('budgetSetupDirty',c),false);assert.equal(status.textContent,'Saved to Excel');
 console.log('Budget save snapshots preserve later edits as unsaved until the next confirmed save.');
})().catch(e=>{console.error(e);process.exitCode=1;});
