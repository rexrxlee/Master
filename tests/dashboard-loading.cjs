// Run with node --experimental-vm-modules tests/dashboard-loading.cjs.
// Exercise the real dashboard import path in a fresh browser-like global scope.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
async function run(insurance) {
  const alerts = [], messages = [], modules = new Map();
  const sheet = { AD2: {v:'Savings'}, ...(insurance ? {AL2:{v:JSON.stringify(insurance)}} : {}) };
  const transactions = {};
  const workbook = {Sheets:{'Budget Setup':sheet,Transactions:transactions}};
  const ctx = vm.createContext({console:{error(){}},setTimeout,clearTimeout,
    localStorage:{getItem:()=>null,setItem(){}},document:{readyState:'loading',addEventListener(){},getElementById:()=>null},
    CONFIG:{sheetName:'Transactions'},downloadExcelFile:async()=>workbook,
    alert:message=>alerts.push(message),log:message=>messages.push(message),clearOutput(){},
    XLSX:{read:x=>x,utils:{sheet_to_json:(s,options)=>{
      if(s===transactions)return [['Date','Transaction','Amount','Main Category','Sub Category','Account']];
      if(options.range==='S2:AC20')return [['Emergency Fund',30000,0,0,'2026-01-01','2029-01-01','Critical'],...(insurance?[['Insurance',130,500,0,'2026-01-01','2027-05-31','Critical']]:[])];
      if(options.range)return [];
      return [[],[null,null,null,null,null,null,null,null,null,'Savings','Savings']];
    }}}
  });
  ctx.window=ctx;ctx.addEventListener=()=>{};
  async function importModule(specifier) {
    const file=specifier.replace(/^\.\//,'').split('?')[0];
    assert.ok(!file.includes('/'), 'Only app modules');
    if(!modules.has(file))modules.set(file,(async()=>{
      const module=new vm.SourceTextModule(fs.readFileSync(path.join(root,file),'utf8'),{context:ctx,identifier:file,importModuleDynamically:importModule});
      await module.link(importModule);await module.evaluate();return module;
    })());
    return modules.get(file);
  }
  vm.runInContext(fs.readFileSync(path.join(root,'business.js'),'utf8'),ctx);
  new vm.Script(fs.readFileSync(path.join(root,'dashboard.js'),'utf8'),{importModuleDynamically:importModule}).runInContext(ctx);
  // Stub UI rendering only; workbook parsing and financial calculation are real.
  vm.runInContext(`setupDashboardTabs=()=>{};loadAssetAccountSelections=()=>{};setupFilters=()=>{};
    getCurrentFilters=()=>({});refreshAllFilters=()=>{};updateDashboard=()=>{};updateActiveFilterBadges=()=>{};`,ctx);
  await ctx.loadDashboard();
  assert.deepEqual(alerts,[], 'Fresh Dashboard must load without alerts');
  assert.equal(messages.at(-1),'Dashboard loaded.');
  assert.equal(vm.runInContext('financePlanEmergencyTarget',ctx),30000);
  assert.ok(Number.isFinite(vm.runInContext('financePlanGoalCash',ctx)));
  assert.ok(ctx.FinTaxStore&&ctx.FinPropertyTaxStore, 'Tax dependencies match Goals');
  if(insurance)assert.equal(ctx.FinInsurance.plan(insurance).amount,4130);
  await ctx.loadDashboard(true);
  assert.deepEqual(alerts,[], 'Reload must reuse modules without scope collisions');
  return modules.size;
}
(async()=>{
  await run(null);
  const policy={number:'TEST',description:'Test premium',premium:130,increment:0,start:'2024-05-31',due:'2027-05-31',annual:true,funding:'Cash'};
  await run({version:1,revision:1,linkedName:'Insurance',history:[],policies:[{...policy,id:'a'},{...policy,id:'b',number:'TEST2',premium:4000,due:'2027-06-10'}]});
  console.log('Dashboard cold load and reload passed with empty and linked insurance records.');
})().catch(e=>{console.error(e);process.exitCode=1;});
