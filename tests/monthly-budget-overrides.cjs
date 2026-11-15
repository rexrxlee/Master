const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
class Clock extends Date{constructor(...a){super(...(a.length?a:['2026-11-15T04:00:00Z']));}}
const c=vm.createContext({Date:Clock,console,clearTimeout,setTimeout,localStorage:{getItem:()=>null},document:{addEventListener(){},getElementById:()=>null},window:{addEventListener(){}},alert(){}});
for(const f of ['cash-rules.js','budget.js'])vm.runInContext(fs.readFileSync(f,'utf8'),c);
const run=s=>vm.runInContext(s,c);
run(`monthlyBudget=[{type:'Monthly Expenses',category:'Others',allocated:200}];billsBudget=[];renderBudget=()=>{};updateBudgetMonthOverride('Monthly Expenses',0,'1000');var metadata=buildBudgetFundingMap();`);
assert.equal(run('monthlyBudget[0].allocated'),200,'Regular budget is never overwritten');assert.equal(run('CashRules.budgetAmount(monthlyBudget[0])'),1000);
assert.equal(run('CashRules.budgetAmount(monthlyBudget[0],new Date(2026,11,1))'),200,'Override automatically stops next month');
assert.equal(run('metadata["monthly expenses|others"].monthOverrides["2026-11"]'),1000,'Monthly overrides persist with metadata');
run(`monthlyBudget=[{type:'Monthly Expenses',category:'Others',allocated:200}];applyBudgetFundingMap({AE2:{v:JSON.stringify(metadata)}});`);
assert.equal(run('CashRules.budgetAmount(monthlyBudget[0])'),1000,'Override survives reload');
run(`updateBudgetMonthOverride('Monthly Expenses',0,'0');`);assert.equal(run('CashRules.budgetAmount(monthlyBudget[0])'),0,'Zero is a valid override');
run(`updateBudgetMonthOverride('Monthly Expenses',0,'');`);assert.equal(run('CashRules.budgetAmount(monthlyBudget[0])'),200);
assert.equal(run('CashRules.budgetAmount({allocated:200,monthOverrides:{"2026-12":1000}},new Date(2027,0,1))'),200,'Year rollover expires override');
console.log('Month overrides preserve regular amounts, survive reload, expire by calendar month, and support zero/clear.');
