const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
class Clock extends Date{constructor(...a){super(...(a.length?a:['2026-11-01T04:00:00Z']));}}
const c=vm.createContext({Date:Clock,console,setTimeout,clearTimeout,Intl,localStorage:{getItem:()=>null},document:{addEventListener(){},getElementById:()=>null},CONFIG:{sheetName:'Transactions'}});c.window=c;c.addEventListener=()=>{};
for(const f of ['business.js','insurance-engine.js','goals.js','goals-compact.js','goals-clarity.js','goals-insights.js'])vm.runInContext(fs.readFileSync(f,'utf8'),c);
const run=s=>vm.runInContext(s,c);
run(`goalsData=[{name:'Test',target:2000,manualSaved:1000,monthlyAlloc:0,startDate:'2026-11-01',endDate:'2027-01-31',goalBuffer:0}];incomeBoosts=[];allTxForGoals=[];historicalStats.avgMonthlyIncome=500;computeFutureMonthBudgetReserve=()=>0;computeDeployableBalance=()=>({deployable:1000,rawSavings:1000,currentCashShortfall:0,futureSalaryGoalBuckets:[]});var original=computeDeployableBalance;`);
assert.equal(c.goalSpendingRoom(0),0,'Dropping assignment cannot leave spent money reusable');
run('historicalStats.avgMonthlyIncome=750;');assert.equal(c.goalSpendingRoom(0),500);assert.equal(c.goalSpendingRoom(),500);assert.equal(run('goalsData[0].manualSaved'),1000);assert.equal(run('computeDeployableBalance===original'),true);
run(`goalsData.push({name:'Other',target:500,manualSaved:0,monthlyAlloc:0,startDate:'2026-11-01',endDate:'2027-01-31',goalBuffer:0});`);assert.equal(c.goalSpendingRoom(0),0,'Other goals reduce spending room');
run(`goalsData[1].endDate='';`);assert.equal(c.goalSpendingRoom(),0,'Missing deadline is not declared safe');
run(`goalsData.pop();goalsData[0].endDate='2026-11-20';`);assert.equal(c.goalSpendingRoom(),0,'Imminent target must have cash now');
run(`goalsData[0].endDate='2027-01-31';var model=buildGoalProjectionModel(18,60);`);assert.ok(c.goalForecastChart(c.model).includes('<svg'));assert.ok(c.goalForecastChart(c.model).includes('Test'));
console.log('Spending simulations: real cash removal, all-goal constraints, near-term protection, missing deadlines and state restoration passed.');
