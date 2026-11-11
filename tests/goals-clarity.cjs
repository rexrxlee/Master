const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
class Clock extends Date{constructor(...a){super(...(a.length?a:['2026-11-01T04:00:00Z']));}static now(){return +new Date('2026-11-01T04:00:00Z');}}
const ctx=vm.createContext({console,Date:Clock,Intl,setTimeout,clearTimeout,localStorage:{getItem:()=>null,setItem(){}},document:{addEventListener(){},getElementById:()=>null,querySelectorAll:()=>[]},cancelAnimationFrame(){},CONFIG:{sheetName:'Transactions'}});ctx.window=ctx;ctx.addEventListener=()=>{};
for(const file of ['cash-rules.js','business.js','insurance-engine.js','goals.js','goals-compact.js','goals-clarity.js'])vm.runInContext(fs.readFileSync(file,'utf8'),ctx);
const run=s=>vm.runInContext(s,ctx);
run(`allAccounts=[{name:'Savings',type:'Savings'},{name:'Card',type:'Credit Card'}];goalSavingsAccts=['Savings'];savingsBalances={Savings:5000};ccOwed=1000;budgetSummary={billsRows:[{category:'Rent',allocated:500,fundingAccount:'Savings'}],monthlyRows:[{category:'Food',allocated:500}],billsTotal:500,monthlyTotal:500};allTxForGoals=[];incomeBoosts=[];historicalStats.avgMonthlyIncome=2000;
computePendingClaimSummary=()=>({total:200,count:1});
var cash=computeDeployableBalance();`);
assert.equal(ctx.cash.deployable,3000,'Pending claims are not cash');assert.equal(ctx.cash.claimReceivableForGoals,200);
run(`goalsData=[{name:'Example',target:4000,manualSaved:500,monthlyAlloc:0,startDate:'2026-11-01',endDate:'2027-05-31',goalBuffer:5}];var model=buildGoalProjectionModel(18,24);var rows=clarityGoalRows(model);`);
assert.equal(ctx.rows[0].target,4200);assert.equal(ctx.rows[0].projected,4200,'Forecast funds buffer as part of target');assert.equal(ctx.rows[0].now,500);assert.equal(ctx.rows[0].proposed,2500);assert.equal(ctx.rows[0].future,1200);assert.equal(ctx.rows[0].status,'On track · forecast','Proposed cash is not already assigned');
run(`goalsData[0].manualSaved=0;savingsBalances.Savings=500;historicalStats.avgMonthlyIncome=2000;model=buildGoalProjectionModel(18,24);`);
assert.equal(ctx.model.monthLedger[0].carryOut,1500);assert.equal(ctx.model.monthLedger[1].available,0);assert.equal(ctx.model.monthLedger[1].carryOut,500);assert.equal(ctx.model.monthLedger[2].available,500,'Shortfall persists until covered');
run(`savingsBalances.Savings=5000;incomeBoosts=[{label:'Reduction',kind:'reduce',amount:5000,frequency:'once',fromMonth:'2026-12',toGoal:'any'}];model=buildGoalProjectionModel(18,24);`);
assert.equal(ctx.model.monthLedger[1].carryOut,4000);assert.equal(ctx.model.monthLedger[2].available,0);assert.equal(ctx.model.monthLedger[2].carryOut,3000);
run(`incomeBoosts=[];goalsInsuranceState={version:1,linkedName:'Insurance',policies:[{id:'a',description:'Example premium',number:'TEST',funding:'Cash',premium:900,start:'2025-06-01',due:'2027-06-01',annual:true,increment:0}],history:[]};goalsData=[{name:'Insurance',insuranceManaged:true,target:900,manualSaved:200,goalBuffer:0}];model=buildGoalProjectionModel(18,36);rows=clarityGoalRows(model);`);
assert.equal(ctx.rows[0].projected,900,'Next cycle never includes next-year funding');assert.equal(ctx.rows[0].now+ctx.rows[0].proposed+ctx.rows[0].future,ctx.rows[0].projected);
for(const ledger of ctx.model.monthLedger)assert.ok(Math.abs(ledger.available-ledger.allocated-ledger.unallocated)<1e-7);
for(const a of [[1/3,1/3,1/3],[100.004,100.004,100.004],[0,0,0]]){const parts=ctx.clarityRoundedParts(a);assert.equal(Math.round(parts.reduce((s,n)=>s+n,0)*100),Math.round(a.reduce((s,n)=>s+n,0)*100));}
const html=ctx.clarityCashHTML(ctx.cash);assert.ok(html.includes('pending claims included in total goal funds'));assert.ok(!html.includes('Claims coming back'));
assert.ok(ctx.clarityGoalCard({name:'<script>',target:900,manualSaved:0},0).includes('&lt;script&gt;'));
console.log('Goals clarity: cash-only reconciliation, buffer funding, proposed vs assigned, deficit carry, insurance cycle isolation, pool conservation and displayed-cent totals passed.');
// Exercise the actual page renderer with its generated IDs and event handlers.
const elements=new Map();function element(id){if(!elements.has(id))elements.set(id,{id,value:'',textContent:'',hidden:false,style:{},classList:{toggle(){}},setAttribute(){},querySelector:()=>null});return elements.get(id);}
const container=element('goalsContainer');Object.defineProperty(container,'innerHTML',{set(html){this.html=html;for(const m of html.matchAll(/\bid="([^"]+)"/g))element(m[1]);},get(){return this.html||'';}});
ctx.document.getElementById=id=>elements.get(id)||null;
run(`renderIncomeBoostsPanel=()=>{};renderClarityPage();`);
assert.ok(container.html.includes('Monthly plan'));assert.ok(element('compactGoalSummary').innerHTML.includes('Cash for goals today'));assert.ok(element('gcBreakdown_0').innerHTML.includes('Premiums'));
run(`compactSmartAssign();`);assert.ok(run('goalsData.reduce((s,g)=>s+g.manualSaved,0)<=computeDeployableBalance().deployable+.001'));
assert.ok(run('compactPlanChanged()'),'Suggestion is a draft');run('resetCompactPlan()');assert.equal(run('compactPlanChanged()'),false);
console.log('Goals UI renderer, forecast, suggested assignment and reset preview passed.');
run(`allAccounts=[{name:'Savings',type:'Savings'},{name:'CDA',type:'Savings'},{name:'Unused card',type:'Credit Card'}];goalSavingsAccts=['Savings'];allTxForGoals=[];budgetSummary={billsRows:[{category:'Selected',allocated:100,fundingAccount:'Savings'},{category:'School',allocated:200,fundingAccount:'CDA'},{category:'Card bill',allocated:50,fundingAccount:'Unused card'},{category:'Unknown',allocated:25}],monthlyRows:[{category:'Food',allocated:300}]};var reserve=getFutureMonthBudgetBreakdown(new Date(2026,11,1));`);
assert.equal(ctx.reserve.total,475,'Future reserve includes configured cards without transactions and unknown funding, excludes CDA');
assert.equal(ctx.reserve.bills.length,3);assert.equal(ctx.reserve.monthly.length,1);
console.log('Future budget reserve uses shared account rules and exposes included rows.');
