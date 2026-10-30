const assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
class Clock extends Date {constructor(...args){super(...(args.length?args:['2026-10-30T04:00:00Z']));}static now(){return +new Date('2026-10-30T04:00:00Z');}}
const ctx=vm.createContext({console,Date:Clock,Intl,performance,setTimeout,clearTimeout,localStorage:{getItem:()=>null,setItem(){}},document:{addEventListener(){},getElementById:()=>null},CONFIG:{sheetName:'Transactions'}});ctx.window=ctx;ctx.addEventListener=()=>{};
for(const f of ['business.js','insurance-engine.js','goals.js','goals-compact.js'])vm.runInContext(fs.readFileSync(f,'utf8'),ctx);
const I=ctx.FinInsurance;
assert.equal(I.fundBy('2027-05-31'),'2027-04-30');assert.equal(I.fundBy('2027-06-03'),'2027-04-30');assert.equal(I.fundBy('2027-09-22'),'2027-07-31');assert.equal(I.fundBy('2028-03-01'),'2028-01-31');
const p={number:'TEST',description:'Example',premium:130,start:'2025-01-01',due:'2027-05-31',annual:true,increment:0,funding:'Cash'};
const state={version:1,revision:0,linkedName:'Insurance',history:[],policies:[{...p,id:'a'},{...p,id:'b',premium:4000,due:'2027-06-03'},{...p,id:'c',premium:166,due:'2027-09-22'},{...p,id:'medisave',funding:'CPF MediSave',premium:600},{...p,id:'missing',premium:''}]};
const plan=I.funding(state,1000,'2026-10-30');assert.equal(plan.length,2);assert.equal(plan[0].amount,4130);assert.equal(plan[0].assigned,1000);assert.equal(plan[1].assigned,0);assert.equal(plan[1].amount,166);
assert.equal(I.funding(state,4200,'2026-10-30')[1].assigned,70);
ctx.inputState=state;
vm.runInContext(`
goalsInsuranceState=inputState;
goalsData=[{name:'Insurance',insuranceManaged:true,target:4296,manualSaved:1000,monthlyAlloc:0,goalBuffer:0,startDate:'2026-10-30',endDate:'2027-07-31',urgency:'Critical',priority:0},{name:'Trip',target:2000,manualSaved:0,monthlyAlloc:0,goalBuffer:0,startDate:'2026-10-30',endDate:'2027-06-30',urgency:'Medium',priority:0}];
allTxForGoals=[];incomeBoosts=[];
computeDeployableBalance=()=>({deployable:1000,currentCashShortfall:0,futureSalaryGoalBuckets:[]});
getPlannedSavingsForMonth=()=>1000;getNextMonthPlannedSavings=()=>1000;
const originalRef=goalsData;
var model=buildGoalProjectionModel(18,36);
`,ctx);
let m=ctx.model;
assert.equal(m.goalState.length,2);assert.equal(vm.runInContext('goalsData===originalRef',ctx),true);
assert.equal(m.goalState[0].effectiveTarget,4296);
assert.equal(m.goalState[0].insuranceGap,0);
assert.equal(m.goalState[0].initialSaved,1000);
assert.ok(m.goalState[0].insuranceMilestones.some(x=>!x.current),'Annual renewals included');
assert.ok(m.poolBoost.every(x=>x===0),'No hidden double deductions');
for(let month=0;month<m.MONTHS;month++)assert.ok(m.allocationData.reduce((s,row)=>s+row[month],0)<=1000.001,'Shared pool cannot be overspent');
const current=m.goalState[0].insuranceMilestones.filter(x=>x.current);
assert.equal(current[0].fundBy,'2027-04-30');assert.equal(current[1].fundBy,'2027-07-31');
assert.ok(current.every(x=>x.gap<.01));
// Every renewal must be funded independently: a fully assigned first cycle does not cover next year.
vm.runInContext(`goalsData[0].manualSaved=4296;computeDeployableBalance=()=>({deployable:4296,currentCashShortfall:0,futureSalaryGoalBuckets:[]});getPlannedSavingsForMonth=()=>0;getNextMonthPlannedSavings=()=>0;model=buildGoalProjectionModel(18,36);`,ctx);
m=ctx.model;assert.ok(m.goalState[0].insuranceMilestones.filter(x=>x.current).every(x=>x.gap<.01));assert.ok(m.goalState[0].insuranceMilestones.some(x=>!x.current&&x.inRange&&x.gap>100));assert.equal(vm.runInContext('compactModelHasShortfall()',ctx),true);
// Cashflow earmarked for Insurance must reach the milestones, not disappear after expansion.
vm.runInContext(`goalsData[0].manualSaved=0;computeDeployableBalance=()=>({deployable:0,currentCashShortfall:0,futureSalaryGoalBuckets:[]});incomeBoosts=[{label:'Insurance grant',kind:'boost',amount:4296,frequency:'once',fromMonth:'2026-11',toGoal:'Insurance'}];model=buildGoalProjectionModel(18,36);`,ctx);
assert.equal(Math.round(ctx.model.allocationCashflowData[0].reduce((s,x)=>s+x,0)),4296);
// Partial cash cannot yield a false 'all on track', even if the final combined date is later.
vm.runInContext(`incomeBoosts=[];getPlannedSavingsForMonth=()=>100;getNextMonthPlannedSavings=()=>100;model=buildGoalProjectionModel(18,36);`,ctx);
assert.ok(ctx.model.goalState[0].insuranceMilestones.find(x=>x.current).gap>0);
assert.equal(vm.runInContext('compactModelHasShortfall()',ctx),true);
console.log('Insurance schedule: buffer dates, saved-money allocation, per-deadline shortfalls, competing goals, renewals, directed cashflow and pool conservation passed.');
const benchStart=performance.now();vm.runInContext('withGoalCalculationCache(()=>{for(let i=0;i<20;i++)buildGoalProjectionModel(18,60);})',ctx);console.log('20 shared schedule projections: '+Math.round(performance.now()-benchStart)+'ms');
