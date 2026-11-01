// One presentation model: cash today, progress already recorded, and future funding
// are separate. All forecast numbers come from the same allocator.
let goalView = 'overview';
let goalClarityModel = null;
function switchGoalView(view) {
  if (!['overview','monthly','settings'].includes(view)) return;
  goalView = view;
  document.querySelectorAll('[data-goal-view]').forEach(el => { el.hidden = el.dataset.goalView !== view; });
  document.querySelectorAll('[data-goal-tab]').forEach(el => { el.setAttribute('aria-selected',String(el.dataset.goalTab === view)); });
}
function clarityCash(dep) {
  const assigned = compactMoney(goalsData.reduce((s,g)=>s+Math.max(0,+g.manualSaved||0),0));
  return {assigned,free:compactMoney(dep.deployable-assigned),reserved:compactMoney(ccOwed+dep.remainingBudget+dep.futureSalaryHold+(dep.futureSalaryGoalBucketTotal||0))};
}
function clarityGoalRows(model) {
  return model.goalState.map((state,idx)=>{
    const goal=goalsData[idx], insurance=!!state.insuranceMilestones;
    const month=Math.min(model.MONTHS-1,Math.max(0,state.deadlineMo??17));
    const sum=a=>compactMoney((a||[]).slice(0,month+1).reduce((s,n)=>s+(n||0),0));
    const base=insurance?state.insuranceCurrentBase:model.allocationBaseData[idx];
    const boosts=insurance?state.insuranceCurrentAdjustments:model.allocationCashflowData[idx];
    const assigned=compactMoney(Math.max(0,+goal.manualSaved||0));
    const recorded=insurance?0:compactRecordedForGoal(idx);
    const now=compactMoney(Math.min(state.effectiveTarget,assigned+recorded));
    const proposed=compactMoney(base?.[0]||0);
    const future=compactMoney(sum(base)-proposed+sum(boosts));
    const projected=compactMoney(now+proposed+future);
    const target=compactMoney(state.effectiveTarget);
    const gap=compactMoney(insurance?state.insuranceGap:Math.max(0,target-projected));
    const unknown=insurance?state.insuranceUnverified:state.deadlineMo===null||state.deadlineMo>=model.MONTHS;
    const status=unknown?'Check deadline':gap>=.01?'Shortfall':now>=target?'Covered now':'On track · forecast';
    return {idx,name:state.name,insurance,state,assigned,recorded,now,proposed,future,projected,target,gap,unknown,status,month};
  });
}
function clarityCashHTML(dep) {
  const cash=clarityCash(dep);
  const line=(label,value,cls='')=>`<div class="gc-ledger-row ${cls}"><span>${label}</span><b>${formatCurrency(value)}</b></div>`;
  return `<section class="gc-cash" aria-label="Cash reconciliation"><div class="gc-kpis">
    <div class="gc-kpi primary"><span>Cash for goals today</span><strong>${formatCurrency(dep.deployable)}</strong><small>After bills, budgets and card debt</small></div>
    <div class="gc-kpi"><span>Assigned to goals</span><strong>${formatCurrency(cash.assigned)}</strong><small>Part of the cash on the left</small></div>
    <div class="gc-kpi ${cash.free<0?'warning':''}"><span>${cash.free<0?'Over-assigned':'Not yet assigned'}</span><strong>${formatCurrency(Math.abs(cash.free))}</strong><small>${cash.free<0?'Reduce assignments before saving':'Available to assign in your plan'}</small></div>
    </div><details class="gc-reconcile"><summary>Where these numbers come from <span>Bank balances → available cash</span></summary><div class="gc-reconcile-grid"><div>
    ${line('Selected account balances',dep.rawSavings)}${line('− Credit cards owed',ccOwed)}${line('− Bills still to pay',dep.unpaidGoalBills)}${line('− Monthly expenses still reserved',dep.remainingMonthlyExpenseBudget)}${line('− Future salary budgets protected',dep.futureSalaryHold)}${line('− Future salary savings protected',dep.futureSalaryGoalBucketTotal||0)}${line('= Cash after protections',dep.rawSavings-cash.reserved,'total')}
    ${dep.currentCashShortfall>.005?`<p class="gc-alert">Cash shortfall: ${formatCurrency(dep.currentCashShortfall)}. Future income must cover this before goals.</p>`:''}</div><div>
    ${line('Cash for goals today',dep.deployable)}${line('− Assigned to goals',cash.assigned)}${line('= Not yet assigned',cash.free,'total')}
    <p>Assignments earmark money already in your accounts. They do not move money.</p><p><b>${formatCurrency(dep.claimReceivableForGoals)} pending claims excluded.</b> Reimbursements become available only after they are recorded as received.</p>
    <p>Early salary stays protected for its recorded month. Spending already reflected in account balances is not deducted twice.</p></div></div></details></section>`;
}
function clarityGoalCard(goal,idx) {
  const recorded=compactRecordedForGoal(idx),assigned=compactTotalAssignedForGoal(idx);
  return `<article class="gc-goal-card" id="compactGoal_${idx}"><div class="gc-goal-top"><div><span class="gc-eyebrow">${escapeHtml(goal.group||'Wants')} · ${escapeHtml(goal.urgency||'Medium')}</span><h3>${escapeHtml(goal.name)}</h3></div><span class="gc-pill" id="cgPill_${idx}">Checking</span></div>
    <div id="gcNumbers_${idx}" class="gc-goal-numbers"></div><div class="gc-progress"><span id="cgBar_${idx}"></span></div><p class="gc-result" id="cgStatus_${idx}"></p>
    <details class="gc-assign"><summary>Adjust assignment <b id="cgAssignedMetric_${idx}">${formatCurrency(assigned)}</b></summary><div class="gc-assign-input"><label>Assigned cash ($)<input id="gcCash_${idx}" type="number" min="0" step="0.01" value="${(+goal.manualSaved||0).toFixed(2)}" oninput="assignClarityCash(${idx},this.value)"></label><input hidden id="cgSlider_${idx}" value="${assigned}"><input hidden id="cgAmount_${idx}" value="${assigned}"></div>${recorded>0?`<p>${formatCurrency(recorded)} of earlier recorded progress is shown separately from assigned cash.</p>`:''}<p>Preview only until you save. This does not transfer money.</p></details>
    <details class="gc-details"><summary>${goal.insuranceManaged?'Premium checkpoints & forecast':'Forecast calculation'}</summary><div id="gcBreakdown_${idx}"></div></details></article>`;
}
function assignClarityCash(idx,value) {
  if(value==='')return;
  const amount=Number(value);if(!Number.isFinite(amount))return;
  assignCompactGoal(idx,Math.max(0,amount)+compactRecordedForGoal(idx));
}
function renderClarityPage() {
  compactNormaliseManualAssignments();if(!compactSavedPlan)captureCompactSavedPlan();
  const today=toDateInputValue(new Date());
  const container=document.getElementById('goalsContainer');
  container.innerHTML=`<nav class="gc-tabs" aria-label="Goal views"><button data-goal-tab="overview" onclick="switchGoalView('overview')">Overview</button><button data-goal-tab="monthly" onclick="switchGoalView('monthly')">Monthly plan</button><button data-goal-tab="settings" onclick="switchGoalView('settings')">Settings</button></nav>
  <div data-goal-view="overview"><div id="compactGoalSummary"></div><div id="gcOverall" aria-live="polite"></div><div class="gc-section-head"><div><h2>Your goals</h2><p>Cash assigned today and projected funding are shown separately.</p></div><button class="btn-primary" onclick="compactSmartAssign()">Suggest assignments</button></div><div id="compactAssignTotal" class="gc-note"></div><div class="gc-goal-grid" id="compactGoalRows">${goalsData.map(clarityGoalCard).join('')||'<p>No goals yet. Add one in Settings.</p>'}</div></div>
  <div data-goal-view="monthly" hidden><section class="gc-panel"><div class="gc-section-head"><div><h2>Where each month’s money goes</h2><p>Today is a proposed allocation of existing cash. Later months are forecasts.</p></div><label>Show<select id="gcHorizon" onchange="renderClarityMonthly(goalClarityModel)"><option value="12">12 months</option><option value="24">24 months</option><option value="60">Full forecast</option></select></label></div><div id="compactGoalTimeline"></div></section></div>
  <div data-goal-view="settings" hidden><section class="gc-panel"><h2>Funding accounts</h2><p>Only selected savings accounts fund this plan.</p><div class="cg-accounts">${allAccounts.filter(a=>a.type==='Savings').map(a=>`<label><input type="checkbox" value="${escapeHtml(a.name)}" ${goalSavingsAccts.includes(a.name)?'checked':''} onchange="toggleGoalAccount(this)">${escapeHtml(a.name)} <strong>${formatCurrency(savingsBalances[a.name]||0)}</strong></label>`).join('')}</div></section>
  <section class="gc-panel"><h2>Forecast adjustments</h2><p>Expected bonuses and changes to future saving capacity. Pending claims are not automatically included.</p><div id="compactAdjustments"></div></section>
  <details class="gc-panel"><summary>Add a goal</summary><form class="cg-add-form" onsubmit="event.preventDefault();addCompactGoal(this)"><label>Goal name<input name="name" required maxlength="100"></label><label>Target ($)<input name="target" type="number" min="0.01" step="0.01" required></label><label>Start<input name="start" type="date" value="${today}" required></label><label>Deadline<input name="end" type="date" required></label><label>Priority<select name="urgency">${compactPriorityOptions('Medium')}</select></label><label>Group<select name="group">${compactGroupOptions('Wants')}</select></label><button type="submit" class="btn-primary">Add goal</button><p class="cg-form-message" role="status"></p></form></details>
  <section class="gc-panel"><h2>Goal settings</h2>${goalsData.map((g,i)=>`<details class="gc-setting"><summary>${escapeHtml(g.name)}</summary>${compactGoalSetupRow(g,i)}</details>`).join('')}</section></div>
  <footer class="gc-save"><span id="goalsAutosaveStatus" role="status"></span><div><button id="compactResetPlan" class="btn-secondary" onclick="resetCompactPlan()">Reset preview</button><button id="compactSavePlan" class="btn-primary" onclick="saveCompactPlan()">Save plan</button></div></footer>`;
  renderIncomeBoostsPanel(document.getElementById('compactAdjustments'));
  switchGoalView(goalView);refreshCompactGoalForecast();updateCompactSaveStatus();
}
function renderClarityForecast() {
  const dep=computeDeployableBalance(),cash=clarityCash(dep),model=buildGoalProjectionModel(18,60);goalClarityModel=model;
  const rows=clarityGoalRows(model);
  document.getElementById('compactGoalSummary').innerHTML=clarityCashHTML(dep);
  const short=rows.filter(r=>r.gap>=.01),unknown=rows.filter(r=>r.unknown);
  document.getElementById('gcOverall').innerHTML=`<div class="gc-status ${cash.free<-.005||short.length?'attention':''}"><b>${cash.free<-.005?'Your assignments exceed available cash':short.length?`${short.length} goal${short.length===1?' needs':'s need'} attention`:unknown.length?'Some deadlines need checking':'Your goals are forecast on track'}</b><span>${cash.free<-.005?`Reduce assigned cash by ${formatCurrency(-cash.free)}.`:short.length?short.map(r=>`${escapeHtml(r.name)}: ${formatCurrency(r.gap)} projected shortfall`).join(' · '):'Based on your current income estimate, budgets and adjustments. Future money is not guaranteed.'}</span></div>`;
  document.getElementById('compactAssignTotal').textContent=`${formatCurrency(Math.max(0,cash.free))} unassigned today. Suggested allocations are a preview until you save.`;
  rows.forEach(r=>{
    const goal=goalsData[r.idx];
    const pill=document.getElementById(`cgPill_${r.idx}`);pill.textContent=cash.free<-.005?'Over-assigned':r.status;pill.classList.toggle('warning',cash.free<-.005||r.gap>=.01||r.unknown);
    const due=r.insurance?'Separate premium deadlines':goal.endDate?`By ${formatDateDisplay(goal.endDate)}`:'No deadline set';
    document.getElementById(`gcNumbers_${r.idx}`).innerHTML=`<div><small>Assigned cash</small><b>${formatCurrency(r.assigned)}</b></div><div><small>Target${goal.goalBuffer?' incl. buffer':''}</small><b>${formatCurrency(r.target)}</b></div><div><small>Still to cover today</small><b>${formatCurrency(Math.max(0,r.target-r.now))}</b></div>`;
    document.getElementById(`cgBar_${r.idx}`).style.width=Math.min(100,100*r.now/Math.max(1,r.target))+'%';
    document.getElementById(`cgStatus_${r.idx}`).textContent=`${due} · ${r.insurance?'Next premium cycle': 'Projected by deadline'}: ${formatCurrency(r.projected)}${r.recorded?` · ${formatCurrency(r.recorded)} earlier recorded progress`:''}`;
    const line=(label,n)=>`<div class="gc-ledger-row"><span>${label}</span><b>${formatCurrency(n)}</b></div>`;
    let detail=`${line('Assigned cash',r.assigned)}${r.recorded?line('+ Earlier recorded progress',r.recorded):''}${line('+ Proposed from unassigned cash today',r.proposed)}${line('+ Future income & adjustments',r.future)}${line('= Projected funding',r.projected)}${line('Target',r.target)}<p>Progress bar shows only assigned cash and recorded progress. Forecast amounts are estimates.</p>`;
    if(r.insurance){detail+=`<div class="gc-table-wrap"><table><thead><tr><th>Ready by</th><th>Premiums</th><th>Projected</th><th>Status</th></tr></thead><tbody>${r.state.insuranceMilestones.filter(m=>m.current||m.inRange&&m.gap>=.01).map(m=>`<tr><td>${escapeHtml(m.fundBy)}${m.current?'':' · renewal'}</td><td>${formatCurrency(m.amount)}</td><td>${formatCurrency(m.projected)}</td><td>${!m.inRange?'Beyond range':m.gap>=.01?formatCurrency(m.gap)+' short':'On track'}</td></tr>`).join('')}</tbody></table></div><p>At least 30 days before payment. The target and calculation above cover only the next premiums. Later renewals are funded separately in the Monthly plan.</p>`;}
    document.getElementById(`gcBreakdown_${r.idx}`).innerHTML=detail;
    document.getElementById(`cgAssignedMetric_${r.idx}`).textContent=formatCurrency(r.assigned);
    const input=document.getElementById(`gcCash_${r.idx}`);if(document.activeElement!==input)input.value=r.assigned.toFixed(2);
  });
  renderClarityMonthly(model);updateCompactSaveStatus();
}
// Round an entire displayed column together so its visible parts equal its total.
function clarityRoundedParts(values) {
  const cents=values.map(n=>Math.floor(Math.max(0,n)*100+1e-7));
  let left=Math.round(values.reduce((s,n)=>s+Math.max(0,n),0)*100)-cents.reduce((s,n)=>s+n,0);
  const order=values.map((n,i)=>({i,f:n*100-cents[i]})).sort((a,b)=>b.f-a.f);
  for(let i=0;i<left;i++)cents[order[i%order.length].i]++;
  return cents.map(n=>n/100);
}
function renderClarityMonthly(model) {
  if(!model)return;
  const months=Math.min(model.MONTHS,Number(document.getElementById('gcHorizon')?.value||12));
  const dep=computeDeployableBalance(),cash=clarityCash(dep);
  let html='<div class="gc-months">';
  for(let m=0;m<months;m++){
    const ledger=model.monthLedger[m];
    const parts=clarityRoundedParts([...model.allocationData.map(a=>a[m]||0),model.unallocatedData[m]||0]);
    const allocated=compactMoney(parts.slice(0,-1).reduce((s,n)=>s+n,0));
    const opening=m===0?cash.assigned:0;
    const available=compactMoney(ledger.available+opening);
    const summary=m===0?'Existing cash · proposed split':'Income forecast · after budget';
    html+=`<details class="gc-month" ${m===0?'open':''}><summary><span><b>${escapeHtml(model.fullLabels[m])}</b><small>${summary}</small></span><strong>${formatCurrency(available)}</strong></summary><div class="gc-month-body"><div class="gc-month-source"><h4>Money available</h4>`;
    if(m===0)html+=`<p>Cash for goals today: <b>${formatCurrency(dep.deployable)}</b></p><p>Already assigned: ${formatCurrency(opening)}<br>Unassigned: ${formatCurrency(Math.max(0,cash.free))}</p>`;
    else html+=`<div class="gc-ledger-row"><span>Income estimate${getRecordedFutureSalaryForMonth(new Date(model.today.getFullYear(),model.today.getMonth()+m,1))?' (received early)':''}</span><b>${formatCurrency(ledger.salary)}</b></div><div class="gc-ledger-row"><span>Budget reserve</span><b>− ${formatCurrency(ledger.budget)}</b></div>`;
    if(ledger.adjustment)html+=`<div class="gc-ledger-row"><span>Cashflow adjustments</span><b>${formatCurrency(ledger.adjustment)}</b></div>`;
    if(ledger.carryIn)html+=`<div class="gc-ledger-row"><span>Earlier cash shortfall</span><b>− ${formatCurrency(ledger.carryIn)}</b></div>`;
    if(m>0&&ledger.salary-ledger.budget>ledger.base+.005)html+=`<p>${formatCurrency(ledger.salary-ledger.budget-ledger.base)} of the early salary surplus is already used or protected.</p>`;
    html+=`<div class="gc-ledger-row total"><span>Available for this split</span><b>${formatCurrency(available)}</b></div>${ledger.carryOut>.005?`<p class="gc-alert">${formatCurrency(ledger.carryOut)} shortfall carries forward.</p>`:''}</div><div class="gc-month-dest"><h4>${m===0?'Proposed assignments':'Planned contributions'}</h4>`;
    model.goalState.forEach((g,i)=>{const n=compactMoney(parts[i]+(m===0?Number(goalsData[i].manualSaved||0):0));if(n>.005)html+=`<div class="gc-ledger-row"><span>${escapeHtml(g.name)}${g.insuranceMilestones?'<small>Includes renewals as their saving windows open</small>':''}</span><b>${formatCurrency(n)}</b></div>`;});
    html+=`<div class="gc-ledger-row"><span>Unallocated</span><b>${formatCurrency(parts.at(-1))}</b></div><div class="gc-ledger-row total"><span>Total</span><b>${formatCurrency(compactMoney(allocated+opening+parts.at(-1)))}</b></div></div></div></details>`;
  }
  document.getElementById('compactGoalTimeline').innerHTML=html+'</div><p class="gc-note">The same forecast powers the goal cards and this monthly split. Amounts are rounded together so displayed totals reconcile. Insurance contributions include later renewals; they do not increase the next-premium target.</p>';
}
renderCompactGoalsPage=renderClarityPage;
renderGoalsPage=renderClarityPage;
renderCompactGoalForecast=renderClarityForecast;
updateCompactAllocationSummary=function(){document.getElementById('compactGoalSummary').innerHTML=clarityCashHTML(computeDeployableBalance());updateCompactSaveStatus();};
// Suggestions use the same deadline/priority allocator as the forecast, not a
// second weighted algorithm. Only existing cash is assigned, never future boosts.
compactSmartAssign=function(){
  goalsData.forEach(g=>{g.manualSaved=0;});
  const model=withGoalCalculationCache(()=>buildGoalProjectionModel(18,60));
  const parts=clarityRoundedParts(model.allocationBaseData.map(a=>a[0]||0));
  goalsData.forEach((g,i)=>{g.manualSaved=parts[i];});
  renderClarityPage();updateCompactSaveStatus();
};
