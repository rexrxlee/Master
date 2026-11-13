// Spending must reduce both the bank cash pool AND the source assignments.
// Reassigning money alone is not a spending simulation.
function goalSpendScenario(amount, sourceIndex=null) {
  const originalCompute=computeDeployableBalance, dep=originalCompute();
  const saved=goalsData.map(g=>Number(g.manualSaved||0));
  const limit=sourceIndex===null?dep.deployable:Math.min(dep.deployable,goalAssignmentFunding(dep)[sourceIndex]?.cash||0);
  if(!Number.isFinite(amount)||amount>limit+.001||(amount<0&&sourceIndex!==null))return null;
  let remaining=amount;
  if(sourceIndex!==null){goalsData[sourceIndex].manualSaved=Math.max(0,saved[sourceIndex]-remaining);remaining=0;}
  else {
    remaining=Math.max(0,remaining-Math.max(0,dep.deployable-saved.reduce((s,n)=>s+n,0)));
    for(const i of getCompactFundingOrder().slice().reverse()){
      const used=Math.min(remaining,saved[i]);goalsData[i].manualSaved=saved[i]-used;remaining-=used;
    }
  }
  try {
    computeDeployableBalance=()=>({...dep,rawSavings:dep.rawSavings-amount,deployable:Math.max(0,dep.deployable-amount)});
    const model=buildGoalProjectionModel(18,60);
    const assigned=goalsData.reduce((s,g)=>s+Number(g.manualSaved||0),0);
    const nearNeed=goalsData.reduce((s,g,i)=>s+Math.max(0,compactRequiredAssignedToday(i)-compactRecordedForGoal(i)),0);
    const ok=assigned<=goalAssignableFunds(dep)-amount+.001 && nearNeed<=goalAssignableFunds(dep)-amount+.001 &&
      !(dep.currentCashShortfall>0.005) && !model.monthLedger.some(m=>m.carryOut>.005) && model.goalState.every((g,i)=>{
        if(g.insuranceMilestones)return !g.insuranceUnverified&&g.insuranceMilestones.filter(m=>m.inRange).every(m=>m.gap<.005);
        if(g.deadlineMo===null||g.deadlineMo>=model.MONTHS)return false;
        return model.progressDollars[i][Math.max(0,g.deadlineMo)]>=g.effectiveTarget-.005;
      });
    return {ok,model,cashGap:Math.max(0,assigned-(goalAssignableFunds(dep)-amount),nearNeed-(goalAssignableFunds(dep)-amount),dep.currentCashShortfall||0),carryGap:Math.max(0,...model.monthLedger.map(m=>m.carryOut||0))};
  } finally {computeDeployableBalance=originalCompute;goalsData.forEach((g,i)=>{g.manualSaved=saved[i];});}
}
function goalSpendingRoom(sourceIndex=null) {
  const dep=computeDeployableBalance();
  if(!goalSpendScenario(0,sourceIndex)?.ok)return 0;
  let lo=0,hi=Math.floor(Math.max(0,sourceIndex===null?dep.deployable:Math.min(dep.deployable,goalAssignmentFunding(dep)[sourceIndex]?.cash||0))*100);
  // Cent precision, checked at every trial against ALL deadlines and renewals.
  while(lo<hi){const mid=Math.ceil((lo+hi)/2);if(goalSpendScenario(mid/100,sourceIndex)?.ok)lo=mid;else hi=mid-1;}
  return lo/100;
}
function goalTiming(g,model) {
  const done=g.completedAt;
  if(done===null)return 'Not fully funded in forecast';
  if(done===0)return model.pendingClaimsForPlan>0?'Covered by plan (includes pending claims)':g.initialSaved>=g.effectiveTarget?'Covered now':'Could be covered from today’s cash';
  const date=model.monthLabel(done,true);
  const early=g.deadlineMo===null?null:g.deadlineMo-done;
  return `${date}${early===null?'':early>0?` · ${early} month${early===1?'':'s'} early`:early===0?' · deadline month':` · ${-early} months late`}`;
}
function goalTimelineEnd(model) {
  return Math.min(model.MONTHS-1,Math.max(12,...model.goalState.map(g=>Math.max(g.deadlineMo||0,g.completedAt||0)))+3);
}
function goalForecastChart(model, axisEnd=null) {
  const end=axisEnd||goalTimelineEnd(model), cell=104, left=218;
  const width=left+(end+1)*cell+12,height=100+model.goalState.length*54;
  const money=v=>escapeHtml(formatCurrency(v));
  const opening=g=>g.insuranceMilestones?g.insuranceMilestones.reduce((sum,m)=>sum+(m.initialSaved||0),0):Number(g.initialSaved||0);
  const amounts=model.goalState.map((g,i)=>Array.from({length:end+1},(_,m)=>Math.round((Math.max(0,model.allocationData?.[i]?.[m]||0)+(m===0?opening(g):0))*100)));
  let svg=`<svg style="width:${width}px;max-width:none" viewBox="0 0 ${width} ${height}" role="img" aria-label="Monthly goal contributions, with totals above each month">`;
  svg+='<text x="8" y="25">Goal / funding timing</text><text x="8" y="48">Assigned now + forecast →</text>';
  for(let m=0;m<=end;m++){
    const x=left+m*cell,total=amounts.reduce((sum,row)=>sum+row[m],0)/100;
    svg+=`<text x="${x+cell/2}" y="25" text-anchor="middle">${escapeHtml(model.monthLabel(m,true))}</text><text x="${x+cell/2}" y="48" text-anchor="middle" style="font-weight:700;fill:#18786b">${money(total)}</text><line x1="${x}" x2="${x}" y1="58" y2="${height-15}" stroke="#e1e9e7"/>`;
  }
  model.goalState.forEach((g,i)=>{
    const y=70+i*54,done=g.completedAt,deadline=g.deadlineMo;
    const late=done===null||(deadline!==null&&done>deadline),color=late?'#b9514e':'#18786b';
    const initial=g.insuranceMilestones?g.insuranceMilestones.reduce((sum,m)=>sum+(m.initialSaved||0),0):g.initialSaved;
    svg+=`<text x="8" y="${y+16}">${escapeHtml(g.name)}</text><text x="8" y="${y+37}" style="font-size:10px">${escapeHtml(goalTiming(g,model))}</text>`;
    if(deadline!==null)svg+=`<rect x="${left}" y="${y-3}" width="${(Math.max(0,Math.min(end,deadline))+1)*cell-4}" height="32" rx="5" fill="none" stroke="#738b83" stroke-width="1.5" stroke-dasharray="3 5"/>`;
    for(let m=0;m<=end;m++){
      const value=amounts[i][m]/100,x=left+m*cell;
      svg+=`<rect x="${x+3}" y="${y+1}" width="${cell-10}" height="24" rx="4" fill="${value>0?color:'#f3f6f5'}"/><text x="${x+cell/2-2}" y="${y+17}" text-anchor="middle" style="fill:${value>0?'#fff':'#83958f'};font-size:11px">${value>0?money(value):'—'}</text>`;
    }
  });
  return `<div class="gi-chart gi-monthly-chart">${svg}</svg></div><p class="gc-note">Scroll horizontally for later months. Each solid chunk shows the <b>assigned cash (current month) or forecast contribution (future months)</b>; totals above sum the goal chunks below. Current assignments include cash and pending claims, counted once in the current month; they are not new income or a history of past deposits. Dotted outlines mark target months. Insurance chunks include later renewals; its timing label refers to the next premium cycle. Zero contributions are shown as —. Slider changes are temporary previews.</p>`;
}
function goalScenarioOutcome(g,i,model) {
  const deadline=g.deadlineMo, done=g.completedAt;
  if(g.insuranceMilestones){
    const missed=g.insuranceMilestones.filter(m=>m.inRange&&m.gap>.005).sort((a,b)=>a.deadlineMo-b.deadlineMo);
    if(missed.length)return {tone:'late',label:'Short '+formatCurrency(missed[0].gap)+' by '+missed[0].fundBy,detail:missed.length+' insurance funding deadline(s) missed'};
    if(g.insuranceUnverified)return {tone:'unknown',label:'Cannot verify',detail:'An insurance deadline is outside the forecast'};
  } else {
    if(deadline===null||deadline>=model.MONTHS)return {tone:'unknown',label:'Cannot verify deadline',detail:goalTiming(g,model)};
    const gap=Math.max(0,g.effectiveTarget-(model.progressDollars[i]?.[Math.max(0,deadline)]||0));
    if(gap>.005)return {tone:'late',label:'Short '+formatCurrency(gap)+' at target',detail:done===null?'Not fully funded within forecast':model.monthLabel(done,true)+' · '+Math.max(0,done-deadline)+' months late'};
  }
  if(done===0)return {tone:'early',label:model.pendingClaimsForPlan>0?'Covered incl. expected claims':'Covered now',detail:model.pendingClaimsForPlan>0?'Conditional on reimbursements arriving':'Target funded from current cash'};
  const early=done===null||deadline===null?0:deadline-done;
  return {tone:early>0?'early':'ontime',label:early>0?early+' month'+(early===1?'':'s')+' early':'On time',detail:goalTiming(g,model)};
}
function goalScenarioHTML(result,baseline,delta) {
  const model=result.model;
  const title=result.ok?'All targets covered in this forecast':'Plan needs attention';
  const notes=[];
  if(result.cashGap>.005)notes.push(formatCurrency(result.cashGap)+' cash requirement not covered');
  if(result.carryGap>.005)notes.push('Monthly cash shortfall reaches '+formatCurrency(result.carryGap));
  const rows=model.goalState.map((g,i)=>{
    const outcome=goalScenarioOutcome(g,i,model),old=baseline.goalState[i]?.completedAt,now=g.completedAt;
    const change=old===now?'No timing change':old===null&&now!==null?'Now forecast to finish':now===null?'No longer forecast to finish':Math.abs(old-now)+' month'+(Math.abs(old-now)===1?'':'s')+(now<old?' earlier':' later')+' than original';
    return `<tr><th scope="row">${escapeHtml(g.name)}</th><td><span class="gi-outcome gi-${outcome.tone}">${escapeHtml(outcome.label)}</span><small>${escapeHtml(outcome.detail)}</small></td><td>${escapeHtml(change)}</td></tr>`;
  }).join('');
  return `<div class="gi-result-head ${result.ok?'gi-result-ok':'gi-result-alert'}"><strong>${title}</strong><span>${delta===0?'Original plan':delta>0?formatCurrency(delta)+' added':formatCurrency(-delta)+' removed'} · temporary preview</span>${notes.length?'<p>'+escapeHtml(notes.join('. '))+'</p>':''}</div><div class="gc-table-wrap"><table class="gi-table gi-outcomes"><thead><tr><th>Goal</th><th>Result at target</th><th>Change from original</th></tr></thead><tbody>${rows}</tbody></table></div><p class="gc-note">Targets include buffers. Insurance checks every funding deadline, including later renewals. Pending claims must arrive before their goal deadlines. Future income and spending remain estimates.</p>`;
}
function setupGoalSandbox(model) {
  const slider=document.getElementById('giCashSlider'), output=document.getElementById('giCashValue'), chart=document.getElementById('giTimeline'), status=document.getElementById('giScenarioStatus');
  if(!slider)return;
  const cash=Math.max(0,computeDeployableBalance().deployable);
  slider.min=-Math.floor(cash);slider.max=Math.max(10000,Math.ceil(cash));slider.step=1;slider.value=0;
  const axisEnd=goalTimelineEnd(model);
  const outcomes=document.getElementById('giScenarioOutcomes');
  const originalResult=withGoalCalculationCache(()=>goalSpendScenario(0));
  function showResult(result,delta){
    status.textContent=result.ok?'All targets covered in this forecast.':'Plan needs attention — see the goal results below.';
    status.classList.toggle('gi-warning',!result.ok);
    if(outcomes)outcomes.innerHTML=goalScenarioHTML(result,model,delta);
  }
  if(originalResult)showResult(originalResult,0);
  let frame=null;
  function apply(){
    const delta=Number(slider.value);
    const result=withGoalCalculationCache(()=>goalSpendScenario(-delta));
    if(!result)return;
    chart.innerHTML=goalForecastChart(result.model,Math.max(axisEnd,goalTimelineEnd(result.model)));
    showResult(result,delta);
  }
  slider.addEventListener('input',()=>{
    const delta=Number(slider.value);
    output.textContent=delta<0?`${formatCurrency(-delta)} removed`:delta>0?`${formatCurrency(delta)} added`:'No change';
    if(frame===null)frame=requestAnimationFrame(()=>{frame=null;apply();});
  });
  slider.addEventListener('change',()=>{if(frame!==null)cancelAnimationFrame(frame);frame=null;apply();});
  document.getElementById('giCashReset').addEventListener('click',()=>{if(frame!==null)cancelAnimationFrame(frame);frame=null;slider.value=0;output.textContent='No change';chart.innerHTML=goalForecastChart(model,axisEnd);if(originalResult)showResult(originalResult,0);});
}
function renderGoalInsights(model) {
  const host=document.getElementById('goalInsights');if(!host)return;
  const room=goalSpendingRoom();
  const rows=model.goalState.map((g,i)=>({g,i,room:goalSpendingRoom(i)}));
  host.innerHTML=`<section class="gc-panel gi-home"><div class="gc-section-head"><div><h2>Your goal forecast</h2><p>When each goal could be funded—and what flexibility remains today.</p></div></div><div class="gi-takeaway"><div><span>Estimated spending room today</span><strong>${formatCurrency(room)}</strong></div><p>${room>0?'The forecast can absorb this one-time spend while still meeting all checked goal targets, buffers and insurance funding dates.':'No spending room is confirmed under the current plan. Check shortfalls, missing deadlines and near-term cash needs.'}<br><b>This is conditional on forecast income and spending. Spending it reduces your safety margin.</b></p></div><div class="gi-sandbox"><div class="gi-slider-head"><label for="giCashSlider">Try a one-time cash change</label><output id="giCashValue" for="giCashSlider">No change</output><button type="button" id="giCashReset">Reset</button></div><input id="giCashSlider" type="range" aria-describedby="giSliderHelp giScenarioStatus"><div class="gi-slider-ends"><span>Remove money</span><span>Add money</span></div><p id="giSliderHelp" class="gc-note">Preview only—never saved. Changes total cash available for goals today. Added money follows your funding priorities; removal uses unassigned cash first, then reverses funding priority. The results immediately below and chart update live. Spending-room estimates further down remain the original plan.</p><p id="giScenarioStatus" role="status" aria-live="polite">Current plan · no hypothetical cash change.</p><div id="giScenarioOutcomes"></div></div><div id="giTimeline">${goalForecastChart(model)}</div><div class="gc-table-wrap"><table class="gi-table"><thead><tr><th>Goal</th><th>Projected fully funded</th><th>Could spend from this assignment</th></tr></thead><tbody>${rows.map(({g,i,room})=>`<tr><td>${escapeHtml(g.name)}</td><td>${escapeHtml(goalTiming(g,model))}${g.insuranceMilestones?'<small>Next premiums; each earlier deadline is also checked</small>':''}</td><td><b>${formatCurrency(room)}</b><small>of ${formatCurrency(goalAssignmentFunding()[i]?.cash||0)} cash-backed assignment</small></td></tr>`).join('')}</tbody></table></div><p class="gc-note"><b>Individual amounts are alternatives, not additive.</b> Each tests spending from that goal alone while protecting every goal. The headline tests one combined withdrawal. These are extra purchases, not payments already included in a goal’s target. Nothing is withdrawn or saved automatically.</p></section>`;
  setupGoalSandbox(model);
}
