// Spending must reduce both the bank cash pool AND the source assignments.
// Reassigning money alone is not a spending simulation.
function goalSpendScenario(amount, sourceIndex=null) {
  const originalCompute=computeDeployableBalance, dep=originalCompute();
  const saved=goalsData.map(g=>Number(g.manualSaved||0));
  const limit=sourceIndex===null?dep.deployable:Math.min(dep.deployable,saved[sourceIndex]||0);
  if(amount<0||amount>limit+.001)return null;
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
    const ok=assigned<=dep.deployable-amount+.001 && nearNeed<=dep.deployable-amount+.001 &&
      !(dep.currentCashShortfall>0.005) && !model.monthLedger.some(m=>m.carryOut>.005) && model.goalState.every((g,i)=>{
        if(g.insuranceMilestones)return !g.insuranceUnverified&&g.insuranceMilestones.filter(m=>m.inRange).every(m=>m.gap<.005);
        if(g.deadlineMo===null||g.deadlineMo>=model.MONTHS)return false;
        return model.progressDollars[i][Math.max(0,g.deadlineMo)]>=g.effectiveTarget-.005;
      });
    return {ok,model};
  } finally {computeDeployableBalance=originalCompute;goalsData.forEach((g,i)=>{g.manualSaved=saved[i];});}
}
function goalSpendingRoom(sourceIndex=null) {
  const dep=computeDeployableBalance();
  if(!goalSpendScenario(0,sourceIndex)?.ok)return 0;
  let lo=0,hi=Math.floor(Math.max(0,sourceIndex===null?dep.deployable:Math.min(dep.deployable,goalsData[sourceIndex]?.manualSaved||0))*100);
  // Cent precision, checked at every trial against ALL deadlines and renewals.
  while(lo<hi){const mid=Math.ceil((lo+hi)/2);if(goalSpendScenario(mid/100,sourceIndex)?.ok)lo=mid;else hi=mid-1;}
  return lo/100;
}
function goalTiming(g,model) {
  const done=g.completedAt;
  if(done===null)return 'Not fully funded in forecast';
  if(done===0)return g.initialSaved>=g.effectiveTarget?'Covered now':'Could be covered from today’s cash';
  const date=model.monthLabel(done,true);
  const early=g.deadlineMo===null?null:g.deadlineMo-done;
  return `${date}${early===null?'':early>0?` · ${early} month${early===1?'':'s'} early`:early===0?' · deadline month':` · ${-early} months late`}`;
}
function goalForecastChart(model) {
  const colors=['#18786b','#4277b9','#b27b29','#8862aa','#bd6262','#5a8570'];
  const end=Math.max(1,Math.min(model.MONTHS-1,Math.max(12,...model.goalState.map(g=>Math.max(g.completedAt||0,Math.min(g.deadlineMo||0,36))))));
  const x=m=>48+m/end*828,y=p=>225-Math.min(100,Math.max(0,p))*1.85;
  let svg='<svg viewBox="0 0 910 280" role="img" aria-label="Projected funding progress for each goal, from today to future months">';
  for(const p of [0,25,50,75,100])svg+=`<line x1="48" x2="876" y1="${y(p)}" y2="${y(p)}" stroke="#e1e9e7"/><text x="38" y="${y(p)+4}" text-anchor="end">${p}%</text>`;
  for(let m=0;m<=end;m+=Math.max(1,Math.ceil(end/6)))svg+=`<text x="${x(m)}" y="253" text-anchor="middle">${escapeHtml(model.monthLabel(m,true))}</text>`;
  model.goalState.forEach((g,i)=>{
    const values=[g.initialSaved,...model.progressDollars[i].slice(1,end+1)];
    const pts=values.map((v,m)=>`${x(m)},${y(v/Math.max(1,g.effectiveTarget)*100)}`).join(' ');
    svg+=`<polyline fill="none" stroke="${colors[i%colors.length]}" stroke-width="2.5" stroke-dasharray="6 3" points="${pts}"><title>${escapeHtml(g.name)}: ${escapeHtml(goalTiming(g,model))}</title></polyline><circle cx="48" cy="${y(g.initialSaved/Math.max(1,g.effectiveTarget)*100)}" r="4" fill="${colors[i%colors.length]}"/>`;
  });
  return `<div class="gi-chart">${svg}</svg></div><div class="gi-legend">${model.goalState.map((g,i)=>`<span><i style="background:${colors[i%colors.length]}"></i>${escapeHtml(g.name)}</span>`).join('')}</div><p class="gc-note">Dots show funding already assigned or recorded. Dashed lines show forecast progress, including proposed assignments. Insurance shows the next premium cycle; later renewals are still checked separately.</p>`;
}
function renderGoalInsights(model) {
  const host=document.getElementById('goalInsights');if(!host)return;
  const room=goalSpendingRoom();
  const rows=model.goalState.map((g,i)=>({g,i,room:goalSpendingRoom(i)}));
  host.innerHTML=`<section class="gc-panel gi-home"><div class="gc-section-head"><div><h2>Your goal forecast</h2><p>When each goal could be funded—and what flexibility remains today.</p></div></div><div class="gi-takeaway"><div><span>Estimated spending room today</span><strong>${formatCurrency(room)}</strong></div><p>${room>0?'The forecast can absorb this one-time spend while still meeting all checked goal targets, buffers and insurance funding dates.':'No spending room is confirmed under the current plan. Check shortfalls, missing deadlines and near-term cash needs.'}<br><b>This is conditional on forecast income and spending. Spending it reduces your safety margin.</b></p></div>${goalForecastChart(model)}<div class="gc-table-wrap"><table class="gi-table"><thead><tr><th>Goal</th><th>Projected fully funded</th><th>Could spend from this assignment</th></tr></thead><tbody>${rows.map(({g,i,room})=>`<tr><td>${escapeHtml(g.name)}</td><td>${escapeHtml(goalTiming(g,model))}${g.insuranceMilestones?'<small>Next premiums; each earlier deadline is also checked</small>':''}</td><td><b>${formatCurrency(room)}</b><small>of ${formatCurrency(goalsData[i].manualSaved||0)} assigned</small></td></tr>`).join('')}</tbody></table></div><p class="gc-note"><b>Individual amounts are alternatives, not additive.</b> Each tests spending from that goal alone while protecting every goal. The headline tests one combined withdrawal. These are extra purchases, not payments already included in a goal’s target. Nothing is withdrawn or saved automatically.</p></section>`;
}
