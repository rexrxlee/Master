let compactGoalChart = null;
let compactForecastFrame = null;
let compactForecastTimer = null;
let compactGoalSort = "original";

function compactMoney(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function compactNearlyZero(value) {
  return Math.abs(Number(value) || 0) < 0.005;
}

function sortCompactGoalRows(value) {
  if (!["original", "priority"].includes(value)) return;
  compactGoalSort = value;
  const container = document.getElementById("compactGoalRows");
  if (!container) return;
  const order = goalsData.map((goal, idx) => ({ ...goal, originalIdx: idx }));
  if (value === "priority") order.sort(compareGoalPriorityOrder);
  order.forEach(goal => {
    const row = document.getElementById(`compactGoal_${goal.originalIdx}`);
    if (row) container.appendChild(row);
  });
}
let compactSavedPlan = null;
let compactSaving = false;

function captureCompactSavedPlan() {
  incomeBoosts.forEach(boost => {
    boost.kind = _normaliseBoostKind(boost);
    boost.frequency = _normaliseBoostFrequency(boost);
    boost.amount = Math.abs(Number(boost.amount || 0) || 0);
    boost.fromMonth = _normaliseBoostMonth(boost.fromMonth);
    boost.toMonth = boost.frequency === "once" ? "" : _normaliseBoostMonth(boost.toMonth);
    if (boost.kind === "reduce") boost.toGoal = "any";
  });
  compactSavedPlan = JSON.stringify({ goals: goalsData, boosts: incomeBoosts });
}

function compactPlanChanged() {
  return compactSavedPlan !== null && compactSavedPlan !== JSON.stringify({ goals: goalsData, boosts: incomeBoosts });
}

function updateCompactSaveStatus() {
  const status = document.getElementById("goalsAutosaveStatus");
  if (status) status.textContent = compactSaving ? "Saving…" : compactPlanChanged() ? "Preview — not saved" : "No unsaved changes";
  const save = document.getElementById("compactSavePlan");
  const reset = document.getElementById("compactResetPlan");
  if (save) save.disabled = compactSaving || !compactPlanChanged();
  if (reset) reset.disabled = compactSaving || !compactPlanChanged();
}

async function saveCompactPlan() {
  if (compactSaving || !compactPlanChanged()) return;
  compactNormaliseManualAssignments();
  compactSaving = true;
  updateCompactSaveStatus();
  try {
    const saved = await persistGoalsToExcel({ silent: false, includeBoosts: true, waitForIdle: true });
    if (!saved) throw new Error("Save is still pending. Please try Save plan again.");
    captureCompactSavedPlan();
    renderCompactGoalsPage();
  } catch (error) {
    const status = document.getElementById("goalsAutosaveStatus");
    if (status) status.textContent = "Save failed — your preview is kept. " + error.message;
  } finally {
    compactSaving = false;
    const save = document.getElementById("compactSavePlan");
    const reset = document.getElementById("compactResetPlan");
    if (save) save.disabled = !compactPlanChanged();
    if (reset) reset.disabled = !compactPlanChanged();
    if (!compactPlanChanged()) updateCompactSaveStatus();
  }
}

function resetCompactPlan() {
  if (compactSaving || !compactSavedPlan) return;
  cancelAnimationFrame(compactForecastFrame);
  clearTimeout(goalInsightsRefreshTimer);
  const saved = JSON.parse(compactSavedPlan);
  goalsData = saved.goals;
  incomeBoosts = saved.boosts;
  incomeBoostsDirty = false;
  renderCompactGoalsPage();
}

function compactGoalUsableBreakdown(dep) {
  const budgetPageTotal = (budgetSummary.billsTotal || 0) + (budgetSummary.monthlyTotal || 0);
  const budgetHoldAmount = futureSalaryBudgetOverride === null ? budgetPageTotal : futureSalaryBudgetOverride;
  const futureDetails = (dep.futureSalaryHoldDetails || []).map(item =>
    `<div><span>${escapeHtml(item.label)}</span><strong>${formatCurrency(item.reserve)}</strong><small>${formatCurrency(item.futureSalary)} salary, ${formatCurrency(item.budgetForMonth || item.reserve)} budget pulled</small></div>`
  ).join("");
  return `
    <section class="cg-usable-now" aria-label="Money available for goals">
      <div class="cg-usable-main">
        <span>Can be used for goals next</span>
        <strong class="${dep.deployable < 0 ? "red" : ""}">${formatCurrency(dep.deployable)}</strong>
      </div>
      <div class="cg-usable-formula">
        <span><b>${formatCurrency(dep.rawSavings)}</b><small>selected goal accounts</small></span>
        <i>−</i>
        <span><b>${formatCurrency(ccOwed)}</b><small>all credit card debt</small></span>
        <i>−</i>
        <span><b>${formatCurrency(dep.remainingBudget)}</b><small>goal-account budget left</small></span>
        <i>−</i>
        <span><b>${formatCurrency(dep.futureSalaryHold)}</b><small>future salary budget hold</small></span>
        <i>+</i>
        <span><b>${formatCurrency(dep.claimReceivableForGoals)}</b><small>pending claims</small></span>
      </div>
      <p class="cg-note">Use this amount for goal assignment. Future salary is not removed in full; only the budget needed from that salary is held back.</p>
      <div class="cg-budget-copy">
        <label><input type="checkbox" ${holdFutureSalaryBudget ? "checked" : ""} onchange="setHoldFutureSalaryBudget(this.checked)"> Copy same Budget page total for future salary</label>
        <button type="button" class="btn-secondary btn-sm" onclick="pullGoalBudgetFromExcel()">Pull Budget</button>
        <span class="cg-note">Budget page total: ${formatCurrency(budgetPageTotal)}</span>
        <label class="cg-budget-manual">Budget hold amount ($)<input type="number" min="0" step="1" value="${budgetHoldAmount}" oninput="setFutureSalaryBudgetOverride(this.value)"></label>
        <span class="cg-note" id="compactBudgetPullStatus"></span>
        <div class="cg-budget-copy-list">${futureDetails || '<span>No future salary budget hold right now.</span>'}</div>
      </div>
    </section>`;
}

function deleteCompactGoal(idx) {
  goalsData.splice(idx, 1);
  renderCompactGoalsPage();
  updateCompactSaveStatus();
}

function compactPriorityOptions(value) {
  return ["Critical", "High", "Medium", "Low"].map(name => `<option ${name === value ? "selected" : ""}>${name}</option>`).join("");
}

function renderCompactGoalsPage() {
  compactNormaliseManualAssignments();
  if (!compactSavedPlan) captureCompactSavedPlan();
  const container = document.getElementById("goalsContainer");
  const adjustmentsOpen = document.getElementById("boostsPanel")?.open || incomeBoostsDirty;
  if (compactGoalChart) { compactGoalChart.destroy(); compactGoalChart = null; }
  const today = toDateInputValue(new Date());
  container.innerHTML = `
    <details class="simple-details cg-account-source" open><summary>Which accounts fund my goals?</summary><div class="simple-details-body cg-accounts">
      ${allAccounts.filter(account => account.type === "Savings").map(account => `<label><input type="checkbox" value="${escapeHtml(account.name)}" ${goalSavingsAccts.includes(account.name) ? "checked" : ""} onchange="toggleGoalAccount(this)">${escapeHtml(account.name)} <strong>${formatCurrency(savingsBalances[account.name] || 0)}</strong></label>`).join("") || '<p>Add a savings account in Accounts & Setup first.</p>'}
    </div></details>
    <section class="cg-summary" id="compactGoalSummary" aria-live="polite"></section>
    <div id="compactUsableBreakdown"></div>
    <details class="simple-details cg-add" ${goalsData.length ? "" : "open"}>
      <summary>＋ Add a goal</summary>
      <form class="cg-add-form" onsubmit="event.preventDefault(); addCompactGoal(this)">
        <label>Goal name<input name="name" required maxlength="100" placeholder="Japan trip"></label>
        <label>Target ($)<input name="target" type="number" min="0.01" step="0.01" required></label>
        <label>Start<input name="start" type="date" value="${today}" required></label>
        <label>Deadline<input name="end" type="date" required></label>
        <label>Priority<select name="urgency">${compactPriorityOptions("Medium")}</select></label>
        <button class="btn-primary" type="submit">Add goal</button>
        <p class="cg-form-message" role="status"></p>
      </form>
    </details>
    <section class="cg-setup-panel">
      <div class="cg-heading"><h2>Goal setup</h2></div>
      <div class="cg-setup-grid">${goalsData.map((goal, idx) => compactGoalSetupRow(goal, idx)).join("") || '<p class="cg-note">Add your first goal above, then use the allocation chart below.</p>'}</div>
    </section>
    <div id="compactAdjustments"></div>
    <div class="cg-workspace">
      <section class="cg-controls">
        <div class="cg-heading"><h2>Assign money</h2><button class="btn-primary" onclick="compactSmartAssign()" ${goalsData.length ? "" : "disabled"}>Smart Assign</button></div>
        <div class="cg-assign-total" id="compactAssignTotal"></div>
        <label class="cg-sort" for="compactGoalSort">Sort by <select id="compactGoalSort" onchange="sortCompactGoalRows(this.value)"><option value="original" ${compactGoalSort === "original" ? "selected" : ""}>Original order</option><option value="priority" ${compactGoalSort === "priority" ? "selected" : ""}>Priority (highest first)</option></select></label>
        <p class="cg-note">Try allocations freely. Sliders and Smart Assign only preview your plan; choose Save plan when ready.</p>
        <div class="cg-chart-scroll"><div class="cg-chart-wrap" id="compactGoalTimeline" role="img" aria-label="Goal allocation timeline"></div></div>
        <div id="compactGoalRows">${goalsData.map((goal, idx) => compactGoalRow(goal, idx)).join("") || '<p class="cg-note">Add your first goal to start planning.</p>'}</div>
        <div id="compactForecastResults" aria-live="polite"></div>
        <details class="cg-method"><summary>How this forecast is calculated</summary><div id="compactForecastMethod"></div></details>
        <div class="cg-save-actions"><button id="compactSavePlan" class="btn-primary" onclick="saveCompactPlan()">Save plan</button><button id="compactResetPlan" class="btn-secondary" onclick="resetCompactPlan()">Reset changes</button><span id="goalsAutosaveStatus" class="goals-autosave-status" role="status">No unsaved changes</span></div>
      </section>
    </div>`;
  renderIncomeBoostsPanel(document.getElementById("compactAdjustments"));
  document.getElementById("boostsPanel").open = adjustmentsOpen;
  sortCompactGoalRows(compactGoalSort);
  refreshCompactGoalForecast();
  updateCompactSaveStatus();
}

function compactGoalRow(goal, idx) {
  const targetWithBuffer = compactGoalTargetWithBuffer(idx);
  const recorded = compactRecordedForGoal(idx);
  const cashAssigned = compactMoney(goal.manualSaved || 0);
  const assigned = Math.min(targetWithBuffer, recorded + cashAssigned);
  return `<article class="cg-goal" id="compactGoal_${idx}">
    <div class="cg-heading"><strong>${escapeHtml(goal.name)}</strong><span id="cgPill_${idx}" class="cg-pill">Preview</span></div>
    <div class="cg-mini-bar" aria-hidden="true"><span id="cgBar_${idx}"></span></div>
    <div class="cg-goal-metrics"><span>Assigned <strong>${formatCurrency(assigned)}</strong></span></div>
    <div class="cg-allocation"><input id="cgSlider_${idx}" type="range" min="${recorded}" max="${compactSliderLimit(idx)}" step="0.01" value="${assigned}" aria-label="Total money assigned to ${escapeHtml(goal.name)}" oninput="assignCompactGoal(${idx}, this.value)"><label>Assigned ($)<input id="cgAmount_${idx}" type="number" min="${recorded}" step="0.01" value="${assigned.toFixed(2)}" oninput="assignCompactGoal(${idx}, this.value)"></label></div>
    ${recorded > 0 ? `<p class="cg-note">Includes ${formatCurrency(recorded)} already recorded for this goal.</p>` : ""}
    <p class="cg-result" id="cgStatus_${idx}"></p>
  </article>`;
}

function compactNormaliseManualAssignments() {
  goalsData.forEach(goal => {
    goal.manualSaved = Math.max(0, compactMoney(goal.manualSaved || 0));
  });
}

function compactGoalSetupRow(goal, idx) {
  return `<article class="cg-setup-goal">
    <div class="cg-heading"><strong>${escapeHtml(goal.name)}</strong><select aria-label="Priority for ${escapeHtml(goal.name)}" onchange="updateCompactGoal(${idx}, 'urgency', this.value)">${compactPriorityOptions(goal.urgency)}</select></div>
    <div class="cg-options">
      <label>Name<input value="${escapeHtml(goal.name)}" onchange="updateCompactGoal(${idx}, 'name', this.value)"></label>
      <label>Target ($)<input type="number" min="0.01" step="0.01" value="${goal.target}" onchange="updateCompactGoal(${idx}, 'target', this.value)"></label>
      <label>Monthly plan ($)<input type="number" min="0" step="0.01" value="${goal.monthlyAlloc}" onchange="updateCompactGoal(${idx}, 'monthlyAlloc', this.value)"></label>
      <label>Start<input type="date" value="${escapeHtml(goal.startDate || "")}" onchange="updateCompactGoal(${idx}, 'startDate', this.value)"></label>
      <label>Deadline<input type="date" value="${escapeHtml(goal.endDate || "")}" onchange="updateCompactGoal(${idx}, 'endDate', this.value)"></label>
      <label>Buffer (%)<input type="number" min="0" max="100" value="${goal.goalBuffer || 0}" onchange="updateCompactGoal(${idx}, 'goalBuffer', this.value)"></label>
      <label><input type="checkbox" ${goal.priority ? "checked" : ""} onchange="updateCompactGoal(${idx}, 'priority', this.checked ? 1 : 0)">Fund before other priorities</label>
      <label>Notes<input value="${escapeHtml(goal.notes || "")}" onchange="updateCompactGoal(${idx}, 'notes', this.value)"></label>
      <a href="add-transaction.html?goal=${encodeURIComponent(goal.name)}">Record goal expense</a>
      <button class="btn-secondary" onclick="deleteCompactGoal(${idx})">Delete goal</button>
    </div>
  </article>`;
}

function addCompactGoal(form) {
  const fields = form.elements;
  const name = fields.namedItem("name").value.trim();
  const target = Number(fields.namedItem("target").value);
  const startDate = fields.namedItem("start").value;
  const endDate = fields.namedItem("end").value;
  const message = form.querySelector(".cg-form-message");
  if (!name || !Number.isFinite(target) || target <= 0 || !startDate || !endDate || endDate < startDate) {
    message.textContent = "Enter a name, positive target, and a deadline on or after the start date."; return;
  }
  if (goalsData.length >= MAX_GOALS || goalsData.some(goal => businessKey(goal.name) === businessKey(name))) {
    message.textContent = goalsData.length >= MAX_GOALS ? `Maximum ${MAX_GOALS} goals.` : "Use a unique goal name."; return;
  }
  goalsData.push({ name, target, startDate, endDate, urgency: fields.namedItem("urgency").value,
    manualSaved: 0, monthlyAlloc: 0, goalBuffer: 0, priority: 0, notes: "", color: getNextGoalColor() });
  renderCompactGoalsPage();
  scheduleGoalsAutoSave();
}

function compactGoalNeed(idx) {
  const goal = goalsData[idx];
  return Math.max(0, compactGoalTargetWithBuffer(idx) - compactRecordedForGoal(idx));
}

function compactSliderLimit(idx) {
  const goal = goalsData[idx];
  const targetWithBuffer = compactGoalTargetWithBuffer(idx);
  return Math.max(1, targetWithBuffer, compactRecordedForGoal(idx) + Number(goal?.manualSaved || 0));
}

function assignCompactGoal(idx, value) {
  if (value === "") return;
  const amount = Number(value);
  if (!Number.isFinite(amount)) return;
  const recorded = compactRecordedForGoal(idx);
  const totalAssigned = compactMoney(Math.max(recorded, Math.min(amount, compactSliderLimit(idx))));
  goalsData[idx].manualSaved = compactMoney(Math.max(0, totalAssigned - recorded));
  document.getElementById(`cgAmount_${idx}`).value = totalAssigned.toFixed(2);
  document.getElementById(`cgSlider_${idx}`).value = totalAssigned;
  updateCompactGoalInstantPreview(idx);
  updateCompactAllocationSummary();
  cancelAnimationFrame(compactForecastFrame);
  clearTimeout(compactForecastTimer);
  compactForecastTimer = setTimeout(() => {
    compactForecastFrame = requestAnimationFrame(refreshCompactGoalForecast);
  }, 140);
  scheduleGoalsAutoSave();
}

function updateCompactGoalInstantPreview(idx) {
  const goal = goalsData[idx];
  if (!goal) return;
  const targetWithBuffer = Math.max(1, compactGoalTargetWithBuffer(idx));
  const assigned = compactTotalAssignedForGoal(idx);
  const pct = Math.min(100, (assigned / targetWithBuffer) * 100);
  const bar = document.getElementById(`cgBar_${idx}`);
  if (bar) bar.style.width = pct + "%";
  const pill = document.getElementById(`cgPill_${idx}`);
  if (pill) {
    const short = Math.max(0, targetWithBuffer - assigned);
    pill.textContent = short < 0.01 ? "On track" : `${formatCurrency(short)} left`;
    pill.classList.toggle("red", short >= 0.01);
  }
}

function compactGoalTargetWithBuffer(idx) {
  const goal = goalsData[idx];
  return goal ? goal.target * (1 + (goal.goalBuffer || 0) / 100) : 0;
}

function compactRecordedForGoal(idx) {
  const goal = goalsData[idx];
  if (!goal) return 0;
  return compactMoney(Math.min(compactGoalTargetWithBuffer(idx), Math.max(0, getSavedViaTransactions(goal.name))));
}

function compactTotalAssignedForGoal(idx) {
  return compactMoney(compactRecordedForGoal(idx) + Number(goalsData[idx]?.manualSaved || 0));
}

function updateCompactAllocationSummary() {
  const summary = document.getElementById("compactGoalSummary");
  if (!summary) return;
  const dep = computeDeployableBalance();
  const assigned = goalsData.reduce((sum, goal) => sum + Number(goal.manualSaved || 0), 0);
  const free = dep.deployable - assigned;
  summary.innerHTML = `<div><small>Available for goals</small><strong>${formatCurrency(dep.deployable)}</strong></div><div><small>Cash assigned with sliders</small><strong>${formatCurrency(assigned)}</strong></div><div><small>${free < 0 ? "Over-assigned" : "Unassigned cash today"}</small><strong class="${free < 0 ? "red" : ""}">${formatCurrency(Math.abs(free))}</strong></div><div><small>12-mo normal savings / month</small><strong>${formatCurrency(Math.max(0, historicalStats.avgMonthlySavings))}</strong></div>`;
  const assignTotal = document.getElementById("compactAssignTotal");
  if (assignTotal) assignTotal.innerHTML = `<span>${free < 0 ? "Reduce slider cash by" : "Unassigned cash today"}</span><strong class="${free < 0 ? "red" : ""}">${formatCurrency(free < 0 ? Math.abs(free) : free)}</strong>`;
  updateCompactSaveStatus();
}

function updateCompactGoal(idx, field, raw) {
  const goal = goalsData[idx];
  let value = raw;
  if (["target", "monthlyAlloc", "goalBuffer", "priority"].includes(field)) {
    value = Number(raw);
    if (!Number.isFinite(value) || value < 0 || (field === "target" && value <= 0) || (field === "goalBuffer" && value > 100)) { renderCompactGoalsPage(); return; }
  }
  if (field === "name") {
    value = String(raw).trim();
    if (!value || goalsData.some((other, i) => i !== idx && businessKey(other.name) === businessKey(value))) { renderCompactGoalsPage(); return; }
    // Transactions link by name; avoid silently orphaning recorded goal progress.
    if (value !== goal.name && allTxForGoals.some(row => businessKey(row["Sub Category"]) === businessKey("Goal: " + goal.name))) {
      alert("This goal has recorded transactions. Keep its name so the transactions remain linked."); renderCompactGoalsPage(); return;
    }
    incomeBoosts.forEach(boost => { if (boost.toGoal === goal.name) { boost.toGoal = value; incomeBoostsDirty = true; } });
  }
  const proposed = { ...goal, [field]: value };
  if (proposed.startDate && proposed.endDate && proposed.endDate < proposed.startDate) {
    alert("The deadline must be on or after the start date."); renderCompactGoalsPage(); return;
  }
  goalsData[idx] = proposed;
  renderCompactGoalsPage();
  scheduleGoalsAutoSave();
}

function compactSmartAssign() {
  let available = Math.max(0, computeDeployableBalance().deployable);
  const today = new Date();
  const order = getCompactFundingOrder();
  goalsData.forEach(goal => { goal.manualSaved = 0; });

  const goalNeed = idx => {
    const goal = goalsData[idx];
    const target = goal.target * (1 + (goal.goalBuffer || 0) / 100);
    return Math.max(0, target - getSavedViaTransactions(goal.name) - Number(goal.manualSaved || 0));
  };

  const urgencyWindow = order.filter(idx => {
    const date = _parseGoalDateValue(goalsData[idx].endDate);
    if (!date) return false;
    const months = Math.ceil((date - today) / (1000 * 60 * 60 * 24 * 30));
    return months <= 1 || isForcePriorityGoal(goalsData[idx]);
  });

  urgencyWindow.forEach(idx => {
    if (compactNearlyZero(available)) return;
    const amount = Math.min(available, goalNeed(idx));
    if (amount <= 0.005) return;
    goalsData[idx].manualSaved = compactMoney(goalsData[idx].manualSaved + amount);
    available = compactMoney(Math.max(0, available - amount));
  });

  while (available > 0.005) {
    const candidates = order
      .map(idx => ({ idx, need: goalNeed(idx), weight: compactSmartAssignWeight(goalsData[idx], today) }))
      .filter(item => item.need > 0.005 && item.weight > 0);
    if (!candidates.length) break;
    const totalWeight = candidates.reduce((sum, item) => sum + item.weight, 0);
    let usedThisRound = 0;
    candidates.forEach(item => {
      if (compactNearlyZero(available)) return;
      const share = available * (item.weight / totalWeight);
      const amount = compactMoney(Math.min(item.need, Math.max(0.01, share), available));
      goalsData[item.idx].manualSaved = compactMoney(goalsData[item.idx].manualSaved + amount);
      available = compactMoney(Math.max(0, available - amount));
      usedThisRound += amount;
    });
    if (usedThisRound <= 0.005) break;
  }

  const forced = order.filter(idx => isForcePriorityGoal(goalsData[idx]));
  forced.forEach(idx => {
    if (compactNearlyZero(available)) return;
    const amount = Math.min(available, goalNeed(idx));
    if (amount <= 0.005) return;
    goalsData[idx].manualSaved = compactMoney(goalsData[idx].manualSaved + amount);
    available = compactMoney(Math.max(0, available - amount));
  });

  compactNormaliseManualAssignments();
  compactReconcileSmartAssignCents();
  renderCompactGoalsPage();
  scheduleGoalsAutoSave(150);
}

function compactReconcileSmartAssignCents() {
  const total = () => compactMoney(goalsData.reduce((sum, goal) => sum + Number(goal.manualSaved || 0), 0));
  const deployable = compactMoney(Math.max(0, computeDeployableBalance().deployable));
  let diff = compactMoney(deployable - total());
  if (Math.abs(diff) > 0.1 || compactNearlyZero(diff)) return;

  if (diff < 0) {
    const idx = getCompactFundingOrder().slice().reverse().find(i => Number(goalsData[i]?.manualSaved || 0) >= Math.abs(diff));
    if (idx !== undefined) goalsData[idx].manualSaved = compactMoney(goalsData[idx].manualSaved + diff);
    return;
  }

  const idx = getCompactFundingOrder().find(i => compactGoalNeed(i) - Number(goalsData[i]?.manualSaved || 0) >= diff);
  if (idx !== undefined) goalsData[idx].manualSaved = compactMoney(goalsData[idx].manualSaved + diff);
}

function compactSmartAssignWeight(goal, today) {
  const deadline = _parseGoalDateValue(goal.endDate);
  const daysLeft = deadline ? Math.max(0, Math.ceil((deadline - today) / (1000 * 60 * 60 * 24))) : 365;
  const urgencyWeight = deadline ? Math.max(1, 120 / Math.max(7, daysLeft + 7)) : 0.6;
  const priorityWeight = ({ Critical: 2.2, High: 1.6, Medium: 1, Low: 0.7 })[goal.urgency] || 1;
  const focusWeight = isForcePriorityGoal(goal) ? 2.5 : 1;
  return urgencyWeight * priorityWeight * focusWeight;
}

function getCompactFundingOrder() {
  return goalsData
    .map((goal, idx) => ({ goal, idx }))
    .sort((a, b) => {
      const forced = (isForcePriorityGoal(b.goal) ? 1 : 0) - (isForcePriorityGoal(a.goal) ? 1 : 0);
      if (forced !== 0) return forced;
      const aDate = _parseGoalDateValue(a.goal.endDate);
      const bDate = _parseGoalDateValue(b.goal.endDate);
      if (aDate && bDate && aDate.getTime() !== bDate.getTime()) return aDate - bDate;
      if (aDate) return -1;
      if (bDate) return 1;
      return goalPriorityValue(a.goal) - goalPriorityValue(b.goal) || a.idx - b.idx;
    })
    .map(item => item.idx);
}

function compactOverallStatus(rows, free, removable) {
  const shortGoals = rows.filter(row => row.gap >= 0.01);
  const unassigned = compactMoney(Math.max(0, free));
  if (free < -0.005) {
    return {
      tone: "bad",
      title: "Over-assigned",
      detail: `Reduce slider cash by ${formatCurrency(Math.abs(free))} or lower one goal before relying on this plan.`
    };
  }
  if (shortGoals.length) {
    const worst = shortGoals.reduce((max, row) => row.gap > max.gap ? row : max, shortGoals[0]);
    const needed = compactCashNeededToday();
    return {
      tone: "bad",
      title: `${shortGoals.length} goal${shortGoals.length === 1 ? "" : "s"} need review`,
      detail: needed > 0.005
        ? `Assign about ${formatCurrency(needed)} more today to stay on track. ${worst.name} is shortest by ${formatCurrency(worst.gap)}.`
        : `${worst.name} is shortest by ${formatCurrency(worst.gap)}. Use Smart Assign, reduce another goal, or lower budget spending.`
    };
  }
  return {
    tone: "good",
    title: unassigned > 0.005 ? "Can add a goal today" : "No spare cash today",
    detail: unassigned > 0.005
      ? `${formatCurrency(unassigned)} is unassigned now and can go to a new goal without changing existing goal sliders.`
      : removable.amount > 0.005
        ? `All current cash is assigned. You can only fund a new goal today by reducing later flexible sliders; up to ${formatCurrency(removable.amount)} may be replaced by forecast/adjustments without missing current dates.`
        : "All current cash is assigned and needed for the current plan. Add a goal only if you lower another goal, increase forecast savings, or add an adjustment."
  };
}

function compactTimingStatus(state, model, gap, capped) {
  if (capped) return "Beyond forecast range";
  if (state.deadlineMo < 0) return gap < 0.01 ? "Target covered" : `Overdue · ${formatCurrency(gap)} short`;
  if (gap >= 0.01) return `${formatCurrency(gap)} short`;
  if (state.deadlineMo === null) return "On track";
  if (state.completedAt === 0) return "Funded now";
  if (state.completedAt !== null) {
    const timing = model.relativeDeadlineText(state.completedAt, state.deadlineMo);
    if (timing === "on deadline") return "On time";
    return timing.includes("early") ? `Early · ${timing}` : `Late · ${timing}`;
  }
  return "On track";
}

function compactModelHasShortfall() {
  const model = buildGoalProjectionModel(18, 60);
  return model.goalState.some((state, idx) => {
    const month = Math.min(model.MONTHS - 1, Math.max(0, state.deadlineMo ?? 17));
    const forecastGap = Math.max(0, state.effectiveTarget - model.progressDollars[idx][month]);
    return forecastGap >= 0.01 || compactNearTermCashGap(idx) >= 0.01;
  });
}

function compactGoalDaysUntilDeadline(idx) {
  const goal = goalsData[idx];
  const deadline = _parseGoalDateValue(goal?.endDate);
  if (!deadline) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  deadline.setHours(0, 0, 0, 0);
  return Math.ceil((deadline - today) / 86400000);
}

function compactGoalNeedsCashNow(idx) {
  const days = compactGoalDaysUntilDeadline(idx);
  return days !== null && days <= 31;
}

function compactRequiredAssignedToday(idx) {
  if (!compactGoalNeedsCashNow(idx)) return 0;
  return compactMoney(Math.max(0, compactGoalTargetWithBuffer(idx)));
}

function compactNearTermCashGap(idx) {
  const required = compactRequiredAssignedToday(idx);
  if (required <= 0) return 0;
  return compactMoney(Math.max(0, required - compactTotalAssignedForGoal(idx)));
}

function compactLockedSliderCash(idx) {
  const required = compactRequiredAssignedToday(idx);
  if (required <= 0) return 0;
  return compactMoney(Math.max(0, required - compactRecordedForGoal(idx)));
}

function compactMaxRemovableToday() {
  if (compactModelHasShortfall()) return { amount: 0, goalName: "" };
  const original = goalsData.map(goal => compactMoney(goal.manualSaved || 0));
  const order = getCompactFundingOrder().slice().reverse();
  let total = 0;
  let firstGoal = "";

  try {
    order.forEach(idx => {
      const current = compactMoney(goalsData[idx]?.manualSaved || 0);
      const locked = compactLockedSliderCash(idx);
      const flexible = compactMoney(Math.max(0, current - locked));
      if (flexible <= 0.005) return;

      let lo = 0;
      let hi = flexible;
      for (let i = 0; i < 18; i++) {
        const mid = compactMoney((lo + hi) / 2);
        goalsData[idx].manualSaved = compactMoney(current - mid);
        if (compactModelHasShortfall()) {
          hi = mid;
        } else {
          lo = mid;
        }
      }

      const removable = compactMoney(lo);
      goalsData[idx].manualSaved = compactMoney(current - removable);
      if (removable > 0.005) {
        total = compactMoney(total + removable);
        if (!firstGoal) firstGoal = goalsData[idx].name;
      }
    });
  } finally {
    goalsData.forEach((goal, idx) => { goal.manualSaved = original[idx]; });
  }

  return { amount: total, goalName: firstGoal || "the lowest-priority funded goal" };
}

function compactCashNeededToday() {
  const nearTermNeed = compactMoney(goalsData.reduce((sum, _goal, idx) => sum + compactNearTermCashGap(idx), 0));
  if (!compactModelHasShortfall()) return nearTermNeed;
  const original = goalsData.map(goal => compactMoney(goal.manualSaved || 0));
  const order = getCompactFundingOrder();
  let needed = nearTermNeed;

  try {
    order.forEach(idx => {
      const gap = compactNearTermCashGap(idx);
      if (gap <= 0.005) return;
      goalsData[idx].manualSaved = compactMoney(goalsData[idx].manualSaved + gap);
    });

    for (let pass = 0; pass < 2 && compactModelHasShortfall(); pass++) {
      order.forEach(idx => {
        if (!compactModelHasShortfall()) return;
        const limit = compactMoney(compactGoalNeed(idx) - Number(goalsData[idx]?.manualSaved || 0));
        if (limit <= 0.005) return;

        let lo = 0;
        let hi = limit;
        for (let i = 0; i < 18; i++) {
          const mid = compactMoney((lo + hi) / 2);
          goalsData[idx].manualSaved = compactMoney(original[idx] + mid);
          if (compactModelHasShortfall()) {
            lo = mid;
          } else {
            hi = mid;
          }
        }

        const add = compactMoney(hi);
        goalsData[idx].manualSaved = compactMoney(original[idx] + add);
        needed = compactMoney(needed + add);
      });
    }
  } finally {
    goalsData.forEach((goal, idx) => { goal.manualSaved = original[idx]; });
  }

  return needed;
}

function compactTimelineMonthLabel(offset) {
  return new Date(new Date().getFullYear(), new Date().getMonth() + offset, 1)
    .toLocaleDateString("en-SG", { month: "short", year: "2-digit" })
    .replace(" ", "-");
}

function compactTimelineRows(rows, model) {
  const lastUsedMonth = Math.max(0, ...rows.map((row, idx) => {
    const series = [model.allocationBaseData[idx] || [], model.allocationCashflowData[idx] || []];
    const last = series.reduce((max, values) => Math.max(max, values.reduce((m, value, month) => Number(value || 0) > 0.005 ? month : m, -1)), -1);
    return Math.max(last, Math.min(model.MONTHS - 1, Math.max(0, model.goalState[idx]?.deadlineMo ?? 0)));
  }));
  const monthCount = Math.min(model.MONTHS, Math.max(6, Math.min(24, lastUsedMonth + 2)));
  const months = Array.from({ length: monthCount }, (_, idx) => ({ idx, label: compactTimelineMonthLabel(idx) }));

  return { months, rows: rows.map((row, idx) => ({
    ...row,
    months: months.map(month => ({
      month: month.idx,
      assigned: month.idx === 0 ? row.now : 0,
      forecast: compactMoney(model.allocationBaseData[idx]?.[month.idx] || 0),
      adjustments: compactMoney(model.allocationCashflowData[idx]?.[month.idx] || 0),
    }))
  })) };
}

function compactTimelineBlock(goalName, monthLabel, label, amount, cls, maxAmount) {
  if (amount <= 0.005) return "";
  const title = `${goalName} · ${monthLabel} · ${label}: ${formatCurrency(amount)}`;
  const width = Math.max(36, Math.min(100, (amount / Math.max(1, maxAmount)) * 100));
  return `<span class="cg-timeline-block ${cls}" style="width:${width}%" title="${escapeHtml(title)}"><b>${formatCurrency(amount)}</b></span>`;
}

function renderCompactTimeline(rows, model) {
  const target = document.getElementById("compactGoalTimeline");
  if (!target) return;
  const timeline = compactTimelineRows(rows, model);
  const template = `minmax(130px, 180px) repeat(${timeline.months.length}, minmax(92px, 1fr))`;
  const maxBlock = Math.max(1, ...timeline.rows.flatMap(row => row.months.flatMap(month => [month.assigned, month.forecast, month.adjustments])));
  target.innerHTML = `
    <div class="cg-timeline-grid" style="grid-template-columns:${template}">
      <div class="cg-timeline-corner">Goal</div>
      ${timeline.months.map(month => `<div class="cg-timeline-head">${escapeHtml(month.label)}</div>`).join("")}
      ${timeline.rows.map(row => `
        <div class="cg-timeline-goal">${escapeHtml(row.name)}<small>${escapeHtml(row.date)}</small></div>
        ${row.months.map(month => {
          const label = timeline.months[month.month]?.label || compactTimelineMonthLabel(month.month);
          return `<div class="cg-timeline-cell">
            ${compactTimelineBlock(row.name, label, "Assigned", month.assigned, "assigned", maxBlock)}
            ${compactTimelineBlock(row.name, label, "Forecast", month.forecast, "forecast", maxBlock)}
            ${compactTimelineBlock(row.name, label, "Adjustments", month.adjustments, "adjustments", maxBlock)}
          </div>`;
        }).join("")}
      `).join("")}
    </div>
    <div class="cg-timeline-legend"><span><i class="assigned"></i>Assigned</span><span><i class="forecast"></i>Forecast</span><span><i class="adjustments"></i>Adjustments</span></div>`;
}

function refreshCompactGoalForecast() {
  const timeline = document.getElementById("compactGoalTimeline");
  if (!timeline) return;
  updateCompactSaveStatus();
  const dep = computeDeployableBalance();
  const assigned = goalsData.reduce((sum, goal) => sum + Number(goal.manualSaved || 0), 0);
  const free = dep.deployable - assigned;
  document.getElementById("compactGoalSummary").innerHTML = `<div><small>Available for goals</small><strong>${formatCurrency(dep.deployable)}</strong></div><div><small>Cash assigned with sliders</small><strong>${formatCurrency(assigned)}</strong></div><div><small>${free < 0 ? "Over-assigned" : "Unassigned cash today"}</small><strong class="${free < 0 ? "red" : ""}">${formatCurrency(Math.abs(free))}</strong></div><div><small>12-mo normal savings / month</small><strong>${formatCurrency(Math.max(0, historicalStats.avgMonthlySavings))}</strong></div>`;
  const assignTotal = document.getElementById("compactAssignTotal");
  if (assignTotal) assignTotal.innerHTML = `<span>${free < 0 ? "Reduce slider cash by" : "Unassigned cash today"}</span><strong class="${free < 0 ? "red" : ""}">${formatCurrency(free < 0 ? Math.abs(free) : free)}</strong>`;
  const usableBreakdown = document.getElementById("compactUsableBreakdown");
  if (usableBreakdown) usableBreakdown.innerHTML = compactGoalUsableBreakdown(dep);
  const model = buildGoalProjectionModel(18, 60);
  const rows = model.goalState.map((state, idx) => {
    const month = Math.min(model.MONTHS - 1, Math.max(0, state.deadlineMo ?? 17));
    const sum = values => values.slice(0, month + 1).reduce((total, value) => total + value, 0);
    const tx = compactRecordedForGoal(idx);
    const extra = compactMoney(Math.max(0, goalsData[idx].manualSaved || 0));
    const now = compactMoney(Math.min(state.effectiveTarget, tx + extra));
    const base = sum(model.allocationBaseData[idx]);
    const adjustments = sum(model.allocationCashflowData[idx]);
    const projected = model.progressDollars[idx][month];
    const gap = Math.max(0, state.effectiveTarget - projected);
    const capped = state.deadlineMo !== null && state.deadlineMo >= model.MONTHS;
    const status = free < -0.005 ? "Over-assigned - rebalance" : compactTimingStatus(state, model, gap, capped);
    const date = state.deadlineMo === null ? model.fullLabels[month] + " (no deadline)" : formatDateDisplay(goalsData[idx].endDate);
    document.getElementById(`cgStatus_${idx}`).textContent = `${status} - ${formatCurrency(projected)} / ${formatCurrency(state.effectiveTarget)} by ${date}`;
    document.getElementById(`cgStatus_${idx}`).classList.toggle("red", gap >= 0.01 || capped || free < -0.005);
    const pill = document.getElementById(`cgPill_${idx}`);
    if (pill) {
      pill.textContent = gap < 0.01 && free >= -0.005 && !capped ? (status.startsWith("Early") ? "Early" : status === "Funded now" ? "Funded" : status === "On time" ? "On time" : "On track") : gap > 0 ? `${formatCurrency(gap)} short` : "Review";
      pill.classList.toggle("red", gap >= 0.01 || capped || free < -0.005);
    }
    const bar = document.getElementById(`cgBar_${idx}`);
    if (bar) bar.style.width = Math.min(100, (projected / Math.max(1, state.effectiveTarget)) * 100) + "%";
    const slider = document.getElementById(`cgSlider_${idx}`);
    slider.min = tx;
    slider.max = compactSliderLimit(idx);
    slider.value = now;
    const amount = document.getElementById(`cgAmount_${idx}`);
    if (amount) amount.min = tx;
    if (amount && document.activeElement !== amount) amount.value = now.toFixed(2);
    return { name: state.name, date, cashAssigned: extra, now, base, adjustments, projected, target: state.effectiveTarget, gap, status };
  });
  const removable = compactMaxRemovableToday();
  const unassignedCash = compactMoney(Math.max(0, free));
  document.getElementById("compactGoalSummary").innerHTML = `<div><small>Available for goals</small><strong>${formatCurrency(dep.deployable)}</strong></div><div><small>Cash assigned with sliders</small><strong>${formatCurrency(assigned)}</strong></div><div><small>${free < -0.005 ? "Over-assigned" : "Unassigned cash today"}</small><strong class="${free < -0.005 ? "red" : ""}">${formatCurrency(free < -0.005 ? Math.abs(free) : unassignedCash)}</strong></div><div><small>12-mo normal savings / month</small><strong>${formatCurrency(Math.max(0, historicalStats.avgMonthlySavings))}</strong></div>`;
  if (assignTotal) assignTotal.innerHTML = `<span>${free < -0.005 ? "Reduce slider cash by" : "Unassigned cash today"}</span><strong class="${free < -0.005 ? "red" : ""}">${formatCurrency(free < -0.005 ? Math.abs(free) : unassignedCash)}</strong>${removable.amount > 0.005 && free >= -0.005 ? `<small>Forecast-replaceable slider cash: ${formatCurrency(removable.amount)}</small>` : ""}`;
  const overall = compactOverallStatus(rows, free, removable);
  const overallHtml = `<div class="cg-overall ${overall.tone}"><strong>${escapeHtml(overall.title)}</strong><span>${escapeHtml(overall.detail)}</span></div>`;
  document.getElementById("compactForecastResults").innerHTML = `${overallHtml}<div class="cg-table-wrap"><table class="cg-breakdown"><thead><tr><th>Goal / deadline</th><th>Assigned</th><th>Cash from sliders</th><th>Forecast</th><th>Adjustments</th><th>Forecast / target</th></tr></thead><tbody>${rows.map(row => `<tr><th>${escapeHtml(row.name)}<small>${escapeHtml(row.date)} · ${escapeHtml(row.status)}</small></th><td>${formatCurrency(row.now)}</td><td>${formatCurrency(row.cashAssigned)}</td><td>${formatCurrency(row.base)}</td><td>${formatCurrency(row.adjustments)}</td><td>${formatCurrency(row.projected)} / ${formatCurrency(row.target)}</td></tr>`).join("")}</tbody></table></div>`;
  document.getElementById("compactForecastMethod").innerHTML = `<p><strong>Money available today:</strong> ${formatCurrency(dep.rawSavings)} selected savings − ${formatCurrency(ccOwed)} personal card debt + ${formatCurrency(dep.claimReceivableForGoals)} pending reimbursements − ${formatCurrency(dep.remainingBudget)} remaining budget − ${formatCurrency(dep.futureSalaryHold)} future salary budget hold = ${formatCurrency(dep.deployable)}.</p><p><strong>Forecast monthly savings:</strong> ${formatCurrency(historicalStats.avgMonthlyIncome)} recurring income − ${formatCurrency(historicalStats.avgMonthlyExpenses)} average expenses = ${formatCurrency(historicalStats.avgMonthlySavings)}; the forecast uses at least $0. Based on the last ${historicalStats.months} completed months, with salary spikes such as bonus income excluded.</p><p><strong>Forecast = assigned now + forecast allocated + positive adjustments allocated.</strong> Assigned now includes progress already recorded to the goal plus any extra amount you add with the slider. Reductions lower the monthly pool before allocation. The same pool is shared across all goals, never counted in full for each one. Contributions begin next month or the goal’s start month, whichever is later. Forced goals go first; other goals follow priority and deadline, with base targets before buffers. No interest or investment return is assumed.</p><p>Pending reimbursements are not cash yet. ${free < 0 ? "Current extra assignments exceed the available pool; reduce a slider or use Smart Assign before relying on this forecast." : "Forecast figures are estimates, not guaranteed savings."}</p>`;
  renderCompactTimeline(rows, model);
}

// Keep existing persistence, account selection, and adjustment handlers on one renderer.
renderGoalsPage = renderCompactGoalsPage;
refreshGoalInsightPanels = refreshCompactGoalForecast;
// All goal edits are drafts. No network writes or blocking autosave while exploring.
scheduleGoalsAutoSave = function() {
  clearTimeout(goalsAutoSaveTimer);
  updateCompactSaveStatus();
};
_saveIncomeBoostsFromPanel = saveCompactPlan;
const redrawCompactAdjustments = _redrawBoostsPanel;
_redrawBoostsPanel = function(panel) {
  redrawCompactAdjustments(panel);
  const button = panel?.querySelector('[onclick="_saveIncomeBoostsFromPanel()"]');
  if (button) button.textContent = "Save plan";
};
window.addEventListener("beforeunload", event => {
  if (!compactPlanChanged()) return;
  event.preventDefault();
  event.returnValue = "";
});