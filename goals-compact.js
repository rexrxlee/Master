let compactGoalChart = null;
let compactForecastFrame = null;
let compactGoalSort = "original";

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

function deleteCompactGoal(idx) {
  goalsData.splice(idx, 1);
  renderCompactGoalsPage();
  updateCompactSaveStatus();
}

function compactPriorityOptions(value) {
  return ["Critical", "High", "Medium", "Low"].map(name => `<option ${name === value ? "selected" : ""}>${name}</option>`).join("");
}

function renderCompactGoalsPage() {
  if (!compactSavedPlan) captureCompactSavedPlan();
  const container = document.getElementById("goalsContainer");
  const adjustmentsOpen = document.getElementById("boostsPanel")?.open || incomeBoostsDirty;
  if (compactGoalChart) { compactGoalChart.destroy(); compactGoalChart = null; }
  const today = toDateInputValue(new Date());
  container.innerHTML = `
    <section class="cg-summary" id="compactGoalSummary" aria-live="polite"></section>
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
    <div class="cg-workspace">
      <section class="cg-controls">
        <div class="cg-heading"><h2>Assign money</h2><button class="btn-primary" onclick="compactSmartAssign()" ${goalsData.length ? "" : "disabled"}>Smart Assign</button></div>
        <label class="cg-sort" for="compactGoalSort">Sort by <select id="compactGoalSort" onchange="sortCompactGoalRows(this.value)"><option value="original" ${compactGoalSort === "original" ? "selected" : ""}>Original order</option><option value="priority" ${compactGoalSort === "priority" ? "selected" : ""}>Priority (highest first)</option></select></label>
        <p class="cg-note">Try allocations freely. Sliders and Smart Assign only preview your plan; choose Save plan when ready.</p>
        <div id="compactGoalRows">${goalsData.map((goal, idx) => compactGoalRow(goal, idx)).join("") || '<p class="cg-note">Add your first goal to start planning.</p>'}</div>
        <div class="cg-save-actions"><button id="compactSavePlan" class="btn-primary" onclick="saveCompactPlan()">Save plan</button><button id="compactResetPlan" class="btn-secondary" onclick="resetCompactPlan()">Reset changes</button><span id="goalsAutosaveStatus" class="goals-autosave-status" role="status">No unsaved changes</span></div>
      </section>
      <section class="cg-forecast">
        <div class="cg-heading"><h2>Can I reach my goals?</h2></div>
        <p class="cg-note">Forecast versus target at each goal’s deadline, including its buffer. Month-end estimates; future savings start next month.</p>
        <div class="cg-chart-scroll"><div class="cg-chart-wrap"><canvas id="compactGoalChart" role="img" aria-label="Goal forecasts compared with targets"></canvas></div></div>
        <div id="compactForecastResults" aria-live="polite"></div>
        <details class="cg-method"><summary>How this forecast is calculated</summary><div id="compactForecastMethod"></div></details>
      </section>
    </div>
    <div id="compactAdjustments"></div>
    <details class="simple-details"><summary>Which accounts fund my goals?</summary><div class="simple-details-body cg-accounts">
      ${allAccounts.filter(account => account.type === "Savings").map(account => `<label><input type="checkbox" value="${escapeHtml(account.name)}" ${goalSavingsAccts.includes(account.name) ? "checked" : ""} onchange="toggleGoalAccount(this)">${escapeHtml(account.name)} <strong>${formatCurrency(savingsBalances[account.name] || 0)}</strong></label>`).join("") || '<p>Add a savings account in Accounts & Setup first.</p>'}
    </div></details>`;
  renderIncomeBoostsPanel(document.getElementById("compactAdjustments"));
  document.getElementById("boostsPanel").open = adjustmentsOpen;
  sortCompactGoalRows(compactGoalSort);
  refreshCompactGoalForecast();
  updateCompactSaveStatus();
}

function compactGoalRow(goal, idx) {
  return `<article class="cg-goal" id="compactGoal_${idx}">
    <div class="cg-heading"><strong>${escapeHtml(goal.name)}</strong><select aria-label="Priority for ${escapeHtml(goal.name)}" onchange="updateCompactGoal(${idx}, 'urgency', this.value)">${compactPriorityOptions(goal.urgency)}</select></div>
    <div class="cg-dates"><label>Start<input type="date" value="${escapeHtml(goal.startDate || "")}" onchange="updateCompactGoal(${idx}, 'startDate', this.value)"></label><label>Deadline<input type="date" value="${escapeHtml(goal.endDate || "")}" onchange="updateCompactGoal(${idx}, 'endDate', this.value)"></label></div>
    <div class="cg-allocation"><input id="cgSlider_${idx}" type="range" min="0" max="1" step="1" value="${goal.manualSaved}" aria-label="Money assigned to ${escapeHtml(goal.name)}" oninput="assignCompactGoal(${idx}, this.value)"><label>Assigned ($)<input id="cgAmount_${idx}" type="number" min="0" step="0.01" value="${goal.manualSaved}" oninput="assignCompactGoal(${idx}, this.value)"></label></div>
    <p class="cg-result" id="cgStatus_${idx}"></p>
    <details class="cg-method"><summary>Edit target & options</summary><div class="cg-options">
      <label>Name<input value="${escapeHtml(goal.name)}" onchange="updateCompactGoal(${idx}, 'name', this.value)"></label>
      <label>Target ($)<input type="number" min="0.01" step="0.01" value="${goal.target}" onchange="updateCompactGoal(${idx}, 'target', this.value)"></label>
      <label>Monthly plan ($)<input type="number" min="0" step="0.01" value="${goal.monthlyAlloc}" onchange="updateCompactGoal(${idx}, 'monthlyAlloc', this.value)"></label>
      <label>Buffer (%)<input type="number" min="0" max="100" value="${goal.goalBuffer || 0}" onchange="updateCompactGoal(${idx}, 'goalBuffer', this.value)"></label>
      <label><input type="checkbox" ${goal.priority ? "checked" : ""} onchange="updateCompactGoal(${idx}, 'priority', this.checked ? 1 : 0)">Fund before other priorities</label>
      <label>Notes<input value="${escapeHtml(goal.notes || "")}" onchange="updateCompactGoal(${idx}, 'notes', this.value)"></label>
      <a href="add-transaction.html?goal=${encodeURIComponent(goal.name)}">Record goal expense</a>
      <button class="btn-secondary" onclick="deleteCompactGoal(${idx})">Delete goal</button>
    </div></details>
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

function compactAssignableLimit(idx) {
  const others = goalsData.reduce((sum, goal, i) => sum + (i === idx ? 0 : Number(goal.manualSaved || 0)), 0);
  const goal = goalsData[idx];
  const need = Math.max(0, goal.target * (1 + (goal.goalBuffer || 0) / 100) - getSavedViaTransactions(goal.name));
  return Math.max(0, Math.min(need, computeDeployableBalance().deployable - others));
}

function assignCompactGoal(idx, value) {
  if (value === "") return;
  const amount = Number(value);
  if (!Number.isFinite(amount)) return;
  goalsData[idx].manualSaved = Math.round(Math.max(0, Math.min(amount, compactAssignableLimit(idx))) * 100) / 100;
  document.getElementById(`cgAmount_${idx}`).value = goalsData[idx].manualSaved;
  document.getElementById(`cgSlider_${idx}`).value = goalsData[idx].manualSaved;
  cancelAnimationFrame(compactForecastFrame);
  compactForecastFrame = requestAnimationFrame(refreshCompactGoalForecast);
  scheduleGoalsAutoSave();
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
  const order = getPrioritizedGoalIndexes();
  goalsData.forEach(goal => { goal.manualSaved = 0; });
  const assign = (indexes, buffer) => indexes.forEach(idx => {
    const goal = goalsData[idx];
    const target = goal.target * (1 + (buffer ? goal.goalBuffer || 0 : 0) / 100);
    const need = Math.max(0, target - getSavedViaTransactions(goal.name) - goal.manualSaved);
    const amount = Math.min(available, need);
    goal.manualSaved += amount;
    available = Math.max(0, available - amount);
  });
  const forced = order.filter(idx => isForcePriorityGoal(goalsData[idx]));
  const regular = order.filter(idx => !isForcePriorityGoal(goalsData[idx]));
  assign(forced, false); assign(forced, true); assign(regular, false); assign(regular, true);
  renderCompactGoalsPage();
  scheduleGoalsAutoSave(150);
}

function refreshCompactGoalForecast() {
  const canvas = document.getElementById("compactGoalChart");
  if (!canvas) return;
  updateCompactSaveStatus();
  const dep = computeDeployableBalance();
  const assigned = goalsData.reduce((sum, goal) => sum + Number(goal.manualSaved || 0), 0);
  const free = dep.deployable - assigned;
  document.getElementById("compactGoalSummary").innerHTML = `<div><small>Available for goals</small><strong>${formatCurrency(dep.deployable)}</strong></div><div><small>Assigned now</small><strong>${formatCurrency(assigned)}</strong></div><div><small>${free < 0 ? "Over-assigned" : "Unassigned"}</small><strong class="${free < 0 ? "red" : ""}">${formatCurrency(Math.abs(free))}</strong></div><div><small>Forecast savings / month</small><strong>${formatCurrency(Math.max(0, historicalStats.avgMonthlySavings))}</strong></div>`;
  const model = buildGoalProjectionModel(18, 240);
  const rows = model.goalState.map((state, idx) => {
    const month = Math.min(model.MONTHS - 1, Math.max(0, state.deadlineMo ?? 17));
    const sum = values => values.slice(0, month + 1).reduce((total, value) => total + value, 0);
    const tx = Math.min(state.effectiveTarget, Math.max(0, getSavedViaTransactions(state.name)));
    const now = Math.min(Math.max(0, state.effectiveTarget - tx), goalsData[idx].manualSaved);
    const base = sum(model.allocationBaseData[idx]);
    const adjustments = sum(model.allocationCashflowData[idx]);
    const projected = model.progressDollars[idx][month];
    const gap = Math.max(0, state.effectiveTarget - projected);
    const capped = state.deadlineMo !== null && state.deadlineMo >= model.MONTHS;
    const status = free < -0.005 ? "Over-assigned — rebalance"
      : capped ? "Beyond forecast range"
      : state.deadlineMo < 0 ? (gap < 0.01 ? "Target covered" : `Overdue · ${formatCurrency(gap)} short`)
      : gap < 0.01 ? "On track" : `${formatCurrency(gap)} short`;
    const date = state.deadlineMo === null ? model.fullLabels[month] + " (no deadline)" : formatDateDisplay(goalsData[idx].endDate);
    document.getElementById(`cgStatus_${idx}`).textContent = `${status} · ${formatCurrency(projected)} / ${formatCurrency(state.effectiveTarget)} by ${date}`;
    document.getElementById(`cgStatus_${idx}`).classList.toggle("red", gap >= 0.01 || capped || free < -0.005);
    const slider = document.getElementById(`cgSlider_${idx}`);
    slider.max = Math.max(compactAssignableLimit(idx), goalsData[idx].manualSaved, 1);
    slider.value = goalsData[idx].manualSaved;
    return { name: state.name, date, tx, now, base, adjustments, projected, target: state.effectiveTarget, gap, status };
  });
  document.getElementById("compactForecastResults").innerHTML = `<div class="cg-table-wrap"><table class="cg-breakdown"><thead><tr><th>Goal / deadline</th><th>Assigned</th><th>Recorded progress</th><th>Future savings</th><th>Adjustments</th><th>Forecast / target</th></tr></thead><tbody>${rows.map(row => `<tr><th>${escapeHtml(row.name)}<small>${escapeHtml(row.date)} · ${escapeHtml(row.status)}</small></th><td>${formatCurrency(row.now)}</td><td>${formatCurrency(row.tx)}</td><td>${formatCurrency(row.base)}</td><td>${formatCurrency(row.adjustments)}</td><td>${formatCurrency(row.projected)} / ${formatCurrency(row.target)}</td></tr>`).join("")}</tbody></table></div>`;
  document.getElementById("compactForecastMethod").innerHTML = `<p><strong>Money available today:</strong> ${formatCurrency(dep.rawSavings)} selected savings − ${formatCurrency(ccOwed)} personal card debt + ${formatCurrency(dep.claimReceivableForGoals)} pending reimbursements − ${formatCurrency(dep.remainingBudget)} remaining budget − ${formatCurrency(dep.futureSalaryHold)} future salary held = ${formatCurrency(dep.deployable)}.</p><p><strong>Future monthly savings:</strong> ${formatCurrency(historicalStats.avgMonthlyIncome)} recurring income − ${formatCurrency(historicalStats.avgMonthlyExpenses)} average expenses = ${formatCurrency(historicalStats.avgMonthlySavings)}; the forecast uses at least $0. Based on ${historicalStats.months} completed historical months, with salary spikes excluded.</p><p><strong>Forecast = assigned now + recorded goal progress + future savings allocated + positive adjustments allocated.</strong> Reductions lower the monthly pool before allocation. The same pool is shared across all goals, never counted in full for each one. Contributions begin next month or the goal’s start month, whichever is later. Forced goals go first; other goals follow priority and deadline, with base targets before buffers. No interest or investment return is assumed.</p><p>Pending reimbursements are not cash yet. ${free < 0 ? "Current assignments exceed the available pool; reduce a slider or use Smart Assign before relying on this forecast." : "Future figures are estimates, not guaranteed savings."}</p>`;
  const datasets = [
    { label: "Assigned now", data: rows.map(row => row.now), backgroundColor: "#2563eb", stack: "forecast" },
    { label: "Recorded progress", data: rows.map(row => row.tx), backgroundColor: "#0f766e", stack: "forecast" },
    { label: "Future savings", data: rows.map(row => row.base), backgroundColor: "#93c5fd", stack: "forecast" },
    { label: "Adjustments", data: rows.map(row => row.adjustments), backgroundColor: "#a78bfa", stack: "forecast" },
    { label: "Target", data: rows.map(row => row.target), backgroundColor: "#e2e8f0", borderColor: "#64748b", borderWidth: 1, stack: "target" }
  ];
  canvas.parentElement.style.height = Math.max(200, rows.length * 75) + "px";
  if (compactGoalChart) {
    compactGoalChart.data.labels = rows.map(row => row.name);
    compactGoalChart.data.datasets = datasets;
    compactGoalChart.update("none");
  } else if (typeof Chart !== "undefined") {
    compactGoalChart = new Chart(canvas, { type: "bar", data: { labels: rows.map(row => row.name), datasets }, options: {
      indexAxis: "y", responsive: true, maintainAspectRatio: false, animation: false,
      scales: { x: { stacked: true, beginAtZero: true, ticks: { callback: value => formatCurrencyShort(value) } }, y: { stacked: true } },
      plugins: { datalabels: { display: false }, legend: { position: "bottom" }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${formatCurrency(ctx.raw)}` } } }
    } });
  }
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
