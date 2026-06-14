// ═══════════════════════════════════════════════════════════════════
//  GOALS PAGE — Fintrack  (goals.js)
//
//  Excel storage: Budget Setup, columns S onwards (S2:AC20)
//  Col S=name, T=target, U=manualSaved, V=monthlyAlloc,
//      W=startDate, X=endDate, Y=urgency, Z=notes, AA=color,
//      AB=goalBuffer (extra $ above target as personal buffer),
//      AC=priority (was AB before)
//
//  "Goal Savings Accounts" stored in Budget Setup col AD2:AD2
//  (pipe-separated list of account names selected as savings-eligible)
// ═══════════════════════════════════════════════════════════════════

const GOALS_RANGE       = "S2:AC20";  // 11 cols × 19 rows
const GOAL_ACCTS_RANGE  = "AD2:AD2";  // single cell, pipe-separated
const MAX_GOALS         = 19;

// Register Chart.js plugins if available
if (typeof Chart !== "undefined" && typeof ChartDataLabels !== "undefined") {
  Chart.register(ChartDataLabels);
}

let goalsData       = [];
let allTxForGoals   = [];
let allAccounts     = [];          // { name, type } from col J/K
let goalSavingsAccts = [];         // names of accounts chosen for goals
let historicalStats = { avgMonthlyIncome:0, avgMonthlyExpenses:0, avgMonthlySavings:0, months:0 };
let budgetSummary   = { billsTotal:0, monthlyTotal:0, billsRows:[], monthlyRows:[] };
let ccOwed          = 0;           // total credit-card balance owed
let savingsBalances = {};          // { accountName: balance } from transactions
let insuranceRenewalsForGoals = [];
let goalsAutoSaveTimer = null;
let goalsAutoSaveInFlight = false;

// ─── Entry Point ──────────────────────────────────────────────────

async function loadGoalsPage() {
  try {
    clearOutput();
    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();
    const workbook    = XLSX.read(arrayBuffer, { type:"array" });
    const budgetSheet = workbook.Sheets["Budget Setup"];
    const txSheet     = workbook.Sheets[CONFIG.sheetName];
    const insuranceSheet = workbook.Sheets["Insurance"];
    if (!budgetSheet) throw new Error("Sheet not found: Budget Setup");
    if (!txSheet)     throw new Error("Sheet not found: " + CONFIG.sheetName);

    // Read accounts
    const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header:1, blankrows:false });
    allAccounts = allRows.slice(1)
      .map(row => ({ name:String(row[9]??"").trim(), type:String(row[10]??"Savings").trim() }))
      .filter(a => a.name !== "");

    // Read which accounts are designated for goals
    const acctCell = budgetSheet["AD2"];
    const acctRaw  = acctCell ? String(acctCell.v ?? "").trim() : "";
    goalSavingsAccts = acctRaw ? acctRaw.split("|").map(s=>s.trim()).filter(Boolean) : [];

    // Budget totals
    const billsRows   = readBudgetSection(budgetSheet, "A2:B13");
    const monthlyRows = readBudgetSection(budgetSheet, "F2:G13");
    budgetSummary.billsTotal   = billsRows.reduce((s,r)=>s+r.allocated,0);
    budgetSummary.monthlyTotal = monthlyRows.reduce((s,r)=>s+r.allocated,0);
    budgetSummary.billsRows    = billsRows;
    budgetSummary.monthlyRows  = monthlyRows;

    // Transactions
    allTxForGoals   = readAllTx(txSheet);
    historicalStats = computeHistoricalStats(allTxForGoals);
    savingsBalances = computeAccountBalances(allTxForGoals);
    ccOwed          = computeCCOwed(allTxForGoals);
    const pendingClaimsForGoals = computePendingClaimSummary(allTxForGoals);
    log(`Pending claims detected: ${pendingClaimsForGoals.count} row(s), ${formatCurrency(pendingClaimsForGoals.total)}.`);

    // Goals
    goalsData = readGoalsFromSheet(budgetSheet);
    insuranceRenewalsForGoals = insuranceSheet ? readInsuranceRenewalsForGoals(insuranceSheet) : [];
    loadIncomeBoosts(budgetSheet);

    renderGoalsPage();
    log("Goals loaded.");
  } catch(err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

// ─── Date Utilities (needed by data readers) ──────────────────────

function _parseGoalDateValue(value) {
  if (!value && value !== 0) return null;
  if (typeof value === "number" && value > 1000) {
    try { const d = XLSX.SSF.parse_date_code(value); return new Date(d.y, d.m-1, d.d); } catch { return null; }
  }
  const s = String(value).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3]);
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return new Date(+ddmm[3], +ddmm[2]-1, +ddmm[1]);
  return null;
}

function _goalDateToInputValue(value) {
  const d = _parseGoalDateValue(value);
  if (!d) return "";
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

// ─── Data Readers ─────────────────────────────────────────────────

function readAllTx(sheet) {
  const rows    = XLSX.utils.sheet_to_json(sheet, { header:1, blankrows:false });
  const headers = rows[0].map(h => String(h??"").trim());
  const canonicalHeaders = [
    "Date",
    "Transaction",
    "Amount",
    "Main Category",
    "Sub Category",
    "Account",
    "Claimable",
    "Claim Status",
    "Claim Amount",
    "Claim Account"
  ];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => {
      if (h) obj[h] = row[i];
    });
    canonicalHeaders.forEach((h,i) => {
      if (obj[h] === undefined && row[i] !== undefined) obj[h] = row[i];
    });
    return obj;
  });
}

function readBudgetSection(sheet, range) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, range, blankrows:false });
  return rows.map(row => ({ category:String(row[0]??"").trim(), allocated:Number(row[1]??0)||0 }))
             .filter(r => r.category !== "");
}

function readGoalsFromSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, range:GOALS_RANGE, blankrows:false });
  return rows.map(row => ({
    name:        String(row[0] ?? "").trim(),
    target:      Number(row[1] ?? 0)  || 0,
    manualSaved: Number(row[2] ?? 0)  || 0,
    monthlyAlloc:Number(row[3] ?? 0)  || 0,
    startDate:   _goalDateToInputValue(row[4]),  // normalize to YYYY-MM-DD
    endDate:     _goalDateToInputValue(row[5]),   // normalize to YYYY-MM-DD
    urgency:     String(row[6] ?? "Medium").trim() || "Medium",
    notes:       String(row[7] ?? "").trim(),
    color:       String(row[8] ?? "#4caf50").trim() || "#4caf50",
    goalBuffer:  Number(row[9] ?? 0)  || 0,   // extra % buffer above target (0-100)
    priority:    Number(row[10] ?? 0) || 0,
  })).filter(g => g.name !== "");
}

function readInsuranceRenewalsForGoals(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, blankrows:false });
  if (rows.length === 0) return [];

  const headerMap = {};
  (rows[0] || []).forEach((header, index) => {
    const key = normaliseInsuranceGoalHeader(header);
    if (key) headerMap[key] = index;
  });

  if (getInsuranceGoalHeaderIndex(headerMap, "Owner") < 0 ||
      getInsuranceGoalHeaderIndex(headerMap, "Policy Name") < 0 ||
      getInsuranceGoalHeaderIndex(headerMap, "Annual Premium") < 0) {
    return [];
  }

  const currentYear = new Date().getFullYear();
  return rows.slice(1)
    .map(row => {
      const policyName = getInsuranceGoalCell(row, headerMap, "Policy Name");
      const owner = getInsuranceGoalCell(row, headerMap, "Owner");
      const policyNo = getInsuranceGoalCell(row, headerMap, "Policy No");
      if (!owner && !policyName && !policyNo) return null;

      const status = getInsuranceGoalCell(row, headerMap, "Status") || "Active";
      if (status.toLowerCase() === "inactive") return null;

      const startDate = parseInsuranceGoalDate(getInsuranceGoalCell(row, headerMap, "Cover Start Date"));
      const renewalMonth = getInsuranceGoalCell(row, headerMap, "Renewal Month");
      let dueDate = computeInsuranceGoalDueDate(startDate, renewalMonth);
      let dueYear = dueDate ? dueDate.getFullYear() : currentYear;
      const annualPremium = parseInsuranceGoalAmount(getInsuranceGoalCell(row, headerMap, "Annual Premium"));
      let paidForDueYear = getInsuranceGoalYearAmount(row, headerMap, "Paid", dueYear, 0);

      while (dueDate && paidForDueYear > 0) {
        dueDate = new Date(dueDate.getFullYear() + 1, dueDate.getMonth(), dueDate.getDate());
        dueYear = dueDate.getFullYear();
        paidForDueYear = getInsuranceGoalYearAmount(row, headerMap, "Paid", dueYear, 0);
      }

      const renewalPremium = getInsuranceGoalYearAmount(row, headerMap, "Premium", dueYear, annualPremium);
      const premiumType = getInsuranceGoalCell(row, headerMap, "Premium Type");
      const isCash = premiumType.toLowerCase().includes("cash");
      const isCpf = premiumType.toLowerCase().includes("cpf") || premiumType.toLowerCase().includes("medisave");

      return {
        owner,
        policyName,
        policyNo,
        premiumType,
        renewalPremium,
        paidForDueYear,
        dueDate,
        dueYear,
        isCash,
        isCpf,
        goalName: getInsuranceGoalCell(row, headerMap, "Goal Name") ||
          (dueDate ? `Insurance ${dueDate.toLocaleString("en-SG", { month: "short", year: "numeric" })}` : `Insurance ${dueYear}`),
        comments: getInsuranceGoalCell(row, headerMap, "Comments"),
        claimsPossible: getInsuranceGoalCell(row, headerMap, "Claims Possible")
      };
    })
    .filter(Boolean)
    .filter(row => row.dueDate && row.renewalPremium > 0)
    .sort((a, b) => a.dueDate - b.dueDate || a.owner.localeCompare(b.owner));
}

function groupInsuranceRenewalsForGoals(rows) {
  const groups = {};
  rows.forEach(row => {
    const key = row.dueDate.getFullYear() + "-" + String(row.dueDate.getMonth() + 1).padStart(2, "0");
    if (!groups[key]) {
      groups[key] = {
        key,
        dueDate: new Date(row.dueDate.getFullYear(), row.dueDate.getMonth(), 1),
        rows: [],
        cashTotal: 0,
        cashPending: 0,
        cashPaid: 0,
        cpfTotal: 0,
        linkedGoalNames: new Set()
      };
    }

    groups[key].rows.push(row);
    if (row.isCash) {
      groups[key].cashTotal += row.renewalPremium;
      if (row.paidForDueYear > 0) groups[key].cashPaid += row.paidForDueYear;
      else groups[key].cashPending += row.renewalPremium;
    }
    if (row.isCpf) groups[key].cpfTotal += row.renewalPremium;
    if (row.goalName) groups[key].linkedGoalNames.add(row.goalName);
  });

  return Object.values(groups).sort((a, b) => a.dueDate - b.dueDate);
}

function renderInsuranceRenewalGoalPanel(container) {
  if (!insuranceRenewalsForGoals.length) return;

  const groups = groupInsuranceRenewalsForGoals(insuranceRenewalsForGoals);
  const totalCashPending = groups.reduce((sum, group) => sum + group.cashPending, 0);
  const totalCashPaid = groups.reduce((sum, group) => sum + group.cashPaid, 0);
  const totalCpf = groups.reduce((sum, group) => sum + group.cpfTotal, 0);

  const groupHtml = groups.map(group => {
    const linkedGoalNames = Array.from(group.linkedGoalNames).filter(Boolean);
    const linkedGoals = linkedGoalNames
      .map(name => goalsData.find(goal => clean(goal.name).toLowerCase() === clean(name).toLowerCase()))
      .filter(Boolean);
    const linkedSaved = linkedGoals.reduce((sum, goal) => sum + goal.manualSaved + getSavedViaTransactions(goal.name), 0);
    const linkedTarget = group.cashTotal;
    const pct = linkedTarget > 0 ? Math.min(100, linkedSaved / linkedTarget * 100) : 100;
    const goalLabel = linkedGoalNames.length
      ? linkedGoalNames.map(escapeHtml).join(", ")
      : `Insurance ${group.dueDate.getFullYear()}`;
    const matchText = linkedGoals.length
      ? `${formatCurrency(linkedSaved)} saved against ${formatCurrency(linkedTarget)} cash premium`
      : `No matching saved goal yet. Use Goal Name "${escapeHtml(goalLabel)}" in Insurance or create that goal.`;

    const policyHtml = group.rows.map(row => {
      const paid = row.paidForDueYear > 0;
      return `
        <div class="insurance-goal-policy">
          <div>
            <strong>${escapeHtml(row.owner)} - ${escapeHtml(row.policyName)}</strong>
            <span>${escapeHtml(row.premiumType)} · ${formatCurrency(row.renewalPremium)}</span>
          </div>
          <span class="insurance-goal-status ${paid ? "paid" : "pending"}">${paid ? "Paid" : "Pending"}</span>
        </div>
      `;
    }).join("");

    return `
      <div class="insurance-goal-group">
        <div class="insurance-goal-group-head">
          <div>
            <h3>${formatInsuranceGoalMonth(group.dueDate)}</h3>
            <p>Linked goal: <strong>${goalLabel}</strong></p>
          </div>
          <div class="insurance-goal-group-money">
            <strong>${formatCurrency(group.cashPending)}</strong>
            <span>cash pending</span>
          </div>
        </div>
        <div class="insurance-goal-progress">
          <div class="insurance-goal-track"><div class="insurance-goal-fill" style="width:${pct.toFixed(1)}%;"></div></div>
          <span>${matchText}</span>
        </div>
        <div class="insurance-goal-metrics">
          <span>Cash total: <strong>${formatCurrency(group.cashTotal)}</strong></span>
          <span>Cash paid: <strong>${formatCurrency(group.cashPaid)}</strong></span>
          <span>CPF: <strong>${formatCurrency(group.cpfTotal)}</strong></span>
        </div>
        <div class="insurance-goal-policies">${policyHtml}</div>
      </div>
    `;
  }).join("");

  container.innerHTML += `
    <div class="goals-panel insurance-goal-panel">
      <div class="panel-header">
        <span class="panel-icon">▥</span>
        <h2>Insurance Renewal Reserve</h2>
        <span class="panel-hint">Grouped by renewal date from the Insurance sheet. CPF premiums are shown but excluded from cash reserve pressure.</span>
      </div>
      <div class="insurance-goal-summary">
        <div><span>Cash pending</span><strong class="${totalCashPending > 0 ? "red" : "green"}">${formatCurrency(totalCashPending)}</strong></div>
        <div><span>Cash paid</span><strong class="green">${formatCurrency(totalCashPaid)}</strong></div>
        <div><span>CPF renewals</span><strong>${formatCurrency(totalCpf)}</strong></div>
        <a href="insurance.html">Open Insurance</a>
      </div>
      <div class="insurance-goal-groups">${groupHtml}</div>
    </div>
  `;
}

function getInsuranceGoalCell(row, headerMap, headerName) {
  const index = getInsuranceGoalHeaderIndex(headerMap, headerName);
  return index >= 0 ? clean(row[index]) : "";
}

function getInsuranceGoalHeaderIndex(headerMap, headerName) {
  const direct = headerMap[normaliseInsuranceGoalHeader(headerName)];
  if (Number.isInteger(direct)) return direct;

  const yearly = clean(headerName).match(/^(Premium|Paid|Remark)\s+(\d{4})$/i);
  if (yearly) {
    const shortYear = yearly[2].slice(-2);
    const shortKey = headerMap[normaliseInsuranceGoalHeader(yearly[1] + shortYear)];
    if (Number.isInteger(shortKey)) return shortKey;
  }

  return -1;
}

function normaliseInsuranceGoalHeader(value) {
  return clean(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function getInsuranceGoalYearAmount(row, headerMap, prefix, year, fallback = 0) {
  const value = parseInsuranceGoalAmount(getInsuranceGoalCell(row, headerMap, `${prefix} ${year}`));
  return value > 0 ? value : fallback;
}

function parseInsuranceGoalAmount(value) {
  const amount = Number(String(value ?? "").replace(/[$,\s]/g, ""));
  return Number.isFinite(amount) ? amount : 0;
}

function parseInsuranceGoalDate(value) {
  const parsed = _parseGoalDateValue(value);
  if (parsed) return parsed;

  const text = clean(value);
  const ddMmm = text.match(/^(\d{1,2})[-\s/]([A-Za-z]{3,})[-\s/](\d{2,4})$/);
  if (ddMmm) {
    const month = insuranceGoalMonthNumber(ddMmm[2]);
    const year = Number(ddMmm[3].length === 2 ? "20" + ddMmm[3] : ddMmm[3]);
    if (month) return new Date(year, month - 1, Number(ddMmm[1]));
  }

  return null;
}

function computeInsuranceGoalDueDate(startDate, renewalMonth) {
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  if (!startDate && !renewalMonth) return null;
  const monthNumber = insuranceGoalMonthNumber(renewalMonth);
  const baseMonth = monthNumber ? monthNumber - 1 : (startDate ? startDate.getMonth() : today.getMonth());
  const baseDay = startDate ? startDate.getDate() : 1;
  const firstYear = startDate ? startDate.getFullYear() + 1 : todayStart.getFullYear();
  let due = new Date(firstYear, baseMonth, Math.min(baseDay, 28));
  while (due < todayStart || (startDate && due < startDate)) {
    due = new Date(due.getFullYear() + 1, baseMonth, Math.min(baseDay, 28));
  }
  return due;
}

function insuranceGoalMonthNumber(value) {
  const text = clean(value);
  if (!text) return null;
  const numeric = Number(text);
  if (Number.isInteger(numeric) && numeric >= 1 && numeric <= 12) return numeric;
  const names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const index = names.findIndex(name => text.toLowerCase().startsWith(name));
  return index >= 0 ? index + 1 : null;
}

function formatInsuranceGoalMonth(date) {
  return date.toLocaleString("en-SG", { month: "short", year: "numeric" });
}

function getSignedAmount(value) {
  if (typeof value === "number") return value;
  const n = Number(String(value).replace(/\$/g, "").replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

function getAccountBalanceImpact(row) {
  const amount = Math.abs(getSignedAmount(row["Amount"]));
  const main = clean(row["Main Category"]).toLowerCase();
  const sub = clean(row["Sub Category"]).toLowerCase();

  if (main === "income") return amount;
  if (main === "transfer") {
    if (sub === "transfer in" || sub === "cc payment in") return amount;
    if (sub === "transfer out" || sub === "cc payment out") return -amount;
    return 0;
  }
  return -amount;
}

function getCreditCardTransferOwedImpact(row) {
  const amount = Math.abs(getSignedAmount(row["Amount"]));
  const sub = clean(row["Sub Category"]).toLowerCase();
  if (sub === "transfer in" || sub === "cc payment in") return -amount;
  if (sub === "transfer out" || sub === "cc payment out") return amount;
  return 0;
}

function computeAccountBalances(txData) {
  // Mirrors dashboard.js: opening balance plus signed account movement from opening date onwards.
  const savingsNames = allAccounts
    .filter(a => clean(a.type || "Savings").toLowerCase() === "savings")
    .map(a => a.name);
  const balances = {};

  savingsNames.forEach(account => {
    const accountRows = txData.filter(row => clean(row["Account"]) === account);

    const openingRow     = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? getSignedAmount(openingRow["Amount"]) : 0;
    const openingDate    = openingRow ? parseExcelDate(openingRow["Date"]) : null;

    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseExcelDate(row["Date"]);
      return d && d >= openingDate;
    });

    const movement = subsequent.reduce((sum, row) => sum + getAccountBalanceImpact(row), 0);

    balances[account] = openingBalance + movement;
  });

  return balances;
}

function computeCCOwed(txData) {
  // Mirrors dashboard.js: opening + charges, adjusted by transfer direction.
  const ccAccounts = getCreditCardAccountsInData(txData);

  let totalOwed = 0;
  ccAccounts.forEach(account => {
    const accountRows = txData.filter(row => clean(row["Account"]) === account);

    const openingRow     = accountRows.find(row => clean(row["Transaction"]) === "Opening Balance");
    const openingBalance = openingRow ? Math.abs(getSignedAmount(openingRow["Amount"])) : 0;
    const openingDate    = openingRow ? parseExcelDate(openingRow["Date"]) : null;

    const subsequent = accountRows.filter(row => {
      if (clean(row["Transaction"]) === "Opening Balance") return false;
      if (!openingDate) return true;
      const d = parseExcelDate(row["Date"]);
      return d && d >= openingDate;
    });

    const charges = subsequent
      .filter(row => { const cat = clean(row["Main Category"]).toLowerCase(); return cat !== "income" && cat !== "transfer"; })
      .reduce((sum, row) => sum + Math.abs(getSignedAmount(row["Amount"])), 0);

    const transferImpact = subsequent
      .filter(row => clean(row["Main Category"]).toLowerCase() === "transfer")
      .reduce((sum, row) => sum + getCreditCardTransferOwedImpact(row), 0);

    totalOwed += openingBalance + charges + transferImpact;
  });

  return Math.max(0, totalOwed);
}

function getCreditCardAccountsInData(txData) {
  const savingsNames = new Set(allAccounts
    .filter(a => clean(a.type || "Savings").toLowerCase() === "savings")
    .map(a => accountKey(a.name)));
  const creditCardNames = allAccounts
    .filter(a => clean(a.type).toLowerCase() === "credit card")
    .map(a => accountKey(a.name));
  const allAccountsInData = [...new Set((txData || []).map(row => clean(row["Account"])).filter(Boolean))];
  return creditCardNames.length > 0
    ? allAccountsInData.filter(a => creditCardNames.includes(accountKey(a)))
    : allAccountsInData.filter(a => !savingsNames.has(accountKey(a)));
}

function computeHistoricalStats(txData) {
  const monthlyIncome = {}, monthlyExpenses = {};
  // Only look at the last 12 complete months (exclude current partial month)
  const today = new Date();
  const cutoff = new Date(today.getFullYear(), today.getMonth() - 12, 1); // 12 months ago (start of that month)
  const currentMonthKey = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0");

  txData.forEach(row => {
    const date = parseExcelDate(row["Date"]);
    if (!date) return;
    if (date < cutoff) return; // older than 12 months ago
    const key = date.getFullYear() + "-" + String(date.getMonth()+1).padStart(2,"0");
    if (key === currentMonthKey) return; // exclude current partial month
    const cat = clean(row["Main Category"]).toLowerCase();
    if (cat === "transfer") return;
    const amt = getAmount(row["Amount"]);
    // Exclude opening balances from income stats
    if (cat === "income" && clean(row["Sub Category"]).toLowerCase() === "opening balance") return;
    if (cat === "income") monthlyIncome[key] = (monthlyIncome[key]||0) + amt;
    else monthlyExpenses[key] = (monthlyExpenses[key]||0) + amt;
  });
  const allMonths = [...new Set([...Object.keys(monthlyIncome),...Object.keys(monthlyExpenses)])].sort();
  const months = allMonths.length || 1;
  const avgIncome   = Object.values(monthlyIncome).reduce((s,v)=>s+v,0) / months;
  const avgExpenses = Object.values(monthlyExpenses).reduce((s,v)=>s+v,0) / months;
  return { avgMonthlyIncome:avgIncome, avgMonthlyExpenses:avgExpenses, avgMonthlySavings:avgIncome-avgExpenses, months };
}

// ─── Deployable Balance Calculation ───────────────────────────────

function computeDeployableBalance() {
  // Step 1: sum balances of goal-eligible savings accounts
  let rawSavings = 0;
  goalSavingsAccts.forEach(name => {
    rawSavings += savingsBalances[name] || 0;
  });

  // Step 2: subtract factual CC owed, then add pending claims as receivable.
  // This keeps debt visible while avoiding claimable spend quietly blocking goals.
  const ccAccounts = getCreditCardAccountsInData(allTxForGoals);
  const ccClaimReceivable = computePendingClaimReceivableForAccounts(allTxForGoals, ccAccounts);
  const savingsClaimReceivable = computePendingClaimReceivableForAccounts(allTxForGoals, goalSavingsAccts);
  const pendingClaimSummary = computePendingClaimSummary(allTxForGoals);
  const claimReceivableForGoals = pendingClaimSummary.total;
  const ccOwedForGoals = Math.max(0, ccOwed - ccClaimReceivable);
  const afterCC = rawSavings - ccOwed + claimReceivableForGoals;

  // Step 3: reserve positive unspent budget for the rest of the month.
  // Budget-page visuals handle category-level overspend reserves; Goals keeps this panel clean.
  const budgetPosition         = computeCurrentMonthBudgetPosition();
  const monthlyBudgetBalance   = budgetPosition.total.balance;
  const remainingBudgetReserve = Math.max(0, monthlyBudgetBalance);
  const deployable             = afterCC - remainingBudgetReserve;

  return {
    rawSavings,
    afterCC,
    ccOwedForGoals,
    ccClaimReceivable,
    savingsClaimReceivable,
    pendingClaimRows: pendingClaimSummary.count,
    claimReceivableForGoals,
    remainingBudget: remainingBudgetReserve,
    monthlyBudgetBalance,
    budgetPosition,
    deployable
  };
}

function computeCurrentMonthBudgetPosition() {
  const today = new Date();
  const isCurrentMonthExpense = row => {
    const d = parseExcelDate(row["Date"]);
    if (!d || d.getMonth() !== today.getMonth() || d.getFullYear() !== today.getFullYear()) return false;
    const cat = clean(row["Main Category"]).toLowerCase();
    return cat !== "income" && cat !== "transfer" && cat !== "saving goals";
  };

  const spentByType = {
    bills: new Map(),
    monthly: new Map()
  };

  allTxForGoals.filter(isCurrentMonthExpense).forEach(row => {
    const main = clean(row["Main Category"]).toLowerCase();
    const bucket = main === "bills"
      ? "bills"
      : main === "monthly expenses"
        ? "monthly"
        : "";
    if (!bucket) return;

    const impactAmount = getClaimAdjustedExpenseAmount(row);
    if (impactAmount <= 0) return;

    const subCategory = clean(row["Sub Category"]) || "(Uncategorised)";
    spentByType[bucket].set(
      subCategory,
      (spentByType[bucket].get(subCategory) || 0) + impactAmount
    );
  });

  const buildSection = (budgetRows, spentMap) => {
    const allocatedByCategory = new Map();
    budgetRows.forEach(row => {
      allocatedByCategory.set(row.category, (allocatedByCategory.get(row.category) || 0) + row.allocated);
    });

    const categoryNames = [...new Set([...allocatedByCategory.keys(), ...spentMap.keys()])];
    const rows = categoryNames.map(category => {
      const allocated = allocatedByCategory.get(category) || 0;
      const spent = spentMap.get(category) || 0;
      const balance = allocated - spent;
      return {
        category,
        allocated,
        spent,
        balance,
        over: Math.max(0, spent - allocated),
        isUnbudgeted: allocated <= 0 && spent > 0
      };
    }).sort((a, b) => b.over - a.over || b.spent - a.spent);

    const allocated = rows.reduce((sum, row) => sum + row.allocated, 0);
    const spent = rows.reduce((sum, row) => sum + row.spent, 0);
    const balance = allocated - spent;
    return { rows, allocated, spent, balance, over: Math.max(0, spent - allocated) };
  };

  const bills = buildSection(budgetSummary.billsRows || [], spentByType.bills);
  const monthly = buildSection(budgetSummary.monthlyRows || [], spentByType.monthly);
  const total = {
    allocated: bills.allocated + monthly.allocated,
    spent: bills.spent + monthly.spent,
    balance: bills.balance + monthly.balance
  };
  total.over = Math.max(0, total.spent - total.allocated);

  return { bills, monthly, total };
}

function computePendingClaimReceivableForAccounts(txData, accountNames) {
  const affectedAccounts = new Set((accountNames || []).map(accountKey));
  return (txData || []).reduce((sum, row) => {
    if (!affectedAccounts.has(accountKey(getClaimReceivableAccount(row)))) return sum;
    return sum + getPendingClaimReceivableAmount(row);
  }, 0);
}

function computePendingClaimSummary(txData) {
  return (txData || []).reduce((summary, row) => {
    const amount = getPendingClaimReceivableAmount(row);
    if (amount <= 0) return summary;
    summary.count += 1;
    summary.total += amount;
    return summary;
  }, { count:0, total:0 });
}

function computeGoalClaimRisk(dep, totalAssigned) {
  const pendingClaims = Math.max(0, dep.claimReceivableForGoals || 0);
  const cashOnlyDeployable = dep.deployable - pendingClaims;
  const cashOnlyUnassigned = cashOnlyDeployable - totalAssigned;
  const assignedFromPendingClaims = Math.min(
    pendingClaims,
    Math.max(0, totalAssigned - cashOnlyDeployable)
  );

  return {
    pendingClaims,
    cashOnlyDeployable,
    cashOnlyUnassigned,
    assignedFromPendingClaims
  };
}

function claimRiskMessage(risk) {
  if (risk.assignedFromPendingClaims > 0) {
    return `${formatCurrency(risk.assignedFromPendingClaims)} of assigned goals depends on pending claims being paid. Without those claims, cash-only unassigned is ${formatCurrency(risk.cashOnlyUnassigned)}.`;
  }
  return `Pending claims are included in deployable balance, but assigned goals are still covered without them. Cash-only unassigned is ${formatCurrency(risk.cashOnlyUnassigned)}.`;
}

function goalPriorityValue(goal) {
  return ({ Critical:0, High:1, Medium:2, Low:3 })[goal?.urgency] ?? 2;
}

function getPrioritizedGoalIndexes() {
  return goalsData
    .map((g, idx) => ({ ...g, originalIdx: idx }))
    .sort((a,b) => {
      const uDiff = goalPriorityValue(a) - goalPriorityValue(b);
      if (uDiff !== 0) return uDiff;
      if (a.endDate && b.endDate) return a.endDate.localeCompare(b.endDate);
      if (a.endDate) return -1;
      if (b.endDate) return  1;
      return a.originalIdx - b.originalIdx;
    })
    .map(g => g.originalIdx);
}

function computeGoalPendingClaimExposure(dep) {
  const pendingClaims = Math.max(0, dep.claimReceivableForGoals || 0);
  let cashPool = Math.max(0, dep.deployable - pendingClaims);
  let pendingPool = pendingClaims;
  const exposure = goalsData.map(() => 0);

  getPrioritizedGoalIndexes().forEach(idx => {
    const assigned = Math.max(0, goalsData[idx]?.manualSaved || 0);
    if (assigned <= 0) return;

    const cashUsed = Math.min(cashPool, assigned);
    cashPool -= cashUsed;
    const uncovered = assigned - cashUsed;
    const pendingUsed = Math.min(pendingPool, uncovered);
    exposure[idx] = pendingUsed;
    pendingPool -= pendingUsed;
  });

  return exposure;
}

function buildAllocStatusTags(remaining, isFullyFunded, pendingExposure) {
  const statusTag = isFullyFunded
    ? `<span class="alloc-done-tag">✓ Fully funded</span>`
    : (remaining > 0
      ? `<span class="alloc-need-tag">${formatCurrency(remaining)} left</span>`
      : `<span class="alloc-done-tag">✓ Done</span>`);
  const claimTag = pendingExposure > 0
    ? `<span class="alloc-claim-tag">${formatCurrency(pendingExposure)} pending claim</span>`
    : "";
  return `<div class="alloc-tag-stack">${statusTag}${claimTag}</div>`;
}

// ─── Render Page ──────────────────────────────────────────────────

function renderGoalsPage() {
  const container = document.getElementById("goalsContainer");
  container.innerHTML = "";

  const today = new Date();
  const dep   = computeDeployableBalance();

  // ── 1. Account Selector ──
  const savingsAccountOptions = allAccounts.filter(a => a.type === "Savings");
  const acctCheckboxes = savingsAccountOptions.map(a => {
    const checked = goalSavingsAccts.includes(a.name) ? "checked" : "";
    return `<label class="acct-check-label">
      <input type="checkbox" value="${escapeHtml(a.name)}" ${checked} onchange="toggleGoalAccount(this)">
      <span>${escapeHtml(a.name)}</span>
      <span class="acct-bal ${(savingsBalances[a.name]||0)>=0?'green':'red'}">${formatCurrency(savingsBalances[a.name]||0)}</span>
    </label>`;
  }).join("");

  container.innerHTML += `
    <div class="goals-panel acct-selector-panel">
      <div class="panel-header">
        <span class="panel-icon">🏦</span>
        <h2>Goal Savings Accounts</h2>
        <span class="panel-hint">Select which accounts fund your goals. Exclude business or restricted accounts.</span>
      </div>
      <div class="acct-checklist">${acctCheckboxes || '<span style="color:var(--muted);font-size:13px;">No savings accounts found. Add them in <a href="accounts.html">Accounts & Setup</a>.</span>'}</div>
    </div>
  `;

  // ── 2. Deployable Balance + Fund Allocator ──
  const depClass = dep.deployable >= 0 ? "green" : "red";
  const totalAllocatedToGoals = goalsData.reduce((s,g)=>s+g.manualSaved,0);
  const unassigned = dep.deployable - totalAllocatedToGoals;
  const claimRisk = computeGoalClaimRisk(dep, totalAllocatedToGoals);
  const budgetBalance = dep.monthlyBudgetBalance || 0;
  const budgetIsOver = budgetBalance < 0;
  const budgetLineLabel = budgetIsOver ? "Budget overrun" : "− Remaining budget reserve";
  const budgetLineValue = budgetIsOver
    ? formatCurrency(budgetBalance)
    : (dep.remainingBudget > 0 ? "−" + formatCurrency(dep.remainingBudget) : formatCurrency(0));
  const claimRiskHtml = claimRisk.pendingClaims > 0 ? `
        <div class="balance-row claim-row">
          <span class="bal-label">+ Pending claims included</span>
          <span class="bal-val amber">${formatCurrency(claimRisk.pendingClaims)}</span>
        </div>` : "";
  const claimCashRowHtml = claimRisk.pendingClaims > 0 ? `
        <div class="balance-row sub-row claim-cash-row ${claimRisk.cashOnlyUnassigned>=0?'':'warn-row'}" id="claimCashRow">
          <span class="bal-label" id="claimCashLabel">${claimRisk.cashOnlyUnassigned>=0 ? "Cash-only unassigned" : "Cash-only shortfall"}</span>
          <span class="bal-val ${claimRisk.cashOnlyUnassigned>=0?'green':'red'}" id="liveCashOnlyUnassigned">${formatCurrency(claimRisk.cashOnlyUnassigned)}</span>
        </div>
        <div class="claim-risk-note ${claimRisk.assignedFromPendingClaims>0?'active':'safe'}" id="claimRiskNote">
          <strong id="claimRiskTitle">${claimRisk.assignedFromPendingClaims>0 ? "Goal allocation uses pending claims" : "Pending claims are visible but not needed yet"}</strong>
          <span id="claimRiskBody">${claimRiskMessage(claimRisk)}</span>
        </div>` : "";
  const claimExposureByGoal = computeGoalPendingClaimExposure(dep);

  // Build per-goal allocator rows
  const allocatorRows = goalsData.map((g, idx) => {
    const savedViaGoalTx  = getSavedViaTransactions(g.name);
    const effectiveTarget = g.target * (1 + (g.goalBuffer || 0) / 100);
    const totalSaved      = g.manualSaved + savedViaGoalTx;
    const remaining       = Math.max(0, effectiveTarget - totalSaved);
    const pct             = effectiveTarget > 0 ? Math.min(100,(totalSaved/effectiveTarget)*100) : 0;
    const urgColor        = { Critical:"#dc2626", High:"#ea580c", Medium:"#ca8a04", Low:"#16a34a" };
    const overTarget      = totalSaved >= effectiveTarget;
    const pendingExposure = claimExposureByGoal[idx] || 0;
    const bufferLabel     = g.goalBuffer > 0 ? ` <span style="color:var(--muted);font-size:11px;">(+${g.goalBuffer}% buffer)</span>` : "";
    return `
      <div class="alloc-row" id="allocRow_${idx}">
        <div class="alloc-name-col">
          <span class="alloc-dot" style="background:${g.color};"></span>
          <div>
            <div class="alloc-name">${escapeHtml(g.name)}${bufferLabel}</div>
            <div class="alloc-sub">${formatCurrency(totalSaved)} / ${formatCurrency(effectiveTarget)} &nbsp;·&nbsp; <span style="color:${urgColor[g.urgency]||'#ca8a04'}">${g.urgency}</span></div>
          </div>
        </div>
        <div class="alloc-bar-col">
          <div class="alloc-pbar"><div class="alloc-pfill" style="width:${pct.toFixed(1)}%;background:${overTarget?'#16a34a':g.color};"></div></div>
          <span class="alloc-pct">${pct.toFixed(0)}%</span>
        </div>
        <div class="alloc-input-col">
          <div style="display:flex;gap:6px;align-items:center;">
            <div class="alloc-input-wrap">
              <span class="alloc-currency">$</span>
              <input class="alloc-input" type="number" step="1" min="0"
                id="allocAmt_${idx}" value="${g.manualSaved}"
                onchange="onAllocChange(${idx}, this.value)"
                oninput="onAllocInput(${idx}, this.value)">
            </div>
            <div class="alloc-buffer-inp-wrap" title="Extra % buffer above target (e.g. 5 = 5% extra)">
              <span class="alloc-currency" style="background:#fff7ed;color:#ea580c;border-color:#fed7aa;">+</span>
              <input class="alloc-input" type="number" step="1" min="0" max="100"
                id="allocBuf_${idx}" value="${g.goalBuffer||0}"
                placeholder="0"
                title="Personal buffer above target as % (e.g. 10 = 10% extra)"
                onchange="onBufferChange(${idx}, this.value)"
                style="width:60px;">
              <span style="font-size:11px;color:#ea580c;margin-left:2px;">%</span>
            </div>
          </div>
          <div id="allocTag_${idx}">
            ${buildAllocStatusTags(remaining, overTarget, pendingExposure)}
          </div>
        </div>
      </div>`;
  }).join("");

  // Overspend alert
  const overspendAlert = unassigned < 0 ? `
    <div class="overspend-alert">
      <div class="overspend-icon">⚠️</div>
      <div class="overspend-body">
        <strong>Over-assigned by ${formatCurrency(Math.abs(unassigned))}</strong>
        <span>Your budget or CC balance changed and goals now exceed your deployable balance. Reduce allocations or add more savings.</span>
        <button class="btn-rebalance" onclick="smartAssign()">✨ Auto-Rebalance Now</button>
      </div>
    </div>` : "";

  container.innerHTML += `
    <div class="goals-panel balance-panel">
      <div class="panel-header">
        <span class="panel-icon">💰</span>
        <h2>Deployable Balance</h2>
        <span class="panel-hint">What you can actually assign to goals right now</span>
      </div>
      ${overspendAlert}
      <div class="balance-flow">
        <div class="balance-row">
          <span class="bal-label">Selected accounts total</span>
          <span class="bal-val">${formatCurrency(dep.rawSavings)}</span>
        </div>
        <div class="balance-row deduct">
          <span class="bal-label">− Credit card owed</span>
          <span class="bal-val red">−${formatCurrency(ccOwed)}</span>
        </div>
        ${claimRiskHtml}
        <div class="balance-row deduct ${budgetIsOver ? 'warn-row' : ''}">
          <span class="bal-label">${budgetLineLabel}</span>
          <span class="bal-val ${budgetIsOver || dep.remainingBudget > 0 ? 'red' : ''}">${budgetLineValue}</span>
        </div>
        <div class="balance-row total-row">
          <span class="bal-label">= Deployable Balance</span>
          <span class="bal-val ${depClass}" style="font-size:22px;">${formatCurrency(dep.deployable)}</span>
        </div>
        <div class="balance-row sub-row">
          <span class="bal-label">Assigned to goals</span>
          <span class="bal-val" id="liveTotalAssigned">${formatCurrency(totalAllocatedToGoals)}</span>
        </div>
        <div class="balance-row sub-row ${unassigned>=0?'':'warn-row'}">
          <span class="bal-label">Unassigned</span>
          <span class="bal-val ${unassigned>=0?'green':'red'}" id="liveUnassigned">${formatCurrency(unassigned)}</span>
        </div>
        ${claimCashRowHtml}
      </div>

      ${goalsData.length > 0 ? `
      <div class="alloc-section">
        <div class="alloc-header">
          <span class="alloc-title">🎯 Fund Allocator</span>
          <div class="alloc-header-actions">
            <button class="btn-smart-assign" onclick="smartAssign()">✨ Smart Assign</button>
          </div>
        </div>
        <div class="alloc-hint">Set allocation ($) and optional per-goal buffer (%) above target. <strong>Smart Assign</strong> uses cash first, then pending claims; claim-backed portions are tagged in amber. Unassigned updates live.</div>
        <div class="alloc-col-labels">
          <span class="alloc-col-label-name">Goal</span>
          <span class="alloc-col-label-bar">Progress</span>
          <span class="alloc-col-label-inp">Allocation &nbsp;&nbsp; + Buffer</span>
        </div>
        <div class="alloc-rows">${allocatorRows}</div>
        <div class="alloc-footer">
          <div class="alloc-unassigned-bar-wrap">
            <div class="alloc-unassigned-label">Unassigned: <strong id="liveUnassignedBar">${formatCurrency(unassigned)}</strong></div>
            <div class="alloc-ubar-track">
              <div class="alloc-ubar-fill ${unassigned>=0?'green':'red'}" id="liveUnassignedFill"
                style="width:${dep.deployable>0?Math.min(100,Math.max(0,(unassigned/dep.deployable)*100)).toFixed(1):0}%;"></div>
            </div>
          </div>
          <span class="panel-hint goals-autosave-status" id="allocAutosaveStatus" style="white-space:nowrap;">Autosaves to Excel</span>
        </div>
      </div>` : ""}
    </div>
  `;

  renderInsuranceRenewalGoalPanel(container);

  // ── 3. Historical snapshot ──
  container.innerHTML += `
    <div class="goals-panel stats-panel">
      <div class="panel-header">
        <span class="panel-icon">📊</span>
        <h2>Your Savings Pattern <span style="font-weight:400;font-size:13px;color:var(--muted);">(${historicalStats.months} months)</span></h2>
      </div>
      <div class="stats-grid">
        <div class="stat-item"><span class="stat-l">Avg Monthly Income</span><span class="stat-v green">${formatCurrency(historicalStats.avgMonthlyIncome)}</span></div>
        <div class="stat-item"><span class="stat-l">Avg Monthly Expenses</span><span class="stat-v red">${formatCurrency(historicalStats.avgMonthlyExpenses)}</span></div>
        <div class="stat-item"><span class="stat-l">Avg Monthly Savings Rate</span><span class="stat-v ${historicalStats.avgMonthlySavings>=0?'green':'red'}">${formatCurrency(historicalStats.avgMonthlySavings)}</span></div>
        <div class="stat-item"><span class="stat-l">Budget Committed</span><span class="stat-v">${formatCurrency(budgetSummary.billsTotal+budgetSummary.monthlyTotal)}/mo</span></div>
      </div>
    </div>
  `;

  // ── 4. Multi-goal realism check ──
  if (goalsData.length > 0) {
    container.innerHTML += `<div id="goalRealismRegion"></div>`;
  }

  // ── 4.5 Budget-to-goal motivation ──
  if (goalsData.length > 0) {
    container.innerHTML += `<div id="goalMotivationRegion"></div>`;
  }

  // ── 4.6 Income Boosts ──
  if (goalsData.length > 0) {
    renderIncomeBoostsPanel(container);
  }

  // ── 4.7 Forecast visual placeholders ──
  if (goalsData.length > 0) {
    container.innerHTML += `<div id="goalForecastRegion"></div>`;
  }

  // ── 5. Add Goal Form ──
  container.innerHTML += `
    <div class="goals-panel add-goal-panel" id="addGoalPanel">
      <div class="panel-header" style="cursor:pointer;" onclick="toggleAddForm()">
        <span class="panel-icon">＋</span>
        <h2>Add New Goal</h2>
        <span class="panel-hint" id="addFormToggleHint">Click to expand</span>
      </div>
      <div id="addGoalFormBody" style="display:none;margin-top:16px;">
        <div class="goal-form-grid">
          <div class="gf-group full"><label>Goal Name</label><input type="text" id="gf_name" placeholder="e.g. Emergency Fund, Japan Trip, New Laptop"></div>
          <div class="gf-group"><label>Target Amount (SGD)</label><input type="number" id="gf_target" step="1" placeholder="10000" oninput="_updateAddGoalCalc()"></div>

          <div class="gf-group">
            <label>Target Deadline</label>
            <input type="date" id="gf_end" oninput="_updateAddGoalCalc()">
          </div>
          <div class="gf-group">
            <label>Start Date
              <span class="gf-label-hint" id="gf_start_hint" style="font-size:11px;font-weight:400;color:var(--muted);margin-left:6px;"></span>
            </label>
            <input type="date" id="gf_start" value="${toDateInputValue(today)}" oninput="_updateAddGoalCalc()">
            <div id="gf_start_suggestions" class="gf-suggestions" style="display:none;"></div>
          </div>

          <div class="gf-group">
            <label>Monthly Allocation (SGD)
              <span style="font-size:11px;font-weight:400;color:var(--muted);margin-left:6px;">auto-calculated</span>
            </label>
            <div class="gf-alloc-display" id="gf_alloc_display">
              <span class="gf-alloc-amount" id="gf_alloc_computed">—</span>
              <span class="gf-alloc-override-link" onclick="_toggleAllocOverride()">override</span>
            </div>
            <input type="number" id="gf_alloc" step="1" placeholder="override amount" style="display:none;margin-top:6px;">
            <input type="hidden" id="gf_alloc_final" value="0">
          </div>

          <div class="gf-group">
            <label>Priority</label>
            <select id="gf_urgency">
              <option value="Critical">🔴 Critical — must hit deadline</option>
              <option value="High">🟠 High — important, some flex</option>
              <option value="Medium" selected>🟡 Medium — steady progress</option>
              <option value="Low">🟢 Low — nice to have</option>
            </select>
          </div>
          <div class="gf-group"><label>Color Tag (auto-assigned)</label><input type="color" id="gf_color" value="#4caf50" style="height:38px;padding:2px 4px;border-radius:6px;border:1px solid var(--border);width:100%;"></div>
          <div class="gf-group full"><label>Notes</label><input type="text" id="gf_notes" placeholder="Optional description or motivation"></div>
        </div>

        <!-- Calc summary -->
        <div id="gf_calc_summary" class="gf-calc-summary" style="display:none;"></div>

        <div style="display:flex;gap:10px;margin-top:16px;">
          <button class="btn-primary" onclick="addGoal()">Add Goal</button>
          <button class="btn-secondary" onclick="toggleAddForm()">Cancel</button>
        </div>
      </div>
    </div>
  `;

  // ── 6. Goal Cards ──
  const urgencyOrder = { Critical:0, High:1, Medium:2, Low:3 };
  const sorted = [...goalsData]
    .map((g,i) => ({...g, originalIdx:i}))
    .sort((a,b) => (urgencyOrder[a.urgency]??2) - (urgencyOrder[b.urgency]??2));

  if (sorted.length === 0) {
    container.innerHTML += `<div class="goals-panel" style="padding:32px;text-align:center;color:var(--muted);">No goals yet — add your first one above!</div>`;
  } else {
    const goalsGrid = document.createElement("div");
    goalsGrid.className = "goals-grid";
    goalsGrid.id = "goalCardsGrid";
    container.appendChild(goalsGrid);
    sorted.forEach(goal => renderGoalCard(goal, goal.originalIdx, goalsGrid));
  }

  // ── 7. Autosave status ──
  container.innerHTML += `
    <div style="margin-top:8px;padding-bottom:40px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
      <span class="panel-hint goals-autosave-status" id="goalsAutosaveStatus">Autosaves to Excel</span>
    </div>
  `;

  refreshGoalInsightPanels();
}

const GOAL_COLORS = [
  "#4caf50","#2196f3","#9c27b0","#ff9800","#e91e63",
  "#00bcd4","#ff5722","#3f51b5","#8bc34a","#f44336",
  "#009688","#673ab7","#cddc39","#795548","#607d8b"
];

function getNextGoalColor() {
  const usedColors = new Set(goalsData.map(g => (g.color||"").toLowerCase()));
  const free = GOAL_COLORS.find(c => !usedColors.has(c.toLowerCase()));
  return free || GOAL_COLORS[goalsData.length % GOAL_COLORS.length];
}

function _refreshAllocUI() {
  const dep = computeDeployableBalance();
  const totalAssigned = goalsData.reduce((s,g)=>s+g.manualSaved, 0);
  const unassigned = dep.deployable - totalAssigned;
  const claimRisk = computeGoalClaimRisk(dep, totalAssigned);
  const claimExposureByGoal = computeGoalPendingClaimExposure(dep);
  const unassignedPct = dep.deployable > 0
    ? Math.min(100, Math.max(0, (unassigned / dep.deployable) * 100))
    : 0;

  const elTotal   = document.getElementById("liveTotalAssigned");
  const elUna     = document.getElementById("liveUnassigned");
  const elUnaBar  = document.getElementById("liveUnassignedBar");
  const elUnaFill = document.getElementById("liveUnassignedFill");
  const elCashOnly = document.getElementById("liveCashOnlyUnassigned");
  const claimCashLabel = document.getElementById("claimCashLabel");
  const claimCashRow = document.getElementById("claimCashRow");
  const claimRiskNote = document.getElementById("claimRiskNote");
  const claimRiskTitle = document.getElementById("claimRiskTitle");
  const claimRiskBody = document.getElementById("claimRiskBody");

  if (elTotal)   elTotal.textContent = formatCurrency(totalAssigned);
  if (elUna)     { elUna.textContent = formatCurrency(unassigned); elUna.className = "bal-val " + (unassigned>=0?"green":"red"); }
  if (elUnaBar)  { elUnaBar.textContent = formatCurrency(unassigned); }
  if (elUnaFill) { elUnaFill.style.width = unassignedPct.toFixed(1) + "%"; elUnaFill.className = "alloc-ubar-fill " + (unassigned>=0?"green":"red"); }
  if (elCashOnly) {
    elCashOnly.textContent = formatCurrency(claimRisk.cashOnlyUnassigned);
    elCashOnly.className = "bal-val " + (claimRisk.cashOnlyUnassigned>=0?"green":"red");
  }
  if (claimCashLabel) {
    claimCashLabel.textContent = claimRisk.cashOnlyUnassigned>=0 ? "Cash-only unassigned" : "Cash-only shortfall";
  }
  if (claimCashRow) {
    claimCashRow.className = "balance-row sub-row claim-cash-row " + (claimRisk.cashOnlyUnassigned>=0 ? "" : "warn-row");
  }
  if (claimRiskNote) {
    claimRiskNote.className = "claim-risk-note " + (claimRisk.assignedFromPendingClaims>0 ? "active" : "safe");
  }
  if (claimRiskTitle) {
    claimRiskTitle.textContent = claimRisk.assignedFromPendingClaims>0
      ? "Goal allocation uses pending claims"
      : "Pending claims are visible but not needed yet";
  }
  if (claimRiskBody) claimRiskBody.textContent = claimRiskMessage(claimRisk);

  // Show/hide overspend alert dynamically
  const alertEl = document.querySelector(".overspend-alert");
  if (alertEl) {
    alertEl.style.display = unassigned < 0 ? "flex" : "none";
    const strong = alertEl.querySelector("strong");
    if (strong) strong.textContent = "Over-assigned by " + formatCurrency(Math.abs(unassigned));
  }

  // Refresh per-goal remaining tags and progress
  goalsData.forEach((g, idx) => {
    const savedViaGoalTx  = getSavedViaTransactions(g.name);
    const effectiveTarget = g.target * (1 + (g.goalBuffer || 0) / 100);
    const totalSaved      = g.manualSaved + savedViaGoalTx;
    const remaining       = effectiveTarget - totalSaved;
    const pct             = effectiveTarget > 0 ? Math.min(100,(totalSaved/effectiveTarget)*100) : 0;
    const pendingExposure = claimExposureByGoal[idx] || 0;
    const row             = document.getElementById("allocRow_" + idx);
    if (!row) return;
    const fill     = row.querySelector(".alloc-pfill");
    const pctLabel = row.querySelector(".alloc-pct");
    const subLabel = row.querySelector(".alloc-sub");
    const tagDiv   = document.getElementById("allocTag_" + idx);
    if (fill)     { fill.style.width = pct.toFixed(1) + "%"; fill.style.background = totalSaved>=effectiveTarget?"#16a34a":g.color; }
    if (pctLabel)  pctLabel.textContent = pct.toFixed(0) + "%";
    if (subLabel)  subLabel.innerHTML = formatCurrency(totalSaved) + " / " + formatCurrency(effectiveTarget) + " &nbsp;·&nbsp; <span style='color:" + ({Critical:"#dc2626",High:"#ea580c",Medium:"#ca8a04",Low:"#16a34a"}[g.urgency]||"#ca8a04") + "'>" + g.urgency + "</span>";
    if (tagDiv) {
      tagDiv.innerHTML = buildAllocStatusTags(remaining, totalSaved >= effectiveTarget, pendingExposure);
    }
  });
}

function refreshGoalCardsGrid() {
  const grid = document.getElementById("goalCardsGrid");
  if (!grid) return;

  const urgencyOrder = { Critical:0, High:1, Medium:2, Low:3 };
  const sorted = [...goalsData]
    .map((g,i) => ({...g, originalIdx:i}))
    .sort((a,b) => (urgencyOrder[a.urgency]??2) - (urgencyOrder[b.urgency]??2));

  grid.innerHTML = "";
  sorted.forEach(goal => renderGoalCard(goal, goal.originalIdx, grid));
}

function onAllocInput(idx, value) {
  const v = Math.max(0, parseFloat(value) || 0);
  goalsData[idx].manualSaved = v;
  // Clamp input to non-negative
  const inp = document.getElementById("allocAmt_" + idx);
  if (inp && parseFloat(inp.value) < 0) inp.value = 0;
  _refreshAllocUI();
  refreshGoalCardsGrid();
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function onAllocChange(idx, value) {
  const v = Math.max(0, parseFloat(value) || 0);
  goalsData[idx].manualSaved = v;
  const inp = document.getElementById("allocAmt_" + idx);
  if (inp) inp.value = v.toFixed(2);
  _refreshAllocUI();
  refreshGoalCardsGrid();
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function onBufferChange(idx, value) {
  const v = Math.min(100, Math.max(0, parseFloat(value) || 0));
  goalsData[idx].goalBuffer = v;
  const inp = document.getElementById("allocBuf_" + idx);
  if (inp) inp.value = v.toFixed(0);
  _refreshAllocUI();
  refreshGoalCardsGrid();
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

/**
 * Smart Assign: fill each goal (target + its own buffer) fully, in priority+deadline order.
 * Cash is consumed first; if pending claims are needed, the affected goal rows are tagged.
 * No global buffer — each goal carries its own buffer if set.
 * Any remaining deployable after all goals are filled stays as unassigned.
 */
function smartAssign() {
  const dep = computeDeployableBalance();
  const deployable = dep.deployable;
  if (deployable <= 0) {
    alert("No deployable balance to assign.");
    return;
  }

  let toDistribute = deployable;
  const prioritizedIndexes = getPrioritizedGoalIndexes();

  // Reset manual savings for goals that aren't already tx-covered
  goalsData.forEach((g, idx) => {
    const savedViaGoalTx  = getSavedViaTransactions(g.name);
    const effectiveTarget = g.target * (1 + (g.goalBuffer || 0) / 100);
    if (savedViaGoalTx < effectiveTarget) {
      goalsData[idx].manualSaved = 0;
    }
  });

  prioritizedIndexes.forEach(originalIdx => {
    const g = goalsData[originalIdx];
    const savedViaGoalTx  = getSavedViaTransactions(g.name);
    const effectiveTarget = g.target * (1 + (g.goalBuffer || 0) / 100);
    const needManual      = Math.max(0, effectiveTarget - savedViaGoalTx);
    if (needManual <= 0) return;

    const assign = Math.min(toDistribute, needManual);
    goalsData[originalIdx].manualSaved = assign;
    toDistribute = Math.max(0, toDistribute - assign);
  });

  // Update all input fields
  goalsData.forEach((g, idx) => {
    const inp = document.getElementById("allocAmt_" + idx);
    if (inp) inp.value = g.manualSaved.toFixed(2);
  });

  _refreshAllocUI();
  refreshGoalCardsGrid();
  refreshGoalInsightPanels();

  const totalAssigned = goalsData.reduce((s,g)=>s+g.manualSaved, 0);
  const remaining = dep.deployable - totalAssigned;
  const finalRisk = computeGoalClaimRisk(dep, totalAssigned);
  log("Smart Assign complete. Unassigned: " + formatCurrency(remaining) + (finalRisk.assignedFromPendingClaims > 0 ? " | pending-claim backed: " + formatCurrency(finalRisk.assignedFromPendingClaims) : ""));
  scheduleGoalsAutoSave(150);
}

async function saveAllocations() {
  try {
    log("Saving allocations...");
    await persistGoalsToExcel({ silent: false, includeBoosts: false });
    log("Allocations saved.");
    const btns = document.querySelectorAll(".alloc-footer .btn-primary");
    btns.forEach(b => { const orig = b.textContent; b.textContent = "✅ Saved!"; setTimeout(()=>{ b.textContent = orig; }, 2000); });
  } catch(err) {
    log("SAVE ERROR: " + err.message);
    alert("Failed to save: " + err.message);
  }
}

// ─── Savings Timeline Chart ────────────────────────────────────────
// Plots % progress (0–100%) per goal so all goals are visible on the same axis.
// Shows: actual progress line, linear "on-track" guide, deadline marker, deficit zone.
// Income boosts (scheduled budget savings) shift the monthly alloc from a set month.

let timelineChart = null;
let goalInsightsRefreshTimer = null;

function refreshGoalInsightPanels() {
  const realismRegion = document.getElementById("goalRealismRegion");
  const motivationRegion = document.getElementById("goalMotivationRegion");
  const forecastRegion = document.getElementById("goalForecastRegion");
  if (!realismRegion && !motivationRegion && !forecastRegion) return;

  const dep = computeDeployableBalance();

  if (realismRegion) {
    realismRegion.innerHTML = "";
    if (goalsData.length > 0) renderRealismCheck(realismRegion, dep);
  }

  if (motivationRegion) {
    motivationRegion.innerHTML = "";
    if (goalsData.length > 0) renderBudgetGoalMotivation(motivationRegion, dep);
  }

  if (forecastRegion) {
    if (timelineChart) {
      timelineChart.destroy();
      timelineChart = null;
    }
    forecastRegion.innerHTML = "";
    if (goalsData.length > 0) {
      renderGoalGanttTimeline(forecastRegion);
      renderSavingsTimeline(forecastRegion, dep);
    }
  }
}

function scheduleGoalInsightRefresh() {
  clearTimeout(goalInsightsRefreshTimer);
  goalInsightsRefreshTimer = setTimeout(refreshGoalInsightPanels, 220);
}

// ── Future Cashflow Changes ───────────────────────────────────────
// Stored in goalsData as a separate list; also saved to Excel in col AE2
// Each item: { label, kind ("boost"|"reduce"), amount, fromMonth (YYYY-MM), toMonth (YYYY-MM or ""), toGoal (name or "any") }
// toMonth is optional — leave blank for a permanent change (e.g. a bill that ends forever).
// Example permanent: "Internet bill ended" → fromMonth: 2026-10, toMonth: "" (runs forever)
// Example temporary: "Annual bonus" → fromMonth: 2026-12, toMonth: 2026-12 (one month only)
let incomeBoosts = [];
const BOOSTS_RANGE = "AE2:AE2"; // single cell, JSON-stringified

function loadIncomeBoosts(budgetSheet) {
  const cell = budgetSheet["AE2"];
  try {
    const raw = cell ? String(cell.v ?? "").trim() : "";
    incomeBoosts = raw ? JSON.parse(raw) : [];
    incomeBoosts = incomeBoosts.map(b => ({
      ...b,
      kind: _normaliseBoostKind(b),
      amount: Math.abs(Number(b.amount || 0) || 0),
      fromMonth: _normaliseBoostMonth(b.fromMonth),
      toMonth: _normaliseBoostMonth(b.toMonth),
      toGoal: _normaliseBoostKind(b) === "reduce" ? "any" : (b.toGoal || "any")
    }));
  } catch { incomeBoosts = []; }
}

async function saveIncomeBoosts() {
  try {
    await writeBudgetSetupRange(BOOSTS_RANGE, [[JSON.stringify(incomeBoosts)]]);
    log("Cashflow changes saved.");
  } catch(err) { log("Cashflow save failed: " + err.message); }
}

function renderIncomeBoostsPanel(container) {
  const panel = document.createElement("div");
  panel.className = "goals-panel";
  panel.id = "boostsPanel";
  _redrawBoostsPanel(panel);
  container.appendChild(panel);
}
function _normaliseBoostMonth(value) {
  if (!value) return "";
  const s = String(value).trim();

  const isoMonth = s.match(/^(\d{4})-(\d{2})$/);
  if (isoMonth) return `${isoMonth[1]}-${isoMonth[2]}`;

  const isoDate = s.match(/^(\d{4})-(\d{2})-\d{2}$/);
  if (isoDate) return `${isoDate[1]}-${isoDate[2]}`;

  const wordMonth = s.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (wordMonth) {
    const months = ["january","february","march","april","may","june","july","august","september","october","november","december"];
    const idx = months.indexOf(wordMonth[1].toLowerCase());
    if (idx >= 0) return `${wordMonth[2]}-${String(idx + 1).padStart(2, "0")}`;
  }

  return "";
}

function _monthToDateValue(monthValue) {
  const m = _normaliseBoostMonth(monthValue);
  return m ? `${m}-01` : "";
}

function _dateValueToMonth(dateValue) {
  return dateValue ? String(dateValue).slice(0, 7) : "";
}

function _normaliseBoostKind(item) {
  const raw = String(item?.kind || item?.type || item?.direction || "").toLowerCase();
  if (raw === "reduce" || raw === "reduction" || raw === "expense" || raw === "decrease") return "reduce";
  if (Number(item?.amount || 0) < 0) return "reduce";
  return "boost";
}

function _boostSignedAmount(item) {
  const amount = Math.abs(Number(item?.amount || 0) || 0);
  return _normaliseBoostKind(item) === "reduce" ? -amount : amount;
}

function _setBoostDate(index, field, dateValue) {
  incomeBoosts[index][field] = _dateValueToMonth(dateValue);
  _redrawBoostsPanel(document.getElementById("boostsPanel"));
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _setBoostKind(index, kind) {
  incomeBoosts[index].kind = kind === "reduce" ? "reduce" : "boost";
  if (incomeBoosts[index].kind === "reduce") incomeBoosts[index].toGoal = "any";
  _redrawBoostsPanel(document.getElementById("boostsPanel"));
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _setBoostLabel(index, value) {
  if (!incomeBoosts[index]) return;
  incomeBoosts[index].label = value;
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _setBoostAmount(index, value) {
  if (!incomeBoosts[index]) return;
  incomeBoosts[index].amount = Math.abs(Number(value) || 0);
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _setBoostGoal(index, value) {
  if (!incomeBoosts[index]) return;
  incomeBoosts[index].toGoal = value;
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _addIncomeBoost(kind) {
  incomeBoosts.push({ label:"", kind, amount:0, fromMonth:"", toMonth:"", toGoal:"any" });
  _redrawBoostsPanel(document.getElementById("boostsPanel"));
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _deleteIncomeBoost(index) {
  incomeBoosts.splice(index, 1);
  _redrawBoostsPanel(document.getElementById("boostsPanel"));
  scheduleGoalInsightRefresh();
  scheduleGoalsAutoSave();
}

function _redrawBoostsPanel(panel) {
  if (!panel) return;
  const rows = incomeBoosts.map((b,i) => {
    b.kind = _normaliseBoostKind(b);
    b.amount = Math.abs(Number(b.amount || 0) || 0);
    b.fromMonth = _normaliseBoostMonth(b.fromMonth);
    b.toMonth = _normaliseBoostMonth(b.toMonth);
    if (b.kind === "reduce") b.toGoal = "any";
    const isPermanent = !b.toMonth;
    const isReduction = b.kind === "reduce";
    return `
      <div class="boost-row" id="boostRow_${i}">
        <div class="boost-cell boost-label-cell">
          <input class="boost-input" value="${escapeHtml(b.label || "")}" placeholder="${isReduction ? "e.g. rent increase" : "e.g. PPHS ended"}"
            oninput="_setBoostLabel(${i}, this.value)">
        </div>
        <div class="boost-cell boost-kind-cell">
          <select class="boost-select boost-kind-select ${isReduction ? "reduce" : "boost"}" onchange="_setBoostKind(${i}, this.value)">
            <option value="boost" ${!isReduction ? "selected" : ""}>Increase savings</option>
            <option value="reduce" ${isReduction ? "selected" : ""}>Reduce savings</option>
          </select>
        </div>
        <div class="boost-cell boost-amt-cell">
          <div class="boost-amt-wrap">
            <span class="boost-currency ${isReduction ? "reduce" : "boost"}">${isReduction ? "− $" : "+ $"}</span>
            <input class="boost-input boost-number" type="number" step="1" min="0" value="${b.amount}"
              oninput="_setBoostAmount(${i}, this.value)" placeholder="0">
            <span class="boost-per-mo">/mo</span>
          </div>
        </div>
        <div class="boost-cell boost-dates-cell">
          <div class="boost-dates-row">
            <div class="boost-date-group">
              <span class="boost-date-label">From</span>
              <input class="boost-month-input" type="date" value="${_monthToDateValue(b.fromMonth)}"
                onchange="_setBoostDate(${i}, 'fromMonth', this.value)">
            </div>
            <div class="boost-date-group">
              <span class="boost-date-label">Until</span>
              <input class="boost-month-input" type="date" value="${_monthToDateValue(b.toMonth)}"
                onchange="_setBoostDate(${i}, 'toMonth', this.value)">
            </div>
            <span class="boost-duration-tag ${isPermanent?'permanent':'temporary'}">
              ${isPermanent ? '∞ Permanent' : _boostDurationLabel(b.fromMonth, b.toMonth)}
            </span>
          </div>
        </div>
        <div class="boost-cell boost-goal-cell">
          <select class="boost-select" ${isReduction ? "disabled" : ""} onchange="_setBoostGoal(${i}, this.value)">
            <option value="any" ${b.toGoal==="any"?"selected":""}>${isReduction ? "Savings pool" : "Any goal (pool)"}</option>
            ${isReduction ? "" : goalsData.map(g=>`<option value="${escapeHtml(g.name)}" ${b.toGoal===g.name?"selected":""}>${escapeHtml(g.name)}</option>`).join("")}
          </select>
        </div>
        <div class="boost-cell boost-del-cell">
          <button class="boost-del-btn" onclick="_deleteIncomeBoost(${i})">✕</button>
        </div>
      </div>`;
  }).join("");

  panel.innerHTML = `
    <div class="panel-header">
      <span class="panel-icon">⚡</span>
      <h2>Future Cashflow Changes</h2>
      <span class="panel-hint">Schedule future changes that increase or reduce forecast savings</span>
    </div>
    <div class="boost-explainer">
      <div class="boost-ex-row"><span class="boost-ex-icon">＋</span><div><strong>Increase savings</strong> — income boost, ended subscription, reduced bill, bonus, etc.</div></div>
      <div class="boost-ex-row"><span class="boost-ex-icon">−</span><div><strong>Reduce savings</strong> — rent increase, new recurring bill, income drop, temporary higher expenses, etc.</div></div>
    </div>
    <div class="boost-table-header">
      <span class="boost-th boost-label-cell">What is it?</span>
      <span class="boost-th boost-kind-cell">Type</span>
      <span class="boost-th boost-amt-cell">Amount</span>
      <span class="boost-th boost-dates-cell">Active Period</span>
      <span class="boost-th boost-goal-cell">Affect</span>
      <span class="boost-th boost-del-cell"></span>
    </div>
    <div id="boostRows">${rows || '<div style="padding:14px 0;font-size:13px;color:var(--muted);">No future cashflow changes yet — add one below.</div>'}</div>
    <div style="display:flex;gap:10px;margin-top:14px;flex-wrap:wrap;align-items:center;border-top:1px solid var(--border);padding-top:14px;">
      <button class="btn-secondary btn-sm" onclick="_addIncomeBoost('boost')">
        + Add Increase
      </button>
      <button class="btn-secondary btn-sm" onclick="_addIncomeBoost('reduce')">
        − Add Reduction
      </button>
      <span class="panel-hint goals-autosave-status" id="boostAutosaveStatus">Autosaves to Excel</span>
    </div>
  `;
}

function _boostDurationLabel(fromMonth, toMonth) {
  fromMonth = _normaliseBoostMonth(fromMonth);
  toMonth = _normaliseBoostMonth(toMonth);

  if (!fromMonth || !toMonth) return "∞ Permanent";

  const [fy, fm] = fromMonth.split("-").map(Number);
  const [ty, tm] = toMonth.split("-").map(Number);
  const months = (ty - fy) * 12 + (tm - fm) + 1;

  if (!Number.isFinite(months) || months <= 0) return "⚠ Invalid";
  if (months === 1) return "1 month";
  return months + " months";
}

function buildGoalProjectionModel(minMonths = 18, maxMonths = 48) {
  const today = new Date();
  const forecastBaseMonthly = Math.max(0, historicalStats.avgMonthlySavings);
  const urgOrder = { Critical:0, High:1, Medium:2, Low:3 };

  function monthOffsetFromToday(dateValue) {
    if (!dateValue) return null;
    const p = String(dateValue).split("-");
    if (p.length < 2) return null;
    const y = Number(p[0]);
    const m = Number(p[1]) - 1;
    if (!Number.isFinite(y) || !Number.isFinite(m)) return null;
    return (y - today.getFullYear()) * 12 + (m - today.getMonth());
  }

  function monthLabel(offset, longForm=false) {
    return new Date(today.getFullYear(), today.getMonth() + offset, 1)
      .toLocaleDateString("en-SG", longForm ? { month:"short", year:"numeric" } : { month:"short", year:"2-digit" });
  }

  function relativeDeadlineText(completedMo, deadlineMo) {
    if (completedMo === null || deadlineMo === null) return "";
    const diff = deadlineMo - completedMo;
    if (diff > 0) return `${diff} month${diff === 1 ? "" : "s"} early`;
    if (diff === 0) return "on deadline";
    return `${Math.abs(diff)} month${Math.abs(diff) === 1 ? "" : "s"} late`;
  }

  function compareProjection(a, b) {
    const u = (urgOrder[a.gs.urgency] ?? 2) - (urgOrder[b.gs.urgency] ?? 2);
    if (u !== 0) return u;
    if (a.gs.deadlineMo !== null && b.gs.deadlineMo !== null) return a.gs.deadlineMo - b.gs.deadlineMo;
    if (a.gs.deadlineMo !== null) return -1;
    if (b.gs.deadlineMo !== null) return 1;
    return a.gi - b.gi;
  }

  const goalState = goalsData.map((g, gi) => {
    const txSaved = getSavedViaTransactions(g.name);
    const initialSaved = g.manualSaved + txSaved;
    const bufferPct = Number(g.goalBuffer || 0) || 0;
    const effectiveTarget = g.target * (1 + bufferPct / 100);
    const rawStartMo = monthOffsetFromToday(g.startDate) ?? 0;
    const startMo = Math.max(0, rawStartMo);
    const deadlineMo = monthOffsetFromToday(g.endDate);
    const savedMonthlyAlloc = Number(g.monthlyAlloc || 0) || 0;
    const suggestedNoDeadlineAlloc = deadlineMo === null && savedMonthlyAlloc <= 0 && forecastBaseMonthly > 0
      ? Math.round(forecastBaseMonthly * 0.3)
      : 0;

    return {
      idx: gi,
      name: g.name,
      color: g.color || "#4caf50",
      urgency: g.urgency || "Medium",
      baseTarget: g.target,
      bufferPct,
      effectiveTarget,
      initialSaved,
      saved: Math.min(initialSaved, effectiveTarget),
      monthlyAlloc: savedMonthlyAlloc || suggestedNoDeadlineAlloc,
      monthlyAllocIsSuggested: savedMonthlyAlloc <= 0 && suggestedNoDeadlineAlloc > 0,
      rawStartMo,
      startMo,
      deadlineMo,
      completedAt: initialSaved >= effectiveTarget ? 0 : null,
      baseCompletedAt: initialSaved >= g.target ? 0 : null,
      plannedTotal: 0,
      waitMonths: 0,
    };
  });

  const maxDeadline = goalState.reduce((mx, gs) => Math.max(mx, gs.deadlineMo ?? 0), 0);
  const maxCompletionGuess = goalState.reduce((mx, gs) => {
    const remaining = Math.max(0, gs.effectiveTarget - gs.initialSaved);
    const monthly = Math.max(1, gs.monthlyAlloc || forecastBaseMonthly || 1);
    return Math.max(mx, gs.startMo + Math.ceil(remaining / monthly));
  }, 0);
  const MONTHS = Math.max(minMonths, Math.min(maxMonths, Math.max(maxDeadline, maxCompletionGuess) + 6));

  const labels = Array.from({ length: MONTHS }, (_, m) => monthLabel(m));
  const fullLabels = Array.from({ length: MONTHS }, (_, m) => monthLabel(m, true));
  const poolBoost = Array(MONTHS).fill(0);
  const goalBoost = {};
  goalState.forEach(gs => { goalBoost[gs.name] = Array(MONTHS).fill(0); });

  incomeBoosts.forEach(b => {
    const fromMonth = _normaliseBoostMonth(b.fromMonth);
    if (!fromMonth || !b.amount) return;
    const fromMoRaw = monthOffsetFromToday(fromMonth);
    if (fromMoRaw === null) return;
    const fromMo = Math.max(0, fromMoRaw);
    const toMoRaw = b.toMonth ? monthOffsetFromToday(_normaliseBoostMonth(b.toMonth)) : MONTHS - 1;
    const toMo = Math.min(MONTHS - 1, toMoRaw === null ? MONTHS - 1 : toMoRaw);
    if (toMo < 0 || fromMo >= MONTHS || toMo < fromMo) return;

    const signedAmount = _boostSignedAmount(b);
    for (let m = fromMo; m <= toMo; m++) {
      if (signedAmount < 0 || b.toGoal === "any") {
        poolBoost[m] += signedAmount;
      } else if (goalBoost[b.toGoal]) {
        goalBoost[b.toGoal][m] += signedAmount;
      }
    }
  });

  const allocationData = goalState.map(() => Array(MONTHS).fill(0));
  const allocationBaseData = goalState.map(() => Array(MONTHS).fill(0));
  const allocationCashflowData = goalState.map(() => Array(MONTHS).fill(0));
  const progressDollars = goalState.map(() => Array(MONTHS).fill(0));
  const requiredMonthlyData = goalState.map(() => Array(MONTHS).fill(null));
  const unallocatedData = Array(MONTHS).fill(0);
  const unallocatedBaseData = Array(MONTHS).fill(0);
  const unallocatedCashflowData = Array(MONTHS).fill(0);
  const cashflowReductionData = Array(MONTHS).fill(0);
  const forecastStartMonth = 1;

  function allocateToGoal(gs, gi, m, amount, source="base") {
    const remaining = Math.max(0, gs.effectiveTarget - gs.saved);
    const used = Math.min(Math.max(0, amount), remaining);
    if (used <= 0) return 0;

    gs.saved += used;
    gs.plannedTotal += used;
    allocationData[gi][m] += used;
    if (source === "cashflow") {
      allocationCashflowData[gi][m] += used;
    } else {
      allocationBaseData[gi][m] += used;
    }

    if (gs.baseCompletedAt === null && gs.saved >= gs.baseTarget) gs.baseCompletedAt = m;
    if (gs.completedAt === null && gs.saved >= gs.effectiveTarget) {
      gs.saved = gs.effectiveTarget;
      gs.completedAt = m;
    }
    return used;
  }

  for (let m = 0; m < MONTHS; m++) {
    const isForecastMonth = m >= forecastStartMonth;
    let basePool = isForecastMonth ? Math.max(0, forecastBaseMonthly + Math.min(0, poolBoost[m])) : 0;
    let cashflowPool = isForecastMonth ? Math.max(0, poolBoost[m]) : 0;
    const monthPoolAvailable = basePool + cashflowPool;
    const monthPoolUsed = Array(goalState.length).fill(0);
    cashflowReductionData[m] = isForecastMonth ? -Math.max(0, -poolBoost[m]) : 0;

    const allocateFromForecastPool = (gs, gi, amount) => {
      let requested = Math.max(0, amount);
      let usedTotal = 0;

      if (requested > 0 && basePool > 0) {
        const baseUsed = allocateToGoal(gs, gi, m, Math.min(basePool, requested), "base");
        basePool -= baseUsed;
        requested -= baseUsed;
        usedTotal += baseUsed;
      }

      if (requested > 0 && cashflowPool > 0) {
        const cashflowUsed = allocateToGoal(gs, gi, m, Math.min(cashflowPool, requested), "cashflow");
        cashflowPool -= cashflowUsed;
        requested -= cashflowUsed;
        usedTotal += cashflowUsed;
      }

      return usedTotal;
    };

    goalState.forEach((gs, gi) => {
      if (!isForecastMonth || m < gs.startMo || gs.completedAt !== null) return;
      const used = allocateToGoal(gs, gi, m, goalBoost[gs.name][m] || 0, "cashflow");
      monthPoolUsed[gi] += used;
    });

    const active = goalState
      .map((gs, gi) => ({ gs, gi }))
      .filter(({ gs }) => m >= gs.startMo && gs.completedAt === null && gs.saved < gs.effectiveTarget);

    active.forEach(({ gs, gi }) => {
      if (gs.deadlineMo === null) return;
      const monthsLeft = Math.max(1, gs.deadlineMo - m + 1);
      requiredMonthlyData[gi][m] = Math.max(0, gs.effectiveTarget - gs.saved) / monthsLeft;
    });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0) return;
        const required = requiredMonthlyData[gi][m] || 0;
        const desired = Math.max(required, gs.monthlyAlloc || 0);
        if (desired <= 0) return;
        const used = allocateFromForecastPool(gs, gi, Math.min(availablePool, desired));
        monthPoolUsed[gi] += used;
      });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0 || gs.completedAt !== null) return;
        const desired = Math.max(0, (gs.monthlyAlloc || 0) - monthPoolUsed[gi]);
        const used = allocateFromForecastPool(gs, gi, Math.min(availablePool, desired));
        monthPoolUsed[gi] += used;
      });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0 || gs.completedAt !== null) return;
        const used = allocateFromForecastPool(gs, gi, availablePool);
        monthPoolUsed[gi] += used;
      });

    unallocatedBaseData[m] = Math.max(0, basePool);
    unallocatedCashflowData[m] = Math.max(0, cashflowPool);
    unallocatedData[m] = unallocatedBaseData[m] + unallocatedCashflowData[m];

    goalState.forEach((gs, gi) => {
      if (m >= gs.startMo && gs.completedAt === null && monthPoolUsed[gi] <= 0 && monthPoolAvailable > 0) {
        gs.waitMonths += 1;
      }
      progressDollars[gi][m] = gs.saved;
    });
  }

  return {
    today,
    forecastBaseMonthly,
    goalState,
    MONTHS,
    labels,
    fullLabels,
    poolBoost,
    allocationData,
    allocationBaseData,
    allocationCashflowData,
    progressDollars,
    requiredMonthlyData,
    unallocatedData,
    unallocatedBaseData,
    unallocatedCashflowData,
    cashflowReductionData,
    forecastStartMonth,
    monthOffsetFromToday,
    monthLabel,
    relativeDeadlineText,
  };
}

function renderGoalGanttTimeline(container) {
  const model = buildGoalProjectionModel(18, 48);
  const { today, goalState, MONTHS, progressDollars, forecastBaseMonthly, poolBoost } = model;

  const finiteMax = values => values.filter(Number.isFinite).reduce((mx, v) => Math.max(mx, v), 0);
  const finiteMin = values => values.filter(Number.isFinite).reduce((mn, v) => Math.min(mn, v), 0);
  const startMoMin = finiteMin(goalState.map(gs => gs.rawStartMo));
  const displayStart = Math.max(-6, Math.min(0, startMoMin));
  const displayEnd = Math.min(
    MONTHS - 1,
    Math.max(12, finiteMax(goalState.map(gs => Math.max(gs.deadlineMo ?? 0, gs.completedAt ?? 0, gs.startMo + 2))) + 2)
  );
  const offsets = [];
  for (let mo = displayStart; mo <= displayEnd; mo++) offsets.push(mo);
  const monthCount = offsets.length || 1;
  const monthWidth = 74;
  const axisWidth = monthCount * monthWidth;
  const labelWidth = 230;
  const statusWidth = 280;
  const ganttMinWidth = labelWidth + axisWidth + statusWidth;
  const ganttStyle = `--gantt-axis-width:${axisWidth}px;--gantt-label-width:${labelWidth}px;--gantt-status-width:${statusWidth}px;min-width:${ganttMinWidth}px;`;
  const gridStyle = `--gantt-months:${monthCount};grid-template-columns:repeat(${monthCount}, minmax(${monthWidth}px, 1fr));`;
  const viewStartDate = new Date(today.getFullYear(), today.getMonth() + displayStart, 1);
  const viewEndDate = new Date(today.getFullYear(), today.getMonth() + displayEnd + 1, 1);
  const viewSpanMs = Math.max(1, viewEndDate - viewStartDate);

  function dateForMonthOffset(mo, mode="start") {
    if (mode === "end") return new Date(today.getFullYear(), today.getMonth() + mo + 1, 0);
    if (mode === "middle") return new Date(today.getFullYear(), today.getMonth() + mo, 15);
    return new Date(today.getFullYear(), today.getMonth() + mo, 1);
  }

  function pctForDate(dateObj) {
    if (!dateObj) return null;
    const d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
    return Math.max(0, Math.min(100, ((d - viewStartDate) / viewSpanMs) * 100));
  }

  function dateWithinView(dateObj) {
    if (!dateObj) return false;
    return dateObj >= viewStartDate && dateObj <= viewEndDate;
  }

  function pctStyle(value) {
    return `${Number(value).toFixed(3)}%`;
  }

  function ganttDateLabel(dateObj) {
    if (!dateObj) return "";
    return dateObj.toLocaleDateString("en-SG", { day:"2-digit", month:"short", year:"numeric" });
  }

  function forecastFundingDateFor(gs, gi) {
    if (gs.completedAt === null) return null;

    const completedMo = Math.max(0, Math.min(MONTHS - 1, gs.completedAt));
    const monthStart = new Date(today.getFullYear(), today.getMonth() + completedMo, 1);
    const daysInMonth = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
    const savedBeforeMonth = completedMo > 0
      ? (progressDollars[gi][completedMo - 1] || gs.initialSaved)
      : gs.initialSaved;
    const remainingAtMonthStart = Math.max(0, gs.effectiveTarget - savedBeforeMonth);
    const allocatedThisMonth = Math.max(0, model.allocationData?.[gi]?.[completedMo] || 0);

    if (remainingAtMonthStart <= 0) return monthStart;
    if (allocatedThisMonth <= 0) return dateForMonthOffset(completedMo, "middle");

    const fundedDay = Math.max(1, Math.min(daysInMonth, Math.ceil((remainingAtMonthStart / allocatedThisMonth) * daysInMonth)));
    return new Date(monthStart.getFullYear(), monthStart.getMonth(), fundedDay);
  }

  function ganttStatus(gs, gi) {
    const completion = gs.completedAt !== null ? model.monthLabel(gs.completedAt, true) : null;
    const deadline = gs.deadlineMo !== null ? model.monthLabel(gs.deadlineMo, true) : null;
    const originalGoal = goalsData[gs.idx] || {};
    const deadlineDate = _parseGoalDateValue(originalGoal.endDate);
    const forecastDate = forecastFundingDateFor(gs, gi);

    if (gs.initialSaved >= gs.effectiveTarget) {
      return { cls:"ok", pill:"Already funded", detail:"Target and buffer are fully covered." };
    }

    if (gs.deadlineMo !== null && gs.deadlineMo < 0) {
      return { cls:"danger", pill:"Overdue", detail:"Deadline has passed and this is not fully funded." };
    }

    if (gs.deadlineMo !== null) {
      const deadlineIdx = Math.max(0, Math.min(MONTHS - 1, gs.deadlineMo));
      const savedByDeadline = progressDollars[gi][deadlineIdx] || gs.initialSaved;
      const baseShort = Math.max(0, gs.baseTarget - savedByDeadline);
      const bufferShort = Math.max(0, gs.effectiveTarget - savedByDeadline);

      if (baseShort > 0) {
        return { cls:"danger", pill:"Misses deadline", detail:`Short ${formatCurrency(baseShort)} by ${deadline}.` };
      }
      if (bufferShort > 0) {
        return { cls:"warn", pill:"Base ok, buffer late", detail:completion ? `Buffer completes ${completion}.` : "Only the extra buffer is at risk." };
      }
      if (forecastDate && deadlineDate) {
        if (forecastDate <= deadlineDate) {
          return { cls:"ok", pill:"Possible", detail:`Forecast funded around ${ganttDateLabel(forecastDate)}, before the ${ganttDateLabel(deadlineDate)} deadline.` };
        }
        if (gs.completedAt !== null && gs.completedAt <= gs.deadlineMo) {
          return { cls:"warn", pill:"Exact-date risk", detail:`Forecast funds around ${ganttDateLabel(forecastDate)}, after the ${ganttDateLabel(deadlineDate)} deadline.` };
        }
      }
      if (gs.completedAt !== null && gs.completedAt <= gs.deadlineMo) {
        const timing = model.relativeDeadlineText(gs.completedAt, gs.deadlineMo);
        const timingLabel = timing === "on deadline" ? "deadline month" : timing;
        return { cls:"ok", pill:"Possible", detail:`Forecast funded ${completion}${timingLabel ? ` (${timingLabel})` : ""}.` };
      }
      return { cls:"ok", pill:"Deadline covered", detail:`Deadline ${deadline} is covered by forecast.` };
    }

    if (gs.completedAt !== null) {
      return { cls:"neutral", pill:"No deadline", detail:forecastDate ? `Forecast funded around ${ganttDateLabel(forecastDate)}.` : `Forecast funded ${completion}.` };
    }

    return { cls:"neutral", pill:"No deadline", detail:"Set a deadline or allocation to judge timing." };
  }

  const displayGoals = goalState
    .map((gs, gi) => ({ gs, gi }))
    .sort((a, b) => {
      if (a.gs.deadlineMo !== null && b.gs.deadlineMo !== null && a.gs.deadlineMo !== b.gs.deadlineMo) return a.gs.deadlineMo - b.gs.deadlineMo;
      if (a.gs.deadlineMo !== null) return -1;
      if (b.gs.deadlineMo !== null) return 1;
      if (a.gs.rawStartMo !== b.gs.rawStartMo) return a.gs.rawStartMo - b.gs.rawStartMo;
      return a.gi - b.gi;
    });

  const rows = displayGoals.map(({ gs, gi }) => {
    const originalGoal = goalsData[gs.idx];
    const status = ganttStatus(gs, gi);
    const parsedStartDate = _parseGoalDateValue(originalGoal.startDate);
    const parsedDeadlineDate = _parseGoalDateValue(originalGoal.endDate);
    const startDateObj = parsedStartDate || today;
    const endDateObj = parsedDeadlineDate || (gs.completedAt !== null ? dateForMonthOffset(gs.completedAt, "end") : dateForMonthOffset(displayEnd, "end"));
    const barStartDate = startDateObj <= endDateObj ? startDateObj : endDateObj;
    const barEndDate = startDateObj <= endDateObj ? endDateObj : startDateObj;
    let barLeftPct = pctForDate(barStartDate) ?? 0;
    const barRightPct = pctForDate(barEndDate) ?? 100;
    let barWidthPct = Math.max(0.6, barRightPct - barLeftPct);
    if (barLeftPct + barWidthPct > 100) barLeftPct = Math.max(0, 100 - barWidthPct);
    const todayPct = dateWithinView(today) ? pctForDate(today) : null;
    const deadlinePct = parsedDeadlineDate && dateWithinView(parsedDeadlineDate) ? pctForDate(parsedDeadlineDate) : null;
    const completionDate = forecastFundingDateFor(gs, gi);
    const completionPct = completionDate && dateWithinView(completionDate) ? pctForDate(completionDate) : null;
    const startLabel = originalGoal.startDate ? formatDateDisplay(originalGoal.startDate) : "Today";
    const endLabel = originalGoal.endDate
      ? formatDateDisplay(originalGoal.endDate)
      : (gs.completedAt !== null ? `Projected ${model.monthLabel(gs.completedAt, true)}` : "No deadline");
    const completionLabel = completionDate ? ganttDateLabel(completionDate) : "";
    const targetLabel = gs.bufferPct > 0
      ? `${formatCurrency(gs.effectiveTarget)} incl. ${gs.bufferPct}% buffer`
      : formatCurrency(gs.effectiveTarget);

    return `
      <div class="gantt-row">
        <div class="gantt-label">
          <div class="gantt-title">
            <span class="gantt-dot" style="background:${gs.color};"></span>
            <strong>${escapeHtml(gs.name)}</strong>
          </div>
          <div class="gantt-meta">
            <span>${escapeHtml(gs.urgency)}</span>
            <span>${targetLabel}</span>
          </div>
        </div>
        <div class="gantt-track" style="${gridStyle}">
          ${todayPct !== null ? `<span class="gantt-line today" style="left:${pctStyle(todayPct)};" title="Today"></span>` : ""}
          <span class="gantt-bar ${status.cls}" style="left:${pctStyle(barLeftPct)};width:${pctStyle(barWidthPct)};background:${gs.color};" title="${escapeHtml(startLabel)} to ${escapeHtml(endLabel)}"></span>
          ${deadlinePct !== null ? `<span class="gantt-line deadline" style="left:${pctStyle(deadlinePct)};color:${gs.color};" title="Deadline: ${escapeHtml(endLabel)}"></span>` : ""}
          ${completionPct !== null ? `<span class="gantt-complete-marker" style="left:${pctStyle(completionPct)};color:${status.cls === "danger" ? "#dc2626" : gs.color};" title="Forecast funded around: ${escapeHtml(completionLabel)}"></span>` : ""}
        </div>
        <div class="gantt-status">
          <span class="gantt-pill ${status.cls}">${status.pill}</span>
          <small>${status.detail}</small>
          <small>${escapeHtml(startLabel)} → ${escapeHtml(endLabel)}</small>
        </div>
      </div>`;
  }).join("");

  const counts = goalState.reduce((acc, gs, gi) => {
    const cls = ganttStatus(gs, gi).cls;
    acc[cls] = (acc[cls] || 0) + 1;
    return acc;
  }, {});
  const totalForecastUnassigned = model.unallocatedData.reduce((sum, v) => sum + v, 0);
  const monthHeader = offsets.map(mo => `<span class="gantt-month">${model.monthLabel(mo)}</span>`).join("");
  const forecastStartsLabel = model.monthLabel(model.forecastStartMonth || 1, true);
  const poolText = `${formatCurrency(forecastBaseMonthly)}/mo${poolBoost.some(v => v > 0) ? " + increases" : ""}${poolBoost.some(v => v < 0) ? " - reductions" : ""}`;

  const panelEl = document.createElement("div");
  panelEl.className = "goals-panel gantt-panel";
  panelEl.innerHTML = `
    <div class="panel-header">
      <span class="panel-icon">▤</span>
      <h2>Goal Date Timeline</h2>
      <span class="panel-hint">Current month is context; forecast funding starts ${forecastStartsLabel}</span>
    </div>
    <div class="gantt-summary">
      <span class="gantt-stat"><strong>${counts.ok || 0}</strong> possible</span>
      <span class="gantt-stat"><strong>${counts.warn || 0}</strong> watch items</span>
      <span class="gantt-stat"><strong>${counts.danger || 0}</strong> deadline risk</span>
      <span class="gantt-stat"><strong>${formatCurrency(totalForecastUnassigned)}</strong> unassigned forecast</span>
      <span class="gantt-stat"><strong>${poolText}</strong> forecast pool from ${forecastStartsLabel}</span>
    </div>
    <div class="gantt-scroll">
      <div class="gantt" style="${ganttStyle}">
        <div class="gantt-months">
          <div class="gantt-corner">Goal</div>
          <div class="gantt-month-grid" style="${gridStyle}">${monthHeader}</div>
          <div class="gantt-status-head">Forecast status</div>
        </div>
        ${rows}
      </div>
    </div>
    <div class="gantt-legend">
      <span><i class="sample-bar"></i> start → deadline window</span>
      <span><i class="sample-line"></i> deadline</span>
      <span><i class="sample-dot"></i> forecast funded estimate</span>
      <span><i class="sample-line" style="border-left-style:solid;border-color:#111827;"></i> today</span>
    </div>
  `;
  container.appendChild(panelEl);
}

function renderSavingsTimelineStacked(container, dep) {
  const today = new Date();
  const forecastBaseMonthly = Math.max(0, historicalStats.avgMonthlySavings);
  const urgOrder = { Critical:0, High:1, Medium:2, Low:3 };

  function monthOffsetFromToday(dateValue) {
    if (!dateValue) return null;
    const p = String(dateValue).split("-");
    if (p.length < 2) return null;
    const y = Number(p[0]);
    const m = Number(p[1]) - 1;
    if (!Number.isFinite(y) || !Number.isFinite(m)) return null;
    return (y - today.getFullYear()) * 12 + (m - today.getMonth());
  }

  function monthLabel(offset, longForm=false) {
    return new Date(today.getFullYear(), today.getMonth() + offset, 1)
      .toLocaleDateString("en-SG", longForm ? { month:"short", year:"numeric" } : { month:"short", year:"2-digit" });
  }

  function relativeDeadlineText(completedMo, deadlineMo) {
    if (completedMo === null || deadlineMo === null) return "";
    const diff = deadlineMo - completedMo;
    if (diff > 0) return `${diff} month${diff === 1 ? "" : "s"} early`;
    if (diff === 0) return "on deadline";
    return `${Math.abs(diff)} month${Math.abs(diff) === 1 ? "" : "s"} late`;
  }

  function compareProjection(a, b) {
    const u = (urgOrder[a.gs.urgency] ?? 2) - (urgOrder[b.gs.urgency] ?? 2);
    if (u !== 0) return u;
    if (a.gs.deadlineMo !== null && b.gs.deadlineMo !== null) return a.gs.deadlineMo - b.gs.deadlineMo;
    if (a.gs.deadlineMo !== null) return -1;
    if (b.gs.deadlineMo !== null) return 1;
    return a.gi - b.gi;
  }

  const goalState = goalsData.map((g, gi) => {
    const txSaved = getSavedViaTransactions(g.name);
    const initialSaved = g.manualSaved + txSaved;
    const bufferPct = Number(g.goalBuffer || 0) || 0;
    const effectiveTarget = g.target * (1 + bufferPct / 100);
    const startMo = Math.max(0, monthOffsetFromToday(g.startDate) ?? 0);
    const deadlineMo = monthOffsetFromToday(g.endDate);
    const savedMonthlyAlloc = Number(g.monthlyAlloc || 0) || 0;
    const suggestedNoDeadlineAlloc = deadlineMo === null && savedMonthlyAlloc <= 0 && forecastBaseMonthly > 0
      ? Math.round(forecastBaseMonthly * 0.3)
      : 0;

    return {
      idx: gi,
      name: g.name,
      color: g.color || "#4caf50",
      urgency: g.urgency || "Medium",
      baseTarget: g.target,
      bufferPct,
      effectiveTarget,
      initialSaved,
      saved: Math.min(initialSaved, effectiveTarget),
      monthlyAlloc: savedMonthlyAlloc || suggestedNoDeadlineAlloc,
      monthlyAllocIsSuggested: savedMonthlyAlloc <= 0 && suggestedNoDeadlineAlloc > 0,
      startMo,
      deadlineMo,
      completedAt: initialSaved >= effectiveTarget ? 0 : null,
      baseCompletedAt: initialSaved >= g.target ? 0 : null,
      plannedTotal: 0,
      waitMonths: 0,
    };
  });

  const maxDeadline = goalState.reduce((mx, gs) => Math.max(mx, gs.deadlineMo ?? 0), 0);
  const MONTHS = Math.max(18, Math.min(36, maxDeadline + 6));
  const labels = Array.from({ length: MONTHS }, (_, m) => monthLabel(m));
  const fullLabels = Array.from({ length: MONTHS }, (_, m) => monthLabel(m, true));

  const poolBoost = Array(MONTHS).fill(0);
  const goalBoost = {};
  goalState.forEach(gs => { goalBoost[gs.name] = Array(MONTHS).fill(0); });

  incomeBoosts.forEach(b => {
    const fromMonth = _normaliseBoostMonth(b.fromMonth);
    if (!fromMonth || !b.amount) return;
    const fromMoRaw = monthOffsetFromToday(fromMonth);
    if (fromMoRaw === null) return;
    const fromMo = Math.max(0, fromMoRaw);
    const toMoRaw = b.toMonth ? monthOffsetFromToday(_normaliseBoostMonth(b.toMonth)) : MONTHS - 1;
    const toMo = Math.min(MONTHS - 1, toMoRaw === null ? MONTHS - 1 : toMoRaw);
    if (toMo < 0 || fromMo >= MONTHS || toMo < fromMo) return;

    const signedAmount = _boostSignedAmount(b);
    for (let m = fromMo; m <= toMo; m++) {
      if (signedAmount < 0 || b.toGoal === "any") {
        poolBoost[m] += signedAmount;
      } else if (goalBoost[b.toGoal]) {
        goalBoost[b.toGoal][m] += signedAmount;
      }
    }
  });

  const allocationData = goalState.map(() => Array(MONTHS).fill(0));
  const allocationBaseData = goalState.map(() => Array(MONTHS).fill(0));
  const allocationCashflowData = goalState.map(() => Array(MONTHS).fill(0));
  const progressDollars = goalState.map(() => Array(MONTHS).fill(0));
  const requiredMonthlyData = goalState.map(() => Array(MONTHS).fill(null));
  const unallocatedData = Array(MONTHS).fill(0);
  const unallocatedBaseData = Array(MONTHS).fill(0);
  const unallocatedCashflowData = Array(MONTHS).fill(0);
  const cashflowReductionData = Array(MONTHS).fill(0);
  const forecastStartMonth = 1;

  function allocateToGoal(gs, gi, m, amount, source="base") {
    const remaining = Math.max(0, gs.effectiveTarget - gs.saved);
    const used = Math.min(Math.max(0, amount), remaining);
    if (used <= 0) return 0;

    gs.saved += used;
    gs.plannedTotal += used;
    allocationData[gi][m] += used;
    if (source === "cashflow") {
      allocationCashflowData[gi][m] += used;
    } else {
      allocationBaseData[gi][m] += used;
    }

    if (gs.baseCompletedAt === null && gs.saved >= gs.baseTarget) gs.baseCompletedAt = m;
    if (gs.completedAt === null && gs.saved >= gs.effectiveTarget) {
      gs.saved = gs.effectiveTarget;
      gs.completedAt = m;
    }
    return used;
  }

  for (let m = 0; m < MONTHS; m++) {
    const isForecastMonth = m >= forecastStartMonth;
    let basePool = isForecastMonth ? Math.max(0, forecastBaseMonthly + Math.min(0, poolBoost[m])) : 0;
    let cashflowPool = isForecastMonth ? Math.max(0, poolBoost[m]) : 0;
    const monthPoolAvailable = basePool + cashflowPool;
    const monthPoolUsed = Array(goalState.length).fill(0);
    cashflowReductionData[m] = isForecastMonth ? -Math.max(0, -poolBoost[m]) : 0;

    const allocateFromForecastPool = (gs, gi, amount) => {
      let requested = Math.max(0, amount);
      let usedTotal = 0;

      if (requested > 0 && basePool > 0) {
        const baseUsed = allocateToGoal(gs, gi, m, Math.min(basePool, requested), "base");
        basePool -= baseUsed;
        requested -= baseUsed;
        usedTotal += baseUsed;
      }

      if (requested > 0 && cashflowPool > 0) {
        const cashflowUsed = allocateToGoal(gs, gi, m, Math.min(cashflowPool, requested), "cashflow");
        cashflowPool -= cashflowUsed;
        requested -= cashflowUsed;
        usedTotal += cashflowUsed;
      }

      return usedTotal;
    };

    goalState.forEach((gs, gi) => {
      if (!isForecastMonth || m < gs.startMo || gs.completedAt !== null) return;
      const used = allocateToGoal(gs, gi, m, goalBoost[gs.name][m] || 0, "cashflow");
      monthPoolUsed[gi] += used;
    });

    const active = goalState
      .map((gs, gi) => ({ gs, gi }))
      .filter(({ gs }) => m >= gs.startMo && gs.completedAt === null && gs.saved < gs.effectiveTarget);

    active.forEach(({ gs, gi }) => {
      if (gs.deadlineMo === null) return;
      const monthsLeft = Math.max(1, gs.deadlineMo - m + 1);
      requiredMonthlyData[gi][m] = Math.max(0, gs.effectiveTarget - gs.saved) / monthsLeft;
    });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0) return;
        const required = requiredMonthlyData[gi][m] || 0;
        const desired = Math.max(required, gs.monthlyAlloc || 0);
        if (desired <= 0) return;
        const used = allocateFromForecastPool(gs, gi, Math.min(availablePool, desired));
        monthPoolUsed[gi] += used;
      });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0 || gs.completedAt !== null) return;
        const desired = Math.max(0, (gs.monthlyAlloc || 0) - monthPoolUsed[gi]);
        const used = allocateFromForecastPool(gs, gi, Math.min(availablePool, desired));
        monthPoolUsed[gi] += used;
      });

    active
      .sort(compareProjection)
      .forEach(({ gs, gi }) => {
        const availablePool = basePool + cashflowPool;
        if (availablePool <= 0 || gs.completedAt !== null) return;
        const used = allocateFromForecastPool(gs, gi, availablePool);
        monthPoolUsed[gi] += used;
      });

    unallocatedBaseData[m] = Math.max(0, basePool);
    unallocatedCashflowData[m] = Math.max(0, cashflowPool);
    unallocatedData[m] = unallocatedBaseData[m] + unallocatedCashflowData[m];

    goalState.forEach((gs, gi) => {
      if (m >= gs.startMo && gs.completedAt === null && monthPoolUsed[gi] <= 0 && monthPoolAvailable > 0) {
        gs.waitMonths += 1;
      }
      progressDollars[gi][m] = gs.saved;
    });
  }

  const forecastMonthIndexes = Array.from(
    { length: Math.max(0, MONTHS - forecastStartMonth) },
    (_, i) => i + forecastStartMonth
  );
  const forecastMonthCount = Math.max(1, forecastMonthIndexes.length);
  const chartLabels = forecastMonthIndexes.map(m => labels[m]);
  const chartFullLabels = forecastMonthIndexes.map(m => fullLabels[m]);
  const totalUnallocatedForecast = unallocatedData.reduce((sum, v) => sum + v, 0);
  const positiveCashflowTotal = allocationCashflowData.reduce((sum, arr) => sum + arr.reduce((s, v) => s + v, 0), 0)
    + unallocatedCashflowData.reduce((sum, v) => sum + v, 0);
  const maxCashflowReduction = Math.max(0, ...cashflowReductionData.map(v => Math.abs(v)));
  const hasCashflowAdditions = positiveCashflowTotal > 0;
  const hasCashflowReductions = maxCashflowReduction > 0;
  const hasCashflowChanges = hasCashflowAdditions || hasCashflowReductions;

  function activeCashflowChangesForMonth(monthIndex) {
    return incomeBoosts
      .map(item => {
        const fromMonth = _normaliseBoostMonth(item.fromMonth);
        if (!fromMonth || !item.amount) return null;
        const fromMoRaw = monthOffsetFromToday(fromMonth);
        if (fromMoRaw === null) return null;
        const fromMo = Math.max(0, fromMoRaw);
        const toMoRaw = item.toMonth ? monthOffsetFromToday(_normaliseBoostMonth(item.toMonth)) : MONTHS - 1;
        const toMo = Math.min(MONTHS - 1, toMoRaw === null ? MONTHS - 1 : toMoRaw);
        if (monthIndex < fromMo || monthIndex > toMo) return null;
        return {
          label: item.label || "Cashflow change",
          amount: _boostSignedAmount(item),
          toGoal: item.toGoal || "any",
        };
      })
      .filter(Boolean);
  }

  let legendHtml = goalState.map(gs => `
    <div class="tl-legend-item">
      <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:${gs.color};flex-shrink:0;"></span>
      <span>${escapeHtml(gs.name)} base</span>
    </div>`).join("") + `
    <div class="tl-legend-item tl-legend-unassigned">
      <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:#94a3b8;flex-shrink:0;"></span>
      <span>Available for new goals</span>
    </div>`;
  if (hasCashflowAdditions) {
    legendHtml += `
      <div class="tl-legend-item">
        <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:#7c3aed;flex-shrink:0;"></span>
        <span>Future cashflow added</span>
      </div>`;
  }
  if (hasCashflowReductions) {
    legendHtml += `
      <div class="tl-legend-item">
        <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:#dc2626;flex-shrink:0;"></span>
        <span>Future cashflow reduction</span>
      </div>`;
  }

  const impactRows = goalState
    .map((gs, gi) => {
      const completion = gs.completedAt !== null ? fullLabels[Math.min(gs.completedAt, MONTHS - 1)] : null;
      const lastSaved = progressDollars[gi][MONTHS - 1] || gs.saved;
      const targetShort = Math.max(0, gs.baseTarget - lastSaved);
      const bufferShort = Math.max(0, gs.effectiveTarget - lastSaved);
      const currentRequired = gs.deadlineMo !== null
        ? Math.max(0, gs.effectiveTarget - gs.initialSaved) / Math.max(1, gs.deadlineMo - forecastStartMonth + 1)
        : null;
      const avgPlanned = gs.plannedTotal / forecastMonthCount;

      let statusClass = "ok";
      let status = "";
      let impact = "";
      const deadlineLabel = gs.deadlineMo !== null ? fullLabels[Math.max(0, Math.min(gs.deadlineMo, MONTHS - 1))] : null;
      const timingText = relativeDeadlineText(gs.completedAt, gs.deadlineMo);

      if (gs.deadlineMo !== null && gs.deadlineMo < 0 && gs.initialSaved < gs.effectiveTarget) {
        statusClass = "danger";
        status = "Deadline already passed";
        impact = "This stays at the front of the forecast queue until it catches up.";
      } else if (gs.deadlineMo !== null && gs.deadlineMo < MONTHS) {
        const deadlineSaved = progressDollars[gi][Math.max(0, gs.deadlineMo)] || gs.initialSaved;
        const baseShortAtDeadline = Math.max(0, gs.baseTarget - deadlineSaved);
        const bufferShortAtDeadline = Math.max(0, gs.effectiveTarget - deadlineSaved);

        if (baseShortAtDeadline > 0) {
          statusClass = "danger";
          status = `Misses deadline by ${formatCurrency(baseShortAtDeadline)}`;
          impact = "It keeps taking forecast savings after the deadline, delaying lower-priority goals.";
        } else if (gs.bufferPct > 0 && bufferShortAtDeadline > 0) {
          statusClass = "warn";
          status = `Base target met; buffer short ${formatCurrency(bufferShortAtDeadline)}`;
          impact = completion
            ? `Buffer completes ${completion}${timingText ? ` (${timingText}; deadline ${deadlineLabel})` : ""}.`
            : "Only the extra buffer is behind, not the original target.";
        } else {
          status = timingText && timingText !== "on deadline" ? `Deadline covered ${timingText}` : "Deadline covered";
          impact = completion
            ? `Completes ${completion}${deadlineLabel ? `; deadline ${deadlineLabel}` : ""}.`
            : "Target remains covered through the forecast.";
        }
      } else if (completion) {
        status = gs.bufferPct > 0 && gs.baseCompletedAt === 0 && gs.completedAt > 0
          ? `Base target met; buffer completes ${completion}`
          : `Completes ${completion}`;
        impact = gs.waitMonths > 0
          ? `Waits ${gs.waitMonths} month${gs.waitMonths === 1 ? "" : "s"} while earlier goals are funded.`
          : "Gets funding immediately in the forecast.";
      } else {
        statusClass = "danger";
        status = `Still short ${formatCurrency(bufferShort || targetShort)}`;
        impact = "Not complete within this forecast window.";
      }

      if (!impact && gs.waitMonths > 0) {
        impact = `Waits ${gs.waitMonths} month${gs.waitMonths === 1 ? "" : "s"} while earlier goals are funded.`;
      }

      return `
        <div class="timeline-impact-row ${statusClass}">
          <div class="ti-goal">
            <span class="ti-dot" style="background:${gs.color};"></span>
            <div>
              <strong>${escapeHtml(gs.name)}</strong>
              <span>${gs.bufferPct > 0 ? `Target + ${gs.bufferPct}% buffer` : "Target"} ${formatCurrency(gs.effectiveTarget)}</span>
            </div>
          </div>
          <div class="ti-metric"><span>Required</span><strong>${currentRequired ? formatCurrency(currentRequired) + "/mo" : "No deadline"}</strong></div>
          <div class="ti-metric"><span>Forecast avg</span><strong>${formatCurrency(avgPlanned)}/mo</strong></div>
          <div class="ti-status">
            <strong>${status}</strong>
            <span>${impact}</span>
          </div>
        </div>`;
    })
    .join("");

  const panelEl = document.createElement("div");
  panelEl.className = "goals-panel timeline-panel";
  panelEl.innerHTML = `
    <div class="panel-header">
      <span class="panel-icon">📈</span>
      <h2>Savings Timeline</h2>
      <span class="panel-hint">Forecast starts next month; current month allocations stay in saved balances</span>
    </div>
    <div class="timeline-forecast-note">
      Base forecast savings from ${chartFullLabels[0] || "next month"}: <strong>${formatCurrency(forecastBaseMonthly)}/mo</strong>
      ${hasCashflowChanges ? " · future cashflow is shown as purple/red sections within the bars" : ""}
      ${totalUnallocatedForecast > 0 ? ` · available for new goals over forecast: <strong>${formatCurrency(totalUnallocatedForecast)}</strong>` : ""}
    </div>
    <div class="timeline-legend-custom" style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;margin-bottom:4px;">${legendHtml}</div>
    <div class="chart-wrap timeline-chart-simple" style="height:300px;position:relative;">
      <canvas id="goalTimelineChart"></canvas>
    </div>
    <div class="timeline-impact-list">${impactRows}</div>
  `;
  container.appendChild(panelEl);

  requestAnimationFrame(() => {
    const ctx = document.getElementById("goalTimelineChart");
    if (!ctx || typeof Chart === "undefined") return;
    if (timelineChart) { timelineChart.destroy(); timelineChart = null; }

    const datasets = [];
    goalState.forEach((gs, gi) => {
      datasets.push({
        label: gs.name,
        data: forecastMonthIndexes.map(m => Number((allocationBaseData[gi][m] || 0).toFixed(2))),
        backgroundColor: gs.color,
        borderWidth: 0,
        borderRadius: 4,
        stack: "forecast",
        _goalIdx: gi,
        _source: "base",
      });
    });
    datasets.push({
      label: "Available for new goals",
      data: forecastMonthIndexes.map(m => Number((unallocatedBaseData[m] || 0).toFixed(2))),
      backgroundColor: "rgba(148,163,184,0.55)",
      borderColor: "#94a3b8",
      borderWidth: 1,
      borderRadius: 4,
      stack: "forecast",
      _isUnassigned: true,
      _source: "base",
    });
    if (hasCashflowAdditions) {
      goalState.forEach((gs, gi) => {
        datasets.push({
          label: `${gs.name} future cashflow`,
          data: forecastMonthIndexes.map(m => Number((allocationCashflowData[gi][m] || 0).toFixed(2))),
          backgroundColor: "rgba(124,58,237,0.82)",
          borderColor: gs.color,
          borderWidth: 1,
          borderRadius: 4,
          stack: "forecast",
          _goalIdx: gi,
          _source: "cashflow",
        });
      });
      datasets.push({
        label: "Future cashflow available",
        data: forecastMonthIndexes.map(m => Number((unallocatedCashflowData[m] || 0).toFixed(2))),
        backgroundColor: "rgba(124,58,237,0.38)",
        borderColor: "#7c3aed",
        borderWidth: 1,
        borderRadius: 4,
        stack: "forecast",
        _isUnassigned: true,
        _source: "cashflow",
      });
    }
    if (hasCashflowReductions) {
      datasets.push({
        label: "Future cashflow reduction",
        data: forecastMonthIndexes.map(m => Number((cashflowReductionData[m] || 0).toFixed(2))),
        backgroundColor: "rgba(220,38,38,0.32)",
        borderColor: "#dc2626",
        borderWidth: 1,
        borderRadius: 4,
        stack: "forecast",
        _isCashflowReduction: true,
      });
    }

    timelineChart = new Chart(ctx, {
      type: "bar",
      data: { labels: chartLabels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { display: false },
          datalabels: { display: false },
          tooltip: {
            backgroundColor: "rgba(15,23,42,0.92)",
            titleFont: { size: 12, weight: "600" },
            bodyFont: { size: 12 },
            padding: 12,
            callbacks: {
              title: items => chartFullLabels[items[0].dataIndex],
              label: item => {
                const rawValue = Number(item.raw ?? item.parsed.y ?? 0);
                const value = Math.abs(rawValue);
                if (value <= 0) return null;
                if (item.dataset._isCashflowReduction) return `Future cashflow reduction: -${formatCurrency(value)}/mo`;
                if (item.dataset._isUnassigned && item.dataset._source === "cashflow") return `Future cashflow available: ${formatCurrency(value)}`;
                if (item.dataset._isUnassigned) return `Available for new goals: ${formatCurrency(value)}`;
                if (item.dataset._source === "cashflow") return `${goalState[item.dataset._goalIdx]?.name || item.dataset.label}: ${formatCurrency(value)} from future cashflow`;
                return `${item.dataset.label}: ${formatCurrency(value)} from base forecast`;
              },
              footer: items => {
                const m = forecastMonthIndexes[items[0].dataIndex] ?? forecastStartMonth;
                const allocated = allocationData.reduce((sum, arr) => sum + (arr[m] || 0), 0);
                const barTotal = allocated + (unallocatedData[m] || 0);
                const baseTotal = allocationBaseData.reduce((sum, arr) => sum + (arr[m] || 0), 0) + (unallocatedBaseData[m] || 0);
                const cashflowTotal = allocationCashflowData.reduce((sum, arr) => sum + (arr[m] || 0), 0) + (unallocatedCashflowData[m] || 0);
                const reduction = Math.abs(cashflowReductionData[m] || 0);
                const footer = [`Projected total in bar: ${formatCurrency(barTotal)}`, `Base forecast: ${formatCurrency(baseTotal)}`];
                if (cashflowTotal > 0) footer.push(`Future cashflow included: +${formatCurrency(cashflowTotal)}`);
                if (reduction > 0) footer.push(`Future cashflow reduction: -${formatCurrency(reduction)}`);
                footer.push(`Allocated to goals: ${formatCurrency(allocated)}`);
                if (unallocatedData[m] > 0) footer.push(`Unallocated: ${formatCurrency(unallocatedData[m])}`);
                const activeChanges = activeCashflowChangesForMonth(m);
                if (activeChanges.length) {
                  footer.push(
                    "",
                    "Active future cashflow:",
                    ...activeChanges.map(change => `${change.amount >= 0 ? "+" : "-"}${formatCurrency(Math.abs(change.amount))}/mo — ${change.label}${change.toGoal !== "any" ? ` to ${change.toGoal}` : ""}`)
                  );
                }
                return footer;
              }
            }
          }
        },
        scales: {
          x: {
            stacked: true,
            ticks: { maxTicksLimit: 12, font: { size: 10 }, color: "#888" },
            grid: { display: false },
          },
          y: {
            stacked: true,
            beginAtZero: true,
            suggestedMin: hasCashflowReductions ? -maxCashflowReduction * 1.15 : 0,
            ticks: {
              callback: v => formatCurrency(v).replace(".00", ""),
              font: { size: 11 },
              color: "#888",
            },
            grid: { color: "rgba(0,0,0,0.05)" },
          }
        }
      }
    });
  });
}

function renderSavingsTimeline(container, dep) {
  return renderSavingsTimelineStacked(container, dep);
  const today = new Date();

  // ── Per-goal state ─────────────────────────────────────────────
  const goalState = goalsData.map((g, gi) => {
    const txSaved = getSavedViaTransactions(g.name);
    const effectiveTarget = g.target * (1 + (g.goalBuffer || 0) / 100);
    const currentSaved = g.manualSaved + txSaved;

    let startMo = 0; // months from now
    if (g.startDate) {
      const p = g.startDate.split("-");
      if (p.length === 3) {
        const s = new Date(+p[0], +p[1]-1, 1);
        startMo = Math.max(0, (s.getFullYear()-today.getFullYear())*12 + (s.getMonth()-today.getMonth()));
      }
    }
    let deadlineMo = null;
    if (g.endDate) {
      const p = g.endDate.split("-");
      if (p.length === 3) {
        const e = new Date(+p[0], +p[1]-1, 1);
        deadlineMo = (e.getFullYear()-today.getFullYear())*12 + (e.getMonth()-today.getMonth());
      }
    }

    return {
      idx: gi,
      name: g.name,
      color: g.color,
      urgency: g.urgency,
      effectiveTarget,
      monthlyAlloc: g.monthlyAlloc || 0,
      startMo,
      deadlineMo,
      saved: currentSaved,
      completedAt: null,
    };
  });

  // ── Project horizon: cover all deadlines + some runway ─────────
  const maxDeadline = goalState.reduce((mx, gs) => Math.max(mx, gs.deadlineMo ?? 0), 0);
  const MONTHS = Math.max(24, Math.min(60, maxDeadline + 6));

  // ── Month labels ───────────────────────────────────────────────
  const labels = [];
  for (let i = 0; i <= MONTHS; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    labels.push(d.toLocaleDateString("en-SG", { month:"short", year:"2-digit" }));
  }

  // ── Build per-month boost amounts per goal ─────────────────────
  // boostForGoal[m][goalName] = extra $ that month (from incomeBoosts)
  const baseMonthlyPool = Array(MONTHS+1).fill(0); // extra to "any" pool
  const boostForGoal = {}; // goalName -> array[MONTHS+1]
  goalState.forEach(gs => { boostForGoal[gs.name] = Array(MONTHS+1).fill(0); });

  incomeBoosts.forEach(b => {
    if (!b.fromMonth || !b.amount) return;
    const [by, bmo] = b.fromMonth.split("-").map(Number);
    const fromMo = Math.max(0, (by - today.getFullYear())*12 + (bmo-1 - today.getMonth()));
    // toMonth blank = permanent (runs to end of horizon); otherwise clamp
    let toMo = MONTHS;
    if (b.toMonth) {
      const [ty, tmo] = b.toMonth.split("-").map(Number);
      toMo = Math.min(MONTHS, (ty - today.getFullYear())*12 + (tmo-1 - today.getMonth()));
    }
    for (let m = fromMo; m <= toMo; m++) {
      if (b.toGoal === "any") {
        baseMonthlyPool[m] += b.amount;
      } else if (boostForGoal[b.toGoal] !== undefined) {
        boostForGoal[b.toGoal][m] += b.amount;
      }
    }
  });

  // ── Simulate month by month ────────────────────────────────────
  // Each goal tracks % progress. We record both "actual" progress and "required" (linear guide).
  const progressData  = goalState.map(() => Array(MONTHS+1).fill(null)); // % 0–100
  const requiredData  = goalState.map(() => Array(MONTHS+1).fill(null)); // required % to be on track
  const deficitFlags  = goalState.map(() => Array(MONTHS+1).fill(false));
  const completions   = [];
  const forecastBaseMonthly = Math.max(0, historicalStats.avgMonthlySavings);
  const forecastAllocData = goalState.map(() => Array(MONTHS + 1).fill(0));
  const requiredMonthlyData = goalState.map(() => Array(MONTHS + 1).fill(null));

const urgencyRank = { Critical: 0, High: 1, Medium: 2, Low: 3 };

function projectionOrder(a, b) {
  const u = (urgencyRank[a.gs.urgency] ?? 2) - (urgencyRank[b.gs.urgency] ?? 2);
  if (u !== 0) return u;
  if (a.gs.deadlineMo !== null && b.gs.deadlineMo !== null) return a.gs.deadlineMo - b.gs.deadlineMo;
  if (a.gs.deadlineMo !== null) return -1;
  if (b.gs.deadlineMo !== null) return 1;
  return a.gi - b.gi;
}

  // Initialise month 0
  goalState.forEach((gs, gi) => {
    if (gs.startMo <= 0) {
      const pct = gs.effectiveTarget > 0 ? Math.min(100, (gs.saved / gs.effectiveTarget) * 100) : 100;
      progressData[gi][0] = parseFloat(pct.toFixed(2));
    }
    if (gs.deadlineMo !== null && gs.startMo <= 0) {
      // required at month 0 = % that should have been saved by now for linear track
      // (start → deadline, linear)
      const totalMonths = gs.deadlineMo - gs.startMo;
      const monthsElapsed = 0 - gs.startMo; // could be negative if started before now
      requiredData[gi][0] = totalMonths > 0
        ? parseFloat(Math.min(100, Math.max(0, (monthsElapsed / totalMonths) * 100)).toFixed(2))
        : 100;
    }
  });

  for (let m = 1; m <= MONTHS; m++) {
    let poolExtra = forecastBaseMonthly + baseMonthlyPool[m]; // forecasted monthly savings + any-goal boosts

    goalState.forEach((gs, gi) => {
      if (m < gs.startMo) { progressData[gi][m] = null; return; }
      if (gs.completedAt !== null) {
        progressData[gi][m] = 100;
        return;
      }

      const directBoost = boostForGoal[gs.name][m] || 0;
const monthsToDeadline = gs.deadlineMo !== null ? Math.max(1, gs.deadlineMo - m + 1) : null;
const remainingNow = Math.max(0, gs.effectiveTarget - gs.saved);
const requiredMonthly = monthsToDeadline ? remainingNow / monthsToDeadline : 0;

requiredMonthlyData[gi][m] = requiredMonthly || null;

const smartNeed = Math.max(gs.monthlyAlloc || 0, requiredMonthly || 0);
const fromPool = Math.min(poolExtra, smartNeed, remainingNow);
const alloc = directBoost + fromPool;

forecastAllocData[gi][m] = alloc;
poolExtra = Math.max(0, poolExtra - fromPool);
      gs.saved += alloc;

      // Absorb any pool extra
      if (poolExtra > 0 && gs.saved < gs.effectiveTarget) {
        const canAbsorb = Math.min(poolExtra, gs.effectiveTarget - gs.saved);
        gs.saved += canAbsorb;
        poolExtra -= canAbsorb;
      }

      if (gs.saved >= gs.effectiveTarget) {
        gs.saved = gs.effectiveTarget;
        gs.completedAt = m;
        completions.push({ m, name: gs.name, color: gs.color, target: gs.effectiveTarget });
      }

      const pct = gs.effectiveTarget > 0 ? Math.min(100, (gs.saved / gs.effectiveTarget) * 100) : 100;
      progressData[gi][m] = parseFloat(pct.toFixed(2));

      // Required % linear guide
      if (gs.deadlineMo !== null) {
        const totalMonths = gs.deadlineMo - gs.startMo;
        const elapsed = m - gs.startMo;
        const reqPct = totalMonths > 0 ? Math.min(100, (elapsed / totalMonths) * 100) : 100;
        requiredData[gi][m] = parseFloat(reqPct.toFixed(2));
        // Flag deficit: actual progress is below required
        deficitFlags[gi][m] = pct < reqPct - 1; // 1% grace
      }
    });
  }

  // ── Build chart datasets ───────────────────────────────────────
  const datasets = [];
  goalState.forEach((gs, gi) => {
    // Progress line — solid, filled under
    datasets.push({
      label: gs.name,
      data: progressData[gi],
      borderColor: gs.color,
      backgroundColor: gs.color + "22",
      fill: "origin",
      tension: 0.3,
      pointRadius: (ctx) => {
        const v = progressData[gi][ctx.dataIndex];
        if (v === null) return 0;
        if (v >= 100) return 8;
        if (gs.deadlineMo !== null && ctx.dataIndex === gs.deadlineMo) return 7;
        return 0; // hide intermediate dots — cleaner look
      },
      pointStyle: (ctx) => {
        if (progressData[gi][ctx.dataIndex] >= 100) return "star";
        if (gs.deadlineMo !== null && ctx.dataIndex === gs.deadlineMo) return "rectRot";
        return "circle";
      },
      pointBackgroundColor: (ctx) => {
        if (gs.deadlineMo !== null && ctx.dataIndex === gs.deadlineMo) {
          return deficitFlags[gi][ctx.dataIndex] ? "#dc2626" : "#16a34a";
        }
        return gs.color;
      },
      pointBorderColor: "#fff",
      pointBorderWidth: 2,
      borderWidth: 2.5,
      spanGaps: true,
      datalabels: { display: false },
      _goalIdx: gi,
      _isProgress: true,
      order: 2,
    });

    // Required / on-track guide (dashed, only for goals with deadline)
    if (gs.deadlineMo !== null) {
      datasets.push({
        label: gs.name + " (on-track guide)",
        data: requiredData[gi],
        borderColor: gs.color + "66",
        backgroundColor: "transparent",
        fill: false,
        tension: 0,
        pointRadius: 0,
        borderWidth: 1.5,
        borderDash: [5, 4],
        spanGaps: true,
        datalabels: { display: false },
        _isGuide: true,
        order: 3,
      });
    }
  });

  // ── Panel HTML ─────────────────────────────────────────────────
  const panelEl = document.createElement("div");
  panelEl.className = "goals-panel timeline-panel";
  panelEl.innerHTML = `
    <div class="panel-header">
      <span class="panel-icon">📈</span>
      <h2>Savings Timeline</h2>
      <span class="panel-hint">Filled area = projected progress · dashed = on-track pace · ◆ at deadline = met ✓ or deficit ✗ · ⚡ = income change period</span>
    </div>
    <div class="timeline-legend-custom" id="timelineLegend" style="display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;margin-bottom:4px;"></div>
    <div class="timeline-forecast-note">
    Forecast available to save: <strong>${formatCurrency(forecastBaseMonthly)}/mo</strong>
    · chart allocates by priority, earliest deadline, required monthly amount, and income boosts
    </div>
    <div class="chart-wrap" style="height:380px;position:relative;">
      <canvas id="goalTimelineChart"></canvas>
    </div>
    <div class="timeline-completion-list" id="timelineCompletions"></div>
  `;
  container.appendChild(panelEl);

  // Completion / deficit summary
  const compList = panelEl.querySelector("#timelineCompletions");
  const summaryRows = goalState.map(gs => {
    const gi = gs.idx;
    const completedAt = gs.completedAt;
    const hasDeadline = gs.deadlineMo !== null;
    const finalPct = progressData[gi][MONTHS] ?? 0;
    const atDeadlinePct = hasDeadline && gs.deadlineMo <= MONTHS ? (progressData[gi][gs.deadlineMo] ?? 0) : null;
    const deficit = atDeadlinePct !== null && atDeadlinePct < 99.5;

    let chip = "";
    if (completedAt !== null) {
      const d = new Date(today.getFullYear(), today.getMonth() + completedAt, 1);
      const lbl = d.toLocaleDateString("en-SG", { month:"short", year:"numeric" });
      const early = hasDeadline && gs.deadlineMo !== null && completedAt < gs.deadlineMo;
      chip = `<div class="timeline-comp-chip" style="border-left:3px solid ${gs.color};">
        <span class="tc-icon">${early ? "🚀" : "✓"}</span>
        <span class="tc-name">${escapeHtml(gs.name)}</span>
        <span class="tc-date">Completes ${lbl}${early ? " (early!)" : ""}</span>
        <span class="tc-val">${formatCurrency(gs.effectiveTarget)}</span>
      </div>`;
    } else if (deficit) {
      const shortfall = gs.effectiveTarget * (1 - atDeadlinePct/100);
      chip = `<div class="timeline-comp-chip" style="border-left:3px solid #dc2626;">
        <span class="tc-icon">⚠️</span>
        <span class="tc-name">${escapeHtml(gs.name)}</span>
        <span class="tc-date">Only ${atDeadlinePct.toFixed(0)}% by deadline — ${formatCurrency(shortfall)} short</span>
        <span class="tc-val" style="color:#dc2626;">Deficit</span>
      </div>`;
    } else if (gs.monthlyAlloc === 0) {
      chip = `<div class="timeline-comp-chip" style="border-left:3px solid #94a3b8;">
        <span class="tc-icon">⏸</span>
        <span class="tc-name">${escapeHtml(gs.name)}</span>
        <span class="tc-date">No monthly allocation set</span>
        <span class="tc-val" style="color:var(--muted);">${finalPct.toFixed(0)}% funded</span>
      </div>`;
    }
    return chip;
  }).join("");
  compList.innerHTML = summaryRows || `<div style="font-size:12.5px;color:var(--muted);padding:8px 0;">Set monthly allocations to see projections.</div>`;

  // ── Draw chart ─────────────────────────────────────────────────
  requestAnimationFrame(() => {
    const ctx = document.getElementById("goalTimelineChart");
    if (!ctx) return;
    if (timelineChart) { timelineChart.destroy(); timelineChart = null; }

    // Deadline annotations as vertical lines drawn on the datasets (no plugin needed)
    // We'll use a custom afterDraw plugin inline
    const deadlineLines = goalState
      .filter(gs => gs.deadlineMo !== null && gs.deadlineMo <= MONTHS)
      .map(gs => ({ mo: gs.deadlineMo, color: gs.color, name: gs.name }));

    const deadlinePlugin = {
      id: "deadlineLines",
      afterDraw(chart) {
        const { ctx: c, chartArea, scales } = chart;
        if (!chartArea) return;
        deadlineLines.forEach(dl => {
          const x = scales.x.getPixelForValue(dl.mo);
          if (x < chartArea.left || x > chartArea.right) return;
          c.save();
          c.strokeStyle = dl.color;
          c.lineWidth = 1.5;
          c.setLineDash([4, 3]);
          c.beginPath();
          c.moveTo(x, chartArea.top);
          c.lineTo(x, chartArea.bottom);
          c.stroke();
          c.setLineDash([]);
          // Label at top
          c.fillStyle = dl.color;
          c.font = "10px DM Sans, sans-serif";
          c.textAlign = "center";
          c.fillText("⏰ " + dl.name, x, chartArea.top - 4);
          c.restore();
        });
      }
    };

    // Income boost markers — show a ⚡ band for the active range
    const boostRanges = [];
    incomeBoosts.forEach(b => {
      if (!b.fromMonth || !b.amount) return;
      const [by, bmo] = b.fromMonth.split("-").map(Number);
      const fromMo = (by - today.getFullYear())*12 + (bmo-1 - today.getMonth());
      let toMo = MONTHS;
      if (b.toMonth) {
        const [ty, tmo] = b.toMonth.split("-").map(Number);
        toMo = (ty - today.getFullYear())*12 + (tmo-1 - today.getMonth());
      }
      if (fromMo >= 0 && fromMo <= MONTHS) boostRanges.push({ fromMo: Math.max(0, fromMo), toMo: Math.min(MONTHS, toMo), label: b.label, amount: b.amount, permanent: !b.toMonth });
    });

    const boostPlugin = {
      id: "boostMarkers",
      afterDraw(chart) {
        const { ctx: c, chartArea, scales } = chart;
        if (!chartArea) return;
        boostRanges.forEach(br => {
          const x1 = scales.x.getPixelForValue(br.fromMo);
          const x2 = scales.x.getPixelForValue(Math.min(br.toMo, MONTHS));
          if (x1 > chartArea.right || x2 < chartArea.left) return;
          const cx1 = Math.max(chartArea.left, x1);
          const cx2 = Math.min(chartArea.right, x2);
          // Shaded band
          c.save();
          c.fillStyle = "rgba(139,92,246,0.07)";
          c.fillRect(cx1, chartArea.top, cx2 - cx1, chartArea.bottom - chartArea.top);
          // Left edge line
          c.strokeStyle = "#8b5cf6";
          c.lineWidth = 1.5;
          c.setLineDash([2, 4]);
          c.beginPath();
          c.moveTo(x1, chartArea.top);
          c.lineTo(x1, chartArea.bottom);
          c.stroke();
          c.setLineDash([]);
          // Label
          c.fillStyle = "#8b5cf6";
          c.font = "bold 10px DM Sans, sans-serif";
          c.textAlign = "left";
          const tag = br.permanent ? "⚡ " + br.label : "⚡ " + br.label + " (ends " + labels[Math.min(br.toMo, MONTHS)] + ")";
          c.fillText(tag, Math.min(cx1 + 3, chartArea.right - 80), chartArea.bottom + 14);
          c.restore();
        });
      }
    };

    timelineChart = new Chart(ctx, {
      type: "line",
      data: { labels, datasets },
      plugins: [deadlinePlugin, boostPlugin],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 32, right: 20, bottom: 24, left: 4 } },
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { display: false }, // we use our custom legend below
          datalabels: { display: false },
          tooltip: {
            backgroundColor: "rgba(15,23,42,0.92)",
            titleFont: { size: 12, weight: "600" },
            bodyFont: { size: 12 },
            padding: 12,
            callbacks: {
              title: items => "📅 " + labels[items[0].dataIndex],
              label: item => {
                if (item.dataset._isGuide) return null;
                const v = item.parsed.y;
                if (v === null || v === undefined) return null;
                const gs = goalState[item.dataset._goalIdx];
                if (!gs) return "  " + item.dataset.label + ": " + v.toFixed(1) + "%";
                const savedAmt = gs.effectiveTarget * (v / 100);
                const guideIdx = datasets.findIndex(d => d._isGuide && d.label === gs.name + " (on-track guide)");
                const reqPct = guideIdx >= 0 ? (datasets[guideIdx].data[item.dataIndex] ?? null) : null;
                let status = "";
                if (reqPct !== null) {
                  const diff = v - reqPct;
                  status = diff >= -1 ? "  ✓ ahead" : `  ⚠ ${Math.abs(diff).toFixed(0)}% behind pace`;
                }
                const planned = forecastAllocData[item.dataset._goalIdx]?.[item.dataIndex] || 0;
const required = requiredMonthlyData[item.dataset._goalIdx]?.[item.dataIndex] || 0;

return `  ${item.dataset.label}: ${v.toFixed(0)}% · ${formatCurrency(savedAmt)} · forecast ${formatCurrency(planned)}/mo${required ? " · required " + formatCurrency(required) + "/mo" : ""}${status}`;
              },
              afterBody: items => {
                const mo = items[0].dataIndex;
                const activeBoosts = incomeBoosts.filter(b => {
                  if (!b.fromMonth || !b.amount) return false;
                  const [by, bmo] = b.fromMonth.split("-").map(Number);
                  const from = (by - today.getFullYear())*12 + (bmo-1 - today.getMonth());
                  let to = MONTHS;
                  if (b.toMonth) { const [ty,tmo] = b.toMonth.split("-").map(Number); to = (ty-today.getFullYear())*12+(tmo-1-today.getMonth()); }
                  return mo >= from && mo <= to;
                });
                if (activeBoosts.length === 0) return [];
                return ["", "⚡ Income changes active:", ...activeBoosts.map(b => `  +${formatCurrency(b.amount)}/mo — ${b.label}`)];
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { maxTicksLimit: 14, font: { size: 10 }, color: "#888" },
            grid: { color: "rgba(0,0,0,0.04)" },
          },
          y: {
            min: 0, max: 105,
            ticks: {
              callback: v => v <= 100 ? v + "%" : "",
              stepSize: 25,
              font: { size: 11 },
              color: "#888",
            },
            grid: {
              color: ctx => ctx.tick.value === 100 ? "rgba(39,174,96,0.35)" : "rgba(0,0,0,0.05)",
              lineWidth: ctx => ctx.tick.value === 100 ? 2 : 1,
            }
          }
        }
      }
    });

    // Build custom legend
    const legendEl = panelEl.querySelector("#timelineLegend");
    if (legendEl) {
      legendEl.innerHTML = goalState.map(gs => {
        const hasDeadline = gs.deadlineMo !== null;
        return `<div class="tl-legend-item">
          <span style="display:inline-block;width:12px;height:12px;border-radius:3px;background:${gs.color};opacity:0.85;flex-shrink:0;"></span>
          <span style="font-size:12.5px;font-weight:500;">${escapeHtml(gs.name)}</span>
          ${hasDeadline ? `<span style="font-size:11px;color:var(--muted);">deadline: ${labels[gs.deadlineMo] || ""}</span>` : `<span style="font-size:11px;color:var(--muted);">no deadline</span>`}
        </div>`;
      }).join("") + `<div class="tl-legend-item tl-legend-guide">
        <span style="display:inline-block;width:18px;height:2px;border-top:2px dashed #999;margin-top:5px;"></span>
        <span style="font-size:12px;color:var(--muted);">On-track pace (dashed)</span>
      </div>`;
    }
  });
}

function renderRealismCheck(container, dep) {
  const today   = new Date();
  const avgSave = Math.max(0, historicalStats.avgMonthlySavings);
  let issues    = [];
  let tips      = [];
  let totalMonthlyNeeded = 0;

  goalsData.forEach(goal => {
    const savedViaGoalTx = getSavedViaTransactions(goal.name);
    const totalSaved     = goal.manualSaved + savedViaGoalTx;
    const effectiveTarget = goal.target * (1 + (goal.goalBuffer || 0) / 100);
    const remaining      = Math.max(0, effectiveTarget - totalSaved);
    let monthsLeft       = null;

    if (goal.endDate) {
      const endParts = goal.endDate.split("-");
      if (endParts.length === 3) {
        const endObj = new Date(+endParts[0], +endParts[1]-1, +endParts[2]);
        if (endObj > today) {
          monthsLeft = (endObj.getFullYear() - today.getFullYear())*12 + (endObj.getMonth()-today.getMonth());
        }
      }
    }

    const reqMonthly = monthsLeft && monthsLeft > 0 ? remaining / monthsLeft : (goal.monthlyAlloc || 0);
    totalMonthlyNeeded += reqMonthly;

    if (goal.urgency === "Critical" && monthsLeft !== null && reqMonthly > avgSave * 0.6) {
      issues.push(`🔴 <strong>${escapeHtml(goal.name)}</strong>: needs ${formatCurrency(reqMonthly)}/mo — that's ${Math.round(reqMonthly/avgSave*100)}% of your avg savings alone.`);
    }
    if (goal.endDate && monthsLeft !== null && monthsLeft <= 0) {
      issues.push(`⏰ <strong>${escapeHtml(goal.name)}</strong>: deadline has passed. Update or archive this goal.`);
    }
  });

  if (totalMonthlyNeeded > avgSave) {
    const gap = totalMonthlyNeeded - avgSave;
    issues.push(`⚡ Goals plus buffers need ${formatCurrency(totalMonthlyNeeded)}/mo but your avg savings is ${formatCurrency(avgSave)}/mo — <strong>${formatCurrency(gap)} gap</strong>.`);
    tips.push(`💡 Reduce your monthly expenses budget by ${formatCurrency(gap)} or increase income to close the gap.`);
    tips.push(`💡 Lower the allocation on Low/Medium priority goals to free up ${formatCurrency(gap)}/mo for Critical ones.`);
  } else if (totalMonthlyNeeded > 0) {
    tips.push(`✅ Your goals are <strong>collectively achievable</strong> based on your avg savings of ${formatCurrency(avgSave)}/mo.`);
  }

  if (dep.deployable < 0) {
    issues.push(`🚨 Deployable balance is negative (${formatCurrency(dep.deployable)}). You may be over-extended — check your CC bill and budget.`);
  }

  if (issues.length === 0 && tips.length === 0) return;

  const issueHtml = issues.map(i => `<div class="rc-issue">${i}</div>`).join("");
  const tipHtml   = tips.map(t => `<div class="rc-tip">${t}</div>`).join("");

  container.innerHTML += `
    <div class="goals-panel realism-panel">
      <div class="panel-header">
        <span class="panel-icon">🧠</span>
        <h2>Realism Check</h2>
        <span class="panel-hint">AI analysis of your goals vs your actual finances</span>
      </div>
      <div class="rc-body">
        ${issueHtml}
        ${tipHtml}
        <div class="rc-summary">
          <span>Total monthly needed for goals + buffers:</span>
          <strong class="${totalMonthlyNeeded <= avgSave ? 'green':'red'}">${formatCurrency(totalMonthlyNeeded)}/mo</strong>
          <span>vs your avg savings</span>
          <strong>${formatCurrency(avgSave)}/mo</strong>
        </div>
      </div>
    </div>
  `;
}

function renderBudgetGoalMotivation(container, dep) {
  if (!container) return;

  const position = dep.budgetPosition || computeCurrentMonthBudgetPosition();
  const trend = computeFlexibleDailyTrend(position);
  const focusGoal = selectGoalForMotivation();
  const goalLabel = focusGoal ? escapeHtml(focusGoal.name) : "your next goal";
  const opportunity = buildDailyTrendOpportunity(trend);
  const impactText = buildGoalMomentumImpact(focusGoal, opportunity.amount, false);
  const chips = buildGoalMomentumChips(position, trend, focusGoal);

  container.innerHTML = `
    <div class="goals-panel motivation-panel">
      <div class="panel-header">
        <span class="panel-icon">↗</span>
        <h2>Goal Momentum</h2>
        <span class="panel-hint">Daily flexible expense pace and possible goal boost</span>
      </div>
      <div class="gm-grid">
        <div class="gm-hero">
          <div class="gm-kicker">${escapeHtml(opportunity.kicker)}</div>
          <h3>${escapeHtml(opportunity.title)}</h3>
          <p>${opportunity.body.replace("{goal}", `<strong>${goalLabel}</strong>`)}</p>
          <div class="gm-impact">${impactText}</div>
        </div>
        <div class="gm-side">
          <div class="gm-stat">
            <span>Actual flexible pace</span>
            <strong class="${trend.actualDaily <= trend.plannedDaily ? "green" : "red"}">${formatCurrency(trend.actualDaily)}/day</strong>
          </div>
          <div class="gm-stat">
            <span>Planned flexible pace</span>
            <strong>${formatCurrency(trend.plannedDaily)}/day</strong>
          </div>
          <div class="gm-stat">
            <span>Rest-of-month budget pace</span>
            <strong>${formatCurrency(trend.remainingDailyBudget)}/day</strong>
          </div>
          <div class="gm-stat">
            <span>Projected month-end balance</span>
            <strong class="${trend.projectedBalance >= 0 ? "green" : "red"}">${formatCurrency(trend.projectedBalance)}</strong>
          </div>
        </div>
      </div>
      ${chips}
    </div>`;
}

function computeFlexibleDailyTrend(position) {
  const today = new Date();
  const daysElapsed = Math.max(1, today.getDate());
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const daysLeft = Math.max(daysInMonth - today.getDate() + 1, 1);
  const allocated = Math.max(0, position.monthly.allocated || 0);
  const spent = Math.max(0, position.monthly.spent || 0);
  const balance = allocated - spent;
  const plannedDaily = allocated > 0 ? allocated / daysInMonth : 0;
  const actualDaily = spent / daysElapsed;
  const projectedSpent = actualDaily * daysInMonth;
  const projectedBalance = allocated - projectedSpent;
  const remainingDailyBudget = Math.max(0, balance / daysLeft);
  const dailyDelta = actualDaily - plannedDaily;
  const nudgeDaily = actualDaily > 0
    ? Math.min(actualDaily, Math.max(5, Math.min(25, actualDaily * 0.12)))
    : 0;

  return {
    allocated,
    spent,
    balance,
    daysElapsed,
    daysLeft,
    daysInMonth,
    plannedDaily,
    actualDaily,
    projectedSpent,
    projectedBalance,
    remainingDailyBudget,
    dailyDelta,
    nudgeDaily,
    nudgeByMonthEnd: nudgeDaily * daysLeft,
    topDailyRows: (position.monthly.rows || [])
      .filter(row => row.spent > 0)
      .sort((a, b) => b.spent - a.spent)
      .slice(0, 3)
      .map(row => ({ ...row, daily: row.spent / daysElapsed }))
  };
}

function buildDailyTrendOpportunity(trend) {
  if (trend.allocated <= 0) {
    return {
      amount: Math.max(0, trend.nudgeByMonthEnd),
      kicker: "Daily expense trend",
      title: "Add a monthly expenses budget to turn daily spending into goal signals",
      body: `Once flexible expenses have an allocation, this panel can compare actual daily spend with planned daily pace and estimate what can move into {goal}.`
    };
  }

  if (trend.spent <= 0) {
    return {
      amount: 0,
      kicker: "Daily expense trend",
      title: "No flexible expenses recorded yet this month",
      body: `As transactions come in, this will show whether your daily expense pace is creating extra room for {goal}.`
    };
  }

  if (trend.projectedBalance >= 0) {
    return {
      amount: trend.projectedBalance,
      kicker: "Under daily pace",
      title: `Flexible expenses are trending ${formatCurrency(trend.projectedBalance)} under budget`,
      body: `You are averaging ${formatCurrency(trend.actualDaily)}/day against a planned ${formatCurrency(trend.plannedDaily)}/day. If that pace holds, the projected leftover can move into {goal} at month end.`
    };
  }

  return {
    amount: trend.nudgeByMonthEnd,
    kicker: "Above daily pace",
    title: `Flexible expenses are running ${formatCurrency(Math.abs(trend.dailyDelta))}/day above plan`,
    body: `No zero-spend challenge needed. A practical ${formatCurrency(trend.nudgeDaily)}/day trim across flexible expenses for the remaining ${trend.daysLeft} day${trend.daysLeft === 1 ? "" : "s"} saves about ${formatCurrency(trend.nudgeByMonthEnd)} while you steer back toward budget.`
  };
}

function selectGoalForMotivation() {
  const urgencyOrder = { Critical:0, High:1, Medium:2, Low:3 };
  return goalsData
    .map((goal, idx) => ({ goal, idx, remaining: getGoalRemainingAmount(goal) }))
    .filter(item => item.remaining > 0)
    .sort((a, b) => {
      const u = (urgencyOrder[a.goal.urgency] ?? 2) - (urgencyOrder[b.goal.urgency] ?? 2);
      if (u !== 0) return u;
      if (a.goal.endDate && b.goal.endDate) return a.goal.endDate.localeCompare(b.goal.endDate);
      if (a.goal.endDate) return -1;
      if (b.goal.endDate) return 1;
      return a.idx - b.idx;
    })[0]?.goal || goalsData[0] || null;
}

function getGoalRemainingAmount(goal) {
  if (!goal) return 0;
  const savedViaGoalTx = getSavedViaTransactions(goal.name);
  const totalSaved = goal.manualSaved + savedViaGoalTx;
  const effectiveTarget = goal.target * (1 + (goal.goalBuffer || 0) / 100);
  return Math.max(0, effectiveTarget - totalSaved);
}

function buildGoalMomentumImpact(goal, amount, recurring) {
  if (!goal || amount <= 0) return "Even a small saved amount makes the next goal easier.";

  const remaining = getGoalRemainingAmount(goal);
  const safeName = escapeHtml(goal.name);
  if (remaining <= 0) return `<strong>${safeName}</strong> is funded. Send this win to the next goal.`;

  const monthlyAlloc = Math.max(0, Number(goal.monthlyAlloc || 0) || 0);
  const pct = Math.min(100, (amount / remaining) * 100);

  if (recurring) {
    const boostedMonthly = monthlyAlloc + amount;
    if (boostedMonthly > 0) {
      const boostedMonths = Math.ceil(remaining / boostedMonthly);
      if (monthlyAlloc > 0) {
        const currentMonths = Math.ceil(remaining / monthlyAlloc);
        const fasterBy = Math.max(0, currentMonths - boostedMonths);
        if (fasterBy > 0) {
          return `<strong>${safeName}</strong> could move ${fasterBy} month${fasterBy === 1 ? "" : "s"} faster.`;
        }
      }
      return `<strong>${safeName}</strong> could finish in about ${boostedMonths} month${boostedMonths === 1 ? "" : "s"}.`;
    }
  }

  if (monthlyAlloc > 0) {
    const currentMonths = Math.ceil(remaining / monthlyAlloc);
    const afterBoostMonths = Math.ceil(Math.max(0, remaining - amount) / monthlyAlloc);
    const fasterBy = Math.max(0, currentMonths - afterBoostMonths);
    if (fasterBy > 0) {
      return `<strong>${safeName}</strong> could move ${fasterBy} month${fasterBy === 1 ? "" : "s"} closer.`;
    }
  }

  return `That covers <strong>${pct >= 10 ? pct.toFixed(0) : pct.toFixed(1)}%</strong> of <strong>${safeName}</strong>'s remaining gap.`;
}

function buildGoalMomentumChips(position, trend, focusGoal) {
  const focusRemaining = getGoalRemainingAmount(focusGoal);
  const chips = trend.topDailyRows.map(row => `
    <span class="gm-chip">${escapeHtml(row.category)} avg <strong>${formatCurrency(row.daily)}/day</strong></span>
  `);

  if (focusGoal) {
    chips.push(`<span class="gm-chip">${escapeHtml(focusGoal.name)} left <strong>${formatCurrency(focusRemaining)}</strong></span>`);
  }
  chips.push(`<span class="gm-chip">Days left <strong>${trend.daysLeft}</strong></span>`);
  chips.push(`<span class="gm-chip">Flexible spent <strong>${formatCurrency(position.monthly.spent)}</strong></span>`);

  return `<div class="gm-list">${chips.join("")}</div>`;
}

// ─── Individual Goal Card ──────────────────────────────────────────

function renderGoalCard(goal, idx, container) {
  const today       = new Date();
  const urgColor    = { Critical:"#dc2626", High:"#ea580c", Medium:"#ca8a04", Low:"#16a34a" };
  const urgIcon     = { Critical:"🔴", High:"🟠", Medium:"🟡", Low:"🟢" };

  const savedViaGoalTx = getSavedViaTransactions(goal.name);
  const totalSaved     = goal.manualSaved + savedViaGoalTx;
  const bufferPct      = Number(goal.goalBuffer || 0) || 0;
  const bufferAmount   = goal.target * (bufferPct / 100);
  const effectiveTarget = goal.target + bufferAmount;
  const remaining      = Math.max(0, effectiveTarget - totalSaved);
  const baseRemaining  = Math.max(0, goal.target - totalSaved);
  const bufferRemaining = Math.max(0, effectiveTarget - Math.max(totalSaved, goal.target));
  const pct            = effectiveTarget > 0 ? Math.min(100, (totalSaved / effectiveTarget) * 100) : 0;
  const baseTargetMet  = totalSaved >= goal.target;

  // Deadline math
  let monthsLeft = null, endDateObj = null;
  if (goal.endDate) {
    const p = goal.endDate.split("-");
    if (p.length === 3) {
      endDateObj = new Date(+p[0], +p[1]-1, +p[2]);
      if (endDateObj > today) {
        monthsLeft = (endDateObj.getFullYear()-today.getFullYear())*12 + (endDateObj.getMonth()-today.getMonth());
      }
    }
  }

  const requiredMonthly = monthsLeft && monthsLeft > 0 ? remaining / monthsLeft : 0;
  let effectiveAlloc    = goal.monthlyAlloc;

  // Auto-suggest allocation
  let recommendation = "";
  const avgSave = Math.max(0, historicalStats.avgMonthlySavings);
  const totalOtherAlloc = goalsData.reduce((s,g,i) => i===idx ? s : s + g.monthlyAlloc, 0);
  const availableForThis = Math.max(0, avgSave - totalOtherAlloc);

  if (requiredMonthly > 0) {
    effectiveAlloc = effectiveAlloc || requiredMonthly;
    if (requiredMonthly > availableForThis) {
      recommendation = `⚠️ Need ${formatCurrency(requiredMonthly)}/mo to hit deadline but only ${formatCurrency(availableForThis)}/mo available after other goals. Consider reducing other allocations or extending deadline.`;
    } else {
      recommendation = `✅ On track — ${formatCurrency(requiredMonthly)}/mo needed to hit deadline, within your savings capacity.`;
    }
  } else if (!goal.monthlyAlloc && avgSave > 0 && availableForThis > 0) {
    const suggested = Math.round(availableForThis * 0.3);
    effectiveAlloc = suggested;
    recommendation = `💡 No deadline set. Suggested ${formatCurrency(suggested)}/mo based on ~30% of your ${formatCurrency(availableForThis)}/mo available capacity.`;
  } else if (!goal.monthlyAlloc) {
    recommendation = `💡 Set a monthly allocation or deadline to get a forecast.`;
  }

  // Forecast
  const monthsToTarget  = effectiveAlloc > 0 ? Math.ceil(remaining / effectiveAlloc) : null;
  const forecastDate    = monthsToTarget != null ? new Date(today.getFullYear(), today.getMonth() + monthsToTarget, 1) : null;
  const forecastDeadlineDelta = forecastDate && endDateObj
    ? (endDateObj.getFullYear() - forecastDate.getFullYear()) * 12 + (endDateObj.getMonth() - forecastDate.getMonth())
    : null;
  const forecastTimingText = forecastDeadlineDelta !== null
    ? (forecastDeadlineDelta > 0
      ? ` · ${forecastDeadlineDelta}mo early`
      : (forecastDeadlineDelta === 0 ? " · on deadline" : ` · ${Math.abs(forecastDeadlineDelta)}mo late`))
    : "";
  const forecastText    = forecastDate
    ? `📅 Projected completion: <strong>${forecastDate.toLocaleDateString("en-SG",{month:"short",year:"numeric"})}</strong> (${monthsToTarget} month${monthsToTarget===1?"":"s"}${forecastTimingText})`
    : `<span style="color:var(--muted);">Set a monthly allocation to get a forecast</span>`;

  // Deadline status
  let deadlineStatus = "";
  if (endDateObj) {
    if (endDateObj <= today) {
      deadlineStatus = `<span class="deadline-badge overdue">OVERDUE</span>`;
    } else if (forecastDate && forecastDate > endDateObj) {
      const monthsBehind = (forecastDate.getFullYear()-endDateObj.getFullYear())*12+(forecastDate.getMonth()-endDateObj.getMonth());
      deadlineStatus = `<span class="deadline-badge behind">⚠️ ${monthsBehind}mo behind deadline</span>`;
    } else if (forecastDate) {
      deadlineStatus = forecastDeadlineDelta > 0
        ? `<span class="deadline-badge ontrack">✓ ${forecastDeadlineDelta}mo early</span>`
        : `<span class="deadline-badge ontrack">✓ On deadline</span>`;
    }
  }

  const daysToDeadline = endDateObj ? Math.ceil((endDateObj - today) / 86400000) : null;
  const deadlineLabel  = endDateObj
    ? `${formatDateDisplay(goal.endDate)} ${daysToDeadline !== null ? `(${daysToDeadline > 0 ? daysToDeadline+" days left" : "past due"})` : ""}`
    : "No deadline";
  const bufferLine = bufferPct > 0
    ? `<div class="goal-buffer-line ${baseTargetMet ? "ok" : ""}">
        ${baseTargetMet
          ? `Base target met. Forecast is now funding the ${formatCurrency(bufferAmount)} buffer.`
          : `Includes ${formatCurrency(bufferAmount)} buffer (${bufferPct}% above base target).`}
      </div>`
    : "";

  const div = document.createElement("div");
  div.className = "goal-card";
  div.dataset.goalName = goal.name;
  div.dataset.goalTarget = String(goal.target || 0);
  div.dataset.goalSaved = String(goal.manualSaved || 0);
  div.dataset.goalAlloc = String(goal.monthlyAlloc || 0);
  div.dataset.goalStart = goal.startDate || "";
  div.dataset.goalEnd = goal.endDate || "";
  div.dataset.goalUrgency = goal.urgency || "Medium";
  div.dataset.goalNotes = goal.notes || "";
  div.dataset.goalColor = goal.color || "";
  div.dataset.goalBuffer = String(goal.goalBuffer || 0);
  div.dataset.goalPriority = String(goal.priority || 0);
  div.style.cssText = `border-top: 3px solid ${urgColor[goal.urgency]||"#ca8a04"};`;
  div.innerHTML = `
    <div class="goal-card-top">
      <div style="display:flex;align-items:center;gap:8px;flex:1;min-width:0;">
        <span style="width:12px;height:12px;border-radius:50%;background:${goal.color};flex-shrink:0;display:inline-block;"></span>
        <strong style="font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(goal.name)}</strong>
        <span style="font-size:12px;font-weight:600;color:${urgColor[goal.urgency]||"#ca8a04"};white-space:nowrap;">${urgIcon[goal.urgency]||"🟡"} ${goal.urgency}</span>
        ${deadlineStatus}
      </div>
      <div style="display:flex;gap:6px;flex-shrink:0;">
        <button class="btn-icon" onclick="toggleGoalEdit(${idx})" title="Edit">✏️</button>
        <button class="btn-icon btn-icon-danger" onclick="deleteGoal(${idx})" title="Delete">🗑️</button>
      </div>
    </div>

    <div class="goal-kpis">
      <div class="gkpi"><span class="gkpi-l">Saved</span><span class="gkpi-v green">${formatCurrency(totalSaved)}</span></div>
      <div class="gkpi"><span class="gkpi-l">${bufferPct > 0 ? "Target + Buffer" : "Target"}</span><span class="gkpi-v">${formatCurrency(effectiveTarget)}</span></div>
      <div class="gkpi"><span class="gkpi-l">Remaining</span><span class="gkpi-v ${remaining>0?'red':''}">${formatCurrency(remaining)}</span></div>
      <div class="gkpi"><span class="gkpi-l">Monthly</span><span class="gkpi-v">${formatCurrency(goal.monthlyAlloc||effectiveAlloc)}</span></div>
    </div>

    <div class="goal-progress-wrap">
      <div class="goal-pbar"><div class="goal-pfill" style="width:${pct.toFixed(1)}%;background:${goal.color};"></div></div>
      <span class="goal-pct-label">${pct.toFixed(1)}%</span>
    </div>

    <div class="goal-meta">
      ${goal.startDate ? `<span>📅 Start: ${formatDateDisplay(goal.startDate)}</span>` : ""}
      <span>🎯 Deadline: ${deadlineLabel}</span>
      ${goal.notes ? `<span>📝 ${escapeHtml(goal.notes)}</span>` : ""}
    </div>

    ${bufferLine}
    <div class="goal-forecast-line">${forecastText}</div>
    ${requiredMonthly > 0 ? `<div class="goal-req-line">📋 Required: ${formatCurrency(requiredMonthly)}/mo to hit deadline</div>` : ""}
    ${bufferPct > 0 && baseRemaining <= 0 && bufferRemaining > 0 ? `<div class="goal-req-line">Buffer left: ${formatCurrency(bufferRemaining)}</div>` : ""}
    ${recommendation ? `<div class="goal-rec-line">${recommendation}</div>` : ""}

    ${savedViaGoalTx > 0 ? `<div class="goal-tx-line">💳 ${formatCurrency(savedViaGoalTx)} tracked via transactions · ${formatCurrency(goal.manualSaved)} manually recorded</div>` : ""}

    <!-- Edit Form -->
    <div id="editGoal_${idx}" class="goal-edit-form" style="display:none;">
      <div class="goal-form-grid" style="margin-top:12px;">
        <div class="gf-group full"><label>Goal Name</label><input type="text" id="eg_name_${idx}" value="${escapeHtml(goal.name)}"></div>
        <div class="gf-group"><label>Target (SGD)</label><input type="number" id="eg_target_${idx}" value="${goal.target}"></div>
        <div class="gf-group"><label>Manually Saved (SGD)</label><input type="number" id="eg_saved_${idx}" value="${goal.manualSaved}" step="0.01"></div>
        <div class="gf-group"><label>Monthly Allocation (SGD)</label><input type="number" id="eg_alloc_${idx}" value="${goal.monthlyAlloc}"></div>
        <div class="gf-group"><label>Start Date</label><input type="date" id="eg_start_${idx}" value="${goal.startDate}"></div>
        <div class="gf-group"><label>Deadline</label><input type="date" id="eg_end_${idx}" value="${goal.endDate}"></div>
        <div class="gf-group">
          <label>Priority</label>
          <select id="eg_urgency_${idx}">
            <option value="Critical" ${goal.urgency==="Critical"?"selected":""}>🔴 Critical</option>
            <option value="High"     ${goal.urgency==="High"    ?"selected":""}>🟠 High</option>
            <option value="Medium"   ${goal.urgency==="Medium"  ?"selected":""}>🟡 Medium</option>
            <option value="Low"      ${goal.urgency==="Low"     ?"selected":""}>🟢 Low</option>
          </select>
        </div>
        <div class="gf-group"><label>Buffer % above target</label><input type="number" id="eg_buffer_${idx}" value="${goal.goalBuffer||0}" step="1" min="0" max="100" placeholder="0"></div>
        <div class="gf-group"><label>Color</label><input type="color" id="eg_color_${idx}" value="${goal.color}" style="height:38px;padding:2px 4px;border-radius:6px;border:1px solid var(--border);width:100%;"></div>
        <div class="gf-group full"><label>Notes</label><input type="text" id="eg_notes_${idx}" value="${escapeHtml(goal.notes)}"></div>
      </div>
      <div style="display:flex;gap:8px;margin-top:12px;">
        <button class="btn-primary btn-sm" onclick="saveEditGoal(${idx})">Save Changes</button>
        <button class="btn-secondary btn-sm" onclick="toggleGoalEdit(${idx})">Cancel</button>
      </div>
    </div>

    <!-- Deduct Form -->
    <div style="margin-top:10px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <button class="btn-deduct btn-sm" onclick="toggleDeductForm(${idx})">💸 Log Expense from Goal</button>
    </div>
    <div id="deductGoal_${idx}" class="goal-edit-form" style="display:none;margin-top:10px;background:#fef2f2;border-color:#fca5a5;">
      <div class="goal-form-grid" style="margin-top:0;">
        <div class="gf-group"><label>Amount (SGD)</label><input type="number" id="dg_amt_${idx}" step="0.01" placeholder="0.00"></div>
        <div class="gf-group"><label>Description</label><input type="text" id="dg_note_${idx}" placeholder="What was it for?"></div>
        <div class="gf-group"><label>Account</label><select id="dg_acct_${idx}"></select></div>
      </div>
      <div style="display:flex;gap:8px;margin-top:10px;">
        <button class="btn-primary btn-sm" style="background:#dc2626;" onclick="saveDeductGoal(${idx})">Save Deduction</button>
        <button class="btn-secondary btn-sm" onclick="toggleDeductForm(${idx})">Cancel</button>
      </div>
    </div>
  `;
  container.appendChild(div);

  // Populate account dropdown for deduct
  const sel = div.querySelector(`#dg_acct_${idx}`);
  allAccounts.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a.name;
    opt.textContent = `${a.name} (${a.type})`;
    sel.appendChild(opt);
  });
}

// ─── Helpers ──────────────────────────────────────────────────────

function getSavedViaTransactions(goalName) {
  return allTxForGoals
    .filter(r => clean(r["Sub Category"]).toLowerCase() === ("goal: " + goalName).toLowerCase())
    .reduce((s,r) => s + getAmount(r["Amount"]), 0);
}

function toggleAddForm() {
  const body = document.getElementById("addGoalFormBody");
  const hint = document.getElementById("addFormToggleHint");
  const open = body.style.display === "none";
  body.style.display = open ? "block" : "none";
  hint.textContent   = open ? "Click to collapse" : "Click to expand";
  if (open) {
    // Auto-assign next unused color
    const colorInput = document.getElementById("gf_color");
    if (colorInput) colorInput.value = getNextGoalColor();
    // Smart start date: today by default (user can override)
    const startInput = document.getElementById("gf_start");
    if (startInput && !startInput.value) startInput.value = toDateInputValue(new Date());
    _updateAddGoalCalc();
  }
}

function _toggleAllocOverride() {
  const input = document.getElementById("gf_alloc");
  const display = document.getElementById("gf_alloc_display");
  if (input.style.display === "none") {
    input.style.display = "block";
    display.style.display = "none";
    const finalVal = document.getElementById("gf_alloc_final").value;
    if (finalVal && Number(finalVal) > 0) input.value = finalVal;
  } else {
    input.style.display = "none";
    display.style.display = "flex";
    _updateAddGoalCalc();
  }
}

function _updateAddGoalCalc() {
  const target = parseFloat(document.getElementById("gf_target")?.value) || 0;
  const saved = 0;
  const endVal = document.getElementById("gf_end")?.value;
  const startVal = document.getElementById("gf_start")?.value;
  const avgSave = Math.max(0, historicalStats.avgMonthlySavings);
  const otherAlloc = goalsData.reduce((s,g) => s + (g.monthlyAlloc||0), 0);
  const available = Math.max(0, avgSave - otherAlloc);

  const remaining = Math.max(0, target - saved);
  const today = new Date();

  let computed = 0;
  let summaryHtml = "";
  const summaryEl = document.getElementById("gf_calc_summary");
  const allocEl   = document.getElementById("gf_alloc_computed");
  const startHintEl = document.getElementById("gf_start_hint");
  const suggestionsEl = document.getElementById("gf_start_suggestions");

  if (target <= 0) {
    if (allocEl) allocEl.textContent = "—";
    if (summaryEl) summaryEl.style.display = "none";
    if (suggestionsEl) suggestionsEl.style.display = "none";
    return;
  }

  // ── Monthly allocation calc ────────────────────────────────────
  if (endVal && startVal) {
    const endDate   = new Date(endVal + "T00:00:00");
    const startDate = new Date(startVal + "T00:00:00");
    const monthsFromStart = Math.max(1, (endDate.getFullYear()-startDate.getFullYear())*12 + (endDate.getMonth()-startDate.getMonth()));
    const monthsFromNow   = Math.max(0, (endDate.getFullYear()-today.getFullYear())*12 + (endDate.getMonth()-today.getMonth()));
    computed = monthsFromStart > 0 ? Math.ceil(remaining / monthsFromStart) : remaining;

    const onTime = computed <= available;
    const feasibility = available > 0
      ? (computed / available * 100).toFixed(0) + "% of your available " + formatCurrency(available) + "/mo savings capacity"
      : "no historical savings data";

    summaryHtml = `
      <div class="gf-calc-row ${onTime?'ok':'warn'}">
        <span class="gf-calc-icon">${onTime?'✅':'⚠️'}</span>
        <div class="gf-calc-body">
          <strong>${formatCurrency(computed)}/mo</strong> needed over ${monthsFromStart} month${monthsFromStart===1?"":"s"} to reach ${formatCurrency(target)} by ${endDate.toLocaleDateString("en-SG",{month:"short",year:"numeric"})}
          <div class="gf-calc-sub">That's ${feasibility}${onTime ? " — you can hit this!" : " — this exceeds your capacity. Extend the deadline or reduce the target."}</div>
          ${monthsFromNow !== monthsFromStart ? `<div class="gf-calc-sub">📌 Starting from your chosen start date (${monthsFromNow} months remain from today to deadline).</div>` : ""}
        </div>
      </div>`;
  } else if (endVal) {
    const endDate = new Date(endVal + "T00:00:00");
    const monthsLeft = Math.max(1, (endDate.getFullYear()-today.getFullYear())*12 + (endDate.getMonth()-today.getMonth()));
    computed = Math.ceil(remaining / monthsLeft);

    // ── Start date suggestions ─────────────────────────────────
    // "Comfortable" = you could do it in half the time, so start that many months later
    // "Latest possible" = start just in time with your full available savings
    const latestMonths = available > 0 ? Math.floor(remaining / available) : monthsLeft;
    const comfortMonths = Math.ceil(monthsLeft * 0.5);

    const latestStart = new Date(today.getFullYear(), today.getMonth() + Math.max(0, monthsLeft - latestMonths), 1);
    const comfortStart = new Date(today.getFullYear(), today.getMonth() + Math.max(0, monthsLeft - comfortMonths), 1);

    const lsStr = toDateInputValue(latestStart);
    const csStr = toDateInputValue(comfortStart);
    const nowStr = toDateInputValue(today);

    if (suggestionsEl) {
      suggestionsEl.style.display = "block";
      suggestionsEl.innerHTML = `
        <div class="gf-suggest-label">Start date suggestions:</div>
        <div class="gf-suggest-btns">
          <button class="gf-suggest-btn gf-suggest-now" onclick="document.getElementById('gf_start').value='${nowStr}';_updateAddGoalCalc();">
            <span class="gfs-icon">🚀</span>
            <span class="gfs-body"><span class="gfs-title">Start Today</span><span class="gfs-sub">${formatCurrency(computed)}/mo — spread it out</span></span>
          </button>
          ${csStr > nowStr ? `<button class="gf-suggest-btn gf-suggest-comfort" onclick="document.getElementById('gf_start').value='${csStr}';_updateAddGoalCalc();">
            <span class="gfs-icon">😌</span>
            <span class="gfs-body"><span class="gfs-title">Start ${comfortStart.toLocaleDateString("en-SG",{month:"short",year:"numeric"})}</span><span class="gfs-sub">halfway point — ${formatCurrency(Math.ceil(remaining/comfortMonths))}/mo</span></span>
          </button>` : ""}
          ${lsStr > csStr ? `<button class="gf-suggest-btn gf-suggest-late" onclick="document.getElementById('gf_start').value='${lsStr}';_updateAddGoalCalc();">
            <span class="gfs-icon">⏰</span>
            <span class="gfs-body"><span class="gfs-title">Start ${latestStart.toLocaleDateString("en-SG",{month:"short",year:"numeric"})}</span><span class="gfs-sub">latest possible — ${formatCurrency(available)}/mo (full capacity)</span></span>
          </button>` : ""}
        </div>
      `;
    }

    const onTime = computed <= available;
    summaryHtml = `
      <div class="gf-calc-row ${onTime?'ok':'warn'}">
        <span class="gf-calc-icon">${onTime?'✅':'⚠️'}</span>
        <div class="gf-calc-body">
          <strong>${formatCurrency(computed)}/mo</strong> needed over ${monthsLeft} months to reach ${formatCurrency(target)}
          <div class="gf-calc-sub">${onTime ? "Within your capacity — pick a start date above to adjust the monthly amount." : "Exceeds your available " + formatCurrency(available) + "/mo. Start sooner, reduce target, or extend deadline."}</div>
        </div>
      </div>`;
  } else {
    // No deadline — suggest based on available capacity
    computed = available > 0 ? Math.round(available * 0.3) : 0;
    if (suggestionsEl) suggestionsEl.style.display = "none";
    summaryHtml = computed > 0
      ? `<div class="gf-calc-row ok"><span class="gf-calc-icon">💡</span><div class="gf-calc-body">No deadline — suggested <strong>${formatCurrency(computed)}/mo</strong> (30% of your ${formatCurrency(available)}/mo available). Set a deadline to get an exact number and start date suggestions.</div></div>`
      : `<div class="gf-calc-row neutral"><span class="gf-calc-icon">💡</span><div class="gf-calc-body">Set a target deadline to calculate the required monthly allocation and get smart start date suggestions.</div></div>`;
  }

  if (allocEl) allocEl.textContent = computed > 0 ? formatCurrency(computed) + "/mo" : "—";
  document.getElementById("gf_alloc_final").value = computed;
  // Keep override input in sync
  const overrideInp = document.getElementById("gf_alloc");
  if (overrideInp && overrideInp.style.display !== "none" && !overrideInp.dataset.manuallySet) {
    overrideInp.value = computed || "";
  }
  if (summaryEl) {
    summaryEl.style.display = summaryHtml ? "block" : "none";
    summaryEl.innerHTML = summaryHtml;
  }
  // Start date hint
  if (startHintEl && endVal) {
    const startDate = new Date((startVal||toDateInputValue(today)) + "T00:00:00");
    const endDate   = new Date(endVal + "T00:00:00");
    const mos = Math.max(0, (endDate.getFullYear()-startDate.getFullYear())*12+(endDate.getMonth()-startDate.getMonth()));
    startHintEl.textContent = mos > 0 ? mos + " months to deadline" : "deadline is this month";
  } else if (startHintEl) {
    startHintEl.textContent = "";
  }
}

function toggleGoalEdit(idx) {
  const el = document.getElementById("editGoal_" + idx);
  el.style.display = el.style.display === "none" ? "block" : "none";
}

function toggleDeductForm(idx) {
  const el = document.getElementById("deductGoal_" + idx);
  el.style.display = el.style.display === "none" ? "block" : "none";
}

async function toggleGoalAccount(checkbox) {
  const name = checkbox.value;
  if (checkbox.checked) {
    if (!goalSavingsAccts.includes(name)) goalSavingsAccts.push(name);
  } else {
    goalSavingsAccts = goalSavingsAccts.filter(n => n !== name);
  }
  renderGoalsPage();
  // Auto-save silently — no alert, just log
  try {
    await writeBudgetSetupRange("AD2:AD2", [[goalSavingsAccts.join("|")]]);
    log("Account selection auto-saved.");
  } catch(err) {
    log("Auto-save failed: " + err.message);
  }
}

function buildGoalsSaveValues() {
  const values = goalsData.map(g => [
    g.name, g.target, g.manualSaved, g.monthlyAlloc,
    g.startDate, g.endDate, g.urgency, g.notes, g.color,
    g.goalBuffer || 0, g.priority
  ]);
  while (values.length < MAX_GOALS) values.push(["","","","","","","","","","",""]);
  return values;
}

function setGoalsAutoSaveStatus(message, tone = "") {
  ["goalsAutosaveStatus", "allocAutosaveStatus", "boostAutosaveStatus"].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = message;
    el.className = "panel-hint goals-autosave-status" + (tone ? " " + tone : "");
  });
}

function scheduleGoalsAutoSave(delay = 700) {
  clearTimeout(goalsAutoSaveTimer);
  setGoalsAutoSaveStatus("Saving soon...");
  goalsAutoSaveTimer = setTimeout(() => {
    persistGoalsToExcel({ silent: true });
  }, delay);
}

async function persistGoalsToExcel(options = {}) {
  if (goalsAutoSaveInFlight) {
    scheduleGoalsAutoSave(1000);
    return false;
  }

  const silent = options.silent === true;
  const includeBoosts = options.includeBoosts !== false;
  goalsAutoSaveInFlight = true;

  try {
    setGoalsAutoSaveStatus("Saving...");
    await writeBudgetSetupRange(GOALS_RANGE, buildGoalsSaveValues());
    if (includeBoosts) await saveIncomeBoosts();
    setGoalsAutoSaveStatus("Saved to Excel", "ok");
    log("Goals saved.");
    return true;
  } catch (err) {
    setGoalsAutoSaveStatus("Save failed", "error");
    log("SAVE ERROR: " + err.message);
    if (!silent) throw err;
    return false;
  } finally {
    goalsAutoSaveInFlight = false;
  }
}

// ─── CRUD ─────────────────────────────────────────────────────────

async function addGoal() {
  const name   = document.getElementById("gf_name").value.trim();
  const target = parseFloat(document.getElementById("gf_target").value) || 0;
  const savedInput = document.getElementById("gf_saved");
  const saved  = savedInput ? (parseFloat(savedInput.value) || 0) : 0;
  const start  = document.getElementById("gf_start").value;
  const end    = document.getElementById("gf_end").value;
  const urg    = document.getElementById("gf_urgency").value;
  const notes  = document.getElementById("gf_notes").value.trim();
  const color  = document.getElementById("gf_color").value;

  // Use override input if visible, otherwise use computed value
  const overrideInp = document.getElementById("gf_alloc");
  const computedVal = parseFloat(document.getElementById("gf_alloc_final").value) || 0;
  const alloc = (overrideInp && overrideInp.style.display !== "none")
    ? (parseFloat(overrideInp.value) || 0)
    : computedVal;

  if (!name)   { alert("Please enter a goal name."); return; }
  if (target <= 0) { alert("Please enter a target amount."); return; }
  if (goalsData.length >= MAX_GOALS) { alert("Maximum " + MAX_GOALS + " goals."); return; }

  goalsData.push({ name, target, manualSaved:saved, monthlyAlloc:alloc, startDate:start, endDate:end, urgency:urg, notes, color, goalBuffer:0, priority:0 });
  renderGoalsPage();
  await persistGoalsToExcel({ silent: true });
}

async function saveEditGoal(idx) {
  goalsData[idx] = {
    name:         document.getElementById("eg_name_"   +idx).value.trim(),
    target:       parseFloat(document.getElementById("eg_target_"+idx).value) || 0,
    manualSaved:  parseFloat(document.getElementById("eg_saved_" +idx).value) || 0,
    monthlyAlloc: parseFloat(document.getElementById("eg_alloc_" +idx).value) || 0,
    startDate:    document.getElementById("eg_start_"  +idx).value,
    endDate:      document.getElementById("eg_end_"    +idx).value,
    urgency:      document.getElementById("eg_urgency_"+idx).value,
    notes:        document.getElementById("eg_notes_"  +idx).value.trim(),
    color:        document.getElementById("eg_color_"  +idx).value,
    goalBuffer:   Math.min(100, Math.max(0, parseFloat(document.getElementById("eg_buffer_"+idx).value) || 0)),
    priority:     goalsData[idx].priority,
  };
  renderGoalsPage();
  await persistGoalsToExcel({ silent: true });
}

async function deleteGoal(idx) {
  if (!confirm(`Delete goal "${goalsData[idx].name}"?`)) return;
  goalsData.splice(idx, 1);
  renderGoalsPage();
  await persistGoalsToExcel({ silent: true });
}

async function saveDeductGoal(idx) {
  const goal   = goalsData[idx];
  const amtVal = document.getElementById("dg_amt_"  +idx).value;
  const note   = document.getElementById("dg_note_" +idx).value.trim() || "Goal expense";
  const acct   = document.getElementById("dg_acct_" +idx).value;

  if (!amtVal) { alert("Please enter an amount."); return; }
  if (!acct)   { alert("Please select an account."); return; }

  const amount = parseFloat(amtVal);
  try {
    log("Saving goal deduction...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data    = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;
    const today   = new Date();
    const dateStr = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");
    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:F${nextRow}`, [[
      dateStr, note, amount, "Savings Goal", "Goal: " + goal.name, acct
    ]]);
    alert("Deduction saved!");
    await loadGoalsPage();
  } catch(err) { alert("Failed: " + err.message); }
}

// ─── Save to Excel ─────────────────────────────────────────────────

async function saveGoalsToExcel(options = {}) {
  try {
    log("Saving goals...");
    const saved = await persistGoalsToExcel(options);
    if (!saved) return;
    if (options.silent !== true) alert("Goals saved to Excel.");
    log("Done.");
  } catch(err) { log("SAVE ERROR: " + err.message); alert(err.message); }
}

// saveGoalAccountsToExcel removed — account selection now auto-saves on toggle

// ─── Utility ──────────────────────────────────────────────────────

function toDateInputValue(d) {
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

/** Parse a goal date value that might be an Excel serial number or a YYYY-MM-DD string */
function parseGoalDateValue(value) { return _parseGoalDateValue(value); }

/** Convert a raw goal date value (serial or string) to YYYY-MM-DD string for input[type=date] */
function goalDateToInputValue(value) { return _goalDateToInputValue(value); }

/** Display a goal date value in a human-readable format */
function formatDateDisplay(value) {
  if (!value) return "";
  const d = _parseGoalDateValue(value);
  if (!d) return String(value);
  return d.toLocaleDateString("en-SG", { day:"2-digit", month:"short", year:"numeric" });
}

function clean(v)    { return String(v??"").trim(); }
function accountKey(v) { return clean(v).toLowerCase().replace(/\s+/g, " "); }
function fieldKey(v) { return clean(v).toLowerCase().replace(/[^a-z0-9]/g, ""); }
function getAmount(v){ const n = Number(String(v).replace(/[$,]/g,"")); return isNaN(n) ? 0 : Math.abs(n); }
function getRowValue(row, fieldName) {
  const wanted = fieldKey(fieldName);
  const key = Object.keys(row || {}).find(k => fieldKey(k) === wanted);
  return key ? row[key] : undefined;
}
function isClaimableRow(row) {
  return clean(getRowValue(row, "Claimable")).toLowerCase() === "yes";
}
function getClaimAmount(row) {
  const expenseAmount = getAmount(getRowValue(row, "Amount"));
  const storedClaimAmount = getAmount(getRowValue(row, "Claim Amount"));
  const claimAmount = storedClaimAmount > 0 ? storedClaimAmount : expenseAmount;
  return Math.min(expenseAmount, claimAmount);
}
function getPendingClaimReceivableAmount(row) {
  if (!isClaimableRow(row)) return 0;
  if (clean(getRowValue(row, "Claim Status")).toLowerCase() !== "pending") return 0;
  return getClaimAmount(row);
}
function getClaimReceivableAccount(row) {
  return clean(getRowValue(row, "Claim Account"));
}
function getClaimAdjustedExpenseAmount(row) {
  const amount = getAmount(row["Amount"]);
  if (!isClaimableRow(row)) return amount;
  return 0;
}
function formatCurrency(v) {
  return v.toLocaleString("en-SG", { style:"currency", currency:"SGD", minimumFractionDigits:2, maximumFractionDigits:2 });
}
function escapeHtml(v) {
  return String(v).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
function parseExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") { const d = XLSX.SSF.parse_date_code(value); return new Date(d.y,d.m-1,d.d); }
  const s = String(value).trim();
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return new Date(+ddmm[3], +ddmm[2]-1, +ddmm[1]);
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3]);
  return null;
}
