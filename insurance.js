const INSURANCE_SHEET = "Insurance";
const INSURANCE_CURRENT_YEAR = new Date().getFullYear();
const INSURANCE_YEARS = Array.from({ length: 6 }, (_, index) => INSURANCE_CURRENT_YEAR - 2 + index);
const INSURANCE_YEAR_HEADERS = INSURANCE_YEARS.flatMap(year => [
  `Premium ${year}`,
  `Paid ${year}`,
  `Remark ${year}`
]);
const INSURANCE_BASE_HEADERS = [
  "Owner",
  "Insurer",
  "Policy Name",
  "Policy No",
  "Cover Start Date",
  "Policy Type",
  "Premium Type",
  "Annual Premium",
  "Renewal Month"
];
const INSURANCE_META_HEADERS = [
  "Claims Possible",
  "Comments",
  "Status",
  "Document Link",
  "Goal Name",
  "Goal Start Date",
  "Goal End Date",
  "Auto Recurring Goal",
  "Last Paid Date",
  "Last Paid Amount",
  "Last Paid Account",
  "Transaction Row"
];
const INSURANCE_HEADERS = [
  ...INSURANCE_BASE_HEADERS,
  ...INSURANCE_YEAR_HEADERS,
  ...INSURANCE_META_HEADERS
];
const INSURANCE_GOALS_FIRST_ROW = 2;
const INSURANCE_GOALS_MAX_ROWS = 19;
const INSURANCE_GOALS_START_COL = 18; // Budget Setup column S
const INSURANCE_GOAL_COLORS = [
  "#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed",
  "#0891b2", "#db2777", "#65a30d", "#ea580c", "#0f766e"
];

let insuranceRows = [];
let insuranceHeaderMap = {};
let insuranceAccounts = [];
let insuranceGoals = [];
let insuranceGroupPayRows = [];
let insuranceSheetReady = false;
let insuranceSheetState = "unknown";
let insurancePolicySaving = false;

async function loadInsurancePage() {
  try {
    clearOutput();
    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const insuranceSheet = workbook.Sheets[INSURANCE_SHEET];
    const budgetSheet = workbook.Sheets["Budget Setup"];

    insuranceAccounts = budgetSheet ? readInsuranceAccounts(budgetSheet) : [];
    insuranceGoals = budgetSheet ? readInsuranceGoals(budgetSheet) : [];

    if (!insuranceSheet) {
      insuranceRows = [];
      insuranceHeaderMap = {};
      insuranceSheetReady = false;
      insuranceSheetState = "missing";
      showInsuranceSetupNotice("missing");
      renderInsurancePage();
      return;
    }

    const parsed = readInsuranceSheet(insuranceSheet);
    insuranceSheetReady = parsed.ready;
    insuranceRows = parsed.rows;
    insuranceHeaderMap = parsed.headerMap;

    if (!insuranceSheetReady) showInsuranceSetupNotice("invalid");
    else hideInsuranceSetupNotice();
    insuranceSheetState = insuranceSheetReady ? "ready" : "invalid";

    renderInsuranceFilters();
    renderInsurancePage();
    log("Insurance page loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    alert(err.message);
    console.error(err);
  }
}

function showInsuranceSetupNotice(type) {
  const notice = document.getElementById("insuranceSetupNotice");
  if (!notice) return;
  notice.style.display = "block";
  notice.innerHTML = type === "missing"
    ? `
      <strong>Insurance sheet not found.</strong>
      <span>Create a normalized Insurance sheet, then paste or enter policies using the template columns.</span>
      <button onclick="createInsuranceSheet()">Create Insurance Sheet</button>
    `
    : `
      <strong>Insurance sheet needs normalized headers.</strong>
      <span>The page expects one row per policy with columns like Owner, Policy Name, Annual Premium, Claims Possible, and Comments.</span>
      <button onclick="writeInsuranceTemplateHeaders()">Write Template Headers</button>
    `;
}

function hideInsuranceSetupNotice() {
  const notice = document.getElementById("insuranceSetupNotice");
  if (notice) notice.style.display = "none";
}

async function createInsuranceSheet() {
  try {
    log("Creating Insurance sheet...");
    await addExcelWorksheet(INSURANCE_SHEET);
    await writeExcelRange(INSURANCE_SHEET, `A1:${columnLetter(INSURANCE_HEADERS.length)}1`, [INSURANCE_HEADERS]);
    setInsuranceHeaderMapFromHeaders(INSURANCE_HEADERS);
    insuranceSheetReady = true;
    insuranceSheetState = "ready";
    log("Insurance sheet created.");
    alert("Insurance sheet created. Add policies in Excel, then reload this page.");
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to create Insurance sheet: " + err.message);
  }
}

async function writeInsuranceTemplateHeaders() {
  if (!confirm("Write the normalized Insurance template headers to row 1? This may overwrite the current first row.")) return;
  try {
    await writeExcelRange(INSURANCE_SHEET, `A1:${columnLetter(INSURANCE_HEADERS.length)}1`, [INSURANCE_HEADERS]);
    setInsuranceHeaderMapFromHeaders(INSURANCE_HEADERS);
    insuranceSheetReady = true;
    insuranceSheetState = "ready";
    alert("Template headers written. Move your policy rows under the headers, then reload.");
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to write template headers: " + err.message);
  }
}

function readInsuranceSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  if (rows.length === 0) return { ready: false, rows: [], headerMap: {} };

  const headers = rows[0].map(cleanInsuranceText);
  const headerMap = {};
  headers.forEach((header, index) => {
    if (header) headerMap[normaliseHeader(header)] = index;
  });

  const hasRequiredHeaders = getHeaderIndex(headerMap, "Owner") >= 0
    && getHeaderIndex(headerMap, "Policy Name") >= 0
    && getHeaderIndex(headerMap, "Annual Premium") >= 0;

  if (!hasRequiredHeaders) return { ready: false, rows: [], headerMap };

  const dataRows = rows.slice(1)
    .map((row, index) => parseInsuranceRow(row, index + 2, headerMap))
    .filter(row => row.owner || row.policyName || row.policyNo);

  return { ready: true, rows: dataRows, headerMap };
}

function parseInsuranceRow(row, excelRow, headerMap) {
  const currentYear = INSURANCE_CURRENT_YEAR;
  const paidHeader = "Paid " + currentYear;
  const remarkHeader = "Remark " + currentYear;
  const policyName = getInsuranceCell(row, headerMap, "Policy Name");
  const insurer = getInsuranceCell(row, headerMap, "Insurer") || inferInsurer(policyName);
  const startDate = parseInsuranceDate(getInsuranceCell(row, headerMap, "Cover Start Date"));
  const renewalMonth = getInsuranceCell(row, headerMap, "Renewal Month");
  const annualPremium = toInsuranceNumber(getInsuranceCell(row, headerMap, "Annual Premium"));
  const paidThisYear = toInsuranceNumber(getInsuranceCell(row, headerMap, paidHeader));
  const status = getInsuranceCell(row, headerMap, "Status") || "Active";
  const premiumType = getInsuranceCell(row, headerMap, "Premium Type");
  let dueDate = computeNextDueDate(startDate, renewalMonth);
  let dueYear = dueDate ? dueDate.getFullYear() : currentYear;
  let paidForDueYear = getInsuranceYearAmount(row, headerMap, "Paid", dueYear, 0);

  while (dueDate && paidForDueYear > 0) {
    dueDate = new Date(dueDate.getFullYear() + 1, dueDate.getMonth(), dueDate.getDate());
    dueYear = dueDate.getFullYear();
    paidForDueYear = getInsuranceYearAmount(row, headerMap, "Paid", dueYear, 0);
  }

  const renewalPremium = getInsuranceYearAmount(row, headerMap, "Premium", dueYear, annualPremium);

  return {
    excelRow,
    owner: getInsuranceCell(row, headerMap, "Owner"),
    insurer,
    policyName,
    policyNo: getInsuranceCell(row, headerMap, "Policy No"),
    coverStartDate: startDate,
    coverStartRaw: getInsuranceCell(row, headerMap, "Cover Start Date"),
    policyType: getInsuranceCell(row, headerMap, "Policy Type") || inferPolicyType(policyName),
    premiumType,
    annualPremium,
    renewalPremium,
    renewalMonth,
    dueYear,
    paidThisYear,
    paidForDueYear,
    remarkThisYear: getInsuranceCell(row, headerMap, remarkHeader),
    claimsPossible: getInsuranceCell(row, headerMap, "Claims Possible"),
    comments: getInsuranceCell(row, headerMap, "Comments"),
    status,
    documentLink: getInsuranceCell(row, headerMap, "Document Link"),
    goalName: getInsuranceCell(row, headerMap, "Goal Name"),
    goalStartDate: parseInsuranceDate(getInsuranceCell(row, headerMap, "Goal Start Date")),
    goalEndDate: parseInsuranceDate(getInsuranceCell(row, headerMap, "Goal End Date")),
    autoRecurringGoal: getHeaderIndex(headerMap, "Auto Recurring Goal") < 0
      ? true
      : isInsuranceTruthy(getInsuranceCell(row, headerMap, "Auto Recurring Goal")),
    lastPaidDate: getInsuranceCell(row, headerMap, "Last Paid Date"),
    lastPaidAmount: toInsuranceNumber(getInsuranceCell(row, headerMap, "Last Paid Amount")),
    lastPaidAccount: getInsuranceCell(row, headerMap, "Last Paid Account"),
    transactionRow: getInsuranceCell(row, headerMap, "Transaction Row"),
    premiumThisYear: getInsuranceYearAmount(row, headerMap, "Premium", currentYear, annualPremium),
    premiumNextYear: getInsuranceYearAmount(row, headerMap, "Premium", currentYear + 1, annualPremium),
    dueDate,
    isCash: premiumType.toLowerCase().includes("cash"),
    isCpf: premiumType.toLowerCase().includes("cpf") || premiumType.toLowerCase().includes("medisave")
  };
}

function readInsuranceAccounts(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return rows.slice(1)
    .map(row => cleanInsuranceText(row[9]))
    .filter(Boolean);
}

function readInsuranceGoals(sheet) {
  const goals = [];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "S2:AC20", blankrows: false });

  rows.forEach((row, index) => {
    goals.push(buildInsuranceGoalFromValues(index + INSURANCE_GOALS_FIRST_ROW, row, index));
  });

  for (let offset = rows.length; offset < INSURANCE_GOALS_MAX_ROWS; offset++) {
    const excelRow = INSURANCE_GOALS_FIRST_ROW + offset;
    goals.push(buildInsuranceGoalFromValues(excelRow, [], offset));
  }

  return goals;
}

function buildInsuranceGoalFromValues(excelRow, values, offset = 0) {
  return {
    excelRow,
    name: cleanInsuranceText(values[0]),
    target: toInsuranceNumber(values[1]),
    manualSaved: toInsuranceNumber(values[2]),
    monthlyAlloc: toInsuranceNumber(values[3]),
    startDate: insuranceGoalDateInputValue(values[4]),
    endDate: insuranceGoalDateInputValue(values[5]),
    urgency: cleanInsuranceText(values[6]) || "Medium",
    notes: cleanInsuranceText(values[7]),
    color: cleanInsuranceText(values[8]) || INSURANCE_GOAL_COLORS[offset % INSURANCE_GOAL_COLORS.length],
    goalBuffer: toInsuranceNumber(values[9]),
    priority: toInsuranceNumber(values[10])
  };
}

function renderInsuranceFilters() {
  setFilterOptions("insuranceOwnerFilter", "All owners", [...new Set(insuranceRows.map(row => row.owner).filter(Boolean))]);
  setFilterOptions("insurancePremiumFilter", "All premium types", [...new Set(insuranceRows.map(row => row.premiumType).filter(Boolean))]);
  setFilterOptions("insuranceStatusFilter", "All statuses", [...new Set(insuranceRows.map(row => row.status).filter(Boolean))]);
}

function setFilterOptions(elementId, allLabel, values) {
  const select = document.getElementById(elementId);
  if (!select) return;
  const current = select.value;
  select.innerHTML = `<option value="">${allLabel}</option>`;
  values.sort().forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  if (values.includes(current)) select.value = current;
}

function getFilteredInsuranceRows() {
  const owner = document.getElementById("insuranceOwnerFilter")?.value || "";
  const premium = document.getElementById("insurancePremiumFilter")?.value || "";
  const status = document.getElementById("insuranceStatusFilter")?.value || "";
  return insuranceRows.filter(row => {
    if (owner && row.owner !== owner) return false;
    if (premium && row.premiumType !== premium) return false;
    if (status && row.status !== status) return false;
    return true;
  });
}

function renderInsurancePage() {
  const rows = getFilteredInsuranceRows();
  renderInsuranceSummary(rows);
  renderInsuranceOwnerBars(rows);
  renderInsuranceDueTable(rows);
  renderInsurancePolicyTable(rows);
  renderInsurancePaymentOptions();
}

function renderInsuranceSummary(rows) {
  const active = rows.filter(row => row.status.toLowerCase() !== "inactive");
  const total = active.reduce((sum, row) => sum + row.renewalPremium, 0);
  const cash = active.filter(row => row.isCash).reduce((sum, row) => sum + row.renewalPremium, 0);
  const cpf = active.filter(row => row.isCpf).reduce((sum, row) => sum + row.renewalPremium, 0);

  setInsuranceMoney("insuranceAnnualTotal", total);
  setInsuranceMoney("insuranceCashTotal", cash);
  setInsuranceMoney("insuranceCpfTotal", cpf);
  setInsuranceMoney("insuranceCashReserve", cash / 12);
}

function renderInsuranceOwnerBars(rows) {
  const container = document.getElementById("insuranceOwnerBars");
  const badge = document.getElementById("insuranceOwnerBadge");
  if (!container) return;

  const totals = {};
  rows.forEach(row => {
    const owner = row.owner || "Unassigned";
    totals[owner] = (totals[owner] || 0) + row.renewalPremium;
  });

  const items = Object.entries(totals).sort((a, b) => b[1] - a[1]);
  const max = Math.max(...items.map(([, total]) => total), 1);
  if (badge) badge.textContent = `${items.length} owner${items.length === 1 ? "" : "s"}`;

  if (!items.length) {
    container.innerHTML = `<p class="insurance-empty">No policies found.</p>`;
    return;
  }

  container.innerHTML = "";
  items.forEach(([owner, total]) => {
    const row = document.createElement("div");
    row.className = "insurance-bar-row";
    row.innerHTML = `
      <div class="insurance-bar-label"><span>${escapeInsuranceHtml(owner)}</span><strong>${formatInsuranceCurrency(total)}</strong></div>
      <div class="insurance-bar-track"><div class="insurance-bar-fill" style="width:${Math.max(5, total / max * 100)}%;"></div></div>
    `;
    container.appendChild(row);
  });
}

function renderInsuranceDueTable(rows) {
  const table = document.getElementById("insuranceDueTable");
  const badge = document.getElementById("insuranceDueBadge");
  if (!table) return;

  const today = startOfDay(new Date());
  const horizon = new Date(today);
  horizon.setDate(today.getDate() + 90);

  const dueRows = rows
    .filter(row => row.status.toLowerCase() !== "inactive")
    .filter(row => row.dueDate && row.dueDate >= today && row.dueDate <= horizon)
    .sort((a, b) => a.dueDate - b.dueDate);

  if (badge) badge.textContent = `${dueRows.length} due soon`;
  table.innerHTML = `<tr><th>Due</th><th>Policy</th><th>Owner</th><th>Amount</th><th>Status</th></tr>`;
  if (!dueRows.length) {
    table.innerHTML += `<tr><td colspan="5" style="text-align:center;color:#999;">No premium due in the next 90 days</td></tr>`;
    return;
  }

  dueRows.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${formatInsuranceDate(row.dueDate)}</td>
      <td>${escapeInsuranceHtml(row.policyName)}</td>
      <td>${escapeInsuranceHtml(row.owner)}</td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;">${formatInsuranceCurrency(row.renewalPremium)}</td>
      <td><span class="insurance-status ${row.paidForDueYear > 0 ? "paid" : "unpaid"}">${row.paidForDueYear > 0 ? "Paid" : "Pending"}</span></td>
    `;
    table.appendChild(tr);
  });
}

function renderInsurancePolicyTable(rows) {
  const table = document.getElementById("insurancePolicyTable");
  if (!table) return;

  table.innerHTML = `<tr>
    <th>Owner</th><th>Policy</th><th>Renewal Premium</th><th>Due</th><th>Paid Due Year</th>
    <th>Claims Possible</th><th>Comments</th><th>Goal</th><th>Actions</th>
  </tr>`;

  if (!insuranceSheetReady) {
    table.innerHTML += `<tr><td colspan="9" style="text-align:center;color:#999;">Create or normalize the Insurance sheet to start tracking policies.</td></tr>`;
    return;
  }
  if (!rows.length) {
    table.innerHTML += `<tr><td colspan="9" style="text-align:center;color:#999;">No policies match these filters.</td></tr>`;
    return;
  }

  rows.forEach(row => {
    const paidClass = row.paidForDueYear > 0 ? "paid" : "unpaid";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeInsuranceHtml(row.owner)}</td>
      <td>
        <strong>${escapeInsuranceHtml(row.policyName)}</strong>
        <div class="insurance-muted">${escapeInsuranceHtml(row.insurer)}${row.policyNo ? " · " + escapeInsuranceHtml(row.policyNo) : ""}</div>
        <div class="insurance-muted">${escapeInsuranceHtml(row.policyType)} · ${escapeInsuranceHtml(row.premiumType)}</div>
      </td>
      <td style="text-align:right;font-family:'DM Mono', monospace;font-weight:bold;">${formatInsuranceCurrency(row.renewalPremium)}</td>
      <td>${row.dueDate ? formatInsuranceDate(row.dueDate) : "-"}</td>
      <td><span class="insurance-status ${paidClass}">${row.paidForDueYear > 0 ? formatInsuranceCurrency(row.paidForDueYear) : "Unpaid " + row.dueYear}</span></td>
      <td>${escapeInsuranceHtml(row.claimsPossible || "-")}</td>
      <td>${escapeInsuranceHtml(row.comments || row.remarkThisYear || "-")}</td>
      <td>${escapeInsuranceHtml(row.goalName || "-")}</td>
      <td>
        <div class="insurance-row-actions">
          <button class="insurance-pay-btn" onclick="openInsurancePayModal(${row.excelRow})">Pay</button>
          <button class="insurance-light-btn" onclick="openInsurancePolicyModal(${row.excelRow})">Edit</button>
          <button class="insurance-danger-btn" onclick="deleteInsurancePolicy(${row.excelRow})">Remove</button>
        </div>
      </td>
    `;
    table.appendChild(tr);
  });
}

function renderInsurancePaymentOptions() {
  populateInsuranceAccountSelect("insurancePayAccount");
  populateInsuranceAccountSelect("insurancePaidThisYearAccount");
  populateInsuranceAccountSelect("insuranceGroupPayAccount");

  renderInsuranceGoalOptions();
}

function populateInsuranceAccountSelect(elementId) {
  const select = document.getElementById(elementId);
  if (!select) return;
  const current = select.value;
  select.innerHTML = `<option value="">Select account</option>`;
  insuranceAccounts.forEach(account => {
    const option = document.createElement("option");
    option.value = account;
    option.textContent = account;
    select.appendChild(option);
  });
  if (current && insuranceAccounts.includes(current)) select.value = current;
}

function openInsurancePayModal(excelRow) {
  const row = insuranceRows.find(item => item.excelRow === excelRow);
  if (!row) return;

  document.getElementById("insurancePayRow").value = String(excelRow);
  document.getElementById("insurancePayTitle").textContent = "Pay " + row.policyName;
  document.getElementById("insurancePaySubtitle").textContent = `${row.owner} · ${row.premiumType} · ${formatInsuranceCurrency(row.renewalPremium)} due ${row.dueYear}`;
  document.getElementById("insurancePayDate").value = toInsuranceDateInput(new Date());
  document.getElementById("insurancePayAmount").value = row.renewalPremium ? row.renewalPremium.toFixed(2) : "";
  document.getElementById("insurancePayGoalName").value = row.goalName || suggestedInsuranceGoalName(row.dueDate || new Date(row.dueYear, 0, 1));
  document.getElementById("insurancePayDescription").value = `Insurance - ${row.owner} - ${row.policyName}`;
  document.getElementById("insurancePayRemark").value = "";

  const modal = document.getElementById("insurancePayModal");
  if (modal) modal.style.display = "flex";
}

function closeInsurancePayModal() {
  const modal = document.getElementById("insurancePayModal");
  if (modal) modal.style.display = "none";
}

function openInsuranceGroupPayModal() {
  if (!insuranceSheetReady) {
    alert("Create or normalize the Insurance sheet first.");
    return;
  }

  insuranceGroupPayRows = insuranceRows
    .filter(row => row.status.toLowerCase() !== "inactive")
    .filter(row => row.isCash)
    .filter(row => row.paidThisYear <= 0)
    .map(row => ({
      ...row,
      currentYearAmount: getInsurancePolicyCurrentYearAmount(row)
    }))
    .filter(row => row.currentYearAmount > 0);

  if (!insuranceGroupPayRows.length) {
    alert(`No unpaid cash insurance policies found for ${INSURANCE_CURRENT_YEAR}.`);
    return;
  }

  document.getElementById("insuranceGroupPayYear").textContent = String(INSURANCE_CURRENT_YEAR);
  document.getElementById("insuranceGroupPayDate").value = toInsuranceDateInput(new Date());
  document.getElementById("insuranceGroupPayDescription").value = `Insurance group payment ${INSURANCE_CURRENT_YEAR}`;
  document.getElementById("insuranceGroupPayRemark").value = "";
  document.getElementById("insuranceGroupPayGoalName").value = getDefaultInsuranceGroupGoalName(insuranceGroupPayRows);
  renderInsuranceGoalOptions(document.getElementById("insuranceGroupPayGoalName").value);

  renderInsuranceGroupPayList();

  const modal = document.getElementById("insuranceGroupPayModal");
  if (modal) modal.style.display = "flex";
}

function closeInsuranceGroupPayModal() {
  const modal = document.getElementById("insuranceGroupPayModal");
  if (modal) modal.style.display = "none";
}

function renderInsuranceGroupPayList() {
  const container = document.getElementById("insuranceGroupPayList");
  if (!container) return;

  container.innerHTML = insuranceGroupPayRows.map(row => `
    <label class="insurance-group-pay-row">
      <input type="checkbox" value="${row.excelRow}" checked onchange="updateInsuranceGroupPayTotal()">
      <span>
        <strong>${escapeInsuranceHtml(row.owner)} - ${escapeInsuranceHtml(row.policyName)}</strong>
        <em>${escapeInsuranceHtml(row.goalName || "No linked goal yet")}</em>
      </span>
      <b>${formatInsuranceCurrency(row.currentYearAmount)}</b>
    </label>
  `).join("");

  updateInsuranceGroupPayTotal();
}

function updateInsuranceGroupPayTotal() {
  const totalEl = document.getElementById("insuranceGroupPayTotal");
  if (!totalEl) return;
  const selected = getSelectedInsuranceGroupPayRows();
  const total = selected.reduce((sum, row) => sum + row.currentYearAmount, 0);
  totalEl.textContent = formatInsuranceCurrency(total);
}

function getSelectedInsuranceGroupPayRows() {
  const checked = Array.from(document.querySelectorAll("#insuranceGroupPayList input[type='checkbox']:checked"))
    .map(input => Number(input.value));
  return insuranceGroupPayRows.filter(row => checked.includes(row.excelRow));
}

async function recordInsuranceGroupPayment() {
  const rows = getSelectedInsuranceGroupPayRows();
  const dateInput = document.getElementById("insuranceGroupPayDate").value;
  const account = document.getElementById("insuranceGroupPayAccount").value;
  const goalName = cleanInsuranceText(document.getElementById("insuranceGroupPayGoalName").value);
  const description = cleanInsuranceText(document.getElementById("insuranceGroupPayDescription").value) || `Insurance group payment ${INSURANCE_CURRENT_YEAR}`;
  const remark = cleanInsuranceText(document.getElementById("insuranceGroupPayRemark").value) || `Paid in group payment ${INSURANCE_CURRENT_YEAR}`;
  const total = rows.reduce((sum, row) => sum + row.currentYearAmount, 0);

  if (!rows.length) { alert("Select at least one policy to pay."); return; }
  if (!dateInput) { alert("Please enter the payment date."); return; }
  if (!account) { alert("Please select the payment account."); return; }
  if (!goalName) { alert("Please select or create a linked goal for this group payment."); return; }
  if (total <= 0) { alert("The selected policies have no payable amount."); return; }

  try {
    log("Recording group insurance payment...");
    const latestDue = rows
      .map(row => row.dueDate)
      .filter(Boolean)
      .sort((a, b) => b - a)[0] || new Date(INSURANCE_CURRENT_YEAR + 1, 0, 1);
    const plan = buildInsuranceGoalPlan({
      owner: "Group",
      policyName: description,
      premiumType: "Cash",
      amount: total,
      dueDate: latestDue
    });

    await syncInsuranceGoalFromPolicy({
      goalName,
      owner: "Group",
      policyName: description,
      autoRecurring: true,
      plan
    });

    const transactionRow = await appendInsuranceGoalTransaction({
      dateInput,
      description,
      amount: total,
      goalName,
      account
    });

    for (const row of rows) {
      await updateInsurancePaidCells({
        row,
        paidYear: INSURANCE_CURRENT_YEAR,
        dateInput,
        amount: row.currentYearAmount,
        account,
        remark,
        transactionRow
      });
    }

    closeInsuranceGroupPayModal();
    alert(`Group payment recorded: ${formatInsuranceCurrency(total)}.`);
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to record group payment: " + err.message);
  }
}

function openInsurancePolicyModal(excelRow = null) {
  if (!insuranceSheetReady && excelRow) {
    alert("Create or normalize the Insurance sheet first.");
    return;
  }

  const row = excelRow ? insuranceRows.find(item => item.excelRow === excelRow) : null;
  document.getElementById("insurancePolicyRow").value = row ? String(row.excelRow) : "";
  document.getElementById("insurancePolicyTitle").textContent = row ? "Edit Policy" : "Add Policy";
  document.getElementById("insurancePolicySubtitle").textContent = row
    ? `${row.owner || "No owner"} · ${row.policyName || "Untitled policy"}`
    : (insuranceSheetReady
      ? "Policy details save to the Insurance worksheet."
      : "Policy details will prepare the Insurance worksheet when saved.");

  const plan = buildInsuranceGoalPlanFromRow(row);
  document.getElementById("insuranceCurrentPremiumLabel").textContent = `Premium ${INSURANCE_CURRENT_YEAR}`;
  document.getElementById("insurancePaidThisYearLabel").textContent = `Paid for ${INSURANCE_CURRENT_YEAR}`;

  document.getElementById("insuranceOwnerInput").value = row?.owner || "";
  document.getElementById("insuranceInsurerInput").value = row?.insurer || "";
  document.getElementById("insurancePolicyNameInput").value = row?.policyName || "";
  document.getElementById("insurancePolicyNoInput").value = row?.policyNo || "";
  document.getElementById("insuranceCoverStartInput").value = row?.coverStartDate ? toInsuranceDateInput(row.coverStartDate) : "";
  document.getElementById("insuranceRenewalMonthInput").value = monthSelectValue(row?.renewalMonth || "");
  document.getElementById("insurancePolicyTypeInput").value = row?.policyType || "";
  document.getElementById("insurancePremiumTypeInput").value = row?.premiumType || "Cash";
  document.getElementById("insuranceAnnualPremiumInput").value = row?.annualPremium ? row.annualPremium.toFixed(2) : "";
  document.getElementById("insuranceCurrentPremiumInput").value = row?.premiumThisYear ? row.premiumThisYear.toFixed(2) : "";
  document.getElementById("insuranceStatusInput").value = row?.status || "Active";
  document.getElementById("insuranceGoalNameInput").value = row?.goalName || "";
  document.getElementById("insuranceGoalStartInput").value = row?.goalStartDate ? toInsuranceDateInput(row.goalStartDate) : plan.startDateInput;
  document.getElementById("insuranceGoalEndInput").value = row?.goalEndDate ? toInsuranceDateInput(row.goalEndDate) : plan.endDateInput;
  document.getElementById("insuranceAutoRecurringInput").checked = row ? row.autoRecurringGoal : true;
  document.getElementById("insurancePaidThisYearInput").checked = !!(row && row.paidThisYear > 0);
  document.getElementById("insurancePaidThisYearInput").dataset.wasPaid = row && row.paidThisYear > 0 ? "yes" : "no";
  document.getElementById("insurancePaidThisYearDate").value = toInsuranceDateInput(new Date());
  document.getElementById("insurancePaidThisYearAccount").value = row?.lastPaidAccount || "";
  toggleInsurancePaidThisYearFields();
  document.getElementById("insuranceDocumentLinkInput").value = row?.documentLink || "";
  document.getElementById("insuranceClaimsInput").value = row?.claimsPossible || "";
  document.getElementById("insuranceCommentsInput").value = row?.comments || "";

  clearInsuranceSmartState();
  renderInsuranceGoalOptions(row?.goalName || plan.goalName);
  updateInsuranceSmartGoalFields({ forceGoalName: !row });

  const modal = document.getElementById("insurancePolicyModal");
  if (modal) modal.style.display = "flex";
  refreshInsuranceGoalsForPicker(row?.goalName || plan.goalName);
}

function closeInsurancePolicyModal() {
  const modal = document.getElementById("insurancePolicyModal");
  if (modal) modal.style.display = "none";
}

function setInsurancePolicySaveBusy(isBusy) {
  insurancePolicySaving = isBusy;
  const button = document.getElementById("insuranceSavePolicyBtn");
  if (!button) return;
  button.disabled = isBusy;
  button.textContent = isBusy ? "Saving..." : "Save Policy";
}

async function saveInsurancePolicy(event) {
  if (event) event.preventDefault();
  if (insurancePolicySaving) return;

  const existingRow = Number(document.getElementById("insurancePolicyRow").value) || null;
  const owner = cleanInsuranceText(document.getElementById("insuranceOwnerInput").value);
  const policyName = cleanInsuranceText(document.getElementById("insurancePolicyNameInput").value);
  const renewalPremiumRaw = cleanInsuranceText(document.getElementById("insuranceCurrentPremiumInput").value);
  const renewalPremium = toInsuranceNumber(renewalPremiumRaw);
  let annualPremium = toInsuranceNumber(document.getElementById("insuranceAnnualPremiumInput").value);
  if (annualPremium <= 0 && renewalPremium > 0) annualPremium = renewalPremium;

  if (!owner) { alert("Please enter the policy owner."); return; }
  if (!policyName) { alert("Please enter the policy name."); return; }
  if (annualPremium <= 0) { alert("Please enter a default annual premium."); return; }

  try {
    setInsurancePolicySaveBusy(true);
    log(existingRow ? "Saving policy..." : "Adding policy...");
    await prepareInsuranceSheetForWrite();
    await ensureInsuranceHeaders();

    const premiumForGoal = renewalPremium || annualPremium;
    const plan = buildInsuranceGoalPlanFromForm({ owner, policyName, annualPremium, renewalPremium: premiumForGoal });
    plan.startDateInput = document.getElementById("insuranceGoalStartInput").value || plan.startDateInput;
    plan.endDateInput = document.getElementById("insuranceGoalEndInput").value || plan.endDateInput;
    plan.monthlyReserve = plan.target > 0 ? plan.target / monthsBetweenInsuranceDates(plan.startDateInput, plan.endDateInput) : 0;
    const targetRow = existingRow || await getNextInsuranceRowNumber();
    const autoRecurring = document.getElementById("insuranceAutoRecurringInput").checked;
    const paidThisYear = document.getElementById("insurancePaidThisYearInput").checked;
    const wasPaidThisYear = document.getElementById("insurancePaidThisYearInput").dataset.wasPaid === "yes";
    const paidDateInput = document.getElementById("insurancePaidThisYearDate").value;
    const paidAccount = document.getElementById("insurancePaidThisYearAccount").value;
    let goalName = getSelectedInsuranceGoalName();
    if (!goalName && document.getElementById("insuranceGoalExistingSelect")?.value === "__new__") {
      alert("Enter the new linked goal name, or choose an existing goal.");
      return;
    }
    if (!goalName && isInsuranceCashPremium(document.getElementById("insurancePremiumTypeInput").value)) {
      goalName = plan.goalName;
      setSelectedInsuranceGoalName(goalName);
      renderInsuranceGoalOptions(goalName);
    }

    if (paidThisYear && !wasPaidThisYear) {
      if (!paidDateInput) { alert("Please enter the paid date."); return; }
      if (!paidAccount) { alert("Please select the account used for this payment."); return; }
      if (!goalName) { alert("Please select or create a linked goal before recording payment."); return; }
    }

    await ensureInsuranceYearHeaders([INSURANCE_CURRENT_YEAR]);

    const valuesByHeader = {
      "Owner": owner,
      "Insurer": cleanInsuranceText(document.getElementById("insuranceInsurerInput").value) || inferInsurer(policyName),
      "Policy Name": policyName,
      "Policy No": cleanInsuranceText(document.getElementById("insurancePolicyNoInput").value),
      "Cover Start Date": formatInsuranceDateForExcelOrBlank(document.getElementById("insuranceCoverStartInput").value),
      "Policy Type": cleanInsuranceText(document.getElementById("insurancePolicyTypeInput").value) || inferPolicyType(policyName),
      "Premium Type": document.getElementById("insurancePremiumTypeInput").value || "Cash",
      "Annual Premium": annualPremium,
      "Renewal Month": document.getElementById("insuranceRenewalMonthInput").value,
      "Claims Possible": cleanInsuranceText(document.getElementById("insuranceClaimsInput").value),
      "Comments": cleanInsuranceText(document.getElementById("insuranceCommentsInput").value),
      "Status": document.getElementById("insuranceStatusInput").value || "Active",
      "Document Link": cleanInsuranceText(document.getElementById("insuranceDocumentLinkInput").value),
      "Goal Name": goalName,
      "Goal Start Date": formatInsuranceDateForExcelOrBlank(document.getElementById("insuranceGoalStartInput").value),
      "Goal End Date": formatInsuranceDateForExcelOrBlank(document.getElementById("insuranceGoalEndInput").value),
      "Auto Recurring Goal": autoRecurring ? "Yes" : "No"
    };
    if (renewalPremiumRaw) valuesByHeader[`Premium ${INSURANCE_CURRENT_YEAR}`] = renewalPremium;

    await writeInsurancePolicyFields(targetRow, valuesByHeader);
    const goalMessage = await syncInsuranceGoalFromPolicy({
      goalName,
      owner,
      policyName,
      autoRecurring,
      plan,
      promptOnMissing: false
    });

    let paymentMessage = "";
    if (paidThisYear && !wasPaidThisYear) {
      const transactionRow = await appendInsuranceGoalTransaction({
        dateInput: paidDateInput,
        description: `Insurance - ${owner} - ${policyName}`,
        amount: premiumForGoal,
        goalName,
        account: paidAccount
      });
      await updateInsurancePaidCells({
        row: {
          excelRow: targetRow,
          annualPremium,
          policyName,
          goalName
        },
        paidYear: INSURANCE_CURRENT_YEAR,
        dateInput: paidDateInput,
        amount: premiumForGoal,
        account: paidAccount,
        remark: `Paid for ${INSURANCE_CURRENT_YEAR}`,
        transactionRow
      });
      paymentMessage = `\nPayment recorded for ${INSURANCE_CURRENT_YEAR}.`;
    }

    closeInsurancePolicyModal();
    alert((existingRow ? "Policy updated." : "Policy added.") + goalMessage + paymentMessage);
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to save policy: " + err.message);
  } finally {
    setInsurancePolicySaveBusy(false);
  }
}

async function deleteInsurancePolicy(excelRow) {
  const row = insuranceRows.find(item => item.excelRow === excelRow);
  if (!row) return;
  if (!confirm(`Remove "${row.policyName}" from the Insurance sheet?`)) return;

  try {
    await ensureInsuranceHeaders();
    const width = Math.max(...Object.values(insuranceHeaderMap)) + 1;
    await writeExcelRange(INSURANCE_SHEET, `A${excelRow}:${columnLetter(width)}${excelRow}`, [Array(width).fill("")]);
    alert("Policy removed.");
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to remove policy: " + err.message);
  }
}

async function ensureInsuranceHeaders() {
  const existingIndexes = Object.values(insuranceHeaderMap).filter(Number.isInteger);
  const missing = INSURANCE_HEADERS.filter(header => getHeaderIndex(insuranceHeaderMap, header) < 0);
  if (!missing.length) return;

  const startIndex = existingIndexes.length ? Math.max(...existingIndexes) + 1 : 0;
  const startCol = columnLetter(startIndex + 1);
  const endCol = columnLetter(startIndex + missing.length);
  await writeExcelRange(INSURANCE_SHEET, `${startCol}1:${endCol}1`, [missing]);

  missing.forEach((header, index) => {
    insuranceHeaderMap[normaliseHeader(header)] = startIndex + index;
  });
}

async function prepareInsuranceSheetForWrite() {
  if (insuranceSheetReady) return;

  if (insuranceSheetState === "missing") {
    log("Creating Insurance sheet...");
    try {
      await addExcelWorksheet(INSURANCE_SHEET);
    } catch (err) {
      if (!/already|exist|name/i.test(err.message || "")) throw err;
    }
    await writeInsuranceHeadersForWrite();
    return;
  }

  const shouldWriteHeaders = confirm(
    "The Insurance worksheet headers are not in the expected format. Write the Insurance template headers to row 1 so this policy can be saved?"
  );
  if (!shouldWriteHeaders) {
    throw new Error("Insurance sheet is not ready for saving.");
  }
  await writeInsuranceHeadersForWrite();
}

async function writeInsuranceHeadersForWrite() {
  await writeExcelRange(INSURANCE_SHEET, `A1:${columnLetter(INSURANCE_HEADERS.length)}1`, [INSURANCE_HEADERS]);
  setInsuranceHeaderMapFromHeaders(INSURANCE_HEADERS);
  insuranceSheetReady = true;
  insuranceSheetState = "ready";
  hideInsuranceSetupNotice();
}

function setInsuranceHeaderMapFromHeaders(headers) {
  insuranceHeaderMap = {};
  headers.forEach((header, index) => {
    if (header) insuranceHeaderMap[normaliseHeader(header)] = index;
  });
}

async function ensureInsuranceYearHeaders(years) {
  const wanted = [];
  years
    .filter(year => Number.isInteger(Number(year)) && Number(year) >= 2000)
    .forEach(year => {
      ["Premium", "Paid", "Remark"].forEach(prefix => {
        const header = `${prefix} ${Number(year)}`;
        if (getHeaderIndex(insuranceHeaderMap, header) < 0 && !wanted.includes(header)) {
          wanted.push(header);
        }
      });
    });

  if (!wanted.length) return;

  const existingIndexes = Object.values(insuranceHeaderMap).filter(Number.isInteger);
  const startIndex = existingIndexes.length ? Math.max(...existingIndexes) + 1 : 0;
  const startCol = columnLetter(startIndex + 1);
  const endCol = columnLetter(startIndex + wanted.length);
  await writeExcelRange(INSURANCE_SHEET, `${startCol}1:${endCol}1`, [wanted]);

  wanted.forEach((header, index) => {
    insuranceHeaderMap[normaliseHeader(header)] = startIndex + index;
  });
}

async function writeInsurancePolicyFields(excelRow, valuesByHeader) {
  for (const [header, value] of Object.entries(valuesByHeader)) {
    const col = getHeaderIndex(insuranceHeaderMap, header);
    if (col >= 0) {
      await writeExcelRange(INSURANCE_SHEET, cellAddress(col, excelRow), [[value]]);
    }
  }
}

async function getNextInsuranceRowNumber() {
  const token = await getToken();
  const encodedPath = getEncodedExcelPath();
  const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
    ":/workbook/worksheets('" + INSURANCE_SHEET.replace(/'/g, "''") + "')/usedRange(valuesOnly=true)";
  const data = await graphGetJson(url, token);
  return Math.max(2, data.rowCount + 1);
}

async function recordInsurancePayment() {
  const excelRow = Number(document.getElementById("insurancePayRow").value);
  const row = insuranceRows.find(item => item.excelRow === excelRow);
  if (!row) return;

  const dateInput = document.getElementById("insurancePayDate").value;
  const amount = Number(document.getElementById("insurancePayAmount").value);
  const account = document.getElementById("insurancePayAccount").value;
  const goalName = cleanInsuranceText(document.getElementById("insurancePayGoalName").value) || row.goalName;
  const description = cleanInsuranceText(document.getElementById("insurancePayDescription").value) || `Insurance - ${row.policyName}`;
  const remark = cleanInsuranceText(document.getElementById("insurancePayRemark").value);

  if (!dateInput) { alert("Please enter a payment date."); return; }
  if (!Number.isFinite(amount) || amount <= 0) { alert("Please enter a valid payment amount."); return; }
  if (!account) { alert("Please select the account used for the payment transaction."); return; }
  if (!goalName) { alert("Please select or create a linked goal for this payment."); return; }

  try {
    log("Recording insurance payment...");
    const plan = buildInsuranceGoalPlanFromRow({ ...row, goalName, renewalPremium: amount });
    await syncInsuranceGoalFromPolicy({
      goalName,
      owner: row.owner,
      policyName: row.policyName,
      autoRecurring: row.autoRecurringGoal,
      plan
    });

    const transactionRow = await appendInsuranceGoalTransaction({
      dateInput,
      description,
      amount,
      goalName,
      account
    });

    await updateInsurancePaidCells({
      row,
      paidYear: row.dueYear,
      dateInput,
      amount,
      account,
      remark,
      transactionRow
    });

    closeInsurancePayModal();
    alert("Insurance paid and goal transaction recorded.");
    await loadInsurancePage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to record payment: " + err.message);
  }
}

async function appendInsuranceGoalTransaction({ dateInput, description, amount, goalName, account }) {
  const token = await getToken();
  const encodedPath = getEncodedExcelPath();
  const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
    ":/workbook/worksheets('" + CONFIG.sheetName.replace(/'/g, "''") + "')/usedRange(valuesOnly=true)";
  const data = await graphGetJson(url, token);
  const nextRow = data.rowCount + 1;
  const row = [formatInsuranceDateForExcel(dateInput), description, amount, "Saving Goals", "Goal: " + goalName, account, "", "", ""];
  await writeExcelRange(CONFIG.sheetName, `A${nextRow}:I${nextRow}`, [row]);
  return nextRow;
}

async function updateInsurancePaidCells({ row, paidYear, dateInput, amount, account, remark, transactionRow }) {
  const year = Number(paidYear || dateInput.split("-")[0]);
  await ensureInsuranceYearHeaders([year, year + 1]);
  const paidCol = getHeaderIndex(insuranceHeaderMap, "Paid " + year);
  const remarkCol = getHeaderIndex(insuranceHeaderMap, "Remark " + year);
  const lastPaidDateCol = getHeaderIndex(insuranceHeaderMap, "Last Paid Date");
  const lastPaidAmountCol = getHeaderIndex(insuranceHeaderMap, "Last Paid Amount");
  const lastPaidAccountCol = getHeaderIndex(insuranceHeaderMap, "Last Paid Account");
  const transactionRowCol = getHeaderIndex(insuranceHeaderMap, "Transaction Row");

  if (paidCol < 0 || remarkCol < 0) {
    throw new Error(`Insurance sheet is missing Paid ${year} or Remark ${year} columns.`);
  }

  const displayDate = formatInsuranceDateForExcel(dateInput);
  const finalRemark = remark || `Paid ${displayDate}${account ? " via " + account : ""}`;

  await writeExcelRange(INSURANCE_SHEET, cellAddress(paidCol, row.excelRow), [[amount]]);
  await writeExcelRange(INSURANCE_SHEET, cellAddress(remarkCol, row.excelRow), [[finalRemark]]);
  if (lastPaidDateCol >= 0) await writeExcelRange(INSURANCE_SHEET, cellAddress(lastPaidDateCol, row.excelRow), [[displayDate]]);
  if (lastPaidAmountCol >= 0) await writeExcelRange(INSURANCE_SHEET, cellAddress(lastPaidAmountCol, row.excelRow), [[amount]]);
  if (lastPaidAccountCol >= 0) await writeExcelRange(INSURANCE_SHEET, cellAddress(lastPaidAccountCol, row.excelRow), [[account || ""]]);
  if (transactionRowCol >= 0) await writeExcelRange(INSURANCE_SHEET, cellAddress(transactionRowCol, row.excelRow), [[transactionRow || ""]]);
}

function inferInsurer(policyName) {
  const value = cleanInsuranceText(policyName).toLowerCase();
  if (value.includes("singlife")) return "Singlife";
  if (value.includes("manulife")) return "Manulife";
  if (value.includes("income")) return "Income";
  return "";
}

function inferPolicyType(policyName) {
  const value = cleanInsuranceText(policyName).toLowerCase();
  if (value.includes("shield") || value.includes("health")) return "Hospitalisation";
  if (value.includes("cancer")) return "Cancer";
  if (value.includes("critical")) return "Critical Illness";
  if (value.includes("term")) return "Term Life";
  if (value.includes("careshield")) return "CareShield";
  if (value.includes("pa ")) return "Accident";
  return "Protection";
}

function computeNextDueDate(startDate, renewalMonth) {
  const today = startOfDay(new Date());
  if (!startDate && !renewalMonth) return null;

  const monthNumber = monthNameToNumber(renewalMonth);
  const baseMonth = monthNumber ? monthNumber - 1 : (startDate ? startDate.getMonth() : today.getMonth());
  const baseDay = startDate ? startDate.getDate() : 1;
  const firstYear = startDate ? startDate.getFullYear() + 1 : today.getFullYear();
  let due = new Date(firstYear, baseMonth, Math.min(baseDay, 28));
  while (due < today || (startDate && due < startDate)) {
    due = new Date(due.getFullYear() + 1, baseMonth, Math.min(baseDay, 28));
  }
  return due;
}

function defaultInsuranceGoalName(row) {
  const year = row?.dueYear || INSURANCE_CURRENT_YEAR + 1;
  return `Insurance ${year}`;
}

function getInsurancePolicyCurrentYearAmount(row) {
  return row.premiumThisYear || row.annualPremium || row.renewalPremium || 0;
}

function getDefaultInsuranceGroupGoalName(rows) {
  const goalNames = [...new Set(rows.map(row => cleanInsuranceText(row.goalName)).filter(Boolean))];
  if (goalNames.length === 1) return goalNames[0];
  const latestDue = rows
    .map(row => row.dueDate)
    .filter(Boolean)
    .sort((a, b) => b - a)[0];
  return latestDue ? suggestedInsuranceGoalName(latestDue) : `Insurance ${INSURANCE_CURRENT_YEAR}`;
}

function toggleInsurancePaidThisYearFields() {
  const checkbox = document.getElementById("insurancePaidThisYearInput");
  const fields = document.getElementById("insurancePaidThisYearFields");
  if (fields) fields.style.display = checkbox && checkbox.checked ? "grid" : "none";
}

function renderInsuranceGoalOptions(selectedValue = "") {
  const datalist = document.getElementById("insuranceGoalOptions");
  const select = document.getElementById("insuranceGoalExistingSelect");

  mergeVisibleGoalsIntoInsuranceGoals();

  const plan = buildInsuranceGoalPlanFromForm();
  const selectedName = cleanInsuranceText(selectedValue || getSelectedInsuranceGoalName());
  const existingNames = insuranceGoals
    .map(goal => goal.name)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  const values = new Set();
  existingNames.forEach(name => values.add(name));

  if (plan.goalName) values.add(plan.goalName);
  if (selectedName) values.add(selectedName);

  if (datalist) {
    datalist.innerHTML = "";
    values.forEach(value => {
      const option = document.createElement("option");
      option.value = value;
      datalist.appendChild(option);
    });
  }

  if (!select) return;

  const existingMatch = existingNames.find(name => normaliseInsuranceGoalName(name) === normaliseInsuranceGoalName(selectedName));
  const suggestedExists = plan.goalName && existingNames.some(name => normaliseInsuranceGoalName(name) === normaliseInsuranceGoalName(plan.goalName));
  select.innerHTML = "";
  appendInsuranceGoalSelectOption(select, "", "No linked goal");

  if (plan.goalName && !suggestedExists) {
    appendInsuranceGoalSelectOption(select, plan.goalName, "Suggested: " + plan.goalName);
  }

  if (existingNames.length) {
    const group = document.createElement("optgroup");
    group.label = "Goals page";
    existingNames.forEach(name => {
      appendInsuranceGoalSelectOption(group, name, name);
    });
    select.appendChild(group);
  }

  if (selectedName && !existingMatch && selectedName !== plan.goalName) {
    appendInsuranceGoalSelectOption(select, selectedName, "New: " + selectedName);
  }

  appendInsuranceGoalSelectOption(select, "__new__", "+ Create a different goal");

  select.value = selectedName && Array.from(select.options).some(option => option.value === selectedName)
    ? selectedName
    : "";
  refreshInsuranceGoalActionState();
}

function appendInsuranceGoalSelectOption(select, value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

async function refreshInsuranceGoalsForPicker(selectedValue = "") {
  const hint = document.getElementById("insuranceGoalHint");
  mergeVisibleGoalsIntoInsuranceGoals();
  const startingGoalCount = countInsuranceGoalNames();

  if (!startingGoalCount) {
    renderInsuranceGoalLoadingOption();
    if (hint) hint.textContent = "Loading goals from Goals page...";
  }

  try {
    await refreshInsuranceGoalsFromWorkbook();
    mergeVisibleGoalsIntoInsuranceGoals();
    const currentSelection = getSelectedInsuranceGoalName() || selectedValue;
    renderInsuranceGoalOptions(currentSelection);
    if (!countInsuranceGoalNames() && hint) {
      hint.textContent = "No saved goals found yet. Use the suggested goal or create a different one.";
    }
  } catch (err) {
    log("Goal picker refresh skipped: " + err.message);
    renderInsuranceGoalOptions(getSelectedInsuranceGoalName() || selectedValue);
    if (!startingGoalCount && hint) {
      hint.textContent = "Could not refresh saved goals. Use the suggested goal or create a different one.";
    }
  }
}

function renderInsuranceGoalLoadingOption() {
  const select = document.getElementById("insuranceGoalExistingSelect");
  if (!select) return;
  select.innerHTML = "";
  const option = document.createElement("option");
  option.value = "";
  option.textContent = "Loading goals from Goals page...";
  option.disabled = true;
  option.selected = true;
  select.appendChild(option);
}

async function refreshInsuranceGoalsFromWorkbook() {
  const arrayBuffer = await downloadExcelFile();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const budgetSheet = workbook.Sheets["Budget Setup"];
  if (!budgetSheet) throw new Error("Sheet not found: Budget Setup");
  insuranceGoals = readInsuranceGoals(budgetSheet);
  log(`Goals available for insurance: ${countInsuranceGoalNames()}.`);
}

function mergeVisibleGoalsIntoInsuranceGoals() {
  if (!document.querySelectorAll) return;
  const cards = Array.from(document.querySelectorAll(".goal-card[data-goal-name]"));
  if (!cards.length) return;

  cards.forEach((card, index) => {
    const name = cleanInsuranceText(card.dataset.goalName);
    if (!name) return;
    const existing = findInsuranceGoalByName(name);
    const goal = {
      excelRow: existing?.excelRow || null,
      name,
      target: toInsuranceNumber(card.dataset.goalTarget),
      manualSaved: toInsuranceNumber(card.dataset.goalSaved),
      monthlyAlloc: toInsuranceNumber(card.dataset.goalAlloc),
      startDate: insuranceGoalDateInputValue(card.dataset.goalStart),
      endDate: insuranceGoalDateInputValue(card.dataset.goalEnd),
      urgency: cleanInsuranceText(card.dataset.goalUrgency) || "Medium",
      notes: cleanInsuranceText(card.dataset.goalNotes),
      color: cleanInsuranceText(card.dataset.goalColor) || INSURANCE_GOAL_COLORS[index % INSURANCE_GOAL_COLORS.length],
      goalBuffer: toInsuranceNumber(card.dataset.goalBuffer),
      priority: toInsuranceNumber(card.dataset.goalPriority)
    };

    if (existing) Object.assign(existing, goal);
    else insuranceGoals.push(goal);
  });
}

function countInsuranceGoalNames() {
  return insuranceGoals.filter(goal => cleanInsuranceText(goal.name)).length;
}

function getSelectedInsuranceGoalName() {
  return cleanInsuranceText(document.getElementById("insuranceGoalNameInput")?.value);
}

function setSelectedInsuranceGoalName(goalName) {
  const input = document.getElementById("insuranceGoalNameInput");
  if (input) input.value = cleanInsuranceText(goalName);
}

function handleInsuranceGoalSelectChange() {
  const select = document.getElementById("insuranceGoalExistingSelect");
  const newGoalRow = document.getElementById("insuranceNewGoalRow");
  const newGoalInput = document.getElementById("insuranceNewGoalNameInput");
  const selectedValue = cleanInsuranceText(select?.value);

  if (selectedValue === "__new__") {
    setSelectedInsuranceGoalName("");
    if (newGoalRow) newGoalRow.style.display = "grid";
    if (newGoalInput) {
      newGoalInput.value = "";
      newGoalInput.focus();
    }
    refreshInsuranceGoalActionState();
    return;
  }

  setSelectedInsuranceGoalName(selectedValue);
  if (newGoalInput) newGoalInput.value = selectedValue && !findInsuranceGoalByName(selectedValue) ? selectedValue : "";

  const existing = findInsuranceGoalByName(selectedValue);
  if (existing) applyInsuranceGoalToForm(existing);
  else updateInsuranceSmartGoalDates();

  renderInsuranceGoalOptions(selectedValue);
}

function selectInsuranceGoalFromDropdown() {
  handleInsuranceGoalSelectChange();
}

function handleInsuranceNewGoalNameInput() {
  const goalName = cleanInsuranceText(document.getElementById("insuranceNewGoalNameInput")?.value);
  setSelectedInsuranceGoalName(goalName);

  const existing = findInsuranceGoalByName(goalName);
  if (existing) {
    applyInsuranceGoalToForm(existing);
    renderInsuranceGoalOptions(existing.name);
    return;
  }

  updateInsuranceSmartGoalDates();
  if (goalName) renderInsuranceGoalOptions(goalName);
  refreshInsuranceGoalActionState();
}

function handleInsurancePolicyNameInput() {
  const policyName = cleanInsuranceText(document.getElementById("insurancePolicyNameInput")?.value);
  const insurerInput = document.getElementById("insuranceInsurerInput");
  const typeInput = document.getElementById("insurancePolicyTypeInput");

  if (insurerInput && (!insurerInput.value || insurerInput.value === insurerInput.dataset.smartValue)) {
    const insurer = inferInsurer(policyName);
    insurerInput.value = insurer;
    insurerInput.dataset.smartValue = insurer;
  }

  if (typeInput && (!typeInput.value || typeInput.value === typeInput.dataset.smartValue)) {
    const type = inferPolicyType(policyName);
    typeInput.value = type;
    typeInput.dataset.smartValue = type;
  }
}

function clearInsuranceSmartState() {
  [
    "insuranceInsurerInput",
    "insurancePolicyTypeInput",
    "insuranceCurrentPremiumInput",
    "insuranceGoalNameInput",
    "insuranceGoalStartInput",
    "insuranceGoalEndInput"
  ].forEach(id => {
    const element = document.getElementById(id);
    if (element) delete element.dataset.smartValue;
  });
}

function updateInsuranceSmartGoalFields(options = {}) {
  const annualPremium = toInsuranceNumber(document.getElementById("insuranceAnnualPremiumInput")?.value);
  const renewalInput = document.getElementById("insuranceCurrentPremiumInput");
  if (renewalInput && annualPremium > 0 && (!renewalInput.value || renewalInput.value === renewalInput.dataset.smartValue)) {
    const smartPremium = annualPremium.toFixed(2);
    renewalInput.value = smartPremium;
    renewalInput.dataset.smartValue = smartPremium;
  }

  const plan = buildInsuranceGoalPlanFromForm();
  const currentLabel = document.getElementById("insuranceCurrentPremiumLabel");
  if (currentLabel) currentLabel.textContent = `Premium ${INSURANCE_CURRENT_YEAR}`;

  const goalInput = document.getElementById("insuranceGoalNameInput");
  if (goalInput && plan.goalName && isInsuranceCashPremium(document.getElementById("insurancePremiumTypeInput")?.value)) {
    const canSetGoal = !options.preserveGoalName && (options.forceGoalName || !goalInput.value || goalInput.value === goalInput.dataset.smartValue);
    if (canSetGoal) {
      goalInput.value = plan.goalName;
      goalInput.dataset.smartValue = plan.goalName;
    }
  }

  const goalName = getSelectedInsuranceGoalName();
  const existing = findInsuranceGoalByName(goalName);
  renderInsuranceGoalOptions(goalName || plan.goalName);
  if (existing) applyInsuranceGoalToForm(existing);
  else updateInsuranceSmartGoalDates(plan, options.forceGoalName);
  refreshInsuranceGoalActionState();
}

function applyInsuranceGoalToForm(goal) {
  if (!goal) return;
  if (goal.startDate) setInsuranceSmartValue("insuranceGoalStartInput", goal.startDate, true);
  if (goal.endDate) setInsuranceSmartValue("insuranceGoalEndInput", goal.endDate, true);
}

function updateInsuranceSmartGoalDates(plan = buildInsuranceGoalPlanFromForm(), force = false) {
  setInsuranceSmartValue("insuranceGoalStartInput", plan.startDateInput, force);
  setInsuranceSmartValue("insuranceGoalEndInput", plan.endDateInput, force);
}

function refreshInsuranceGoalActionState(existingGoal = null) {
  const button = document.getElementById("insuranceCreateGoalBtn");
  const hint = document.getElementById("insuranceGoalHint");
  const select = document.getElementById("insuranceGoalExistingSelect");
  const newGoalRow = document.getElementById("insuranceNewGoalRow");
  const newGoalInput = document.getElementById("insuranceNewGoalNameInput");
  const creatingNew = select?.value === "__new__";
  const goalName = getSelectedInsuranceGoalName();
  const existing = existingGoal || findInsuranceGoalByName(goalName);

  if (!goalName) {
    if (newGoalRow) newGoalRow.style.display = creatingNew ? "grid" : "none";
    if (button) button.style.display = "none";
    if (hint) hint.textContent = creatingNew ? "Type the new goal name, then create it or save the policy." : "";
    return;
  }

  if (existing) {
    if (newGoalRow) newGoalRow.style.display = "none";
    if (newGoalInput) newGoalInput.value = "";
    if (button) button.style.display = "none";
    if (hint) hint.textContent = existing.endDate
      ? `Using existing goal ending ${formatInsuranceInputDateLabel(existing.endDate)}.`
      : "Using existing goal.";
    return;
  }

  if (newGoalRow) newGoalRow.style.display = "grid";
  if (newGoalInput && !newGoalInput.value) newGoalInput.value = goalName;
  if (button) button.style.display = "inline-flex";
  if (hint) hint.textContent = "No matching goal yet. Save will create it, or create it now.";
}

async function createInsuranceGoalFromEntry() {
  const goalName = cleanInsuranceText(document.getElementById("insuranceGoalNameInput")?.value);
  if (!goalName) { alert("Enter a linked goal name first."); return; }

  const existing = findInsuranceGoalByName(goalName);
  if (existing) {
    applyInsuranceGoalToForm(existing);
    renderInsuranceGoalOptions(existing.name);
    refreshInsuranceGoalActionState(existing);
    alert(`Goal already exists: ${existing.name}`);
    return;
  }

  const owner = cleanInsuranceText(document.getElementById("insuranceOwnerInput")?.value);
  const policyName = cleanInsuranceText(document.getElementById("insurancePolicyNameInput")?.value);
  const annualPremium = toInsuranceNumber(document.getElementById("insuranceAnnualPremiumInput")?.value);
  const renewalPremium = toInsuranceNumber(document.getElementById("insuranceCurrentPremiumInput")?.value);
  const plan = buildInsuranceGoalPlanFromForm({ owner, policyName, annualPremium, renewalPremium: renewalPremium || annualPremium });
  plan.startDateInput = document.getElementById("insuranceGoalStartInput")?.value || plan.startDateInput || toInsuranceDateInput(startOfDay(new Date()));
  plan.endDateInput = document.getElementById("insuranceGoalEndInput")?.value || plan.endDateInput;
  plan.monthlyReserve = plan.target > 0 ? plan.target / monthsBetweenInsuranceDates(plan.startDateInput, plan.endDateInput) : 0;

  if (plan.target <= 0) { alert("Enter the premium amount before creating the linked goal."); return; }
  if (!plan.endDateInput) { alert("Enter the goal renewal date before creating the linked goal."); return; }

  const targetRow = findFirstEmptyInsuranceGoalRow();
  if (!targetRow) { alert("No empty goal slot was available in Budget Setup."); return; }

  try {
    const autoRecurring = document.getElementById("insuranceAutoRecurringInput")?.checked !== false;
    const values = buildInsuranceGoalValues({
      goalName,
      owner,
      policyName,
      plan,
      autoRecurring
    });
    await writeInsuranceGoal(targetRow, values);
    upsertInsuranceGoalFromValues(targetRow, values);
    renderInsuranceGoalOptions(goalName);
    selectInsuranceGoalFromDropdown();
    alert(`Goal created: ${goalName}`);
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to create linked goal: " + err.message);
  }
}

function setInsuranceSmartValue(elementId, value, force = false) {
  const element = document.getElementById(elementId);
  if (!element) return;
  const nextValue = cleanInsuranceText(value);
  if (force || !element.value || element.value === element.dataset.smartValue) {
    element.value = nextValue;
  }
  element.dataset.smartValue = nextValue;
}

function buildInsuranceGoalPlanFromRow(row) {
  if (!row) {
    return buildInsuranceGoalPlan({});
  }

  return buildInsuranceGoalPlan({
    owner: row.owner,
    policyName: row.policyName,
    premiumType: row.premiumType,
    amount: row.renewalPremium || row.annualPremium,
    dueDate: row.dueDate
  });
}

function buildInsuranceGoalPlanFromForm(fallback = {}) {
  const coverStartValue = document.getElementById("insuranceCoverStartInput")?.value;
  const coverStartDate = parseInsuranceDate(coverStartValue);
  const renewalMonth = document.getElementById("insuranceRenewalMonthInput")?.value || "";
  const annualPremium = fallback.annualPremium ?? toInsuranceNumber(document.getElementById("insuranceAnnualPremiumInput")?.value);
  const renewalPremium = fallback.renewalPremium ?? toInsuranceNumber(document.getElementById("insuranceCurrentPremiumInput")?.value);
  const amount = renewalPremium || annualPremium;

  return buildInsuranceGoalPlan({
    owner: fallback.owner ?? cleanInsuranceText(document.getElementById("insuranceOwnerInput")?.value),
    policyName: fallback.policyName ?? cleanInsuranceText(document.getElementById("insurancePolicyNameInput")?.value),
    premiumType: document.getElementById("insurancePremiumTypeInput")?.value || "Cash",
    amount,
    dueDate: computeNextDueDate(coverStartDate, renewalMonth)
  });
}

function buildInsuranceGoalPlan({ owner = "", policyName = "", premiumType = "Cash", amount = 0, dueDate = null }) {
  const target = Number(amount) > 0 ? Number(amount) : 0;
  const startDate = dueDate ? suggestInsuranceGoalStartDate(dueDate) : null;
  const startDateInput = startDate ? toInsuranceDateInput(startDate) : "";
  const endDateInput = dueDate ? toInsuranceDateInput(dueDate) : "";
  const monthlyReserve = target > 0 ? target / monthsBetweenInsuranceDates(startDateInput, endDateInput) : 0;

  return {
    owner,
    policyName,
    premiumType,
    dueDate,
    goalName: dueDate ? suggestedInsuranceGoalName(dueDate) : "",
    target,
    startDateInput,
    endDateInput,
    monthlyReserve
  };
}

function suggestedInsuranceGoalName(dueDate) {
  return "Insurance " + dueDate.toLocaleString("en-SG", { month: "short", year: "numeric" });
}

function suggestInsuranceGoalStartDate(dueDate) {
  const today = startOfDay(new Date());
  const firstOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  if (!dueDate || dueDate <= today) return today;
  return firstOfThisMonth;
}

function monthsBetweenInsuranceDates(startInput, endInput) {
  const start = parseInsuranceDate(startInput) || startOfDay(new Date());
  const end = parseInsuranceDate(endInput) || start;
  const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  return Math.max(1, months);
}

async function syncInsuranceGoalFromPolicy({ goalName, owner, policyName, autoRecurring, plan, promptOnMissing = true }) {
  if (!goalName) return "";

  const finalPlan = { ...plan };

  const existing = findInsuranceGoalByName(goalName);
  if (!existing) {
    const shouldCreate = !promptOnMissing || confirm(
      `No matching goal "${goalName}" was found. Create it now?\n\n` +
      `Premium target: ${formatInsuranceCurrency(finalPlan.target)}\n` +
      `Start: ${formatInsuranceInputDateLabel(finalPlan.startDateInput)}\n` +
      `Renewal: ${formatInsuranceInputDateLabel(finalPlan.endDateInput)}\n` +
      `Type: ${autoRecurring ? "Recurring" : "Non-recurring"}`
    );
    if (!shouldCreate) return "";

    const targetRow = findFirstEmptyInsuranceGoalRow();
    if (!targetRow) return "\nNo empty goal slot was available.";
    const values = buildInsuranceGoalValues({
      goalName,
      owner,
      policyName,
      plan: finalPlan,
      autoRecurring
    });
    await writeInsuranceGoal(targetRow, values);
    upsertInsuranceGoalFromValues(targetRow, values);
    return `\nGoal created: ${goalName}.`;
  }

  if (autoRecurring && insuranceGoalNeedsUpdate(existing, finalPlan)) {
    const shouldUpdate = confirm(
      `Update existing goal "${goalName}" with the insurance renewal plan?\n\n` +
      `Premium target: ${formatInsuranceCurrency(finalPlan.target)}\n` +
      `Start: ${formatInsuranceInputDateLabel(finalPlan.startDateInput)}\n` +
      `Renewal: ${formatInsuranceInputDateLabel(finalPlan.endDateInput)}\n` +
      `Type: Recurring`
    );
    if (shouldUpdate) {
      const values = buildInsuranceGoalValues({
        goalName,
        owner,
        policyName,
        plan: finalPlan,
        existing,
        autoRecurring
      });
      await writeInsuranceGoal(existing.excelRow, values);
      upsertInsuranceGoalFromValues(existing.excelRow, values);
      return `\nGoal updated: ${goalName}.`;
    }
  }

  return "";
}

function buildInsuranceGoalValues({ goalName, owner, policyName, plan, existing = null, autoRecurring = true }) {
  return [
    goalName,
    plan.target,
    existing ? existing.manualSaved : 0,
    autoRecurring ? plan.monthlyReserve : (existing?.monthlyAlloc || 0),
    plan.startDateInput,
    plan.endDateInput,
    existing?.urgency || suggestedInsuranceGoalUrgency(plan.endDateInput),
    existing?.notes || `Insurance renewal - ${owner} - ${policyName}`,
    existing?.color || getNextInsuranceGoalColor(),
    existing?.goalBuffer || 0,
    existing?.priority || 0
  ];
}

async function writeInsuranceGoal(excelRow, values) {
  await writeBudgetSetupRange(`S${excelRow}:AC${excelRow}`, [values]);
}

function upsertInsuranceGoalFromValues(excelRow, values) {
  const goal = {
    excelRow,
    name: cleanInsuranceText(values[0]),
    target: toInsuranceNumber(values[1]),
    manualSaved: toInsuranceNumber(values[2]),
    monthlyAlloc: toInsuranceNumber(values[3]),
    startDate: insuranceGoalDateInputValue(values[4]),
    endDate: insuranceGoalDateInputValue(values[5]),
    urgency: cleanInsuranceText(values[6]) || "Medium",
    notes: cleanInsuranceText(values[7]),
    color: cleanInsuranceText(values[8]),
    goalBuffer: toInsuranceNumber(values[9]),
    priority: toInsuranceNumber(values[10])
  };
  const index = insuranceGoals.findIndex(item => item.excelRow === excelRow);
  if (index >= 0) insuranceGoals[index] = goal;
  else insuranceGoals.push(goal);
}

function findInsuranceGoalByName(goalName) {
  const target = normaliseInsuranceGoalName(goalName);
  return insuranceGoals.find(goal => goal.name && normaliseInsuranceGoalName(goal.name) === target) || null;
}

function findFirstEmptyInsuranceGoalRow() {
  const empty = insuranceGoals.find(goal => !goal.name);
  return empty ? empty.excelRow : null;
}

function insuranceGoalNeedsUpdate(goal, plan) {
  if (plan.target > 0 && Math.abs((goal.target || 0) - plan.target) > 0.01) return true;
  if (plan.monthlyReserve > 0 && Math.abs((goal.monthlyAlloc || 0) - plan.monthlyReserve) > 0.01) return true;
  if (plan.startDateInput && goal.startDate !== plan.startDateInput) return true;
  if (plan.endDateInput && goal.endDate !== plan.endDateInput) return true;
  return false;
}

function suggestedInsuranceGoalUrgency(endDateInput) {
  const end = parseInsuranceDate(endDateInput);
  if (!end) return "Medium";
  const months = monthsBetweenInsuranceDates(toInsuranceDateInput(startOfDay(new Date())), endDateInput);
  if (months <= 2) return "High";
  if (months >= 9) return "Low";
  return "Medium";
}

function getNextInsuranceGoalColor() {
  const used = new Set(insuranceGoals
    .filter(goal => goal.name)
    .map(goal => cleanInsuranceText(goal.color).toLowerCase())
    .filter(Boolean));
  return INSURANCE_GOAL_COLORS.find(color => !used.has(color.toLowerCase())) ||
    INSURANCE_GOAL_COLORS[insuranceGoals.filter(goal => goal.name).length % INSURANCE_GOAL_COLORS.length];
}

function isInsuranceCashPremium(value) {
  return cleanInsuranceText(value).toLowerCase().includes("cash");
}

function normaliseInsuranceGoalName(value) {
  return cleanInsuranceText(value).toLowerCase().replace(/\s+/g, " ");
}

function formatInsuranceInputDateLabel(value) {
  const date = parseInsuranceDate(value);
  return date ? formatInsuranceDate(date) : (cleanInsuranceText(value) || "-");
}

function getInsuranceYearAmount(row, headerMap, prefix, year, fallback = 0) {
  const value = toInsuranceNumber(getInsuranceCell(row, headerMap, `${prefix} ${year}`));
  return value > 0 ? value : fallback;
}

function monthSelectValue(value) {
  const monthNumber = monthNameToNumber(value);
  if (!monthNumber) return "";
  return ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][monthNumber - 1];
}

function monthNameToNumber(value) {
  const text = cleanInsuranceText(value);
  if (!text) return null;
  const numeric = Number(text);
  if (Number.isInteger(numeric) && numeric >= 1 && numeric <= 12) return numeric;
  const names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const index = names.findIndex(name => text.toLowerCase().startsWith(name));
  return index >= 0 ? index + 1 : null;
}

function parseInsuranceDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    return new Date(date.y, date.m - 1, date.d);
  }
  const text = cleanInsuranceText(value);
  const ddMmm = text.match(/^(\d{1,2})[-\s/]([A-Za-z]{3,})[-\s/](\d{2,4})$/);
  if (ddMmm) {
    const year = Number(ddMmm[3].length === 2 ? "20" + ddMmm[3] : ddMmm[3]);
    const month = monthNameToNumber(ddMmm[2]);
    if (month) return new Date(year, month - 1, Number(ddMmm[1]));
  }
  const ddmm = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return new Date(Number(ddmm[3]), Number(ddmm[2]) - 1, Number(ddmm[1]));
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  return null;
}

function formatInsuranceDateForExcel(dateInput) {
  const [year, month, day] = dateInput.split("-").map(Number);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return dateInput;
  return String(day).padStart(2, "0") + "/" + String(month).padStart(2, "0") + "/" + year;
}

function formatInsuranceDateForExcelOrBlank(dateInput) {
  return dateInput ? formatInsuranceDateForExcel(dateInput) : "";
}

function toInsuranceDateInput(date) {
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}

function insuranceGoalDateInputValue(value) {
  const date = parseInsuranceDate(value);
  return date ? toInsuranceDateInput(date) : "";
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function getInsuranceCell(row, headerMap, headerName) {
  const index = getHeaderIndex(headerMap, headerName);
  return index >= 0 ? cleanInsuranceText(row[index]) : "";
}

function getHeaderIndex(headerMap, headerName) {
  const direct = headerMap[normaliseHeader(headerName)];
  if (Number.isInteger(direct)) return direct;

  const yearly = cleanInsuranceText(headerName).match(/^(Premium|Paid|Remark)\s+(\d{4})$/i);
  if (yearly) {
    const shortYear = yearly[2].slice(-2);
    const shortKey = headerMap[normaliseHeader(yearly[1] + shortYear)];
    if (Number.isInteger(shortKey)) return shortKey;
  }

  return Number.isInteger(direct) ? direct : -1;
}

function normaliseHeader(value) {
  return cleanInsuranceText(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cleanInsuranceText(value) {
  return String(value ?? "").trim();
}

function toInsuranceNumber(value) {
  const number = Number(String(value ?? "").replace(/[$,\s]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

function isInsuranceTruthy(value) {
  return ["yes", "true", "1", "y"].includes(cleanInsuranceText(value).toLowerCase());
}

function getWorksheetCellValue(sheet, zeroBasedCol, oneBasedRow) {
  const cell = sheet[XLSX.utils.encode_cell({ c: zeroBasedCol, r: oneBasedRow - 1 })];
  return cell ? cell.v : "";
}

function setInsuranceMoney(elementId, value) {
  const element = document.getElementById(elementId);
  if (element) element.textContent = formatInsuranceCurrency(value);
}

function formatInsuranceCurrency(value) {
  return Number(value || 0).toLocaleString("en-SG", {
    style: "currency",
    currency: "SGD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatInsuranceDate(date) {
  if (!date) return "-";
  return date.toLocaleDateString("en-SG", { day: "2-digit", month: "short", year: "numeric" });
}

function columnLetter(oneBasedIndex) {
  return XLSX.utils.encode_col(oneBasedIndex - 1);
}

function cellAddress(zeroBasedCol, oneBasedRow) {
  const cell = columnLetter(zeroBasedCol + 1) + oneBasedRow;
  return `${cell}:${cell}`;
}

function escapeInsuranceHtml(value) {
  return cleanInsuranceText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
