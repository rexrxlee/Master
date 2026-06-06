// ── State ─────────────────────────────────────────────────────────────────────
let categories = {};
let accounts = [];
let incomeSubCategories = [];
let currentMode = "transaction";
let isClaimable = false;
let recentRows = [];               // raw rows from sheet (last 50)
let filteredRecentRows = [];       // after applying search/category filter

// ── Page load ─────────────────────────────────────────────────────────────────
async function loadAddTransactionPage() {
  try {
    clearOutput();
    log("Loading data...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Load categories from Budget Setup
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (budgetSheet) {
      categories = readCategories(budgetSheet);
      incomeSubCategories = readIncomeSubCategories(budgetSheet);
    }

    // Load accounts
    if (budgetSheet) {
      const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, blankrows: false });
      accounts = allRows.slice(1)
        .map(row => ({ name: String(row[9] ?? "").trim(), type: String(row[10] ?? "Savings").trim() }))
        .filter(a => a.name !== "");
    }

    // Load recent transactions
    const txSheet = workbook.Sheets[CONFIG.sheetName];
    if (txSheet) {
      const allTxRows = XLSX.utils.sheet_to_json(txSheet, { header: 1, blankrows: false });
      // headers: A=Date B=Transaction C=Amount D=MainCategory E=SubCategory F=Account G=Claimable H=ClaimStatus I=ClaimAmount
      recentRows = allTxRows.slice(1)
        .map((row, i) => ({
          _rowIndex: i + 2, // 1-based Excel row (header is row 1)
          date:        row[0] ?? "",
          transaction: String(row[1] ?? ""),
          amount:      row[2] ?? "",
          mainCat:     String(row[3] ?? ""),
          subCat:      String(row[4] ?? ""),
          account:     String(row[5] ?? ""),
          claimable:   String(row[6] ?? "").trim(),   // "Yes" or ""
          claimStatus: String(row[7] ?? "").trim(),   // "Pending" | "Claimed" | ""
          claimAmount: row[8] ?? "",
        }))
        .filter(r => r.date !== "" && r.transaction !== "Opening Balance");
    }

    populateMainCategories();
    populateAccountDropdowns();
    populateIncomeSubCategories();
    setDefaultDates();
    renderRecentTransactions();
    populateTxFilterCategories();
    renderClaimsTracker();

    log("Ready.");
  } catch (err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

// ── Read helpers ──────────────────────────────────────────────────────────────
function readCategories(sheet) {
  const map = {};
  // Main cats in col C (index 2), sub cats in col D (index 3) — rows 2..13 (Bills) and rows 2..13 (Monthly)
  // Actually read from Budget Setup: cols A (sub), D (main) grouping
  // We'll parse col D = main cat, col E = sub cat (if that's your setup)
  // ── fallback: derive from transaction sheet header structure ──
  // Use cols A–B for Bills, F–G for Monthly; main categories are the section headers
  const billsRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "A2:A13", blankrows: false });
  const monthlyRows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "F2:F13", blankrows: false });

  map["Bills"] = billsRows.map(r => String(r[0] ?? "").trim()).filter(Boolean);
  map["Monthly Expenses"] = monthlyRows.map(r => String(r[0] ?? "").trim()).filter(Boolean);
  return map;
}

function readIncomeSubCategories(sheet) {
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => String(row[16] ?? "").trim())   // column Q = index 16
    .filter(v => v !== "");
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function setDefaultDates() {
  const today = new Date();
  const val = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");
  ["txDate","inDate","trDate","ccPayDate"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = val;
  });
}

function populateMainCategories() {
  const sel = document.getElementById("txMainCategory");
  if (!sel) return;
  sel.innerHTML = `<option value="">-- Select --</option>`;
  Object.keys(categories).forEach(mainCat => {
    const opt = document.createElement("option");
    opt.value = mainCat;
    opt.textContent = mainCat;
    sel.appendChild(opt);
  });
}

function onMainCategoryChange() {
  const mainCat = document.getElementById("txMainCategory").value;
  const subSel = document.getElementById("txSubCategory");
  subSel.innerHTML = `<option value="">-- Select --</option>`;
  (categories[mainCat] || []).forEach(sub => {
    const opt = document.createElement("option");
    opt.value = sub;
    opt.textContent = sub;
    subSel.appendChild(opt);
  });
}

function populateAccountDropdowns() {
  const savingsAccounts  = accounts.filter(a => a.type !== "Credit Card");
  const ccAccounts       = accounts.filter(a => a.type === "Credit Card");
  const allAccounts      = accounts;

  fillSelect("txAccount",        allAccounts);
  fillSelect("inAccount",        savingsAccounts);
  fillSelect("trFromAccount",    allAccounts);
  fillSelect("trToAccount",      allAccounts);
  fillSelect("ccPayFromAccount", savingsAccounts);
  fillSelect("ccPayToAccount",   ccAccounts);
}

function fillSelect(id, list) {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = "";
  list.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a.name;
    opt.textContent = a.name + (a.type === "Credit Card" ? " (CC)" : "");
    sel.appendChild(opt);
  });
}

function populateIncomeSubCategories() {
  const sel = document.getElementById("inSubCategory");
  if (!sel) return;
  sel.innerHTML = `<option value="">-- Select --</option>`;
  incomeSubCategories.forEach(sub => {
    const opt = document.createElement("option");
    opt.value = sub;
    opt.textContent = sub;
    sel.appendChild(opt);
  });
}

// ── Mode switching ─────────────────────────────────────────────────────────────
function setMode(mode) {
  currentMode = mode;
  ["transaction","income","transfer","ccpay"].forEach(m => {
    const formId = m === "transaction" ? "transactionForm"
                 : m === "income"      ? "incomeForm"
                 : m === "transfer"    ? "transferForm"
                 :                       "ccPayForm";
    const btn = document.getElementById("btn" + m.charAt(0).toUpperCase() + m.slice(1));
    const form = document.getElementById(formId);
    if (form) form.style.display = m === mode ? "block" : "none";
    if (btn)  btn.classList.toggle("mode-active", m === mode);
  });
}

// ── Claimable toggle (expense form) ──────────────────────────────────────────
function toggleClaimable() {
  isClaimable = !isClaimable;
  const btn = document.getElementById("claimableToggle");
  if (btn) {
    btn.textContent  = isClaimable ? "🔖 Claimable: ON" : "🔖 Mark as Claimable";
    btn.style.background = isClaimable ? "#e67e22" : "";
    btn.style.color      = isClaimable ? "white"   : "";
    btn.style.borderColor = isClaimable ? "#e67e22" : "";
  }
  const hint = document.getElementById("claimableHint");
  if (hint) hint.style.display = isClaimable ? "block" : "none";
  const details = document.getElementById("claimDetails");
  if (details) details.classList.toggle("open", isClaimable);
  if (isClaimable) syncClaimAmountCap();
}

function setClaimAmountPreset(mode) {
  const amount = parseFloat(document.getElementById("txAmount")?.value);
  const claimInput = document.getElementById("txClaimAmount");
  if (!claimInput || isNaN(amount) || amount <= 0) return;
  claimInput.value = (mode === "half" ? amount / 2 : amount).toFixed(2);
}

function syncClaimAmountCap() {
  const amount = parseFloat(document.getElementById("txAmount")?.value);
  const claimInput = document.getElementById("txClaimAmount");
  if (!claimInput) return;
  if (!isNaN(amount) && amount > 0) claimInput.max = String(amount);
  const claimAmount = parseFloat(claimInput.value);
  if (!isNaN(amount) && !isNaN(claimAmount) && claimAmount > amount) {
    claimInput.value = amount.toFixed(2);
  }
}

function resetClaimableUi() {
  isClaimable = false;
  const btn = document.getElementById("claimableToggle");
  if (btn) { btn.textContent = "🔖 Mark as Claimable"; btn.style.background = ""; btn.style.color = ""; btn.style.borderColor = ""; }
  const hint = document.getElementById("claimableHint");
  if (hint) hint.style.display = "none";
  const details = document.getElementById("claimDetails");
  if (details) details.classList.remove("open");
  const claimInput = document.getElementById("txClaimAmount");
  if (claimInput) claimInput.value = "";
}

// ── Date formatting helper ────────────────────────────────────────────────────
function formatDateForExcel(dateInputValue) {
  // Write as YYYY-MM-DD. Slash-based formats (DD/MM/YYYY) are ambiguous —
  // Excel / Graph API interprets slashes as MM/DD/YYYY (American), causing
  // e.g. 10 June to be stored as 6 October. ISO format is unambiguous.
  return dateInputValue; // input[type="date"] already yields YYYY-MM-DD
}

// ── Save transaction ──────────────────────────────────────────────────────────
async function saveTransaction() {
  const dateVal  = document.getElementById("txDate").value;
  const mainCat  = document.getElementById("txMainCategory").value;
  const subCat   = document.getElementById("txSubCategory").value;
  const txDesc   = document.getElementById("txTransaction").value.trim();
  const amount   = parseFloat(document.getElementById("txAmount").value);
  const account  = document.getElementById("txAccount").value;

  if (!dateVal || !mainCat || !subCat || !txDesc || isNaN(amount) || !account) {
    alert("Please fill in all fields."); return;
  }

  const dateStr     = formatDateForExcel(dateVal);
  const claimFlag   = isClaimable ? "Yes" : "";
  const claimStatus = isClaimable ? "Pending" : "";
  let claimAmount = "";

  if (isClaimable) {
    const rawClaimAmount = parseFloat(document.getElementById("txClaimAmount")?.value);
    claimAmount = isNaN(rawClaimAmount) ? amount : rawClaimAmount;
    if (claimAmount <= 0) { alert("Claim amount must be more than $0."); return; }
    if (claimAmount > amount) { alert("Claim amount cannot be more than the expense amount."); return; }
    claimAmount = Number(claimAmount.toFixed(2));
  }

  // Columns: A=Date, B=Transaction, C=Amount, D=MainCategory, E=SubCategory, F=Account, G=Claimable, H=ClaimStatus, I=ClaimAmount
  const row = [dateStr, txDesc, amount, mainCat, subCat, account, claimFlag, claimStatus, claimAmount];

  try {
    log("Saving transaction...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await ensureClaimAmountHeader();
    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:I${nextRow}`, [row]);

    resetClaimableUi();

    document.getElementById("txTransaction").value = "";
    document.getElementById("txAmount").value = "";
    document.getElementById("transactionSuccess").style.display = "block";
    setTimeout(() => document.getElementById("transactionSuccess").style.display = "none", 4000);

    log("Saved. Reloading recent transactions...");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save income ───────────────────────────────────────────────────────────────
async function saveIncome() {
  const dateVal = document.getElementById("inDate").value;
  const txDesc  = document.getElementById("inTransaction").value.trim();
  const subCat  = document.getElementById("inSubCategory").value;
  const amount  = parseFloat(document.getElementById("inAmount").value);
  const account = document.getElementById("inAccount").value;

  if (!dateVal || !txDesc || isNaN(amount) || !account) {
    alert("Please fill in all fields."); return;
  }

  const dateStr = formatDateForExcel(dateVal);
  // Income rows: MainCategory = "Income", Claimable = ""
  const row = [dateStr, txDesc, amount, "Income", subCat || "", account, "", ""];

  try {
    log("Saving income...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow}`, [row]);

    document.getElementById("inTransaction").value = "";
    document.getElementById("inAmount").value = "";
    document.getElementById("incomeSuccess").style.display = "block";
    setTimeout(() => document.getElementById("incomeSuccess").style.display = "none", 4000);

    log("Income saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save transfer ─────────────────────────────────────────────────────────────
async function saveTransfer() {
  const dateVal     = document.getElementById("trDate").value;
  const fromAccount = document.getElementById("trFromAccount").value;
  const toAccount   = document.getElementById("trToAccount").value;
  const amount      = parseFloat(document.getElementById("trAmount").value);
  const note        = document.getElementById("trNote").value.trim() || "Transfer";

  if (!dateVal || !fromAccount || !toAccount || isNaN(amount)) {
    alert("Please fill in all fields."); return;
  }
  if (fromAccount === toAccount) { alert("From and To accounts must be different."); return; }

  const dateStr = formatDateForExcel(dateVal);
  const rows = [
    [dateStr, note, amount, "Transfer", "Transfer Out", fromAccount, "", ""],
    [dateStr, note, amount, "Transfer", "Transfer In",  toAccount,   "", ""],
  ];

  try {
    log("Saving transfer...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow + 1}`, rows);

    document.getElementById("trAmount").value = "";
    document.getElementById("trNote").value = "";
    document.getElementById("transferSuccess").style.display = "block";
    setTimeout(() => document.getElementById("transferSuccess").style.display = "none", 4000);

    log("Transfer saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Save CC payment ───────────────────────────────────────────────────────────
async function saveCcPayment() {
  const dateVal     = document.getElementById("ccPayDate").value;
  const fromAccount = document.getElementById("ccPayFromAccount").value;
  const toAccount   = document.getElementById("ccPayToAccount").value;
  const amount      = parseFloat(document.getElementById("ccPayAmount").value);

  if (!dateVal || !fromAccount || !toAccount || isNaN(amount)) {
    alert("Please fill in all fields."); return;
  }

  const dateStr = formatDateForExcel(dateVal);
  const rows = [
    [dateStr, "CC Payment",  amount, "Transfer", "CC Payment Out", fromAccount, "", ""],
    [dateStr, "CC Payment",  amount, "Transfer", "CC Payment In",  toAccount,   "", ""],
  ];

  try {
    log("Saving CC payment...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();
    const url = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const data = await graphGetJson(url, token);
    const nextRow = data.rowCount + 1;

    await writeExcelRange(CONFIG.sheetName, `A${nextRow}:H${nextRow + 1}`, rows);

    document.getElementById("ccPayAmount").value = "";
    document.getElementById("ccPaySuccess").style.display = "block";
    setTimeout(() => document.getElementById("ccPaySuccess").style.display = "none", 4000);

    log("CC payment saved.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Recent transactions table ─────────────────────────────────────────────────
// Convert any date value (serial number or DD/MM/YYYY string) to YYYY-MM-DD for <input type="date">
function rawDateToInputValue(value) {
  if (!value && value !== 0) return "";
  // Excel serial number (numeric or stringified numeric)
  if (typeof value === "number" || (typeof value === "string" && /^\d{5}$/.test(value.trim()))) {
    try {
      const d = XLSX.SSF.parse_date_code(Number(value));
      return d.y + "-" + String(d.m).padStart(2,"0") + "-" + String(d.d).padStart(2,"0");
    } catch { return ""; }
  }
  const s = String(value).trim();
  // DD/MM/YYYY
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return ddmm[3] + "-" + ddmm[2].padStart(2,"0") + "-" + ddmm[1].padStart(2,"0");
  // YYYY-MM-DD already
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  return "";
}

function renderRecentTransactions() {
  applyTxFilter();
}

function populateTxFilterCategories() {
  const sel = document.getElementById("txFilterMainCat");
  if (!sel) return;
  const current = sel.value;
  sel.innerHTML = `<option value="">All Categories</option>`;
  const allCats = [...new Set(recentRows.map(r => r.mainCat).filter(Boolean))].sort();
  allCats.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    if (cat === current) opt.selected = true;
    sel.appendChild(opt);
  });
}

function applyTxFilter() {
  const text    = (document.getElementById("txFilterText")?.value ?? "").toLowerCase().trim();
  const mainCat = (document.getElementById("txFilterMainCat")?.value ?? "");

  const sorted = [...recentRows]
    .filter(r => rawDateToInputValue(r.date))
    .sort((a, b) => rawDateToInputValue(b.date).localeCompare(rawDateToInputValue(a.date)))
    .slice(0, 50);

  filteredRecentRows = sorted.filter(r => {
    const matchesCat  = !mainCat || r.mainCat === mainCat;
    const matchesText = !text ||
      r.transaction.toLowerCase().includes(text) ||
      r.subCat.toLowerCase().includes(text) ||
      r.account.toLowerCase().includes(text);
    return matchesCat && matchesText;
  });

  renderRecentTransactionsTable();
}

function clearTxFilter() {
  const txt = document.getElementById("txFilterText");
  const sel = document.getElementById("txFilterMainCat");
  if (txt) txt.value = "";
  if (sel) sel.value = "";
  applyTxFilter();
}

function renderRecentTransactionsTable() {
  const container = document.getElementById("txTableBody");
  if (!container) return;

  const badge = document.getElementById("recentTxBadge");
  if (badge) badge.textContent = filteredRecentRows.length + " shown";

  if (filteredRecentRows.length === 0) {
    container.innerHTML = '<div class="tx-empty">No transactions match the filter.</div>';
    return;
  }

  const allMainCats = [...new Set(Object.keys(categories))].sort();
  const extraCats   = ["Income", "Transfer", "Saving Goals"].filter(c => !allMainCats.includes(c));
  const allCats     = [...allMainCats, ...extraCats];
  const allAccounts = accounts.map(a => a.name);

  let html = `
    <div class="tx-row-header">
      <span>Date</span>
      <span>Description</span>
      <span>Amount</span>
      <span>Category</span>
      <span>Account</span>
      <span></span>
    </div>`;

  filteredRecentRows.forEach(row => {
    const idx     = row._rowIndex;
    const dateVal = rawDateToInputValue(row.date);
    const dispDate = formatDisplayDate(row.date);
    const amt     = parseFloat(String(row.amount).replace(/[$,]/g, "")) || 0;
    const amtStr  = amt.toLocaleString("en-SG", { style: "currency", currency: "SGD", minimumFractionDigits: 2 });
    const catLabel = [row.mainCat, row.subCat].filter(Boolean).join(" › ");
    const claimAmt = getClaimAmount(row);
    const isPartialClaim = row.claimable === "Yes" && claimAmt > 0 && claimAmt < amt;
    const claimBadge = row.claimable === "Yes"
      ? `<span style="font-size:10px;background:#fff3e0;color:#e67e22;border:1px solid #f0c080;border-radius:4px;padding:1px 5px;margin-left:4px;" title="Claim amount: ${formatAmount(claimAmt)}">${isPartialClaim ? "🔖 Partial" : "🔖"}</span>`
      : "";

    const catOptions   = allCats.map(c =>
      `<option value="${c}" ${c === row.mainCat ? "selected" : ""}>${c}</option>`).join("");
    const acctOptions  = allAccounts.map(a =>
      `<option value="${a}" ${a === row.account ? "selected" : ""}>${a}</option>`).join("");

    html += `
      <div class="tx-row" id="tx-row-${idx}" onclick="toggleEditPanel(${idx})">
        <span class="tx-row-date">${dispDate}</span>
        <span class="tx-row-desc">${escapeHtml(row.transaction)}${claimBadge}</span>
        <span class="tx-row-amount">${amtStr}</span>
        <span class="tx-row-cat">${escapeHtml(catLabel)}</span>
        <span class="tx-row-account">${escapeHtml(row.account)}</span>
        <span class="tx-row-actions"><span class="tx-row-chevron">▾</span></span>
      </div>
      <div class="tx-edit-panel" id="tx-edit-${idx}">
        <div class="tx-edit-grid">
          <div class="ef"><label>Date</label><input type="date" id="edit-date-${idx}" value="${dateVal}"></div>
          <div class="ef"><label>Description</label><input type="text" id="edit-desc-${idx}" value="${escapeHtml(row.transaction)}"></div>
          <div class="ef"><label>Amount</label><input type="number" step="0.01" id="edit-amt-${idx}" value="${amt || ""}"></div>
          <div class="ef"><label>Account</label><select id="edit-acc-${idx}">${acctOptions}</select></div>
          <div class="ef"><label>Main Category</label><select id="edit-main-${idx}">${catOptions}</select></div>
          <div class="ef"><label>Sub Category</label><input type="text" id="edit-sub-${idx}" value="${escapeHtml(row.subCat)}"></div>
          <div class="ef"><label>Claimable</label>
            <select id="edit-claim-${idx}">
              <option value="" ${row.claimable !== "Yes" ? "selected" : ""}>—</option>
              <option value="Yes" ${row.claimable === "Yes" ? "selected" : ""}>Yes</option>
            </select>
          </div>
          <div class="ef"><label>Claim Status</label>
            <select id="edit-claimst-${idx}">
              <option value="" ${!row.claimStatus ? "selected" : ""}>—</option>
              <option value="Pending" ${row.claimStatus === "Pending" ? "selected" : ""}>Pending</option>
              <option value="Claimed" ${row.claimStatus === "Claimed" ? "selected" : ""}>Claimed</option>
            </select>
          </div>
          <div class="ef"><label>Claim Amount</label><input type="number" step="0.01" min="0" id="edit-claimamt-${idx}" value="${row.claimable === "Yes" ? getClaimAmount(row) : ""}" placeholder="Full amount if blank"></div>
        </div>
        <div class="tx-edit-actions">
          <button class="btn-edit-save" onclick="saveRow(${idx}); event.stopPropagation();">💾 Save</button>
          <button class="btn-edit-delete" onclick="deleteRow(${idx}); event.stopPropagation();">🗑 Delete</button>
          <button class="btn-edit-cancel" onclick="toggleEditPanel(${idx}); event.stopPropagation();">Cancel</button>
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

let openEditPanel = null;

function toggleEditPanel(idx) {
  const panel = document.getElementById("tx-edit-" + idx);
  const row   = document.getElementById("tx-row-"  + idx);
  if (!panel) return;

  const isOpen = panel.classList.contains("open");

  // Close previously open panel
  if (openEditPanel !== null && openEditPanel !== idx) {
    const prev = document.getElementById("tx-edit-" + openEditPanel);
    const prevRow = document.getElementById("tx-row-" + openEditPanel);
    if (prev) prev.classList.remove("open");
    if (prevRow) prevRow.classList.remove("expanded");
  }

  panel.classList.toggle("open", !isOpen);
  row.classList.toggle("expanded", !isOpen);
  openEditPanel = isOpen ? null : idx;
}

// ── Save an edited row back to Excel ─────────────────────────────────────────
async function saveRow(excelRowNumber) {
  const dateInput = document.getElementById("edit-date-" + excelRowNumber).value;
  const desc      = document.getElementById("edit-desc-" + excelRowNumber).value.trim();
  const amt       = parseFloat(document.getElementById("edit-amt-" + excelRowNumber).value);
  const mainCat   = document.getElementById("edit-main-" + excelRowNumber).value;
  const subCat    = document.getElementById("edit-sub-" + excelRowNumber).value.trim();
  const account   = document.getElementById("edit-acc-" + excelRowNumber).value;
  const claimable = document.getElementById("edit-claim-" + excelRowNumber).value;
  const claimSt   = document.getElementById("edit-claimst-" + excelRowNumber).value;
  let claimAmount = "";

  if (!dateInput) { alert("Please enter a date."); return; }
  if (!desc)      { alert("Please enter a description."); return; }
  if (isNaN(amt)) { alert("Please enter a valid amount."); return; }
  if (claimable === "Yes") {
    const rawClaimAmount = parseFloat(document.getElementById("edit-claimamt-" + excelRowNumber).value);
    claimAmount = isNaN(rawClaimAmount) ? amt : rawClaimAmount;
    if (claimAmount <= 0) { alert("Claim amount must be more than $0."); return; }
    if (claimAmount > amt) { alert("Claim amount cannot be more than the expense amount."); return; }
    claimAmount = Number(claimAmount.toFixed(2));
  }

  try {
    log("Saving row " + excelRowNumber + "...");
    await ensureClaimAmountHeader();
    await writeExcelRange(
      CONFIG.sheetName,
      `A${excelRowNumber}:I${excelRowNumber}`,
      [[dateInput, desc, amt, mainCat, subCat, account, claimable, claimSt, claimAmount]]
    );
    log("Row " + excelRowNumber + " saved.");
    // Flash the row green briefly
    const tr = document.getElementById("tx-row-" + excelRowNumber);
    if (tr) { tr.style.background = "#f0fdf4"; setTimeout(() => { tr.style.background = ""; }, 1500); }
    // Reload so recentRows stay in sync
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed to save: " + err.message);
  }
}

// ── Claims tracker section ────────────────────────────────────────────────────
function renderClaimsTracker() {
  const container = document.getElementById("claimsContainer");
  if (!container) return;

  const claimRows = recentRows.filter(r => r.claimable === "Yes");

  const badge = document.getElementById("claimsBadge");
  if (badge) {
    const pending = claimRows.filter(r => r.claimStatus !== "Claimed").length;
    badge.textContent = pending > 0 ? pending + " pending" : "";
    badge.style.display = pending > 0 ? "" : "none";
  }

  if (claimRows.length === 0) {
    container.innerHTML = `<p style="color:#aaa;font-size:14px;">No claimable transactions yet.</p>`;
    return;
  }

  const pending = claimRows.filter(r => r.claimStatus !== "Claimed");
  const claimed = claimRows.filter(r => r.claimStatus === "Claimed");

  const pendingTotal = pending.reduce((s, r) => s + getClaimAmount(r), 0);

  let html = `
    <div style="background:#fff8f0;border:1px solid #f0c080;border-radius:10px;padding:16px;margin-bottom:16px;max-width:680px;">
      <strong style="color:#e67e22;">⏳ Pending Claims</strong>
      <span style="float:right;font-weight:bold;color:#e67e22;">${formatAmount(pendingTotal)} to recover</span>
    </div>`;

  if (pending.length > 0) {
    html += `<table class="data-table" style="max-width:680px;margin-bottom:24px;"><tr>
      <th>Date</th><th>Description</th><th>Claim</th><th>Expense</th><th>Category</th><th>Account</th><th></th></tr>`;
    pending.forEach(row => {
      html += `<tr>
        <td>${formatDisplayDate(row.date)}</td>
        <td>${escapeHtml(row.transaction)}</td>
        <td style="color:#e67e22;font-weight:bold;">${formatAmount(getClaimAmount(row))}</td>
        <td>${formatAmount(row.amount)}</td>
        <td>${escapeHtml(row.subCat)}</td>
        <td>${escapeHtml(row.account)}</td>
        <td><button onclick="markClaimed(${row._rowIndex})"
          style="background:#27ae60;color:white;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-size:12px;">
          ✓ Mark Claimed</button></td>
      </tr>`;
    });
    html += `</table>`;
  }

  if (claimed.length > 0) {
    html += `<details style="max-width:680px;"><summary style="cursor:pointer;color:#888;font-size:13px;margin-bottom:8px;">
      ✅ ${claimed.length} claimed transaction(s)</summary>
      <table class="data-table"><tr><th>Date</th><th>Description</th><th>Claim</th><th>Expense</th><th>Account</th></tr>`;
    claimed.forEach(row => {
      html += `<tr style="opacity:0.6;">
        <td>${formatDisplayDate(row.date)}</td>
        <td>${escapeHtml(row.transaction)}</td>
        <td>${formatAmount(getClaimAmount(row))}</td>
        <td>${formatAmount(row.amount)}</td>
        <td>${escapeHtml(row.account)}</td>
      </tr>`;
    });
    html += `</table></details>`;
  }

  container.innerHTML = html;
}

// ── Mark a row as Claimed (writes col H) ─────────────────────────────────────
async function markClaimed(excelRowNumber) {
  try {
    log("Marking row " + excelRowNumber + " as Claimed...");
    // Write "Claimed" to column H of that row
    await writeExcelRange(CONFIG.sheetName, `H${excelRowNumber}:H${excelRowNumber}`, [["Claimed"]]);
    log("Marked as Claimed.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Delete row ────────────────────────────────────────────────────────────────
async function deleteRow(excelRowNumber) {
  if (!confirm("Delete this transaction? This cannot be undone.")) return;
  try {
    log("Deleting row " + excelRowNumber + "...");
    // Overwrite with blank row
    await writeExcelRange(CONFIG.sheetName, `A${excelRowNumber}:I${excelRowNumber}`, [["","","","","","","","",""]]);
    log("Row cleared.");
    await loadAddTransactionPage();
  } catch (err) {
    log("ERROR: " + err.message);
    alert("Failed: " + err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function formatDisplayDate(value) {
  const iso = rawDateToInputValue(value);
  if (!iso) return String(value ?? "");
  const [y, m, d] = iso.split("-");
  return d + "/" + m + "/" + y;
}

function formatAmount(value) {
  const n = parseFloat(String(value).replace(/[$,]/g, ""));
  if (isNaN(n)) return "—";
  return n.toLocaleString("en-SG", { style: "currency", currency: "SGD", minimumFractionDigits: 2 });
}

function parseMoney(value) {
  const n = parseFloat(String(value ?? "").replace(/[$,]/g, ""));
  return isNaN(n) ? 0 : Math.abs(n);
}

function getClaimAmount(row) {
  if (!row || row.claimable !== "Yes") return 0;
  const expenseAmount = parseMoney(row.amount);
  const storedClaimAmount = parseMoney(row.claimAmount);
  const claimAmount = storedClaimAmount > 0 ? storedClaimAmount : expenseAmount;
  return Math.min(expenseAmount, claimAmount);
}

async function ensureClaimAmountHeader() {
  await writeExcelRange(CONFIG.sheetName, "I1:I1", [["Claim Amount"]]);
}

function escapeHtml(v) {
  return String(v ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
