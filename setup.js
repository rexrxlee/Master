let setupAccounts = [];

async function loadSetupPage() {
  try {
    clearOutput();
    log("Loading accounts...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (!budgetSheet) throw new Error("Sheet not found: Budget Setup");

    const allRows = XLSX.utils.sheet_to_json(budgetSheet, { header: 1, blankrows: false });
    setupAccounts = allRows.slice(1)
      .map(row => ({ name: String(row[9] ?? "").trim(), type: String(row[10] ?? "Savings").trim() }))
      .filter(a => a.name !== "");

    const txSheet = workbook.Sheets[CONFIG.sheetName];
    const existingEntries = txSheet ? detectExistingOpeningBalances(txSheet) : [];

    renderSetupForm(existingEntries);
    log("Ready.");
  } catch (err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

function detectExistingOpeningBalances(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  const headers = rows[0] ? rows[0].map(h => String(h ?? "").trim()) : [];
  const dateIdx    = headers.indexOf("Date");
  const subCatIdx  = headers.indexOf("Sub Category");
  const amountIdx  = headers.indexOf("Amount");
  const accountIdx = headers.indexOf("Account");
  if (subCatIdx === -1) return [];
  return rows.slice(1)
    .filter(row => String(row[subCatIdx] ?? "").trim() === "Opening Balance")
    .map(row => ({
      date:    row[dateIdx]    ?? "",
      account: row[accountIdx] ?? "",
      amount:  row[amountIdx]  ?? ""
    }));
}

function renderOpeningBalanceSetup(accounts, txSheet) {
  setupAccounts = (accounts || [])
    .map(account => ({
      name: String(account.name ?? "").trim(),
      type: String(account.type ?? "Savings").trim() || "Savings"
    }))
    .filter(account => account.name !== "");

  const existingEntries = txSheet ? detectExistingOpeningBalances(txSheet) : [];
  renderSetupForm(existingEntries);
}

function renderSetupForm(existingEntries = []) {
  const container = document.getElementById("setupForm");
  if (setupAccounts.length === 0) {
    container.innerHTML = `<p style="color:var(--red);">No accounts found. Add accounts above first.</p>`;
    return;
  }

  const today = new Date();
  const todayStr = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");

  // Warning banner if existing entries found
  let warningHtml = "";
  if (existingEntries.length > 0) {
    const rowsHtml = existingEntries.map(e =>
      `<tr><td>${escapeHtml(excelDateToDisplay(e.date))}</td><td>${escapeHtml(String(e.account))}</td><td style="font-family:'DM Mono',monospace;">$${Number(e.amount).toLocaleString("en-SG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td></tr>`
    ).join("");
    warningHtml = `
      <div class="existing-ob-warning">
        <div class="ob-warning-header"><span class="ob-warning-icon">⚠️</span><strong>Existing opening balances detected</strong></div>
        <p class="ob-warning-body">The entries below already exist. Saving will <strong>delete them first</strong> and write the new values — no duplicates.</p>
        <table class="ob-existing-table">
          <tr><th>Date</th><th>Account</th><th>Amount</th></tr>
          ${rowsHtml}
        </table>
        <p class="ob-warning-tip">✏️ Update the amounts and dates below, then click <strong>Replace Opening Balances</strong>.</p>
      </div>`;
  }

  // Per-account rows with individual date fields
  let tableHtml = `
    <table class="budget-table" style="max-width:640px;">
      <tr>
        <th>Account</th>
        <th>Type</th>
        <th>Date</th>
        <th>Opening Balance (SGD)</th>
      </tr>
  `;
  setupAccounts.forEach((account, idx) => {
    const existing = existingEntries.find(e =>
      String(e.account).trim().toLowerCase() === account.name.toLowerCase()
    );
    const prefillAmt  = existing ? Number(existing.amount).toFixed(2) : "";
    const prefillDate = existing ? excelDateToInputValue(existing.date) : todayStr;

    tableHtml += `
      <tr>
        <td style="font-weight:500;">${escapeHtml(account.name)}</td>
        <td style="color:var(--muted);font-size:13px;">${escapeHtml(account.type)}</td>
        <td>
          <input type="date" id="setupDate_${idx}" value="${prefillDate}"
            style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:13px;font-family:inherit;width:150px;">
        </td>
        <td>
          <input type="number" step="0.01" id="setupAmt_${idx}" placeholder="0.00" value="${prefillAmt}"
            style="width:150px;padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:13px;font-family:inherit;">
        </td>
      </tr>
    `;
  });
  tableHtml += `</table>`;

  container.innerHTML = warningHtml + tableHtml;
  document.getElementById("saveSetupBtn").style.display = "inline-block";
  document.getElementById("saveSetupBtn").textContent =
    existingEntries.length > 0 ? "🔄 Replace Opening Balances" : "💾 Save Opening Balances";
}

async function saveOpeningBalances(options = {}) {
  const newRows = [];
  let hasError = false;
  let missingDateAccount = "";

  setupAccounts.forEach((account, idx) => {
    const dateInput = document.getElementById("setupDate_" + idx).value;
    const val       = parseFloat(document.getElementById("setupAmt_" + idx).value);

    if (isNaN(val) || val === 0) return; // skip blank rows

    if (!dateInput) {
      missingDateAccount = missingDateAccount || account.name;
      hasError = true;
      return;
    }

    // Write as YYYY-MM-DD — slash formats are read as MM/DD/YYYY (American) by Excel/Graph API
    newRows.push([dateInput, "Opening Balance", val, "Income", "Opening Balance", account.name]);
  });

  if (hasError) {
    const message = "Please select a date for " + missingDateAccount + ".";
    if (options.rethrow) throw new Error(message);
    alert(message);
    return false;
  }
  if (newRows.length === 0 && !(options.allowEmpty && options.replaceExistingWhenEmpty)) {
    if (options.allowEmpty) {
      log("No opening balances entered; skipping opening balance save.");
      return false;
    }
    alert("Please enter at least one opening balance.");
    return false;
  }

  try {
    log("Reading current transaction sheet...");
    const token = await getToken();
    const encodedPath = getEncodedExcelPath();

    const usedRangeUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" + encodedPath +
      ":/workbook/worksheets('" + CONFIG.sheetName + "')/usedRange(valuesOnly=true)";
    const usedRange = await graphGetJson(usedRangeUrl, token);
    const allValues = usedRange.values;

    if (!allValues || allValues.length === 0) throw new Error("Transaction sheet appears empty.");

    const headers = allValues[0].map(h => String(h ?? "").trim());
    const subCatIdx = headers.indexOf("Sub Category");
    const accountIdx = headers.indexOf("Account");
    if (subCatIdx !== 4 || accountIdx !== 5 || headers[0] !== "Date" ||
        (usedRange.rowIndex || 0) !== 0 || (usedRange.columnIndex || 0) !== 0) {
      throw new Error("Transaction headers are not in the expected A1:F1 layout. No opening balances were changed.");
    }
    const key = value => String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    const positions = new Map();
    allValues.slice(1).forEach((row, index) => {
      if (key(row[subCatIdx]) !== "opening balance") return;
      const account = key(row[accountIdx]);
      if (!positions.has(account)) positions.set(account, []);
      positions.get(account).push(index + 2);
    });
    // Validate the complete plan before any write. Never compact/rewrite the
    // transaction sheet or remove balances merely because an input is blank.
    const writes = newRows.map(row => {
      const matches = positions.get(key(row[5])) || [];
      if (matches.length > 1) throw new Error("Multiple opening balances exist for " + row[5] + ". Review those rows before editing.");
      return { row, position: matches[0] || null };
    });
    let appendRow = allValues.length + 1;
    for (const write of writes) {
      const position = write.position || appendRow++;
      await writeExcelRange(CONFIG.sheetName, `A${position}:F${position}`, [write.row]);
    }

    log("Opening balances saved.");
    if (!options.silentSuccess) {
      const successEl = document.getElementById("setupSuccess");
      const saveBtn = document.getElementById("saveSetupBtn");
      if (successEl) successEl.style.display = "block";
      if (saveBtn) saveBtn.style.display = "none";
      if (!options.skipReload) {
        setTimeout(() => {
          if (successEl) successEl.style.display = "none";
          if (typeof loadAccountsPage === "function") loadAccountsPage();
          else loadSetupPage();
        }, 3000);
      }
    }
    return true;

  } catch (err) {
    log("ERROR: " + err.message);
    if (!options.silentError) alert("Failed to save: " + err.message);
    console.error(err);
    if (options.rethrow) throw err;
    return false;
  }
}

/** Convert a raw Excel date value (number or DD/MM/YYYY string) to YYYY-MM-DD for <input type="date"> */
function excelDateToInputValue(value) {
  if (!value) return "";
  if (typeof value === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(value);
      return d.y + "-" + String(d.m).padStart(2,"0") + "-" + String(d.d).padStart(2,"0");
    } catch { return ""; }
  }
  const s = String(value).trim();
  // DD/MM/YYYY — our stored format
  const ddmm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmm) return ddmm[3] + "-" + ddmm[2].padStart(2,"0") + "-" + ddmm[1].padStart(2,"0");
  // Already YYYY-MM-DD
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  return "";
}

/** Convert a raw Excel date value to a human-readable display string */
function excelDateToDisplay(value) {
  const iso = excelDateToInputValue(value);
  if (!iso) return String(value);
  const [y, m, d] = iso.split("-");
  return d + "/" + m + "/" + y;
}

function colIndexToLetter(idx) {
  let letter = "";
  idx += 1;
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    idx = Math.floor((idx - 1) / 26);
  }
  return letter;
}

function escapeHtml(v) {
  return String(v).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
