let accountsList = [];
let incomeSubCategoryList = [];
let openingBalanceTxSheet = null;
let existingOpeningEntries = [];
let openingBalanceDrafts = {};
let accountsAutoSaveTimer = null;
let accountsAutoSaveInFlight = false;

const ACCOUNTS_FULL_RANGE = "J2:K10";
const INCOME_SUBS_RANGE = "Q2:Q10";
const MAX_ACCOUNTS = 9;
const MAX_INCOME_SUBS = 9;

async function loadAccountsPage() {
  try {
    clearOutput();
    log("Downloading Excel file...");
    const arrayBuffer = await downloadExcelFile();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const budgetSheet = workbook.Sheets["Budget Setup"];
    if (!budgetSheet) throw new Error("Sheet not found: Budget Setup");
    openingBalanceTxSheet = workbook.Sheets[CONFIG.sheetName] || null;
    existingOpeningEntries = typeof detectExistingOpeningBalances === "function" && openingBalanceTxSheet
      ? detectExistingOpeningBalances(openingBalanceTxSheet)
      : [];
    openingBalanceDrafts = {};
    accountsList = readAccountsFromSheet(budgetSheet);
    incomeSubCategoryList = readIncomeSubsFromSheet(budgetSheet);
    renderAccountsTable();
    renderIncomeSubsTable();
    log("Accounts loaded.");
  } catch (err) {
    log("ERROR: " + err.message);
    console.error(err);
  }
}

function readAccountsFromSheet(sheet) {
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => ({ name: String(row[9] ?? "").trim(), type: String(row[10] ?? "Savings").trim() }))
    .filter(a => a.name !== "");
}

function readIncomeSubsFromSheet(sheet) {
  // Column Q = index 16
  const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
  return allRows.slice(1)
    .map(row => String(row[16] ?? "").trim())
    .filter(v => v !== "");
}

function addAccount() {
  captureOpeningBalanceDrafts();
  const name = document.getElementById("accountNameInput").value.trim();
  const type = document.getElementById("accountTypeInput").value;
  if (!name) { alert("Please enter an account name."); return; }
  if (accountsList.length >= MAX_ACCOUNTS) { alert("Maximum " + MAX_ACCOUNTS + " accounts allowed."); return; }
  if (accountsList.find(a => a.name.toLowerCase() === name.toLowerCase())) { alert("Account already exists."); return; }
  accountsList.push({ name, type });
  document.getElementById("accountNameInput").value = "";
  renderAccountsTable();
  scheduleAccountsAutoSave();
}

function deleteAccount(index) {
  captureOpeningBalanceDrafts();
  const key = accountKey(accountsList[index]?.name);
  if (key) delete openingBalanceDrafts[key];
  accountsList.splice(index, 1);
  renderAccountsTable();
  scheduleAccountsAutoSave();
}

function updateAccountName(index, value) {
  const oldName = accountsList[index]?.name || "";
  const oldKey = accountKey(oldName);
  const draft = readOpeningBalanceRow(index);
  accountsList[index].name = value.trim();
  const newKey = accountKey(accountsList[index].name);
  if (newKey) openingBalanceDrafts[newKey] = draft;
  if (oldKey && oldKey !== newKey) delete openingBalanceDrafts[oldKey];
  scheduleAccountsAutoSave();
}

function updateAccountType(index, value) {
  accountsList[index].type = value;
  scheduleAccountsAutoSave();
}

function renderAccountsTable() {
  const table = document.getElementById("accountsTable");
  table.innerHTML = `
    <tr>
      <th class="acct-name-col">Account</th>
      <th class="acct-type-col">Type</th>
      <th class="acct-date-col">Opening Date</th>
      <th class="acct-balance-col">Opening Balance</th>
      <th class="acct-action-col"></th>
    </tr>`;
  accountsList.forEach((account, index) => {
    const opening = getOpeningBalanceForAccount(account);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeHtml(account.name)}" onchange="updateAccountName(${index}, this.value)"></td>
      <td>
        <select onchange="updateAccountType(${index}, this.value)">
          <option value="Savings" ${account.type === "Savings" ? "selected" : ""}>Savings</option>
          <option value="Credit Card" ${account.type === "Credit Card" ? "selected" : ""}>Credit Card</option>
        </select>
      </td>
      <td>
        <input type="date" id="setupDate_${index}" value="${escapeHtml(opening.date)}"
          onchange="updateOpeningDraft(${index}, 'date', this.value)">
      </td>
      <td>
        <input type="number" step="0.01" id="setupAmt_${index}" placeholder="0.00"
          value="${escapeHtml(opening.amount)}" oninput="updateOpeningDraft(${index}, 'amount', this.value)">
      </td>
      <td><button class="acct-delete-btn" onclick="deleteAccount(${index})">Delete</button></td>
    `;
    table.appendChild(tr);
  });
  syncOpeningBalanceAccounts();
}

function accountKey(name) {
  return String(name || "").trim().toLowerCase();
}

function todayInputValue() {
  const today = new Date();
  return today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
}

function readOpeningBalanceRow(index) {
  return {
    date: document.getElementById("setupDate_" + index)?.value || "",
    amount: document.getElementById("setupAmt_" + index)?.value || ""
  };
}

function updateOpeningDraft(index, field, value) {
  const key = accountKey(accountsList[index]?.name);
  if (!key) return;
  const draft = openingBalanceDrafts[key] || readOpeningBalanceRow(index);
  draft[field] = value;
  openingBalanceDrafts[key] = draft;
  scheduleAccountsAutoSave();
}

function captureOpeningBalanceDrafts() {
  accountsList.forEach((account, index) => {
    const key = accountKey(account.name);
    if (!key) return;
    const row = readOpeningBalanceRow(index);
    if (row.date || row.amount) openingBalanceDrafts[key] = row;
  });
}

function getOpeningBalanceForAccount(account) {
  const key = accountKey(account.name);
  const draft = openingBalanceDrafts[key];
  if (draft) return { date: draft.date || todayInputValue(), amount: draft.amount || "" };

  const existing = existingOpeningEntries.find(entry =>
    accountKey(entry.account) === key
  );
  if (!existing) return { date: todayInputValue(), amount: "" };

  return {
    date: typeof excelDateToInputValue === "function" ? excelDateToInputValue(existing.date) : "",
    amount: Number(existing.amount || 0) ? Number(existing.amount).toFixed(2) : ""
  };
}

function syncOpeningBalanceAccounts() {
  if (typeof setupAccounts === "undefined") return;
  setupAccounts = accountsList
    .map(account => ({
      name: String(account.name || "").trim(),
      type: String(account.type || "Savings").trim() || "Savings"
    }))
    .filter(account => account.name !== "");
}

function hasOpeningBalanceInput() {
  return accountsList.some((_, index) => {
    const val = parseFloat(document.getElementById("setupAmt_" + index)?.value);
    return !isNaN(val) && val !== 0;
  });
}

function getCurrentOpeningBalanceEntries() {
  return accountsList.map((account, index) => {
    const val = parseFloat(document.getElementById("setupAmt_" + index)?.value);
    if (isNaN(val) || val === 0) return null;
    return {
      date: document.getElementById("setupDate_" + index)?.value || todayInputValue(),
      account: account.name,
      amount: val
    };
  }).filter(Boolean);
}

function addIncomeSub() {
  const name = document.getElementById("incomeSubInput").value.trim();
  if (!name) { alert("Please enter a sub-category name."); return; }
  if (incomeSubCategoryList.length >= MAX_INCOME_SUBS) { alert("Maximum " + MAX_INCOME_SUBS + " income sub-categories allowed."); return; }
  if (incomeSubCategoryList.find(s => s.toLowerCase() === name.toLowerCase())) { alert("Already exists."); return; }
  incomeSubCategoryList.push(name);
  document.getElementById("incomeSubInput").value = "";
  renderIncomeSubsTable();
  scheduleAccountsAutoSave();
}

function deleteIncomeSub(index) {
  incomeSubCategoryList.splice(index, 1);
  renderIncomeSubsTable();
  scheduleAccountsAutoSave();
}
function updateIncomeSub(index, value) {
  incomeSubCategoryList[index] = value.trim();
  scheduleAccountsAutoSave();
}

function renderIncomeSubsTable() {
  const table = document.getElementById("incomeSubsTable");
  table.innerHTML = `<tr><th>Income Sub Category</th><th></th></tr>`;
  incomeSubCategoryList.forEach((sub, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input value="${escapeHtml(sub)}" onchange="updateIncomeSub(${index}, this.value)"></td>
      <td><button onclick="deleteIncomeSub(${index})">Delete</button></td>
    `;
    table.appendChild(tr);
  });
}

function scheduleAccountsAutoSave(delay = 900) {
  clearTimeout(accountsAutoSaveTimer);
  setAccountsAutoSaveStatus("Saving soon...");
  accountsAutoSaveTimer = setTimeout(() => {
    saveAccountsToExcel({ silent: true });
  }, delay);
}

function setAccountsAutoSaveStatus(message, tone = "") {
  const el = document.getElementById("accountsAutosaveStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "accounts-autosave-status" + (tone ? " " + tone : "");
}

async function saveAccountsToExcel(options = {}) {
  if (accountsAutoSaveInFlight) {
    scheduleAccountsAutoSave(1200);
    return false;
  }

  const silent = options.silent === true;
  accountsAutoSaveInFlight = true;
  try {
    setAccountsAutoSaveStatus("Saving...");
    captureOpeningBalanceDrafts();
    syncOpeningBalanceAccounts();

    log("Saving accounts to Excel...");
    const acctValues = accountsList.map(a => [a.name, a.type]);
    while (acctValues.length < MAX_ACCOUNTS) acctValues.push(["", ""]);
    await writeBudgetSetupRange(ACCOUNTS_FULL_RANGE, acctValues);
    log("Accounts saved.");

    log("Saving income sub-categories to column Q...");
    const subValues = incomeSubCategoryList.map(s => [s]);
    while (subValues.length < MAX_INCOME_SUBS) subValues.push([""]);
    await writeBudgetSetupRange(INCOME_SUBS_RANGE, subValues);

    let openingSaved = false;
    const shouldSaveOpeningBalances = typeof saveOpeningBalances === "function" &&
      (hasOpeningBalanceInput() || (!silent && existingOpeningEntries.length > 0));
    if (shouldSaveOpeningBalances) {
      openingSaved = await saveOpeningBalances({
        allowEmpty: true,
        replaceExistingWhenEmpty: existingOpeningEntries.length > 0,
        rethrow: true,
        silentError: silent,
        silentSuccess: true,
        skipReload: true
      });
      if (openingSaved) existingOpeningEntries = getCurrentOpeningBalanceEntries();
    }

    const successEl = document.getElementById("setupSuccess");
    if (successEl && !silent) {
      successEl.textContent = openingSaved
        ? "Accounts, categories, and opening balances saved."
        : "Accounts and categories saved.";
      successEl.style.display = "block";
      setTimeout(() => { successEl.style.display = "none"; }, 3000);
    }

    setAccountsAutoSaveStatus("Saved to Excel", "ok");
    if (!silent) alert(openingSaved ? "Accounts, categories, and opening balances saved." : "Accounts and categories saved.");
    log("Done.");
    return true;
  } catch (err) {
    setAccountsAutoSaveStatus("Save failed", "error");
    log("SAVE ERROR: " + err.message);
    if (!silent) alert(err.message);
    console.error(err);
    return false;
  } finally {
    accountsAutoSaveInFlight = false;
  }
}

function escapeHtml(value) {
  return String(value).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
}
