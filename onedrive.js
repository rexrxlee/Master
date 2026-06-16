const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: CONFIG.authority,
    redirectUri: CONFIG.redirectUri
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let msalReady = false;
let excelBusyDepth = 0;
let excelBusyOverlay = null;
let excelBusyMessage = null;
let excelBusyPreviouslyFocused = null;

const EXCEL_BUSY_ACTIONS = {
  loadDashboard: "Loading dashboard from Excel...",
  loadBudgetPage: "Loading budget from Excel...",
  loadAccountsPage: "Loading accounts from Excel...",
  loadSetupPage: "Loading setup from Excel...",
  loadAddTransactionPage: "Loading transactions from Excel...",
  loadGoalsPage: "Loading goals from Excel...",
  loadGoalItemsTab: "Loading goal items from Excel...",
  loadInsurancePage: "Loading insurance from Excel...",
  refreshInsuranceGoalsForPicker: "Refreshing goals from Excel...",
  refreshInsuranceGoalsFromWorkbook: "Refreshing goals from Excel...",
  saveBudgetSetupToExcel: "Autosaving budget to Excel...",
  saveAccountsToExcel: "Autosaving accounts to Excel...",
  saveAllocations: "Saving allocations to Excel...",
  saveIncomeBoosts: "Saving cashflow changes to Excel...",
  persistRewardStateToExcel: "Saving card rewards to Excel...",
  saveRewardActual: "Saving cashback comparison to Excel...",
  resetRewardSettings: "Saving card reward rules to Excel...",
  persistGoalsToExcel: "Autosaving goals to Excel...",
  saveGoalsToExcel: "Saving goals to Excel...",
  persistGoalItemsToExcel: "Autosaving goal items to Excel...",
  saveAllGoalItems: "Saving goal items to Excel...",
  writeInsuranceHeadersForWrite: "Saving insurance headers to Excel...",
  writeInsurancePolicyFields: "Saving insurance policy to Excel...",
  getNextInsuranceRowNumber: "Reading insurance sheet from Excel...",
  appendInsuranceGoalTransaction: "Saving goal transaction to Excel...",
  updateInsurancePaidCells: "Saving insurance payment to Excel...",
  writeInsuranceGoal: "Saving insurance goal to Excel..."
};

function ensureExcelBusyOverlay() {
  if (excelBusyOverlay) return excelBusyOverlay;

  excelBusyOverlay = document.createElement("div");
  excelBusyOverlay.id = "excelBusyOverlay";
  excelBusyOverlay.className = "excel-busy-overlay";
  excelBusyOverlay.setAttribute("role", "status");
  excelBusyOverlay.setAttribute("aria-live", "polite");
  excelBusyOverlay.setAttribute("aria-label", "Excel operation in progress");
  excelBusyOverlay.setAttribute("tabindex", "-1");
  excelBusyOverlay.innerHTML = `
    <div class="excel-busy-dialog">
      <div class="excel-busy-spinner" aria-hidden="true"></div>
      <div>
        <div class="excel-busy-title">Working with Excel</div>
        <div class="excel-busy-message" id="excelBusyMessage">Please wait...</div>
      </div>
    </div>`;

  document.body.appendChild(excelBusyOverlay);
  excelBusyMessage = document.getElementById("excelBusyMessage");
  return excelBusyOverlay;
}

function setExcelBusyMessage(message) {
  ensureExcelBusyOverlay();
  if (excelBusyMessage) excelBusyMessage.textContent = message || "Please wait...";
}

function showExcelBusy(message = "Syncing with Excel...") {
  if (!document.body) return;
  ensureExcelBusyOverlay();
  excelBusyDepth += 1;
  setExcelBusyMessage(message);
  excelBusyOverlay.classList.add("open");
  document.body.setAttribute("aria-busy", "true");
  if (excelBusyDepth === 1) {
    excelBusyPreviouslyFocused = document.activeElement;
    excelBusyOverlay.focus({ preventScroll: true });
  }
}

function hideExcelBusy() {
  if (excelBusyDepth > 0) excelBusyDepth -= 1;
  if (excelBusyDepth !== 0 || !excelBusyOverlay) return;

  excelBusyOverlay.classList.remove("open");
  document.body.removeAttribute("aria-busy");
  if (excelBusyPreviouslyFocused && typeof excelBusyPreviouslyFocused.focus === "function") {
    try { excelBusyPreviouslyFocused.focus({ preventScroll: true }); } catch (_) {}
  }
  excelBusyPreviouslyFocused = null;
}

async function withExcelBusy(message, work) {
  showExcelBusy(message);
  try {
    return await work();
  } finally {
    hideExcelBusy();
  }
}

function isExcelBusy() {
  return excelBusyDepth > 0;
}

function blockExcelBusyInteraction(event) {
  if (!isExcelBusy()) return;
  if (event.type === "keydown" && (event.metaKey || event.ctrlKey || event.altKey)) return;
  event.preventDefault();
  event.stopImmediatePropagation();
}

function installExcelBusyInteractionBlockers() {
  ["click", "dblclick", "pointerdown", "submit", "keydown"].forEach(type => {
    document.addEventListener(type, blockExcelBusyInteraction, true);
  });
}

function installExcelBusyActionGuards() {
  Object.entries(EXCEL_BUSY_ACTIONS).forEach(([name, message]) => {
    const original = window[name];
    if (typeof original !== "function" || original.__excelBusyWrapped) return;

    const wrapped = function(...args) {
      return withExcelBusy(message, () => original.apply(this, args));
    };
    wrapped.__excelBusyWrapped = true;
    wrapped.__excelBusyOriginal = original;
    window[name] = wrapped;
  });
}

installExcelBusyInteractionBlockers();

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", installExcelBusyActionGuards);
} else {
  installExcelBusyActionGuards();
}
window.addEventListener("load", installExcelBusyActionGuards);

async function initializeMsal() {
  if (msalReady) return;

  await msalInstance.initialize();

  const response = await msalInstance.handleRedirectPromise();

  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
  } else {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }

  const account = msalInstance.getActiveAccount();

  if (account) {
    const status = document.getElementById("status");

    if (status) {
      status.innerText = "Logged in as " + account.username;
    }
  }

  msalReady = true;
}

function log(message) {
  const debug = document.getElementById("debug");

  if (debug) {
    debug.textContent += message + "\n";
  }
}

function clearOutput() {
  const debug = document.getElementById("debug");

  if (debug) {
    debug.textContent = "";
  }
}

async function login() {
  await initializeMsal();

  await msalInstance.loginRedirect({
    scopes: CONFIG.scopes
  });
}

async function getToken() {
  await initializeMsal();

  let account = msalInstance.getActiveAccount();

  if (!account) {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      account = accounts[0];
      msalInstance.setActiveAccount(account);
    }
  }

  if (!account) {
    await login();
    return;
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.scopes,
      account: account
    });

    return result.accessToken;
  } catch (err) {
    await msalInstance.acquireTokenRedirect({
      scopes: CONFIG.scopes,
      account: account
    });
  }
}

function getEncodedExcelPath() {
  return CONFIG.filePath
    .split("/")
    .map(part => encodeURIComponent(part))
    .join("/");
}

async function graphFetch(url, token) {
  return withExcelBusy("Reading from Excel...", async () => {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(response.status + " " + errorText);
    }

    return response;
  });
}

async function graphGetJson(url, token) {
  return withExcelBusy("Reading from Excel...", async () => {
    const response = await graphFetch(url, token);
    return response.json();
  });
}

async function graphPatch(url, token, body) {
  return withExcelBusy("Saving to Excel...", async () => {
    const response = await fetch(url, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(response.status + " " + errorText);
    }

    return response.json();
  });
}

async function graphPost(url, token, body) {
  return withExcelBusy("Saving to Excel...", async () => {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(response.status + " " + errorText);
    }

    return response.json();
  });
}

async function downloadExcelFile() {
  return withExcelBusy("Loading from Excel...", async () => {
    const token = await getToken();

    if (!token) {
      throw new Error("No access token available yet. Please wait for login redirect to complete.");
    }

    const encodedPath = getEncodedExcelPath();

    const downloadUrl =
      "https://graph.microsoft.com/v1.0/me/drive/root:/" +
      encodedPath +
      ":/content";

    const response = await graphFetch(downloadUrl, token);

    return response.arrayBuffer();
  });
}

async function readExcelRange(sheetName, rangeAddress) {
  return withExcelBusy("Reading from Excel...", async () => {
    const token = await getToken();

    if (!token) {
      throw new Error("No access token available yet. Please wait for login redirect to complete.");
    }

    const encodedPath = getEncodedExcelPath();

    const url =
      "https://graph.microsoft.com/v1.0/me/drive/root:/" +
      encodedPath +
      ":/workbook/worksheets('" +
      sheetName.replace(/'/g, "''") +
      "')/range(address='" +
      rangeAddress +
      "')";

    return graphGetJson(url, token);
  });
}

async function writeExcelRange(sheetName, rangeAddress, values) {
  return withExcelBusy("Saving to Excel...", async () => {
    const token = await getToken();

    if (!token) {
      throw new Error("No access token available yet. Please wait for login redirect to complete.");
    }

    const encodedPath = getEncodedExcelPath();

    const url =
      "https://graph.microsoft.com/v1.0/me/drive/root:/" +
      encodedPath +
      ":/workbook/worksheets('" +
      sheetName.replace(/'/g, "''") +
      "')/range(address='" +
      rangeAddress +
      "')";

    return graphPatch(url, token, {
      values: values
    });
  });
}

async function readBudgetSetupRange(rangeAddress) {
  return readExcelRange("Budget Setup", rangeAddress);
}

async function writeBudgetSetupRange(rangeAddress, values) {
  return writeExcelRange("Budget Setup", rangeAddress, values);
}

async function addExcelWorksheet(sheetName) {
  return withExcelBusy("Preparing Excel sheet...", async () => {
    const token = await getToken();

    if (!token) {
      throw new Error("No access token available yet. Please wait for login redirect to complete.");
    }

    const encodedPath = getEncodedExcelPath();

    const url =
      "https://graph.microsoft.com/v1.0/me/drive/root:/" +
      encodedPath +
      ":/workbook/worksheets/add";

    return graphPost(url, token, { name: sheetName });
  });
}
