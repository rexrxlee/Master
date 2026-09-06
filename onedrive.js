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
let workbookDownloadPromise = null;
let excelWriteQueue = Promise.resolve();
let excelCacheGeneration = 0;
let excelCacheQueue = Promise.resolve();

const EXCEL_CACHE_NAME = "fintrack-workbook-v1";
const EXCEL_CACHE_MAX_AGE_MS = 2 * 60 * 1000;
const EXCEL_CACHE_TIME_KEY = "fintrack.workbookCacheTime.v1";

const EXCEL_BUSY_ACTIONS = {
  loadDashboard: "Loading dashboard from Excel...",
  loadBudgetPage: "Loading budget from Excel...",
  loadAccountsPage: "Loading accounts from Excel...",
  loadSetupPage: "Loading setup from Excel...",
  loadAddTransactionPage: "Loading transactions from Excel...",
  loadGoalsPage: "Loading goals from Excel...",
  saveBudgetSetupToExcel: "Autosaving budget to Excel...",
  saveAccountsToExcel: "Autosaving accounts to Excel...",
  saveAllocations: "Saving allocations to Excel...",
  saveIncomeBoosts: "Saving cashflow changes to Excel...",
  persistGoalsToExcel: "Autosaving goals to Excel...",
  saveGoalsToExcel: "Saving goals to Excel...",
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

async function graphPatch(url, token, body, preserveCache = false) {
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

    const result = await response.json();
    if (!preserveCache) await invalidateExcelDownloadCache();
    return result;
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

    const result = await response.json();
    await invalidateExcelDownloadCache();
    return result;
  });
}

function getExcelCacheRequest() {
  if (typeof Request === "undefined" || !window.location?.origin) return null;
  const cacheUrl = new URL("/__fintrack_workbook_cache__", window.location.origin);
  cacheUrl.searchParams.set("path", CONFIG.filePath);
  return new Request(cacheUrl.toString());
}

async function readExcelDownloadCache() {
  if (!("caches" in window)) return null;
  const savedAt = Number(localStorage.getItem(EXCEL_CACHE_TIME_KEY) || 0);
  if (!savedAt || Date.now() - savedAt > EXCEL_CACHE_MAX_AGE_MS) return null;

  const request = getExcelCacheRequest();
  if (!request) return null;
  const cache = await caches.open(EXCEL_CACHE_NAME);
  const cached = await cache.match(request);
  return cached ? cached.arrayBuffer() : null;
}

function storeExcelDownloadCache(arrayBuffer, savedAt = Date.now(), generation = excelCacheGeneration) {
  const store = excelCacheQueue.then(async () => {
    if (!("caches" in window) || generation !== excelCacheGeneration) return;
    const request = getExcelCacheRequest();
    if (!request) return;
    const cache = await caches.open(EXCEL_CACHE_NAME);
    await cache.put(request, new Response(arrayBuffer.slice(0)));
    if (generation === excelCacheGeneration) {
      localStorage.setItem(EXCEL_CACHE_TIME_KEY, String(savedAt));
    }
  });
  excelCacheQueue = store.catch(() => {});
  return store;
}

async function invalidateExcelDownloadCache() {
  excelCacheGeneration++;
  try { localStorage.removeItem(EXCEL_CACHE_TIME_KEY); } catch (_) {}
  const clear = excelCacheQueue.then(async () => {
    if (!("caches" in window)) return;
    const request = getExcelCacheRequest();
    if (!request) return;
    const cache = await caches.open(EXCEL_CACHE_NAME);
    await cache.delete(request);
  });
  excelCacheQueue = clear.catch(() => {});
  await excelCacheQueue;
}

async function downloadExcelFile(forceRefresh = false) {
  if (!forceRefresh) {
    try {
      const cached = await readExcelDownloadCache();
      if (cached) return cached;
    } catch (_) {
      // Cache failures must never prevent a normal workbook download.
    }
  }

  if (workbookDownloadPromise) return workbookDownloadPromise;

  workbookDownloadPromise = withExcelBusy("Loading from Excel...", async () => {
    const generation = excelCacheGeneration;
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
    const arrayBuffer = await response.arrayBuffer();
    if (generation === excelCacheGeneration) {
      try { await storeExcelDownloadCache(arrayBuffer, Date.now(), generation); } catch (_) {}
    }
    return arrayBuffer;
  });

  try {
    return await workbookDownloadPromise;
  } finally {
    workbookDownloadPromise = null;
  }
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

function writeExcelRange(sheetName, rangeAddress, values) {
  const write = excelWriteQueue.then(() => writeExcelRangeNow(sheetName, rangeAddress, values));
  excelWriteQueue = write.catch(() => {});
  return write;
}

async function writeExcelRangeNow(sheetName, rangeAddress, values) {
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

    // Retain the downloaded snapshot after confirmed range writes. This avoids
    // downloading the entire workbook again when navigating to another page.
    let cached = null;
    let savedAt = 0;
    try {
      cached = await readExcelDownloadCache();
      savedAt = Number(localStorage.getItem(EXCEL_CACHE_TIME_KEY));
    } catch (_) {}
    const result = await graphPatch(url, token, { values }, true);
    await invalidateExcelDownloadCache();
    if (cached) {
      try {
        const workbook = XLSX.read(cached, { type: "array" });
        const sheet = workbook.Sheets[sheetName];
        // Formula dependencies need a server refresh, not a partial local update.
        const hasFormulas = Object.values(workbook.Sheets).some(ws =>
          Object.keys(ws).some(key => !key.startsWith("!") && ws[key]?.f));
        const writesFormula = values.some(row => row.some(value =>
          typeof value === "string" && value.startsWith("=")));
        if (sheet && !hasFormulas && !writesFormula) {
          XLSX.utils.sheet_add_aoa(sheet, values, { origin: rangeAddress.split(":")[0] });
          const updated = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
          // Keep the original freshness deadline so external edits are still checked.
          await storeExcelDownloadCache(updated, savedAt);
        }
      } catch (_) {
        // The server save succeeded; cache failures must not turn it into a failed save.
      }
    }
    return result;
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
