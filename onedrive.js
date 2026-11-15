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

const EXCEL_CACHE_NAME = "fintrack-workbook-v2";
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

// Only replay idempotent reads and fixed-range writes, never row-append POSTs.
let excelSession = null;
let excelSessionPromise = null;
let excelSessionUnsupported = false;
const excelPause = ms => new Promise(resolve => setTimeout(resolve, ms));
async function excelRequest(url, options, retrySafe = false) {
  for (let attempt = 0; ; attempt++) {
    const response = await fetch(url, { ...options, ...(!options?.method || options.method === "GET" ? {cache:"no-store"} : {}) });
    if (response.ok) return response;
    const detail = await response.text();
    const transient = [429,502,503,504].includes(response.status) ||
      (response.status === 500 && /GenericFileOpenError|internalServerError/i.test(detail));
    if (retrySafe && transient && attempt < 2) {
      const retry = Number(response.headers?.get('Retry-After'));
      const delay = Number.isFinite(retry) && retry > 0 ? retry * 1000 : 1000 * 2 ** attempt;
      if (delay <= 30000) {
        setExcelBusyMessage('Excel is temporarily unavailable. Retrying safely…');
        await excelPause(delay);
        continue;
      }
    }
    const error = new Error(/GenericFileOpenError/i.test(detail)
      ? 'Microsoft Excel could not open the workbook. Your changes are not confirmed saved. Close any Excel editing window, then retry. If this continues, open the workbook in Excel Online to check it.'
      : response.status + ' ' + detail);
    error.status = response.status; error.detail = detail;
    throw error;
  }
}
async function excelSessionHeaders(url, token) {
  const match = url.match(/^(.*\/workbook)(?:\/|$)/);
  if (!match || excelSessionUnsupported) return {};
  const base = match[1];
  if (excelSession?.base === base && Date.now() - excelSession.used < 180000) {
    excelSession.used = Date.now();
    return {'workbook-session-id': excelSession.id};
  }
  if (!excelSessionPromise) excelSessionPromise = (async () => {
    try {
      const response = await excelRequest(base + '/createSession', {
        method:'POST', headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
        body:JSON.stringify({persistChanges:true})
      }, true);
      const session = await response.json();
      if (!session.id) throw new Error('Excel did not return a workbook session. Please retry.');
      excelSession = {base,id:session.id,used:Date.now()};
    } catch(error) {
      // Some account/workbook types do not support sessions; retain range access.
      if ([400,403,405,501].includes(error.status) && /notSupported|not supported|unsupported/i.test(error.detail||'')) {
        excelSessionUnsupported = true; return;
      }
      throw error;
    } finally { excelSessionPromise = null; }
  })();
  await excelSessionPromise;
  return excelSession ? {'workbook-session-id':excelSession.id} : {};
}
async function excelRangeRequest(url, token, method='GET', body) {
  for (let attempt=0;attempt<2;attempt++) {
    const sessionHeaders=await excelSessionHeaders(url,token);
    try {
      return await excelRequest(url,{method,headers:{Authorization:`Bearer ${token}`,
        'Content-Type':'application/json',...sessionHeaders},...(body===undefined?{}:{body:JSON.stringify(body)})},true);
    } catch(error) {
      if(attempt===0 && /invalidSessionReCreatable|invalidSessionNotFound/i.test(error.detail||'')) {excelSession=null;continue;}
      throw error;
    }
  }
}

async function graphFetch(url, token) {
  return withExcelBusy("Reading from Excel...", async () => {
    const response = await excelRangeRequest(url, token);

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
    const response = await excelRangeRequest(url, token, 'PATCH', body);

    const result = await response.json();
    if (!preserveCache) await invalidateExcelDownloadCache();
    return result;
  });
}

async function graphPost(url, token, body) {
  return withExcelBusy("Saving to Excel...", async () => {
    const sessionHeaders = await excelSessionHeaders(url, token);
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...sessionHeaders
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

function excelFileScope() {
  const account=msalInstance.getActiveAccount?.() || msalInstance.getAllAccounts?.()[0];
  return encodeURIComponent((account?.homeAccountId || account?.username || 'signed-out')+'|'+CONFIG.filePath);
}
function excelCacheTimeKey(){return EXCEL_CACHE_TIME_KEY+':'+excelFileScope();}
function excelRecentRangesKey(){return 'fintrack.recentRanges.v1:'+excelFileScope();}
function readExcelRecentRanges(){
  try{return JSON.parse(localStorage.getItem(excelRecentRangesKey())||'[]');}catch(_){return [];}
}
function recordExcelRecentRange(update){
  const previous=readExcelRecentRanges().filter(item=>item.sheetName!==update.sheetName||item.rangeAddress!==update.rangeAddress);
  previous.push({sheetName:update.sheetName,rangeAddress:update.rangeAddress,rows:update.values.length,cols:Math.max(...update.values.map(row=>row.length)),id:Date.now()+':'+Math.random()});
  localStorage.setItem(excelRecentRangesKey(),JSON.stringify(previous));
}
function retireExcelRecentRanges(ids){
  localStorage.setItem(excelRecentRangesKey(),JSON.stringify(readExcelRecentRanges().filter(item=>!ids.includes(item.id))));
}
function withExcelFileLock(work){
  return typeof navigator!=='undefined'&&navigator.locks?.request
    ? navigator.locks.request('fintrack:'+excelFileScope(),work):work();
}
function excelCellEqual(a,b){
  if(a===null||a===undefined)a='';if(b===null||b===undefined)b='';
  if(a===b)return true;
  // Excel converts ISO date inputs into numeric serials on write.
  const date=value=>typeof value==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(value)?(Date.parse(value+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000:value;
  const aa=date(a),bb=date(b);
  return typeof aa==='number'&&typeof bb==='number'&&Math.abs(aa-bb)<1e-8;
}
function excelValuesMatch(actual,expected){
  return expected.every((row,r)=>row.every((value,c)=>excelCellEqual(actual?.[r]?.[c],value)));
}
async function reconcileExcelDownload(buffer,fromServer){
  const ranges=readExcelRecentRanges();if(!ranges.length)return buffer;
  setExcelBusyMessage('Checking recent changes in Excel…');
  const workbook=XLSX.read(buffer,{type:'array'}),retired=[];
  for(const range of ranges){
    const sheet=workbook.Sheets[range.sheetName];
    if(!sheet)throw new Error('Could not verify recently edited sheet '+range.sheetName+'. Reload before editing.');
    const live=await readExcelRange(range.sheetName,range.rangeAddress);
    const values=Array.from({length:range.rows},(_,r)=>Array.from({length:range.cols},(_,c)=>live.values?.[r]?.[c]??''));
    const origin=XLSX.utils.decode_cell(range.rangeAddress.split(':')[0]);
    const downloaded=values.map((row,r)=>row.map((_,c)=>sheet[XLSX.utils.encode_cell({r:origin.r+r,c:origin.c+c})]?.v??''));
    if(fromServer&&excelValuesMatch(downloaded,values))retired.push(range.id);
    // Live Excel cells win, including blank/deleted cells and external edits.
    // Never replay old desired values over a newer external change.
    XLSX.utils.sheet_add_aoa(sheet,values,{origin:range.rangeAddress.split(':')[0]});
  }
  if(retired.length)retireExcelRecentRanges(retired);
  return XLSX.write(workbook,{type:'array',bookType:'xlsx'});
}

function getExcelCacheRequest() {
  if (typeof Request === "undefined" || !window.location?.origin) return null;
  const cacheUrl = new URL("/__fintrack_workbook_cache__", window.location.origin);
  cacheUrl.searchParams.set("path", CONFIG.filePath);
  cacheUrl.searchParams.set("account",excelFileScope());
  return new Request(cacheUrl.toString());
}

async function readExcelDownloadCache() {
  if (!("caches" in window)) return null;
  const savedAt = Number(localStorage.getItem(excelCacheTimeKey()) || 0);
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
      localStorage.setItem(excelCacheTimeKey(), String(savedAt));
    }
  });
  excelCacheQueue = store.catch(() => {});
  return store;
}

async function invalidateExcelDownloadCache() {
  excelCacheGeneration++;
  try { localStorage.removeItem(excelCacheTimeKey()); } catch (_) {}
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

// Graph /content redirects cannot reliably follow an authenticated CORS
// preflight in browsers. Resolve a short-lived download URL via metadata,
// then fetch bytes without forwarding the Graph bearer token.
async function downloadExcelBytes(encodedPath, token) {
  const metadataUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" +
    encodedPath;
  const metadata = await graphGetJson(metadataUrl, token);
  const downloadUrl = metadata["@microsoft.graph.downloadUrl"];
  if (typeof downloadUrl !== "string" || !downloadUrl.startsWith("https://")) {
    throw new Error("OneDrive did not provide a workbook download link. Please retry loading.");
  }
  let response;
  try {
    response = await excelRequest(downloadUrl, {method:"GET", credentials:"omit"}, true);
  } catch (error) {
    // Never include the preauthenticated URL in logs or error messages.
    throw new Error(error.status
      ? "Workbook download failed (HTTP " + error.status + "). Please retry loading."
      : "Could not download the workbook from OneDrive. Check your connection and retry loading. Your saved data has not changed.");
  }
  return response.arrayBuffer();
}

async function downloadExcelFile(forceRefresh = false) {
  if(workbookDownloadPromise){
    if(!forceRefresh)return workbookDownloadPromise;
    try{await workbookDownloadPromise;}catch(_){} // Explicit reload starts fresh even after a failed load.
  }
  await excelWriteQueue;
  const download=withExcelFileLock(async()=>{
    return withExcelBusy('Loading from Excel...',async()=>{
      if(!forceRefresh){
        let cached=null;try{cached=await readExcelDownloadCache();}catch(_){}
        if(cached)return reconcileExcelDownload(cached,false);
      }
      const generation=excelCacheGeneration,token=await getToken();
      if(!token)throw new Error('No access token available yet. Please wait for login redirect to complete.');
      const bytes=await downloadExcelBytes(getEncodedExcelPath(),token);
      const reconciled=await reconcileExcelDownload(bytes,true);
      // Cache the untouched server snapshot: patched copies must not be mistaken
      // for proof that the downloadable file has caught up with Excel edits.
      if(generation===excelCacheGeneration){try{await storeExcelDownloadCache(bytes,Date.now(),generation);}catch(_){}}
      return reconciled;
    });
  });
  workbookDownloadPromise=download;
  try{return await download;}finally{if(workbookDownloadPromise===download)workbookDownloadPromise=null;}
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
  return writeExcelRanges([{ sheetName, rangeAddress, values }]).then(results => results[0]);
}

function writeExcelRanges(updates) {
  // Snapshot edits before queuing; one workbook rebuild per save, not per range.
  const pending = updates.map(update => ({ ...update, values: update.values.map(row => row.slice()) }));
  const write = excelWriteQueue.then(()=>withExcelFileLock(async()=>{
    // Persist range markers before sending edits, including partially failed batches.
    // A reload must verify their live values, never quietly use a stale snapshot.
    for(const update of pending)recordExcelRecentRange(update);
    await invalidateExcelDownloadCache();
    const results=[];
    for(const update of pending){
      const result=await writeExcelRangeNow(update.sheetName,update.rangeAddress,update.values);
      const live=await readExcelRange(update.sheetName,update.rangeAddress);
      if(!excelValuesMatch(live.values,update.values))throw new Error('Save could not be verified in Excel. Some changes may already be saved. Reload before retrying.');
      results.push(result);
    }
    return results;
  }));
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

    const result = await graphPatch(url, token, { values }, true);
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
