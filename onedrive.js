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

async function initializeMsal() {
  if (msalReady) return;

  await msalInstance.initialize();

  const accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);

    const status = document.getElementById("status");

    if (status) {
      status.innerText = "Logged in as " + accounts[0].username;
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

  const result = await msalInstance.loginPopup({
    scopes: CONFIG.scopes
  });

  msalInstance.setActiveAccount(result.account);

  const status = document.getElementById("status");

  if (status) {
    status.innerText = "Logged in as " + result.account.username;
  }

  log("Login successful");
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
    account = msalInstance.getActiveAccount();
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.scopes,
      account: account
    });

    return result.accessToken;
  } catch (err) {
    const result = await msalInstance.acquireTokenPopup({
      scopes: CONFIG.scopes,
      account: account
    });

    return result.accessToken;
  }
}

function getEncodedExcelPath() {
  return CONFIG.filePath
    .split("/")
    .map(part => encodeURIComponent(part))
    .join("/");
}

async function graphFetch(url, token) {
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
}

async function graphGetJson(url, token) {
  const response = await graphFetch(url, token);
  return response.json();
}

async function graphPatch(url, token, body) {
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
}

async function downloadExcelFile(forceRefresh = false) {
  const cacheKey = "masterExcelFileBase64";

  if (!forceRefresh) {
    const cached = sessionStorage.getItem(cacheKey);

    if (cached) {
      return base64ToArrayBuffer(cached);
    }
  }

  const token = await getToken();
  const encodedPath = getEncodedExcelPath();

  const downloadUrl =
    "https://graph.microsoft.com/v1.0/me/drive/root:/" +
    encodedPath +
    ":/content";

  const response = await graphFetch(downloadUrl, token);

  const arrayBuffer = await response.arrayBuffer();

  sessionStorage.setItem(
    cacheKey,
    arrayBufferToBase64(arrayBuffer)
  );

  return arrayBuffer;
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);

  bytes.forEach(byte => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes.buffer;
}

async function readExcelRange(sheetName, rangeAddress) {
  const token = await getToken();
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
}

async function writeExcelRange(sheetName, rangeAddress, values) {
  const token = await getToken();
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
}

async function readBudgetSetupRange(rangeAddress) {
  return readExcelRange("Budget Setup", rangeAddress);
}

async function writeBudgetSetupRange(rangeAddress, values) {
  return writeExcelRange("Budget Setup", rangeAddress, values);
}
