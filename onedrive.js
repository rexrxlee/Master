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

async function downloadExcelFile() {
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
}

async function readExcelRange(sheetName, rangeAddress) {
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
}

async function writeExcelRange(sheetName, rangeAddress, values) {
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
}

async function readBudgetSetupRange(rangeAddress) {
  return readExcelRange("Budget Setup", rangeAddress);
}

async function writeBudgetSetupRange(rangeAddress, values) {
  return writeExcelRange("Budget Setup", rangeAddress, values);
}