const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: CONFIG.authority,
    redirectUri: CONFIG.redirectUri
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

function log(message) {
  document.getElementById("debug").textContent += message + "\n";
}

function clearOutput() {
  document.getElementById("debug").textContent = "";
}

async function login() {
  await msalInstance.initialize();

  const result = await msalInstance.loginPopup({
    scopes: CONFIG.scopes
  });

  document.getElementById("status").innerText =
    "Logged in as " + result.account.username;

  log("Login successful");
}

async function getToken() {
  await msalInstance.initialize();

  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    await login();
  }

  const account = msalInstance.getAllAccounts()[0];

  const result = await msalInstance.acquireTokenSilent({
    scopes: CONFIG.scopes,
    account: account
  });

  return result.accessToken;
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

async function downloadExcelFile() {
  const token = await getToken();

  const encodedPath = CONFIG.filePath
    .split("/")
    .map(part => encodeURIComponent(part))
    .join("/");

  const downloadUrl =
    "https://graph.microsoft.com/v1.0/me/drive/root:/" +
    encodedPath +
    ":/content";

  const response = await graphFetch(downloadUrl, token);

  return response.arrayBuffer();
}