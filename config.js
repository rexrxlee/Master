const CONFIG = {
  clientId: "768e8544-9eab-40cb-940e-5ba478af614f",
  authority: "https://login.microsoftonline.com/consumers",
  redirectUri: window.location.origin + window.location.pathname,
  filePath: "Master.xlsx",
  sheetName: "Transaction",
  scopes: ["Files.Read"]
};