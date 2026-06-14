const CONFIG = {
  clientId: "768e8544-9eab-40cb-940e-5ba478af614f",
  authority: "https://login.microsoftonline.com/consumers",

  redirectUri:
    window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1"
      ? "http://localhost:5500"
      : "https://rexrxlee.github.io/Master/",

  filePath: "Master.xlsx",
  sheetName: "Transaction",
  scopes: ["Files.ReadWrite"]
};
