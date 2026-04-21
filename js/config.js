// js/config.js
// Azure AD アプリ登録後に clientId と tenantId を書き換えてください
window.APP_CONFIG = Object.freeze({
  clientId:    "e38ab014-d46d-406e-970a-a5e816f502fe",
  tenantId:    "4705d2d2-adc0-4eb4-9423-fc878bdcb170",
  authority:   "https://login.microsoftonline.com/4705d2d2-adc0-4eb4-9423-fc878bdcb170",
  redirectUri: "https://flect.sharepoint.com/sites/CI_AE/Shared%20Documents/pptx-viewer/index.html",
  scopes:      ["Files.Read.All"],
  graphBaseUrl: "https://graph.microsoft.com/v1.0",

  // ツリー表示の起点フォルダ
  // searchSite:       Graph API サイトパス形式（hostname:/sites/sitename）
  // searchFolderPath: ドキュメントライブラリルートからの相対パス
  searchSite:       "flect.sharepoint.com:/sites/CI_AE",
  searchFolderPath: "02_個別案件",
});
