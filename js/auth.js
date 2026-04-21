// js/auth.js
// MSAL.js v3 の薄いラッパー。他モジュールは window.Auth を介してのみ認証を扱う。
(function () {
  "use strict";

  let msalInstance = null;

  // ページ起動時に一度だけ呼び出す。既存セッションがあればアカウントを返す。
  async function initAuth() {
    const cfg = window.APP_CONFIG;

    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId:    cfg.clientId,
        authority:   cfg.authority,
        redirectUri: cfg.redirectUri,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });

    // MSAL v3 では initialize() が必須
    await msalInstance.initialize();

    // リダイレクトフロー用（今回は popup を使うが念のため）
    await msalInstance.handleRedirectPromise();

    // キャッシュ済みアカウントを復元
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
      return accounts[0];
    }

    // サイレント SSO で既存 Entra ID セッションを利用
    try {
      const res = await msalInstance.ssoSilent({ scopes: cfg.scopes });
      msalInstance.setActiveAccount(res.account);
      return res.account;
    } catch {
      return null; // ログインが必要
    }
  }

  async function login() {
    const res = await msalInstance.loginPopup({
      scopes: window.APP_CONFIG.scopes,
      prompt: "select_account",
    });
    msalInstance.setActiveAccount(res.account);
    return res.account;
  }

  async function logout() {
    await msalInstance.logoutPopup({
      account: msalInstance.getActiveAccount(),
    });
  }

  // 有効なアクセストークン文字列を返す。期限切れ近くなれば自動更新。
  async function getToken() {
    const request = {
      scopes:  window.APP_CONFIG.scopes,
      account: msalInstance.getActiveAccount(),
    };
    try {
      const res = await msalInstance.acquireTokenSilent(request);
      return res.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        const res = await msalInstance.acquireTokenPopup(request);
        return res.accessToken;
      }
      throw err;
    }
  }

  window.Auth = { initAuth, login, logout, getToken };
})();
