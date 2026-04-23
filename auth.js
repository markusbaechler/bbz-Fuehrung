// ==========================================================
// auth.js - Microsoft Authentication (MSAL.js)
// ==========================================================

const AuthService = (() => {
  let msalInstance = null;
  let currentAccount = null;

  const SCOPES = [
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/User.Read'
  ];

  async function init() {
    if (!window.msal) {
      throw new Error('MSAL library nicht geladen - prüfe index.html');
    }

    const cfg = window.BBZ_CONFIG.auth;
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
        postLogoutRedirectUri: window.location.origin + window.location.pathname
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
      }
    });

    await msalInstance.initialize();

    // Handle redirect response on page load
    const resp = await msalInstance.handleRedirectPromise();
    if (resp) {
      currentAccount = resp.account;
      msalInstance.setActiveAccount(currentAccount);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        currentAccount = accounts[0];
        msalInstance.setActiveAccount(currentAccount);
      }
    }

    return currentAccount;
  }

  async function signIn() {
    if (!msalInstance) await init();
    try {
      const resp = await msalInstance.loginPopup({
        scopes: SCOPES,
        prompt: 'select_account'
      });
      currentAccount = resp.account;
      msalInstance.setActiveAccount(currentAccount);
      return currentAccount;
    } catch (e) {
      // Fallback to redirect if popup blocked
      if (e.errorCode === 'popup_window_error' || e.errorCode === 'empty_window_error') {
        await msalInstance.loginRedirect({ scopes: SCOPES });
      }
      throw e;
    }
  }

  async function signOut() {
    if (!msalInstance) return;
    await msalInstance.logoutPopup({ account: currentAccount });
    currentAccount = null;
  }

  async function getToken() {
    if (!msalInstance) await init();
    if (!currentAccount) throw new Error('Nicht angemeldet');

    try {
      const resp = await msalInstance.acquireTokenSilent({
        scopes: SCOPES,
        account: currentAccount
      });
      return resp.accessToken;
    } catch (e) {
      // Silent acquisition failed -> try interactive
      if (e instanceof msal.InteractionRequiredAuthError) {
        const resp = await msalInstance.acquireTokenPopup({ scopes: SCOPES });
        return resp.accessToken;
      }
      throw e;
    }
  }

  function getAccount() {
    return currentAccount;
  }

  function isAuthenticated() {
    return !!currentAccount;
  }

  return { init, signIn, signOut, getToken, getAccount, isAuthenticated };
})();
