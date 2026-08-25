/**
 * Gorelo Ticket - Auth Helper
 * Handles MSAL B2C authentication against Gorelo's Azure AD B2C tenant.
 *
 * Gorelo uses Azure AD B2C with a custom policy. The token acquired here
 * is the same Bearer token their web app uses against gw.usw.gorelo.tech.
 */

const GORELO_MSAL_CONFIG = {
  auth: {
    clientId: '6a628336-23e6-4bcf-b6b2-c16af729a828',  // Gorelo's app client ID
    authority: 'https://login.gorelo.io/5860777d-7a0c-4a21-926b-91762e9a662c/b2c_1a_gorelo_susi_totp_dev',
    knownAuthorities: ['login.gorelo.io'],
    redirectUri: window.location.origin + '/src/auth/auth-redirect.html',
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

// The scope that grants access to gw.usw.gorelo.tech
const GORELO_SCOPES = [
  'https://gorelo.onmicrosoft.com/gorelo-api-geteway/gorelo-api-gateway.read',
  'openid',
  'profile',
  'offline_access',
];

let _msalInstance = null;
let _tokenCache = null;
let _tokenExpiry = null;

/**
 * Initialise MSAL. Called once on task pane load.
 */
export async function initAuth() {
  // Dynamically load MSAL from CDN
  if (!window.msal) {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/microsoft-authentication-library/3.11.0/msal-browser.min.js');
  }

  _msalInstance = new window.msal.PublicClientApplication(GORELO_MSAL_CONFIG);
  await _msalInstance.initialize();

  // Handle redirect response if returning from login
  try {
    const response = await _msalInstance.handleRedirectPromise();
    if (response) {
      _tokenCache = response.accessToken;
      _tokenExpiry = response.expiresOn;
    }
  } catch (e) {
    console.warn('Redirect handling error:', e);
  }

  return _msalInstance;
}

/**
 * Get a valid access token, acquiring silently or via popup as needed.
 * Returns the Bearer token string.
 */
export async function getToken() {
  // Return cached token if still valid (with 5 min buffer)
  if (_tokenCache && _tokenExpiry && new Date() < new Date(_tokenExpiry.getTime() - 300000)) {
    return _tokenCache;
  }

  if (!_msalInstance) {
    await initAuth();
  }

  const accounts = _msalInstance.getAllAccounts();

  // Try silent acquisition first
  if (accounts.length > 0) {
    try {
      const response = await _msalInstance.acquireTokenSilent({
        scopes: GORELO_SCOPES,
        account: accounts[0],
      });
      _tokenCache = response.accessToken;
      _tokenExpiry = response.expiresOn;
      return _tokenCache;
    } catch (silentError) {
      console.warn('Silent token acquisition failed, trying popup:', silentError);
    }
  }

  // Fall back to popup login
  try {
    const response = await _msalInstance.acquireTokenPopup({
      scopes: GORELO_SCOPES,
    });
    _tokenCache = response.accessToken;
    _tokenExpiry = response.expiresOn;
    return _tokenCache;
  } catch (popupError) {
    throw new Error('Authentication failed: ' + popupError.message);
  }
}

/**
 * Get the current logged-in account info.
 */
export function getCurrentAccount() {
  if (!_msalInstance) return null;
  const accounts = _msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : null;
}

/**
 * Sign out of Gorelo.
 */
export async function signOut() {
  if (!_msalInstance) return;
  const account = getCurrentAccount();
  if (account) {
    await _msalInstance.logoutPopup({ account });
  }
  _tokenCache = null;
  _tokenExpiry = null;
}

/**
 * Check if the user is currently authenticated.
 */
export function isAuthenticated() {
  if (!_msalInstance) return false;
  return _msalInstance.getAllAccounts().length > 0;
}

/**
 * Utility: dynamically load a script tag.
 */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}
