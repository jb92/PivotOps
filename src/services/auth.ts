/**
 * PivotOps — Authentication Service
 * Default: PAT (Personal Access Token) — zero setup, works everywhere.
 * Optional: OAuth/MSAL — for users who want auto-refresh tokens.
 */

import { storageGetItem, storageSetItem, storageRemoveItem } from "./storage";

export type AuthMode = "pat" | "oauth";

const PAT_KEY = "pivotops_pat";
const AUTH_MODE_KEY = "pivotops_auth_mode";
const OAUTH_CLIENT_ID_KEY = "pivotops_client_id";
const OAUTH_AUTHORITY_KEY = "pivotops_authority";

// ── PAT Auth (Default) ────────────────────────────────────────────────────

export async function savePat(pat: string): Promise<void> {
  await storageSetItem(PAT_KEY, pat.trim());
  await storageSetItem(AUTH_MODE_KEY, "pat");
}

export async function getPat(): Promise<string | null> {
  return storageGetItem(PAT_KEY);
}

export async function clearPat(): Promise<void> {
  await storageRemoveItem(PAT_KEY);
}

// ── OAuth Config ─────────────────────────────────────────────

export async function saveOAuthConfig(clientId: string, tenantId?: string): Promise<void> {
  await storageSetItem(OAUTH_CLIENT_ID_KEY, clientId);
  const authority = tenantId
    ? `https://login.microsoftonline.com/${tenantId}`
    : "https://login.microsoftonline.com/common";
  await storageSetItem(OAUTH_AUTHORITY_KEY, authority);
  await storageSetItem(AUTH_MODE_KEY, "oauth");
}

export async function getOAuthConfig(): Promise<{ clientId: string; authority: string } | null> {
  const clientId = await storageGetItem(OAUTH_CLIENT_ID_KEY);
  if (!clientId) return null;
  return {
    clientId,
    authority: (await storageGetItem(OAUTH_AUTHORITY_KEY)) || "https://login.microsoftonline.com/common",
  };
}

// ── Auth Mode ──────────────────────────────────────────────────────────────

export async function getAuthMode(): Promise<AuthMode> {
  return ((await storageGetItem(AUTH_MODE_KEY)) as AuthMode) || "pat";
}

export async function hasAuthConfig(): Promise<boolean> {
  const mode = await getAuthMode();
  if (mode === "pat") return !!(await getPat());
  return !!(await storageGetItem(OAUTH_CLIENT_ID_KEY));
}

// ── OAuth (lazy-loaded to avoid bundling MSAL for PAT users) ───────────────

let msalModule: typeof import("@azure/msal-browser") | null = null;
let msalInstance: InstanceType<typeof import("@azure/msal-browser").PublicClientApplication> | null = null;

const ADO_SCOPES = ["499b84ac-1321-427f-aa17-267ca6975798/.default"];

async function getMsalInstance() {
  if (msalInstance) return msalInstance;

  const config = await getOAuthConfig();
  if (!config) throw new Error("OAuth not configured. Set Client ID in Settings.");

  if (!msalModule) {
    msalModule = await import("@azure/msal-browser");
  }

  msalInstance = new msalModule.PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: config.authority,
      redirectUri: window.location.origin + "/taskpane.html",
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  });
  await msalInstance.initialize();
  return msalInstance;
}

// ── Unified API ────────────────────────────────────────────────────────────

export async function signIn(): Promise<{ displayName: string }> {
  const mode = await getAuthMode();

  if (mode === "pat") {
    const pat = await getPat();
    if (!pat) throw new Error("No PAT configured. Please add your Personal Access Token in Settings.");
    return { displayName: "PAT User" };
  }

  // OAuth flow
  const instance = await getMsalInstance();
  const accounts = instance.getAllAccounts();
  if (accounts.length > 0) {
    return { displayName: accounts[0].username };
  }

  const result = await instance.loginPopup({ scopes: ADO_SCOPES, prompt: "select_account" });
  if (!result.account) throw new Error("Authentication succeeded but no account was returned.");
  return { displayName: result.account.username };
}

export async function signOut(): Promise<void> {
  const mode = await getAuthMode();
  if (mode === "pat") {
    await clearPat();
    return;
  }

  const instance = await getMsalInstance();
  const accounts = instance.getAllAccounts();
  if (accounts.length > 0) {
    await instance.logoutPopup({ account: accounts[0] });
  }
  msalInstance = null;
}

/**
 * Returns the Authorization header value for ADO API calls.
 * PAT → "Basic base64(:token)"
 * OAuth → "Bearer <access_token>"
 */
export async function getAuthHeader(): Promise<string> {
  const mode = await getAuthMode();

  if (mode === "pat") {
    const pat = await getPat();
    if (!pat) throw new Error("No PAT configured. Add it in Settings.");
    return `Basic ${btoa(`:${pat}`)}`;
  }

  // OAuth — acquire token silently or via popup
  const instance = await getMsalInstance();
  const accounts = instance.getAllAccounts();
  if (accounts.length === 0) throw new Error("Not signed in. Please sign in first.");

  try {
    const result = await instance.acquireTokenSilent({ scopes: ADO_SCOPES, account: accounts[0] });
    return `Bearer ${result.accessToken}`;
  } catch (error) {
    if (!msalModule) msalModule = await import("@azure/msal-browser");
    if (error instanceof msalModule.InteractionRequiredAuthError) {
      const result = await instance.acquireTokenPopup({ scopes: ADO_SCOPES, account: accounts[0] });
      return `Bearer ${result.accessToken}`;
    }
    throw error;
  }
}

export async function getAccount(): Promise<{ displayName: string } | null> {
  const mode = await getAuthMode();

  if (mode === "pat") {
    return (await getPat()) ? { displayName: "PAT User" } : null;
  }

  try {
    const instance = await getMsalInstance();
    const accounts = instance.getAllAccounts();
    return accounts.length > 0 ? { displayName: accounts[0].username } : null;
  } catch {
    return null;
  }
}

export async function isSignedIn(): Promise<boolean> {
  const mode = await getAuthMode();
  if (mode === "pat") return !!(await getPat());
  const config = await getOAuthConfig();
  if (!config) return false;
  const key = `msal.${config.clientId}.active-account`;
  return !!(await storageGetItem(key));
}
