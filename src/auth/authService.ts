/**
 * authService.ts
 * ──────────────
 * Singleton wrapper around MSAL's **Nested App Authentication (NAA)**
 * public client.  Provides:
 *
 *  1. Silent & interactive token acquisition.
 *  2. Automatic claims-challenge handling (CAE).
 *  3. Account management helpers.
 *
 * NAA uses `createNestablePublicClientApplication` which lets the
 * add-in delegate auth to the host Office app (Outlook, Word, etc.)
 * instead of opening a popup or redirect.
 */

import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  type IPublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type PopupRequest,
} from "@azure/msal-browser";

import { msalConfig, loginRequest, buildGraphRequest } from "./authConfig";


/* ================================================================== */
/*  Singleton MSAL instance                                           */
/* ================================================================== */

let msalInstance: IPublicClientApplication | null = null;

/**
 * Initialise (or return) the singleton MSAL NAA instance.
 *
 * Must be called **after** `Office.onReady()` because NAA relies on
 * the Office host bridge being available.
 */
export async function getMsalInstance(): Promise<IPublicClientApplication> {
  if (msalInstance) {
    return msalInstance;
  }

  // NAA entry-point – the host Office app acts as the broker
  msalInstance = await createNestablePublicClientApplication(msalConfig);
  return msalInstance;
}

/* ================================================================== */
/*  Account helpers                                                   */
/* ================================================================== */

/**
 * Return the currently-active account, falling back to the first
 * cached account.
 */
export function getActiveAccount(): AccountInfo | null {
  if (!msalInstance) return null;

  const active = msalInstance.getActiveAccount();
  if (active) return active;

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  return null;
}

/* ================================================================== */
/*  Login / Logout                                                    */
/* ================================================================== */

/**
 * Interactive login.  With NAA the host shows the consent/login UI
 * inline — no popups or redirects.
 */
export async function login(): Promise<AuthenticationResult> {
  const pca = await getMsalInstance();
  const result = await pca.acquireTokenPopup(loginRequest);
  console.log(result.accessToken);

  if (result.account) {
    pca.setActiveAccount(result.account);
  }

  return result;
}

/**
 * Clear the local session (there is no "server-side logout" with NAA
 * since the host manages the session).
 */
export async function logout(): Promise<void> {
  const pca = await getMsalInstance();
  const account = getActiveAccount();

  if (account) {
    // Clear the token cache for the account
    await pca.clearCache({ account });
  }
}

/* ================================================================== */
/*  Token Acquisition — with Claims-Challenge support                 */
/* ================================================================== */

/**
 * Acquire an access token for `scopes`, silently if possible.
 *
 * If a **claims challenge** is provided, it is attached to the request
 * so Azure AD can satisfy the Conditional-Access requirement.
 *
 * @param scopes   The OAuth scopes to request.
 * @param claims   Optional claims challenge string (from WWW-Authenticate header).
 */
export async function acquireToken(
  scopes: string[],
  claims?: string
): Promise<AuthenticationResult> {
  const pca = await getMsalInstance();
  const account = getActiveAccount();

  // Build request, injecting claims challenge if provided
  const request: PopupRequest = buildGraphRequest(scopes, claims);

  if (account) {
    request.account = account;
  }

  // LOG: Track claims state and request details
  console.log("[AUTH] acquireToken called", {
    scopes,
    hasClaims: !!claims,
    claimsLength: claims?.length,
    accountId: account?.localAccountId,
    timestamp: new Date().toISOString()
  });

  // When a claims challenge is present, MSAL must skip the cache and
  // go to the network so the new token (satisfying the challenge) is
  // fetched AND cached, replacing the old one.
  if (claims) {
    
    console.log("[AUTH] Claims challenge present", {
      claimsPreview: claims.substring(0, 100)
    });
  }

  try {
    // ① Try silent acquisition first
    console.log("[AUTH] Attempting acquireTokenSilent", {
      hasClaims: !!request.claims
    });

    const result = await pca.acquireTokenSilent(request);

    // LOG: Token details with snippet to verify it changed
    const tokenSnippet = `${result.accessToken.substring(0, 15)}...${result.accessToken.substring(result.accessToken.length - 15)}`;
    console.log("[AUTH] ✓ Token acquired silently", {
      tokenSnippet,
      tokenLength: result.accessToken.length,
      expiresOn: result.expiresOn,
      scopes: result.scopes,
      hadClaims: !!claims,
      accountId: result.account?.localAccountId
    });

    // After a claims-challenged silent refresh, update the active
    // account so the fresh token is properly associated and cached.
    if (result.account) {
      pca.setActiveAccount(result.account);
    }

    return result;
  } catch (error: unknown) {
    // ② If interaction is required (or claims challenge forces it),
    //    fall back to interactive (popup via NAA bridge).
    if (error instanceof InteractionRequiredAuthError) {
      // The error itself may carry a claims string
      console.log("[AUTH] InteractionRequiredAuthError caught", {
        errorMessage: error.message,
        errorCode: error.errorCode,
        hasClaims: !!error.claims
      });

      // If the error contains claims, use them
      if (error.claims) {
        request.claims = error.claims;
        console.log("[AUTH] Using claims from error", {
          claimsPreview: error.claims.substring(0, 100)
        });
      }

      console.log("[AUTH] Attempting acquireTokenPopup (interactive)");
      const result = await pca.acquireTokenPopup(request);

      // LOG: Token details from popup
      const tokenSnippet = `${result.accessToken.substring(0, 15)}...${result.accessToken.substring(result.accessToken.length - 15)}`;
      console.log("[AUTH] ✓ Token acquired via popup", {
        tokenSnippet,
        tokenLength: result.accessToken.length,
        expiresOn: result.expiresOn,
        accountId: result.account?.localAccountId
      });

      if (result.account) {
        pca.setActiveAccount(result.account);
      }

      return result;
    }

    console.error("[AUTH] Token acquisition failed with unexpected error", error);
    throw error;
  }
}

/* ================================================================== */
/*  Graph API call with automatic claims-challenge retry              */
/* ================================================================== */

/**
 * Call the Microsoft Graph API.  If Graph returns a 401 with a claims
 * challenge (CAE), the function:
 *
 *  1. Parses the challenge from the WWW-Authenticate header.
 *  2. Re-acquires a token with the challenge attached.
 *  3. Retries the Graph call **once**.
 *
 * @param endpoint  Graph endpoint (e.g. "/me", "/me/messages").
 * @param scopes    Scopes for the access token.
 * @returns         The parsed JSON response from Graph.
 */
export async function callGraphWithClaimsRetry<T = unknown>(
  endpoint: string,
  scopes: string[]
): Promise<T> {
  const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

  // First attempt - no claims
  let tokenResult = await acquireToken(scopes);

  let response = await fetch(`${GRAPH_BASE}${endpoint}`, {
    headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
  });

  // If 401 with claims challenge → parse and retry once
  if (response.status === 401) {
    const wwwAuth = response.headers.get("WWW-Authenticate");

    if (wwwAuth) {
      console.log("[AUTH] 401 received, parsing claims challenge from WWW-Authenticate header");
      const claims = parseClaimsChallengeFromHeader(wwwAuth);

      if (claims) {
        console.log("[AUTH] Claims challenge found, re-acquiring token with claims");
        // Re-acquire token with the claims challenge
        tokenResult = await acquireToken(scopes, claims);

        response = await fetch(`${GRAPH_BASE}${endpoint}`, {
          headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
        });
      } else {
        console.log("[AUTH] No claims challenge found in WWW-Authenticate header");
      }
    }
  }

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(
      `Graph API error ${response.status}: ${response.statusText}\n${errorBody}`
    );
  }

  return (await response.json()) as T;
}

/* ================================================================== */
/*  Debug utilities                                                   */
/* ================================================================== */

/**
 * Debug utility: Inspect current token cache and claims state.
 * Call from browser console to diagnose caching issues.
 */
export async function inspectTokenCache(): Promise<void> {
  console.log("=".repeat(60));
  console.log("[CACHE INSPECTOR] Token Cache Diagnostic");
  console.log("=".repeat(60));

  if (!msalInstance) {
    console.log("[CACHE] ⚠️  MSAL not initialized");
    return;
  }

  // Check accounts
  const accounts = msalInstance.getAllAccounts();
  console.log(`[CACHE] Accounts in cache: ${accounts.length}`);
  accounts.forEach((acc, idx) => {
    console.log(`[CACHE] Account ${idx}:`, {
      username: acc.username,
      localAccountId: acc.localAccountId,
      environment: acc.environment
    });
  });

  const activeAccount = msalInstance.getActiveAccount();
  console.log("[CACHE] Active account:", activeAccount?.username || "none");

  // Check localStorage for MSAL entries
  const cacheKeys = Object.keys(localStorage).filter(k =>
    k.includes('msal') || k.includes('token')
  );
  console.log(`[CACHE] MSAL entries in localStorage: ${cacheKeys.length}`);
  if (cacheKeys.length > 0 && cacheKeys.length <= 5) {
    cacheKeys.forEach(key => {
      const value = localStorage.getItem(key);
      if (value && value.length < 200) {
        console.log(`  ${key}: ${value.substring(0, 100)}...`);
      } else {
        console.log(`  ${key}: [${value?.length} chars]`);
      }
    });
  }

  console.log("=".repeat(60));
}


/**
 * Attempts to extract the `claims` value from the `WWW-Authenticate`
 * header of an HTTP 401 response.
 *
 * Example header:
 *   Bearer realm="", authorization_uri="…", client_id="…",
 *   error="insufficient_claims",
 *   claims="eyJhY2Nlc3NfdG9rZW4iOnsi…"   ← base-64 encoded JSON
 *
 * @returns The *decoded* claims JSON string, or `undefined` if no
 *          claims directive was found.
 */
function parseClaimsChallengeFromHeader(
  wwwAuthenticateHeader: string
): string | undefined {
  // The header may contain multiple challenges separated by commas
  // that are *not* inside double-quotes.  We look for the `claims=`
  // directive and grab the quoted value.
  const claimsRegex = /claims="([^"]+)"/i;
  const match = claimsRegex.exec(wwwAuthenticateHeader);

  if (!match?.[1]) {
    return undefined;
  }

  try {
    // The value is base-64-encoded JSON → decode it
    return atob(match[1]);
  } catch {
    // If decoding fails, try using the raw value (some providers send
    // the JSON directly without base-64 encoding).
    return match[1];
  }
}