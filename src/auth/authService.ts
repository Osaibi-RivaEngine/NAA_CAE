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
import {
  getStoredClaimsChallenge,
  storeClaimsChallenge,
  clearClaimsChallenge,
  handleClaimsChallengeFromResponse,
} from "./claimsManager";

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
 * If a **claims challenge** is stored for `resource`, it is
 * automatically attached to the request so Azure AD can satisfy the
 * Conditional-Access requirement.
 *
 * @param scopes   The OAuth scopes to request.
 * @param resource A key identifying the target API (e.g. "graph").
 *                 Used to look up previously stored claims challenges.
 */
export async function acquireToken(
  scopes: string[],
  resource = "graph"
): Promise<AuthenticationResult> {
  const pca = await getMsalInstance();
  const account = getActiveAccount();

  // Build request, injecting claims challenge if one was stored
  const storedClaims = getStoredClaimsChallenge(resource);
  const request: PopupRequest = buildGraphRequest(scopes, storedClaims);

  if (account) {
    request.account = account;
  }

  // When a claims challenge is present, MSAL must skip the cache and
  // go to the network so the new token (satisfying the challenge) is
  // fetched AND cached, replacing the old one.
  if (storedClaims) {
    request.forceRefresh = true;
  }

  try {
    // ① Try silent acquisition first
    const result = await pca.acquireTokenSilent(request);

    // After a claims-challenged silent refresh, update the active
    // account so the fresh token is properly associated and cached.
    if (result.account) {
      pca.setActiveAccount(result.account);
    }

    // Clear claims after successful acquisition
    if (storedClaims) {
      clearClaimsChallenge(resource);
    }

    return result;
  } catch (error: unknown) {
    // ② If interaction is required (or claims challenge forces it),
    //    fall back to interactive (popup via NAA bridge).
    if (error instanceof InteractionRequiredAuthError) {
      // The error itself may carry a claims string
      if (error.claims) {
        request.claims = error.claims;
        storeClaimsChallenge(resource, error.claims);
      }

      const result = await pca.acquireTokenPopup(request);

      if (result.account) {
        pca.setActiveAccount(result.account);
      }

      // Clear claims after interactive success
      clearClaimsChallenge(resource);
      return result;
    }

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
 *  1. Parses and stores the challenge.
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
  const RESOURCE = "graph";
  const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

  // First attempt
  let tokenResult = await acquireToken(scopes, RESOURCE);

  let response = await fetch(`${GRAPH_BASE}${endpoint}`, {
    headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
  });

  // If 401 with claims challenge → handle and retry once
  if (response.status === 401) {
    const claims = handleClaimsChallengeFromResponse(response, RESOURCE);

    if (claims) {
      // Re-acquire token including the claims challenge
      tokenResult = await acquireToken(scopes, RESOURCE);

      response = await fetch(`${GRAPH_BASE}${endpoint}`, {
        headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
      });
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
