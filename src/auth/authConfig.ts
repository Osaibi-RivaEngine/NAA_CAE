/**
 * authConfig.ts
 * ─────────────
 * MSAL configuration for Nested App Authentication (NAA) in an
 * Office Add-in.  Update the clientId and (optionally) tenantId
 * before deploying.
 *
 * NAA docs: https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
 */

import type { Configuration, PopupRequest } from "@azure/msal-browser";

/* ------------------------------------------------------------------ */
/*  Azure AD app registration values – CHANGE THESE                   */
/* ------------------------------------------------------------------ */

/**
 * The Application (client) ID from your Azure AD app registration.
 */
export const CLIENT_ID = "CLIENT ID";

/**
 * Authority URL. Use "common" for multi-tenant, or replace with your
 * tenant ID / domain for single-tenant.
 */
export const AUTHORITY = "https://login.microsoftonline.com/common";

/* ------------------------------------------------------------------ */
/*  MSAL Configuration                                                */
/* ------------------------------------------------------------------ */

export const msalConfig: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    authority: AUTHORITY,
    supportsNestedAppAuth: true, // <<< Enable Nested App Auth (NAA)
    clientCapabilities: ["CP1"],  // <<< Opt-in to CAE claims challenges
  },
  cache: {
    cacheLocation: "localStorage", // Recommended for Office add-ins
  },
  // Uncomment for verbose logging during development:
  // system: {
  //   loggerOptions: {
  //     logLevel: LogLevel.Verbose,
  //     loggerCallback: (_level, message) => console.log(message),
  //   },
  // },
};

/* ------------------------------------------------------------------ */
/*  Scopes & Request Objects                                          */
/* ------------------------------------------------------------------ */

/** Default login request – requests an ID token only. */
export const loginRequest: PopupRequest = {
  scopes: ["User.Read"],
};

/** Microsoft Graph scopes used by the app. */
export const graphScopes = {
  userRead: ["User.Read"],
  mailRead: ["Mail.Read"],
};

/**
 * Build a Graph API request object, optionally injecting a claims
 * challenge string (base-64-decoded JSON).
 */
export function buildGraphRequest(
  scopes: string[],
  claimsChallenge?: string
): PopupRequest {
  const request: PopupRequest = { scopes };

  if (claimsChallenge) {
    request.claims = claimsChallenge;
  }

  return request;
}
