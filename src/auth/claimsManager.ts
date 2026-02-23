/**
 * claimsManager.ts
 * ────────────────
 * Utilities for handling Continuous Access Evaluation (CAE) claims
 * challenges returned by Microsoft Graph (or other APIs protected by
 * Azure AD Conditional-Access policies).
 *
 * A "claims challenge" is a base-64-encoded JSON payload returned in
 * the WWW-Authenticate header of a 401 response.  MSAL must be given
 * the decoded JSON string via the `claims` property on the next token
 * request so Azure AD can satisfy the challenge.
 *
 * References:
 *  - https://learn.microsoft.com/entra/identity-platform/claims-challenge
 *  - https://learn.microsoft.com/entra/msal/dotnet/advanced/exceptions/claims-challenge
 */

/* ------------------------------------------------------------------ */
/*  Session-scoped claims store (per-resource)                        */
/* ------------------------------------------------------------------ */

const STORAGE_KEY_PREFIX = "naa_cae_claims_";

/**
 * Persist the decoded claims JSON string for a given resource so it
 * survives task-pane reloads inside the same session.
 */
export function storeClaimsChallenge(resource: string, claims: string): void {
  sessionStorage.setItem(`${STORAGE_KEY_PREFIX}${resource}`, claims);
}

/**
 * Retrieve a previously stored claims challenge for a resource.
 * Returns `undefined` when no challenge is stored.
 */
export function getStoredClaimsChallenge(
  resource: string
): string | undefined {
  return sessionStorage.getItem(`${STORAGE_KEY_PREFIX}${resource}`) ?? undefined;
}

/**
 * Clear the stored claims challenge for a resource (e.g. after a
 * successful token acquisition with the challenge).
 */
export function clearClaimsChallenge(resource: string): void {
  sessionStorage.removeItem(`${STORAGE_KEY_PREFIX}${resource}`);
}

/* ------------------------------------------------------------------ */
/*  Parsing helpers                                                   */
/* ------------------------------------------------------------------ */

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
export function parseClaimsChallengeFromHeader(
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

/**
 * Convenience: inspect a `Response` object for claims challenges.
 * Stores the challenge if found and returns it.
 */
export function handleClaimsChallengeFromResponse(
  response: Response,
  resource: string
): string | undefined {
  if (response.status !== 401) {
    return undefined;
  }

  const wwwAuth = response.headers.get("WWW-Authenticate");
  if (!wwwAuth) {
    return undefined;
  }

  const claims = parseClaimsChallengeFromHeader(wwwAuth);
  if (claims) {
    storeClaimsChallenge(resource, claims);
  }

  return claims;
}
