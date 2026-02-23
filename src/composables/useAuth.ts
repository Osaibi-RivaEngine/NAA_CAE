/**
 * useAuth.ts
 * ──────────
 * Vue composable that exposes reactive authentication state and
 * actions powered by the MSAL NAA auth service.
 */

import { ref, readonly, onMounted } from "vue";
import type { AccountInfo } from "@azure/msal-browser";
import {
  getMsalInstance,
  getActiveAccount,
  login as authLogin,
  logout as authLogout,
  acquireToken,
  callGraphWithClaimsRetry,
  graphScopes,
} from "@/auth";

/* ------------------------------------------------------------------ */
/*  Shared reactive state (singleton across components)               */
/* ------------------------------------------------------------------ */

const isAuthenticated = ref(false);
const account = ref<AccountInfo | null>(null);
const isLoading = ref(false);
const error = ref<string | null>(null);

/**
 * Sync local refs with the MSAL cache.
 */
function syncAccountState(): void {
  const active = getActiveAccount();
  account.value = active;
  isAuthenticated.value = active !== null;
}

/* ------------------------------------------------------------------ */
/*  Composable                                                        */
/* ------------------------------------------------------------------ */

export function useAuth() {
  /* ── Initialise MSAL on first mount ── */
  onMounted(async () => {
    try {
      await getMsalInstance();
      syncAccountState();
    } catch (e: unknown) {
      error.value = (e as Error).message;
    }
  });

  /* ── Actions ── */

  async function login(): Promise<void> {
    isLoading.value = true;
    error.value = null;
    try {
      await authLogin();
      syncAccountState();
    } catch (e: unknown) {
      error.value = (e as Error).message;
    } finally {
      isLoading.value = false;
    }
  }

  async function logout(): Promise<void> {
    isLoading.value = true;
    error.value = null;
    try {
      await authLogout();
      account.value = null;
      isAuthenticated.value = false;
    } catch (e: unknown) {
      error.value = (e as Error).message;
    } finally {
      isLoading.value = false;
    }
  }

  /**
   * Acquire an access token for the given scopes.
   * Claims challenges are handled automatically inside `acquireToken`.
   */
  async function getToken(scopes?: string[]): Promise<string | null> {
    isLoading.value = true;
    error.value = null;
    try {
      const result = await acquireToken(
        scopes ?? graphScopes.userRead,
        "graph"
      );
      syncAccountState();
      return result.accessToken;
    } catch (e: unknown) {
      error.value = (e as Error).message;
      return null;
    } finally {
      isLoading.value = false;
    }
  }

  /**
   * Call Microsoft Graph handling claims challenges transparently.
   */
  async function callGraph<T = unknown>(
    endpoint: string,
    scopes?: string[]
  ): Promise<T | null> {
    isLoading.value = true;
    error.value = null;
    try {
      const data = await callGraphWithClaimsRetry<T>(
        endpoint,
        scopes ?? graphScopes.userRead
      );
      return data;
    } catch (e: unknown) {
      error.value = (e as Error).message;
      return null;
    } finally {
      isLoading.value = false;
    }
  }

  return {
    // State (readonly to consumers)
    isAuthenticated: readonly(isAuthenticated),
    account: readonly(account),
    isLoading: readonly(isLoading),
    error: readonly(error),

    // Actions
    login,
    logout,
    getToken,
    callGraph,
  };
}
