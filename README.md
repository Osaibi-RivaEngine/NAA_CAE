# NAA CAE — Outlook Office.js Add-in

A **Vue 3 + TypeScript** Outlook web add-in that authenticates via **MSAL.js Nested App Authentication (NAA)** and handles **Continuous Access Evaluation (CAE) claims challenges**.

## Architecture

```
NAA_CAE/
├── manifest.json              # Office Unified Manifest (Teams JSON)
├── taskpane.html              # Task-pane entry (loads Vue app)
├── commands.html              # Ribbon command runtime
├── src/
│   ├── main.ts                # Office.onReady → Vue bootstrap
│   ├── App.vue                # App shell (header + router-view)
│   ├── router.ts              # Vue Router routes
│   ├── auth/
│   │   ├── authConfig.ts      # MSAL config (client ID, scopes)
│   │   ├── authService.ts     # NAA init, login, token, Graph calls
│   │   ├── claimsManager.ts   # CAE claims parse / store / clear
│   │   └── index.ts           # Barrel re-exports
│   ├── composables/
│   │   └── useAuth.ts         # Reactive Vue composable for auth
│   ├── views/
│   │   ├── HomeView.vue       # Sign-in / mailbox read
│   │   └── ProfileView.vue    # Graph /me profile with claims retry
│   └── commands/
│       └── commands.ts        # Ribbon function commands
└── public/assets/             # Add-in icons
```

## Key Features

| Feature | Details |
|---|---|
| **Nested App Auth (NAA)** | Uses `createNestablePublicClientApplication` so the Office host (Outlook) brokers auth — no popups or redirects. |
| **Claims Challenge (CAE)** | `claimsManager.ts` parses `WWW-Authenticate` 401 headers, stores the decoded claims, and `authService.ts` retries token acquisition with the challenge. |
| **Silent → Interactive fallback** | `acquireToken()` tries silent first; on `InteractionRequiredAuthError` (including claims), it falls back to interactive. |
| **Graph helper** | `callGraphWithClaimsRetry()` makes a Graph call, detects 401 claims, re-acquires a token, and retries automatically. |

## Prerequisites

- **Node.js** ≥ 18
- **Office desktop or Office on the web** that supports NAA (Outlook ≥ build 16.0.16000, or Outlook on the web)
- An **Azure AD app registration** with:
  - `User.Read` (and optionally `Mail.Read`) delegated permissions
  - **SPA** redirect URI: `brk-multihub://CLIENT ID` (required for NAA)
  - **`crossOriginIsolated`** enabled (Entra portal → Authentication → advanced)

## Quick Start

### 1. Install dependencies

```bash
npm install
```

### 2. Generate HTTPS dev certificates

Office Add-ins require HTTPS. Run once:

```bash
npx office-addin-dev-certs install
```

### 3. Configure your Azure AD app

Open `src/auth/authConfig.ts` and replace the placeholder values:

```ts
export const CLIENT_ID = "<your-app-client-id>";
export const AUTHORITY  = "https://login.microsoftonline.com/<your-tenant-id>";
```

Also update `manifest.json`:
- Set `"webApplicationInfo.id"` to your client ID.
- Set `"webApplicationInfo.resource"` to `api://localhost:3000/<your-client-id>`.
- Replace `"id"` (top-level) with a real GUID.

### 4. Start the dev server

```bash
npm run dev
```

The add-in will be served at `https://localhost:3000`.

### 5. Side-load into Outlook

- **Outlook on the web**: Upload `manifest.json` via the *Integrated Apps* admin center or use [Teams Toolkit](https://learn.microsoft.com/microsoftteams/platform/toolkit/toolkit-v4/teams-toolkit-fundamentals-vs-code-v4).
- **Outlook desktop**: Use `npx office-addin-debugging start manifest.json` or side-load manually.

## Claims Challenge Flow

```
┌──────────────┐    ①  acquireTokenSilent(scopes)
│  Vue App     │───────────────────────────────────────►│  MSAL NAA  │
│  (useAuth)   │◄──── token ─────────────────────────── │            │
│              │                                        └────────────┘
│              │    ②  GET graph.microsoft.com/v1.0/me
│              │───────────────────────────────────────►│  MS Graph  │
│              │◄──── 401 + WWW-Authenticate: claims="…"│            │
│              │                                        └────────────┘
│              │    ③  parse claims from header
│              │    ④  store in sessionStorage
│              │    ⑤  acquireTokenSilent/Popup({ claims })
│              │───────────────────────────────────────►│  MSAL NAA  │
│              │◄──── new token ────────────────────────│            │
│              │                                        └────────────┘
│              │    ⑥  retry GET /me with new token
│              │───────────────────────────────────────►│  MS Graph  │
│              │◄──── 200 OK ──────────────────────────│            │
└──────────────┘                                        └────────────┘
```

## Build for Production

```bash
npm run build
```

Output goes to `dist/`. Deploy the contents to any static host that serves HTTPS.