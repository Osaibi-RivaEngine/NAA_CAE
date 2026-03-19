import { createApp } from "vue";
import { createRouter, createMemoryHistory } from "vue-router";
import App from "./App.vue";
import { routes } from "./router";
import { inspectTokenCache } from "./auth";

/* ------------------------------------------------------------------ */
/*  Debug utilities — expose to browser console                       */
/* ------------------------------------------------------------------ */

// Make cache inspector available in browser console for debugging
if (typeof window !== "undefined") {
  (window as any).inspectTokenCache = inspectTokenCache;
}

/* ------------------------------------------------------------------ */
/*  Bootstrap Vue — wait for Office, then mount once                  */
/* ------------------------------------------------------------------ */

let mounted = false;

function mountApp(): void {
  if (mounted) return;
  mounted = true;

  const router = createRouter({
    history: createMemoryHistory(),
    routes,
  });

  const app = createApp(App);
  app.use(router);
  app.mount("#app");
}

// Wait for Office.js, but guarantee the app boots within 3 seconds.
const officeReady =
  typeof Office !== "undefined" && Office.onReady
    ? Office.onReady().then(() => mountApp())
    : Promise.resolve().then(() => mountApp());

const timeout = new Promise<void>((resolve) =>
  setTimeout(() => {
    mountApp();
    resolve();
  }, 3000)
);

Promise.race([officeReady, timeout]);
