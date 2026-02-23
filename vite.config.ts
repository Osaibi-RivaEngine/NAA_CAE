import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";
import { resolve } from "path";
import * as devCerts from "office-addin-dev-certs";

// https://vitejs.dev/config/
export default defineConfig(async () => {
  // Generate dev HTTPS certs (required by Office Add-ins)
  const httpsOptions = await getHttpsOptions();

  return {
    plugins: [vue()],
    resolve: {
      alias: {
        "@": resolve(__dirname, "src"),
      },
    },
    server: {
      https: httpsOptions,
      port: 3000,
      headers: {
        // Required for Office Add-ins / NAA cross-origin messaging
        "Access-Control-Allow-Origin": "*",
      },
    },
    build: {
      outDir: "dist",
      rollupOptions: {
        input: {
          index: resolve(__dirname, "index.html"),
          taskpane: resolve(__dirname, "taskpane.html"),
          commands: resolve(__dirname, "commands.html"),
        },
      },
    },
  };
});

async function getHttpsOptions() {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return httpsOptions;
  } catch {
    console.warn(
      "Unable to get HTTPS certs. Using default (you may need to run: npx office-addin-dev-certs install)"
    );
    return undefined;
  }
}
