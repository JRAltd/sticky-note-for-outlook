import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";
import { resolve } from "path";

export default defineConfig({
  base: "./",
  plugins: [react(), basicSsl()],
  server: {
    port: 3000,
    https: true,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
        commands: resolve(__dirname, "commands.html"),
      },
    },
  },
});
