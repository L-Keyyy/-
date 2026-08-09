import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  root: resolve(projectRoot, "standalone"),
  base: "./",
  plugins: [react()],
  build: {
    outDir: resolve(projectRoot, ".standalone-dist"),
    emptyOutDir: true,
    cssCodeSplit: false,
    target: "es2020",
    minify: "esbuild",
    rolldownOptions: {
      output: {
        codeSplitting: false,
      },
    },
  },
});
