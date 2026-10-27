import { defineConfig } from "vite";
import { resolve } from "node:path";

const repo = resolve(__dirname, "../..");

export default defineConfig({
  clearScreen: false,
  server: { port: 5175, strictPort: true, fs: { allow: [repo] } },
  build: { target: "es2022", outDir: "dist", emptyOutDir: true },
});
