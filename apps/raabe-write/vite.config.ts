import { defineConfig, type Plugin } from "vite";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join, resolve } from "node:path";

const repo = resolve(__dirname, "../..");

/**
 * The equation editor (MathLive) loads its fonts at run time: they are served by the dev
 * server and copied into the build, so formulas work offline.
 */
function mathliveFonts(): Plugin {
  const dir = join(repo, "node_modules/mathlive/fonts");
  return {
    name: "mathlive-fonts",
    configureServer(server) {
      server.middlewares.use("/mathlive-fonts", (req, res, next) => {
        const name = decodeURIComponent((req.url ?? "").split("?")[0]).replace(/^\/+/, "");
        const file = join(dir, name);
        if (!name || name.includes("..") || !existsSync(file)) return next();
        res.setHeader("Content-Type", "font/woff2");
        res.end(readFileSync(file));
      });
    },
    generateBundle() {
      for (const f of readdirSync(dir)) this.emitFile({ type: "asset", fileName: `mathlive-fonts/${f}`, source: readFileSync(join(dir, f)) });
    },
  };
}

export default defineConfig({
  clearScreen: false,
  plugins: [mathliveFonts()],
  server: { port: 5174, strictPort: true, fs: { allow: [repo] } },
  build: { target: "es2022", outDir: "dist", emptyOutDir: true },
});
