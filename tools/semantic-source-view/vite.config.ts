import { readFileSync } from "node:fs";
import path from "node:path";
import react from "@vitejs/plugin-react";
import { defineConfig, type Plugin } from "vite";

const SEMANTIC_INDEX_PATH = path.resolve(
  __dirname,
  "..",
  "..",
  ".tools",
  "onair_dump",
  "index",
  "semantic_index.json"
);

function semanticIndexPlugin(): Plugin {
  const route = "/api/semantic-index";

  const handler = (req: any, res: any, next: () => void): void => {
    if (!req.url || !req.url.startsWith(route)) {
      next();
      return;
    }

    try {
      const payload = readFileSync(SEMANTIC_INDEX_PATH, "utf-8");
      res.statusCode = 200;
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.end(payload);
    } catch (error) {
      res.statusCode = 500;
      res.setHeader("Content-Type", "application/json; charset=utf-8");
      res.end(
        JSON.stringify({
          error: "Failed to load semantic index",
          path: SEMANTIC_INDEX_PATH,
          detail: String(error),
        })
      );
    }
  };

  return {
    name: "semantic-index-endpoint",
    configureServer(server) {
      server.middlewares.use(handler);
    },
    configurePreviewServer(server) {
      server.middlewares.use(handler);
    },
  };
}

export default defineConfig({
  plugins: [react(), semanticIndexPlugin()],
});
