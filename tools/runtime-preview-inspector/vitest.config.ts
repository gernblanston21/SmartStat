import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

export default defineConfig({
  plugins: [react()],
  test: {
    environment: "node",
    include: ["src/tests/**/*.test.ts", "src/tests/**/*.test.tsx"],
  },
  server: {
    fs: {
      // Allow importing the frozen fixture from repo tests/ for local scaffold rendering.
      allow: [resolve(__dirname, "..", "..", "..")],
    },
  },
});
