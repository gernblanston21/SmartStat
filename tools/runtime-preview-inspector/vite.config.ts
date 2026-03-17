import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

export default defineConfig({
  plugins: [react()],
  server: {
    fs: {
      // Allow loading frozen committed fixture from repo tests tree.
      allow: [resolve(__dirname, "..", "..", "..")],
    },
  },
});
