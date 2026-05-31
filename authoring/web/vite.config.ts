import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Dev server proxies /api to the FastAPI backend (python -m authoring.server).
// Production build lands in dist/, which the backend serves as static files.
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      "/api": { target: "http://127.0.0.1:8099", changeOrigin: true },
    },
  },
  build: { outDir: "dist", emptyOutDir: true },
});
