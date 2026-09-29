import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  server: { host: true },
  build: { emptyOutDir: false },
  plugins: [react()],
});
