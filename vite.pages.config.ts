// GitHub Pages 배포용 정적 빌드 설정 (Cloudflare 빌드는 vite.config.ts)
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  root: "static",
  base: process.env.PAGES_BASE ?? "/voc-report/",
  publicDir: "../public",
  plugins: [react()],
  build: { outDir: "../dist-pages", emptyOutDir: true },
});
