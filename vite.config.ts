/// <reference types="vitest/config" />
import { defineConfig } from "vite";
import { viteStaticCopy } from "vite-plugin-static-copy";

const host = process.env.TAURI_DEV_HOST;

// pdf.js 运行时需要的资源（中文 CMap、标准字体、wasm 解码器、ICC 色彩配置）随应用一起打包，离线可用
const pdfjsAssets = ["cmaps", "standard_fonts", "wasm", "iccs"].map((dir) => ({
  src: `node_modules/pdfjs-dist/${dir}/*`,
  dest: `pdfjs/${dir}`,
  rename: { stripBase: true as const },
}));

export default defineConfig({
  plugins: [viteStaticCopy({ targets: pdfjsAssets })],
  clearScreen: false,
  build: { target: "es2022", chunkSizeWarningLimit: 4000 },
  server: {
    port: 1430,
    strictPort: true,
    host: host || false,
    hmr: host ? { protocol: "ws", host, port: 1431 } : undefined,
    watch: { ignored: ["**/src-tauri/**"] },
  },
  test: { include: ["tests/**/*.test.ts"] },
});
