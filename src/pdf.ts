/** 浏览器（WebView）端的 pdf.js 配置：worker、中文 CMap、标准字体等资源都从本地加载。 */
import * as pdfjs from "pdfjs-dist";
import type { PDFPageProxy } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

const assets = `${import.meta.env.BASE_URL}pdfjs/`;

export function loadPdf(data: Uint8Array) {
  return pdfjs.getDocument({
    data,
    cMapUrl: `${assets}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${assets}standard_fonts/`,
    wasmUrl: `${assets}wasm/`,
    iccUrl: `${assets}iccs/`,
    verbosity: 0,
  });
}

/** 把一页渲染成白底 PNG。scale=1 约等于 72 DPI。 */
export async function renderPage(page: PDFPageProxy, scale: number): Promise<Blob> {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#fff"; // 透明背景的 PDF 也输出白底
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvas, viewport, background: "#fff" }).promise;
  return new Promise<Blob>((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("渲染失败"))), "image/png"),
  );
}

/** 把第 1 页（封面）渲染成 PNG */
export async function renderCover(data: Uint8Array, scale: number): Promise<Blob> {
  const task = loadPdf(data);
  try {
    const doc = await task.promise;
    return await renderPage(await doc.getPage(1), scale);
  } finally {
    await task.destroy();
  }
}
