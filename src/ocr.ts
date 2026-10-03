/** 扫描件 OCR：页面渲染成图片后交给 Windows 自带的 OCR 引擎（见 src-tauri/src/ocr.rs）。 */
import { invoke } from "@tauri-apps/api/core";
import type { PDFPageProxy } from "pdfjs-dist";
import { glyphsFromOcr, type Glyph, type OcrWord } from "./extractor";
import { renderPage } from "./pdf";

export interface OcrOutput {
  language: string;
  lines: OcrWord[][];
}

/** 约 250 DPI；长边不超过 3500 像素（Windows OCR 上限 10000，再大只会更慢） */
function ocrScale(page: PDFPageProxy) {
  const [x0, y0, x1, y1] = page.view;
  return Math.min(250 / 72, 3500 / Math.max(x1 - x0, y1 - y0));
}

export async function ocrPage(page: PDFPageProxy, onLanguage?: (lang: string) => void): Promise<Glyph[]> {
  const scale = ocrScale(page);
  const png = new Uint8Array(await (await renderPage(page, scale)).arrayBuffer());
  const out = await invoke<OcrOutput>("ocr_image", png);
  onLanguage?.(out.language);
  return glyphsFromOcr(out.lines, scale);
}
