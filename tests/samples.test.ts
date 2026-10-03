/**
 * 用本地真实论文检验识别效果（不会提交到仓库）：
 *   SAMPLES_DIR="D:\论文" npx vitest run tests/samples.test.ts
 */
import { readdirSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import { describe, it } from "vitest";
import { extract, needsCheck } from "../src/extractor";

const dir = process.env.SAMPLES_DIR;
const pdfjsRoot = dirname(createRequire(import.meta.url).resolve("pdfjs-dist/package.json"));

describe.skipIf(!dir)("本地样本", () => {
  it("逐个识别并打印结果", async () => {
    const files = readdirSync(dir!).filter((f) => f.toLowerCase().endsWith(".pdf"));
    for (const f of files) {
      const task = getDocument({
        data: new Uint8Array(readFileSync(join(dir!, f))),
        cMapUrl: join(pdfjsRoot, "cmaps") + "/",
        cMapPacked: true,
        standardFontDataUrl: join(pdfjsRoot, "standard_fonts") + "/",
        verbosity: 0,
      });
      const r = await extract(await task.promise, f);
      await task.destroy();
      console.log(`${needsCheck(r) ? "⚠" : "✓"} ${f}\n   ${JSON.stringify(r)}`);
    }
  }, 120_000);
});
