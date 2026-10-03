import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import { describe, expect, it } from "vitest";
import { extract, glyphsFromOcr, Method, needsCheck, parseFileName, type ExtractOptions } from "../src/extractor";

const pdfjsRoot = dirname(createRequire(import.meta.url).resolve("pdfjs-dist/package.json"));
const fixtures = join(dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, "$1")), "fixtures");

async function run(file: string, options?: ExtractOptions) {
  const data = new Uint8Array(readFileSync(join(fixtures, file)));
  const task = getDocument({
    data,
    cMapUrl: join(pdfjsRoot, "cmaps") + "/",
    cMapPacked: true,
    standardFontDataUrl: join(pdfjsRoot, "standard_fonts") + "/",
    verbosity: 0,
  });
  const doc = await task.promise;
  try {
    return await extract(doc, file, options);
  } finally {
    await task.destroy();
  }
}

describe("extract", () => {
  it("表格式封面：题目第一行高于「题目」标签，姓名与学号同一行", async () => {
    const r = await run("2210000001+王五+论文原文.pdf");
    expect(r).toMatchObject({
      title: "人工智能对制造业企业劳动力结构的影响研究",
      method: Method.Label,
      name: "王五",
      studentId: "2210000001",
      advisor: "赵老师",
    });
  });

  it("「题目：」标签 + 副标题换行", async () => {
    const r = await run("毕业论文_张三.pdf");
    expect(r).toMatchObject({
      title: "数字普惠金融对农村居民消费的影响研究——基于省级面板数据的实证分析",
      method: Method.Label,
      name: "张三",
      studentId: "20220101",
    });
  });

  it("硕士论文封面：无题目标签，取最大字号并合并两行", async () => {
    const r = await run("scan_0002.pdf");
    expect(r).toMatchObject({
      title: "大学生创业意愿的影响因素研究：以长三角高校为例",
      method: Method.Font,
      name: "李四",
      advisor: "王教授",
    });
  });

  it("英文期刊论文：跳过页眉，保留标题里的冒号", async () => {
    const r = await run("paper3.pdf");
    expect(r.title).toBe("Green Credit Policy and Corporate Innovation: Evidence from Chinese Listed Firms");
    expect(r.method).toBe(Method.Font);
  });

  it("中文期刊论文：跳过刊名、卷期页眉", async () => {
    const r = await run("文献4.pdf");
    expect(r.title).toBe("乡村振兴背景下农村电商发展路径研究");
  });

  it("没有文字的扫描件：退回文件名并标记需核对", async () => {
    const r = await run("企业ESG表现与融资约束_孙八.pdf");
    expect(r).toMatchObject({ title: "企业ESG表现与融资约束", method: Method.Scan, name: "孙八" });
    expect(needsCheck(r)).toBe(true);
  });
});

describe("扫描件 OCR", () => {
  // 扫描件样例.ocr.json 是 Windows OCR 对「扫描件样例.pdf」的真实输出：
  // 「题　目」两个字被漏掉，「学号」被认成了「子」
  const ocr = JSON.parse(readFileSync(join(fixtures, "扫描件样例.ocr.json"), "utf8"));
  const fakeOcr = async () => glyphsFromOcr(ocr.lines, ocr.scale);

  it("漏掉题目标签时按封面版式找到题目，并截断姓名", async () => {
    const r = await run("扫描件样例.pdf", { ocr: fakeOcr });
    expect(r).toMatchObject({
      title: "新质生产力赋能乡村产业振兴的路径研究",
      method: Method.OcrLayout,
      name: "钱七",
      studentId: "2210000003",
      advisor: "孙老师",
    });
    expect(needsCheck(r)).toBe(true);
  });

  it("不开 OCR 时退回文件名", async () => {
    const r = await run("扫描件样例.pdf");
    expect(r.method).toBe(Method.Scan);
  });

  it("有文字层的页面不会调用 OCR", async () => {
    let called = 0;
    await run("毕业论文_张三.pdf", { ocr: async () => (called++, []) });
    expect(called).toBe(0);
  });
});

describe("parseFileName", () => {
  it("学号+姓名+论文原文", () => {
    expect(parseFileName("C:\\x\\2210000009+周九+论文原文.pdf")).toMatchObject({
      studentId: "2210000009",
      name: "周九",
    });
  });
  it("题目_姓名", () => {
    expect(parseFileName("乡村振兴研究_李雷.pdf")).toEqual({
      title: "乡村振兴研究",
      studentId: "",
      name: "李雷",
    });
  });
});
