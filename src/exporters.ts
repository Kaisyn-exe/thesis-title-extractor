/** 导出：Excel 汇总表、封面合集 PDF、封面图片。 */
import ExcelJS from "exceljs";
import { PDFDocument, PDFHexString, PDFName, type PDFRef } from "pdf-lib";
import { needsCheck } from "./extractor";
import { basename, joinPath, readFile, writeFile } from "./native";
import { renderCover } from "./pdf";
import type { RosterResult } from "./roster";

export interface PaperRecord {
  path: string;
  file: string; // 相对所选文件夹的路径
  title: string;
  method: string;
  name: string;
  studentId: string;
  advisor: string;
  advisorTitle: string;
  college: string;
  major: string;
  pages: number;
}

type Progress = (done: number, total: number) => void;

/** Excel 可导出的列，顺序即表格中的顺序 */
export const EXCEL_COLUMNS = [
  { key: "no", header: "序号", width: 6 },
  { key: "title", header: "论文题目", width: 56 },
  { key: "name", header: "学生姓名", width: 11 },
  { key: "studentId", header: "学号", width: 14 },
  { key: "college", header: "学院", width: 20 },
  { key: "major", header: "专业年级", width: 26 },
  { key: "advisor", header: "指导教师", width: 11 },
  { key: "advisorTitle", header: "职称", width: 10 },
  { key: "pages", header: "页数", width: 6 },
  { key: "file", header: "文件名", width: 38 },
  { key: "method", header: "识别方式", width: 18 },
] as const;
export type ExcelColumn = (typeof EXCEL_COLUMNS)[number]["key"];
export type SortKey = "file" | "studentId" | "name" | "title";

export interface ExcelOptions {
  columns: ExcelColumn[];
  sort: SortKey;
  /** 需核对的行标黄 */
  highlight: boolean;
  roster?: RosterResult;
}

export const DEFAULT_EXCEL_OPTIONS: ExcelOptions = {
  columns: ["no", "title", "name", "studentId", "college", "major", "advisor", "advisorTitle", "file", "method"],
  sort: "file",
  highlight: true,
};

const zh = new Intl.Collator("zh-CN", { numeric: true });

export function sortRecords<T extends PaperRecord>(records: T[], sort: SortKey): T[] {
  if (sort === "file") return [...records];
  // 空值排到最后
  return [...records].sort((a, b) => {
    const x = a[sort];
    const y = b[sort];
    if (!x !== !y) return x ? -1 : 1;
    return zh.compare(x, y);
  });
}

const fill = (argb: string): ExcelJS.Fill => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

function styleHeader(ws: ExcelJS.Worksheet) {
  const header = ws.getRow(1);
  header.height = 22;
  header.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = fill("FF217346");
    c.alignment = { vertical: "middle", horizontal: "center" };
  });
}

export async function buildExcel(records: PaperRecord[], opts: ExcelOptions = DEFAULT_EXCEL_OPTIONS): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  const ws = wb.addWorksheet("论文题目", { views: [{ state: "frozen", ySplit: 1 }] });
  const cols = EXCEL_COLUMNS.filter((c) => opts.columns.includes(c.key));
  ws.columns = cols.map((c) => ({ header: c.header, key: c.key, width: c.width }));
  styleHeader(ws);
  sortRecords(records, opts.sort).forEach((r, i) => {
    const row = ws.addRow({ ...r, no: i + 1 });
    row.alignment = { vertical: "middle" };
    if (opts.columns.includes("title")) row.getCell("title").alignment = { vertical: "middle", wrapText: true };
    if (opts.columns.includes("studentId")) row.getCell("studentId").numFmt = "@"; // 学号按文本保存，避免变成科学计数法
    if (opts.highlight && needsCheck(r)) row.eachCell((c) => (c.fill = fill("FFFFF2CC")));
  });
  ws.autoFilter = { from: "A1", to: { row: records.length + 1, column: cols.length } };
  if (opts.roster) addRosterSheet(wb, opts.roster);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/** 名单核对结果：未交、信息不一致、不在名单、已交 */
function addRosterSheet(wb: ExcelJS.Workbook, r: RosterResult) {
  const ws = wb.addWorksheet("名单核对", { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = [
    { header: "情况", key: "kind", width: 12 },
    { header: "学号", key: "id", width: 14 },
    { header: "姓名", key: "name", width: 11 },
    { header: "论文题目", key: "title", width: 50 },
    { header: "文件名", key: "file", width: 38 },
    { header: "说明", key: "note", width: 36 },
  ];
  styleHeader(ws);
  const add = (v: Record<string, string>, argb?: string) => {
    const row = ws.addRow(v);
    row.getCell("id").numFmt = "@";
    if (argb) row.eachCell((c) => (c.fill = fill(argb)));
  };
  for (const s of r.missing) add({ kind: "未交", id: s.id, name: s.name }, "FFFDE2E1");
  for (const m of r.mismatched) {
    add({ kind: "信息不一致", id: m.student.id, name: m.student.name, title: m.paper.title, file: m.paper.file, note: m.reason }, "FFFFF2CC");
  }
  for (const p of r.extra) {
    add({ kind: "不在名单", id: p.studentId, name: p.name, title: p.title, file: p.file, note: "名单里找不到这篇论文的学生" }, "FFDDEBF7");
  }
  for (const s of r.submitted) {
    for (const p of s.papers) {
      add({ kind: "已交", id: s.student.id, name: s.student.name, title: p.title, file: p.file, note: s.papers.length > 1 ? `交了 ${s.papers.length} 份` : "" });
    }
  }
  ws.autoFilter = { from: "A1", to: { row: ws.rowCount, column: 6 } };
}

/** 给合集 PDF 加书签：每篇论文一条，点击跳到对应封面。 */
function addOutline(doc: PDFDocument, entries: { title: string; pageIndex: number }[]) {
  if (!entries.length) return;
  const ctx = doc.context;
  const root = ctx.nextRef();
  const refs: PDFRef[] = entries.map(() => ctx.nextRef());
  entries.forEach((e, i) => {
    const item = ctx.obj({
      Title: PDFHexString.fromText(e.title),
      Parent: root,
      Dest: [doc.getPage(e.pageIndex).ref, "Fit"],
      ...(i > 0 ? { Prev: refs[i - 1] } : {}),
      ...(i < refs.length - 1 ? { Next: refs[i + 1] } : {}),
    });
    ctx.assign(refs[i], item);
  });
  ctx.assign(root, ctx.obj({ Type: "Outlines", First: refs[0], Last: refs[refs.length - 1], Count: refs.length }));
  doc.catalog.set(PDFName.of("Outlines"), root);
  doc.catalog.set(PDFName.of("PageMode"), PDFName.of("UseOutlines"));
}

/** 每篇论文的第 1 页合并成一个 PDF。加密或损坏的 PDF 改为嵌入渲染后的图片。 */
export async function exportCoverPdf(records: PaperRecord[], target: string, progress: Progress) {
  const out = await PDFDocument.create();
  out.setTitle("论文封面合集");
  const entries: { title: string; pageIndex: number }[] = [];
  const failed: string[] = [];
  for (const [i, r] of records.entries()) {
    try {
      const bytes = await readFile(r.path);
      try {
        const src = await PDFDocument.load(bytes, { ignoreEncryption: false, updateMetadata: false });
        const [page] = await out.copyPages(src, [0]);
        out.addPage(page);
      } catch {
        const png = await renderCover(bytes, 2);
        const img = await out.embedPng(new Uint8Array(await png.arrayBuffer()));
        const page = out.addPage([img.width / 2, img.height / 2]);
        page.drawImage(img, { x: 0, y: 0, width: img.width / 2, height: img.height / 2 });
      }
      entries.push({ title: `${i + 1}. ${r.title || basename(r.path)}`, pageIndex: out.getPageCount() - 1 });
    } catch {
      failed.push(r.file);
    }
    progress(i + 1, records.length);
  }
  addOutline(out, entries);
  if (out.getPageCount()) await writeFile(target, await out.save());
  return { count: entries.length, failed };
}

const safeName = (s: string) => s.replace(/[\\/:*?"<>|\r\n\t]+/g, "_").replace(/^[\s._]+|[\s._]+$/g, "").slice(0, 80) || "未命名";

/** 每篇论文的第 1 页存为 PNG（约 150 DPI），文件名「序号_原文件名.png」。 */
export async function exportCoverImages(records: PaperRecord[], dir: string, progress: Progress) {
  const failed: string[] = [];
  const width = String(records.length).length < 3 ? 3 : String(records.length).length;
  for (const [i, r] of records.entries()) {
    try {
      const png = await renderCover(await readFile(r.path), 150 / 72);
      const stem = safeName(basename(r.path).replace(/\.pdf$/i, ""));
      const name = `${String(i + 1).padStart(width, "0")}_${stem}.png`;
      await writeFile(joinPath(dir, name), new Uint8Array(await png.arrayBuffer()));
    } catch {
      failed.push(r.file);
    }
    progress(i + 1, records.length);
  }
  return { count: records.length - failed.length, failed };
}
