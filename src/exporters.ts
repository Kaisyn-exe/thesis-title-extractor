/** 导出：Excel 汇总表、封面合集 PDF、封面图片。 */
import ExcelJS from "exceljs";
import { PDFDocument, PDFHexString, PDFName, type PDFRef } from "pdf-lib";
import { needsCheck } from "./extractor";
import { basename, joinPath, readFile, writeFile } from "./native";
import { renderCover } from "./pdf";

export interface PaperRecord {
  path: string;
  file: string; // 相对所选文件夹的路径
  title: string;
  method: string;
  name: string;
  studentId: string;
  advisor: string;
  pages: number;
}

type Progress = (done: number, total: number) => void;

export async function buildExcel(records: PaperRecord[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  const ws = wb.addWorksheet("论文题目", { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = [
    { header: "序号", key: "no", width: 6 },
    { header: "论文题目", key: "title", width: 60 },
    { header: "学生姓名", key: "name", width: 11 },
    { header: "学号", key: "studentId", width: 14 },
    { header: "指导教师", key: "advisor", width: 11 },
    { header: "页数", key: "pages", width: 6 },
    { header: "文件名", key: "file", width: 40 },
    { header: "识别方式", key: "method", width: 20 },
  ];
  const header = ws.getRow(1);
  header.height = 22;
  header.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B6FD4" } };
    c.alignment = { vertical: "middle", horizontal: "center" };
  });
  records.forEach((r, i) => {
    const row = ws.addRow({ ...r, no: i + 1, studentId: r.studentId });
    row.alignment = { vertical: "middle" };
    row.getCell("title").alignment = { vertical: "middle", wrapText: true };
    row.getCell("studentId").numFmt = "@"; // 学号按文本保存，避免变成科学计数法
    if (needsCheck(r)) {
      row.eachCell((c) => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
      });
    }
  });
  ws.autoFilter = { from: "A1", to: { row: records.length + 1, column: ws.columns.length } };
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
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
