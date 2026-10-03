/** 学生名单：读取 Excel / CSV，并与识别出的论文比对，找出未交、不在名单、信息不一致的情况。 */
import ExcelJS from "exceljs";

export interface Student {
  id: string;
  name: string;
}

export interface RosterPaper {
  id: number;
  studentId: string;
  name: string;
  title: string;
  file: string;
}

export interface RosterResult {
  submitted: { student: Student; papers: RosterPaper[] }[];
  missing: Student[];
  /** 论文对上了名单，但学号或姓名不一致 */
  mismatched: { student: Student; paper: RosterPaper; reason: string }[];
  /** 名单里找不到的论文 */
  extra: RosterPaper[];
}

const normId = (s: string) => s.replace(/\D/g, "");
const normName = (s: string) => s.replace(/[\s　·•.]/g, "");

function cellText(v: ExcelJS.CellValue): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if ("richText" in v) return v.richText.map((t) => t.text).join("");
    if ("text" in v) return String(v.text);
    if ("result" in v) return String(v.result ?? "");
    return "";
  }
  return String(v);
}

async function readXlsx(data: Uint8Array): Promise<string[][]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer);
  // 取第一个有数据的工作表
  const ws = wb.worksheets.find((w) => w.actualRowCount > 0);
  if (!ws) return [];
  const rows: string[][] = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const cells: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell, col) => (cells[col - 1] = cellText(cell.value).trim()));
    rows.push(Array.from(cells, (c) => c ?? ""));
  });
  return rows;
}

function readCsv(data: Uint8Array): string[][] {
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(data);
  } catch {
    text = new TextDecoder("gbk").decode(data); // 中文 Windows 上 Excel 另存的 CSV 是 GBK 编码
  }
  text = text.replace(/^﻿/, "");
  const rows: string[][] = [];
  for (const line of text.split(/\r?\n/)) {
    if (!line.trim()) continue;
    const cells: string[] = [];
    let cur = "";
    let quoted = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (quoted) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') quoted = false;
        else cur += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === "," || ch === "\t") { cells.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    cells.push(cur.trim());
    rows.push(cells);
  }
  return rows;
}

/** 找表头里的「学号」「姓名」列；没有表头时按内容猜：大多是长数字的列是学号，大多是 2-4 个汉字的列是姓名 */
function locateColumns(rows: string[][]): { start: number; idCol: number; nameCol: number } {
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const idCol = rows[r].findIndex((c) => /学号|学生编号|学籍号/.test(c));
    const nameCol = rows[r].findIndex((c) => /姓名|名字|学生$/.test(c));
    if (idCol >= 0 || nameCol >= 0) return { start: r + 1, idCol, nameCol };
  }
  const width = Math.max(0, ...rows.map((r) => r.length));
  const score = (test: (s: string) => boolean) =>
    Array.from({ length: width }, (_, c) => rows.filter((r) => test(r[c] ?? "")).length);
  const ids = score((s) => /^\d{6,13}$/.test(s.trim()));
  const names = score((s) => /^[一-鿿·]{2,5}$/.test(normName(s)));
  const pick = (arr: number[]) => (Math.max(...arr) > 0 ? arr.indexOf(Math.max(...arr)) : -1);
  return { start: 0, idCol: pick(ids), nameCol: pick(names) };
}

export async function parseRoster(data: Uint8Array, fileName: string): Promise<Student[]> {
  const rows = /\.xlsx$/i.test(fileName) ? await readXlsx(data) : readCsv(data);
  const { start, idCol, nameCol } = locateColumns(rows);
  if (idCol < 0 && nameCol < 0) throw new Error("没有找到「学号」或「姓名」列");
  const seen = new Set<string>();
  const students: Student[] = [];
  for (const row of rows.slice(start)) {
    const id = idCol >= 0 ? normId(row[idCol] ?? "") : "";
    const name = nameCol >= 0 ? normName(row[nameCol] ?? "") : "";
    if (!id && !name) continue;
    if (/学号|姓名/.test(name)) continue; // 重复出现的表头行
    const key = id || name;
    if (seen.has(key)) continue;
    seen.add(key);
    students.push({ id, name });
  }
  if (!students.length) throw new Error("名单里没有读到学生");
  return students;
}

export function matchRoster(students: Student[], papers: RosterPaper[]): RosterResult {
  const used = new Set<number>();
  const res: RosterResult = { submitted: [], missing: [], mismatched: [], extra: [] };

  for (const s of students) {
    // 依次按学号、姓名、文件名里的学号/姓名找这位学生的论文
    let found = s.id ? papers.filter((p) => normId(p.studentId) === s.id) : [];
    let by: "id" | "name" | "file" = "id";
    if (!found.length && s.name) {
      found = papers.filter((p) => normName(p.name) === s.name);
      by = "name";
    }
    if (!found.length) {
      found = papers.filter((p) => (s.id && p.file.includes(s.id)) || (s.name.length >= 2 && p.file.includes(s.name)));
      by = "file";
    }
    if (!found.length) {
      res.missing.push(s);
      continue;
    }
    found.forEach((p) => used.add(p.id));
    res.submitted.push({ student: s, papers: found });
    for (const p of found) {
      if (by === "id" && s.name && p.name && normName(p.name) !== s.name) {
        res.mismatched.push({ student: s, paper: p, reason: `学号相同，但论文上的姓名是「${p.name}」` });
      } else if (by === "name" && s.id && p.studentId && normId(p.studentId) !== s.id) {
        res.mismatched.push({ student: s, paper: p, reason: `姓名相同，但论文上的学号是「${p.studentId}」` });
      }
    }
  }
  res.extra = papers.filter((p) => !used.has(p.id));
  return res;
}
