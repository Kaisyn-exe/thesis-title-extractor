/**
 * 从论文 PDF 中识别题目（以及封面上的学生姓名、学号、指导教师）。
 * 只依赖 pdf.js 的文档对象，不依赖界面和 Tauri，可在 Node 中单独测试。
 *
 * 识别顺序（可信度从高到低）：
 *   1. 封面上的「题目 / 论文题目 / Title」栏——按坐标取标签右侧的文字，
 *      能处理表格式封面里题目分两行、第一行高于标签的情况；
 *   2. 页面上字号明显最大的一段文字（期刊论文、英文论文、正文首页）；
 *   3. PDF 文件属性里的标题；
 *   4. 文件名。
 * 前几页没有文字（扫描件）时，可以传入 OCR 函数，把识别出的文字按同样的规则处理。
 */
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import type { TextItem } from "pdfjs-dist/types/src/display/api";

export const Method = {
  Label: "封面题目栏",
  Font: "大字号标题",
  Layout: "封面版式",
  OcrLabel: "OCR·封面题目栏",
  OcrFont: "OCR·大字号标题",
  OcrLayout: "OCR·封面版式",
  Meta: "PDF属性(需核对)",
  FileName: "文件名(需核对)",
  Scan: "扫描件无文字(需核对)",
  Manual: "人工修改",
  Confirmed: "人工确认",
} as const;

// OCR 可能认错个别字，同样请老师对照封面预览确认
const NEEDS_CHECK = new Set<string>([
  Method.OcrLabel, Method.OcrFont, Method.OcrLayout, Method.Meta, Method.FileName, Method.Scan,
]);

export interface ExtractResult {
  title: string;
  method: string;
  name: string;
  studentId: string;
  advisor: string;
  /** 导师职称 */
  advisorTitle: string;
  college: string;
  /** 专业 / 专业年级 / 专业班级 */
  major: string;
  pages: number;
}

export function needsCheck(r: Pick<ExtractResult, "title" | "method">): boolean {
  return !r.title || NEEDS_CHECK.has(r.method) || r.method.startsWith("打开失败");
}

// ---------------------------------------------------------------- 规则

const CJK = "\\u3400-\\u9fff\\uf900-\\ufaff\\u3000-\\u303f\\uff00-\\uffef";
const CJK_CHAR_RE = new RegExp(`[${CJK}]`);
const CJK_GAP_RE = new RegExp(`(?<=[${CJK}])\\s+(?=[${CJK}])`, "g");

const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
/** 「题目」-> 「题\s*目」，允许字间有空格 */
const lab = (word: string) => Array.from(word).map(escape).join("\\s*");
const alt = (words: string[]) =>
  "(?:" + [...words].sort((a, b) => b.length - a.length).map(lab).join("|") + ")";

const TITLE_LABEL =
  alt(["论文题目", "中文题目", "设计题目", "课题名称", "论文名称", "题目", "题名", "标题"]) +
  "(?:\\s*[（(]\\s*中\\s*文\\s*[)）])?|(?:paper\\s+|thesis\\s+)?title";
const NAME_LABEL = alt(["研究生姓名", "学生姓名", "姓名", "作者"]);
const ID_LABEL = alt(["学号"]);
const ADVISOR_LABEL = alt(["指导教师", "指导老师", "导师"]);
const RANK_LABEL = alt(["职称"]);
const COLLEGE_LABEL = alt(["教学学院", "所在学院", "学院名称", "学院", "院系", "系别"]) + "|院\\s*[（(]\\s*系\\s*[)）]";
const MAJOR_LABEL = alt(["专业年级", "专业班级", "年级专业", "专业名称", "专业"]);

/** 封面上常见的栏目名：用来判断「这一行是另一个栏目」或「值到这里结束」 */
const FIELD_ANY = alt([
  "论文题目", "中文题目", "设计题目", "课题名称", "论文名称", "题目", "题名", "标题",
  "教学学院", "所在学院", "学院", "院系", "系别", "专业年级", "专业班级", "专业", "年级", "班级",
  "研究生姓名", "学生姓名", "姓名", "作者", "学号", "指导教师", "指导老师", "导师", "职称",
  "完成日期", "答辩日期", "提交日期", "日期", "英文题目",
]);
const FIELD_END = "(?:\\s*[:：]|\\s+|$)";
const ROW_IS_FIELD_RE = new RegExp(`^\\s*${FIELD_ANY}${FIELD_END}`, "i");
const FIELD_IN_ROW_RE = new RegExp(`(?:^|\\s)${FIELD_ANY}${FIELD_END}`, "i");
const FIELD_AFTER_RE = new RegExp(`\\s${FIELD_ANY}${FIELD_END}`, "i");

/** 整行是这些内容时，不可能是题目 */
const NOISE_RE = new RegExp(
  "^\\s*(?:" +
    [
      ".{0,20}(学位论文|毕业论文|毕业设计|学士论文|硕士论文|博士论文|学年论文|课程论文|结课论文|设计（论文）|设计\\(论文\\))\\s*([（(].*[)）])?",
      ".{0,20}(大学|学院|学校|研究院|研究所|研究生院)",
      "(硕士|博士|学士|本科|专业)(学位)?(研究生)?",
      "(摘\\s*要|abstract|目\\s*录|contents|致\\s*谢|诚信承诺书|承诺书|声\\s*明|原创性声明|关键词|keywords?)(?![a-z]).*",
      ".*(issn|isbn|doi|vol\\.|no\\.|journal|proceedings|学报|期刊|杂志|第\\s*\\d+\\s*[卷期]).*",
      "(分类号|密\\s*级|udc|编\\s*号|学校代码).*",
      "[\\d\\s年月日届.\\-/〇一二三四五六七八九十零]+",
      "(arxiv|preprint|accepted|received|published|copyright|©).*",
    ].join("|") +
    ")\\s*$",
  "i",
);
const BAD_META_RE = /(microsoft|word|\.docx?|\.pdf|untitled|无标题)/i;

export function clean(text: string): string {
  return text
    .replace(/[　 ]/g, " ")
    .replace(/[_＿]{2,}/g, " ") // 封面上的填写下划线
    .replace(CJK_GAP_RE, "") // 中文字间被拉开的空格
    .replace(/\s+/g, " ")
    .replace(/^[\s:：]+|[\s:：]+$/g, "");
}

const plainLen = (s: string) => s.replace(/\s/g, "").length;
const isNoise = (s: string) => NOISE_RE.test(clean(s));

function join(a: string, b: string): string {
  a = a.trimEnd();
  b = b.trimStart();
  if (!a || !b) return a || b;
  if (a.endsWith("-") && /^[a-z]/.test(b)) return a.slice(0, -1) + b;
  if (/[A-Za-z0-9,;:]$/.test(a) && /^[A-Za-z0-9(]/.test(b)) return a + " " + b;
  return a + b;
}

// ---------------------------------------------------------------- 版面

export interface Glyph {
  c: string;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  cy: number;
  size: number;
  spaceBefore: boolean;
}

/** 页面上视觉意义的一行。text 的每个 UTF-16 单元对应 pos 中的一个字（插入的空格对应 null）。 */
export class Row {
  readonly text: string;
  readonly pos: (Glyph | null)[];
  readonly size: number;
  readonly cy: number;
  readonly y0: number;
  readonly y1: number;
  readonly x0: number;

  /** fuzzy：坐标和字号来自 OCR，会有几个像素的误差，比较字号时放宽容差 */
  constructor(glyphs: Glyph[], readonly fuzzy = false) {
    glyphs.sort((a, b) => a.x0 - b.x0);
    this.size = Math.max(...glyphs.map((g) => g.size));
    this.cy = glyphs.reduce((s, g) => s + g.cy, 0) / glyphs.length;
    this.y0 = Math.min(...glyphs.map((g) => g.y0));
    this.y1 = Math.max(...glyphs.map((g) => g.y1));
    this.x0 = glyphs[0].x0;
    let text = "";
    const pos: (Glyph | null)[] = [];
    let prev: Glyph | null = null;
    for (const g of glyphs) {
      if (prev) {
        const gap = g.x0 - prev.x1;
        const cjk = CJK_CHAR_RE.test(g.c) || CJK_CHAR_RE.test(prev.c);
        if (g.spaceBefore || gap > Math.min(g.size, prev.size) * (cjk ? 0.6 : 0.2)) {
          text += " ";
          pos.push(null);
        }
      }
      text += g.c;
      for (let k = 0; k < g.c.length; k++) pos.push(g);
      prev = g;
    }
    this.text = text;
    this.pos = pos;
  }

  /** 横坐标落在 [xFrom, xTo) 内的文字；遇到比 maxGap 更宽的空白就停止（说明已经到了下一个栏目） */
  sliceX(xFrom: number, xTo: number, maxGap = Infinity): string {
    let out = "";
    let prev: Glyph | null = null;
    for (let i = 0; i < this.text.length; i++) {
      const p = this.pos[i];
      if (p === null) out += " ";
      else if (p.x0 >= xFrom - 0.5 && p.x0 < xTo) {
        if (prev && prev !== p && p.x0 - prev.x1 > maxGap) break;
        out += this.text[i];
        prev = p;
      }
    }
    return out.trim();
  }

  /** 第 index 个字符之后（含）第一个真实字的左边界 */
  xAt(index: number): number {
    for (let i = index; i < this.pos.length; i++) {
      const p = this.pos[i];
      if (p) return p.x0;
    }
    return Infinity;
  }
}

const charWeight = (ch: string) => (/\s/.test(ch) ? 0.3 : CJK_CHAR_RE.test(ch) ? 1 : 0.55);

/** pdf.js 给的是一段段文字（run），这里按字宽权重拆成单字，便于按坐标切分 */
async function textGlyphs(page: PDFPageProxy): Promise<Glyph[]> {
  const top = page.view[3];
  const content = await page.getTextContent();
  const glyphs: Glyph[] = [];
  for (const item of content.items) {
    const it = item as TextItem;
    if (!it.str || !it.transform) continue;
    const [a, b, c, d, e, f] = it.transform as number[];
    // 只要正向横排的文字（跳过旋转的水印、竖排文字）
    if (a <= 0 || d <= 0 || Math.abs(b) > a * 0.05 || Math.abs(c) > d * 0.05) continue;
    const size = d;
    const baseline = top - f;
    const chars = Array.from(it.str);
    const total = chars.reduce((s, ch) => s + charWeight(ch), 0) || 1;
    let x = e;
    let space = false;
    for (const ch of chars) {
      const w = (it.width * charWeight(ch)) / total;
      if (/\s/.test(ch)) {
        space = true;
      } else {
        glyphs.push({
          c: ch, x0: x, x1: x + w,
          y0: baseline - size * 0.88, y1: baseline + size * 0.12, cy: baseline - size * 0.38,
          size, spaceBefore: space,
        });
        space = false;
      }
      x += w;
    }
  }
  return glyphs;
}

/** Windows OCR 的输出：按行分组的单词，坐标单位是渲染图片的像素 */
export interface OcrWord { text: string; x: number; y: number; w: number; h: number }

/** 把 OCR 结果换算回 PDF 坐标（除以渲染倍率），并拆成单字 */
export function glyphsFromOcr(lines: OcrWord[][], scale: number): Glyph[] {
  const glyphs: Glyph[] = [];
  for (const line of lines) {
    // 同一行用统一的字号（取最高的字框），避免个别字框高低不一
    const size = Math.max(...line.map((w) => w.h)) / scale;
    let prevRight = -Infinity;
    for (const word of line) {
      const chars = Array.from(word.text).filter((ch) => !/\s/.test(ch));
      const total = chars.reduce((s, ch) => s + charWeight(ch), 0) || 1;
      const y0 = word.y / scale;
      const y1 = (word.y + word.h) / scale;
      let x = word.x / scale;
      chars.forEach((ch, i) => {
        const w = (word.w / scale) * (charWeight(ch) / total);
        // OCR 把英文按单词切分，单词之间本来就有空格
        const spaceBefore = i === 0 && !CJK_CHAR_RE.test(ch) && x > prevRight;
        glyphs.push({ c: ch, x0: x, x1: x + w, y0, y1, cy: (y0 + y1) / 2, size, spaceBefore });
        x += w;
      });
      prevRight = (word.x + word.w) / scale;
    }
  }
  return glyphs;
}

export function rowsFromGlyphs(glyphs: Glyph[], fuzzy = false): Row[] {
  glyphs = [...glyphs].sort((p, q) => p.cy - q.cy);
  const groups: { cy: number; size: number; glyphs: Glyph[] }[] = [];
  for (const g of glyphs) {
    const last = groups[groups.length - 1];
    if (last && Math.abs(g.cy - last.cy) < Math.max(g.size, last.size) * 0.45) {
      last.glyphs.push(g);
      last.size = Math.max(last.size, g.size);
    } else {
      groups.push({ cy: g.cy, size: g.size, glyphs: [g] });
    }
  }
  return groups.map((grp) => new Row(grp.glyphs, fuzzy));
}

// ---------------------------------------------------------------- 识别

/**
 * 在封面上找「标签 + 值」。值可以在标签右侧，也可以在表格单元格里与标签上下错开。
 * 单行栏目（姓名、学号等）遇到超过 2.5 个字宽的空白就截断——OCR 偶尔认不出下一个栏目名，
 * 不截断会把「学号」后面的数字也算进姓名。
 */
function findField(
  rows: Row[], labelPat: string, window = 0.6, multiline = false, midRow = false,
): { value: string; row: Row } | null {
  // midRow：标签也可以出现在一行中间，如「指导教师 张三 职称 讲师」里的「职称」
  const labelRe = midRow
    ? new RegExp(`(?:^|\\s)(?:${labelPat})${FIELD_END}\\s*`, "i")
    : new RegExp(`^\\s*(?:${labelPat})${FIELD_END}\\s*`, "i");
  const fieldRows = rows.filter((r) => ROW_IS_FIELD_RE.test(r.text));
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const m = labelRe.exec(row.text);
    if (!m) continue;
    const valueStart = m.index + m[0].length;
    let k = valueStart - 1;
    while (k >= 0 && row.pos[k] === null) k--;
    const xFrom = row.pos[k]!.x1;
    const after = FIELD_AFTER_RE.exec(row.text.slice(valueStart));
    const xTo = after ? row.xAt(valueStart + after.index) : Infinity;

    // 标签所在行，以及上下各约一个字高内的行（表格单元格垂直居中时，题目第一行会高于标签）
    const parts: string[] = [];
    let last = i;
    rows.forEach((r, j) => {
      const dist = Math.abs(r.cy - row.cy);
      if (dist > row.size * window) return;
      if (j !== i && ROW_IS_FIELD_RE.test(r.text)) return;
      // 夹在两个栏目之间的行，归离它更近的那个栏目
      if (j !== i && fieldRows.some((f) => f !== row && Math.abs(r.cy - f.cy) < dist)) return;
      let stop = xTo;
      const other = j !== i ? FIELD_IN_ROW_RE.exec(r.text) : null;
      if (other) stop = Math.min(stop, r.xAt(other.index));
      const part = r.sliceX(xFrom, stop, multiline ? Infinity : row.size * 2.5);
      if (part) {
        parts.push(part);
        last = Math.max(last, j);
      }
    });
    let value = parts.reduce(join, "");

    // 题目换行续写到下面几行，或者标签单独一行、题目写在下一行
    if (multiline) {
      for (const r of rows.slice(last + 1, last + 4)) {
        const prev = rows[last];
        if (
          r.y0 - prev.y1 > Math.max(r.size, row.size) * 1.8 ||
          ROW_IS_FIELD_RE.test(r.text) ||
          isNoise(r.text) ||
          (value && r.x0 < xFrom - row.size)
        ) break;
        value = join(value, r.text);
        last++;
      }
    }

    value = clean(value);
    if (value) return { value, row };
  }
  return null;
}

/**
 * 没有题目标签时按版式找：表格式封面上，题目就在第一个栏目（教学学院、学生姓名等）正上方。
 * 主要用于 OCR——孤立、间距很大的「题　目」两个字常常被 OCR 漏掉。
 */
function aboveFields(rows: Row[]): string {
  const first = rows.findIndex((r) => ROW_IS_FIELD_RE.test(r.text));
  if (first <= 0) return "";
  const lines: Row[] = [];
  let below = rows[first];
  for (let i = first - 1; i >= 0 && lines.length < 4; i--) {
    const r = rows[i];
    const gap = below.y0 - r.y1;
    const limit = Math.max(r.size, below.size) * (lines.length ? 1.8 : 6);
    if (gap > limit || isNoise(r.text) || plainLen(r.text) < 2) break;
    lines.unshift(r);
    below = r;
  }
  const title = clean(lines.map((r) => r.text).reduce(join, ""));
  return plainLen(title) >= 6 ? title : "";
}

/** 没有题目栏时：取页面上字号明显最大的一段文字 */
function byFont(rows: Row[]): string {
  if (!rows.length) return "";
  const cands = rows.filter(
    (r) => !isNoise(r.text) && !ROW_IS_FIELD_RE.test(r.text) && plainLen(r.text) >= 2,
  );
  if (!cands.length) return "";
  const sizes = rows.map((r) => r.size).sort((p, q) => p - q);
  const body = sizes[Math.floor(sizes.length / 2)]; // 正文的常见字号
  const topSize = Math.max(...cands.map((r) => r.size));
  // OCR 估出的字号有误差，按比例放宽；PDF 文字层的字号是精确的
  const tol = rows[0].fuzzy ? topSize * 0.12 : 0.6;
  if (topSize < body + Math.max(1, tol) && rows.length > 5) return ""; // 没有明显突出的大字

  const start = rows.findIndex((r) => cands.includes(r) && r.size >= topSize - tol);
  let title = rows[start].text;
  let prev = rows[start];
  for (const r of rows.slice(start + 1)) {
    if (
      Math.abs(r.size - topSize) > tol ||
      r.y0 - prev.y1 > topSize * 1.6 ||
      isNoise(r.text) ||
      ROW_IS_FIELD_RE.test(r.text)
    ) break;
    title = join(title, r.text);
    prev = r;
  }
  title = clean(title);
  return plainLen(title) >= 4 ? title : "";
}

/** 从「2210000001+张三+论文题目.pdf」这类文件名里猜学号、姓名、题目 */
export function parseFileName(fileName: string): { title: string; studentId: string; name: string } {
  const stem = fileName.replace(/^.*[\\/]/, "").replace(/\.pdf$/i, "");
  const sid = /(?<!\d)\d{8,13}(?!\d)/.exec(stem)?.[0] ?? "";
  const parts = stem.split(/[_＿+＋\-—\s]+/).map((p) => p.trim()).filter(Boolean);
  // 只有一段（没有分隔符）时，整个文件名更可能是题目而不是姓名
  const name =
    parts.length > 1 && parts.find(
      (p) => /^[一-鿿·]{2,5}$/.test(p) &&!/论文|原文|定稿|终稿|初稿|修改|毕业|设计/.test(p),
    ) || "";
  const words = parts.filter(
    (p) =>
      !/^[\d.]+$/.test(p) && p !== name &&
      !/^.{0,4}(论文|原文|定稿|终稿|初稿|修改稿|最终版|v\d+)$/i.test(p),
  );
  const title = words.length ? words.reduce((a, b) => (b.length > a.length ? b : a)) : stem;
  return { title: clean(title), studentId: sid, name };
}

export interface ExtractOptions {
  /** 最多读取前几页 */
  maxPages?: number;
  /** 对没有文字的页面做 OCR，返回识别出的字；不传则不做 OCR */
  ocr?: (page: PDFPageProxy) => Promise<Glyph[]>;
  /** 最多对前几页做 OCR（封面通常在第 1 页） */
  maxOcrPages?: number;
}

type Detected = Omit<ExtractResult, "pages"> & { page: number };

function detect(pages: Row[][]): Detected {
  const compact = (s: string) => s.replace(/\s/g, "");
  const d: Detected = {
    title: "", method: "", page: -1, name: "", studentId: "", advisor: "", advisorTitle: "", college: "", major: "",
  };
  for (const rows of pages) {
    const name = findField(rows, NAME_LABEL);
    d.name ||= compact(name?.value ?? "");
    d.studentId ||= compact(findField(rows, ID_LABEL)?.value ?? "");
    // 「学号」两个字没认出来时，学号通常和姓名在同一行
    const nameRow = (name?.row.text ?? "").replace(/(?<=\d)\s+(?=\d)/g, ""); // OCR 常把数字逐个分开
    d.studentId ||= /(?<!\d)\d{8,13}(?!\d)/.exec(nameRow)?.[0] ?? "";
    d.advisor ||= compact(findField(rows, ADVISOR_LABEL)?.value ?? "");
    d.advisorTitle ||= compact(findField(rows, RANK_LABEL, 0.6, false, true)?.value ?? "");
    // 学院、专业在表格式封面里可能分两行写，取值范围和题目一样放宽
    d.college ||= compact(findField(rows, COLLEGE_LABEL, 1.6)?.value ?? "");
    d.major ||= compact(findField(rows, MAJOR_LABEL, 1.6)?.value ?? "");
  }
  const steps: [(rows: Row[]) => string, string][] = [
    [(rows) => findField(rows, TITLE_LABEL, 1.6, true)?.value ?? "", Method.Label],
    [byFont, Method.Font],
    [aboveFields, Method.Layout],
  ];
  for (const [find, method] of steps) {
    for (const [i, rows] of pages.entries()) {
      const t = find(rows);
      if (plainLen(t) >= 4 && !NOISE_RE.test(t)) return { ...d, title: t, method, page: i };
    }
  }
  return d;
}

export async function extract(
  doc: PDFDocumentProxy,
  fileName: string,
  { maxPages = 3, ocr, maxOcrPages = 2 }: ExtractOptions = {},
): Promise<ExtractResult> {
  const count = Math.min(maxPages, doc.numPages);
  const pages: Row[][] = [];
  for (let i = 1; i <= count; i++) {
    const page = await doc.getPage(i);
    pages.push(rowsFromGlyphs(await textGlyphs(page)));
    page.cleanup();
  }
  let d = detect(pages);

  // 扫描件：逐页 OCR 没有文字的页面，直到识别出题目
  const ocrPages = new Set<number>();
  if (!d.title && ocr) {
    for (let i = 0; i < Math.min(maxOcrPages, count); i++) {
      if (pages[i].length) continue;
      const page = await doc.getPage(i + 1);
      pages[i] = rowsFromGlyphs(await ocr(page), true);
      page.cleanup();
      ocrPages.add(i);
      d = detect(pages);
      if (d.title) break;
    }
  }
  if (ocrPages.has(d.page)) {
    const ocrMethod: Record<string, string> = {
      [Method.Label]: Method.OcrLabel, [Method.Font]: Method.OcrFont, [Method.Layout]: Method.OcrLayout,
    };
    d.method = ocrMethod[d.method] ?? d.method;
  }

  const { page: _page, ...found } = d;
  const res: ExtractResult = { ...found, pages: doc.numPages };
  const fromName = parseFileName(fileName);
  res.studentId ||= fromName.studentId;
  res.name ||= fromName.name;
  if (!res.title) {
    const info = (await doc.getMetadata().catch(() => null))?.info as { Title?: string } | undefined;
    const meta = clean(info?.Title ?? "");
    if (meta && !BAD_META_RE.test(meta) && plainLen(meta) >= 4) {
      Object.assign(res, { title: meta, method: Method.Meta });
    } else {
      const hasText = pages.some((rows) => rows.length > 0);
      Object.assign(res, { title: fromName.title, method: hasText ? Method.FileName : Method.Scan });
    }
  }
  return res;
}
