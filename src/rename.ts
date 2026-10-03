/** 批量重命名：按「{学号}_{姓名}_{题目}」这样的模板生成新文件名，并在执行前找出冲突。 */

export interface RenamePaper {
  id: number;
  path: string;
  title: string;
  name: string;
  studentId: string;
  advisor: string;
  college: string;
  major: string;
}

export const TOKENS: Record<string, { label: string; get: (p: RenamePaper, no: string) => string }> = {
  "{学号}": { label: "学号", get: (p) => p.studentId },
  "{姓名}": { label: "姓名", get: (p) => p.name },
  "{题目}": { label: "题目", get: (p) => p.title },
  "{导师}": { label: "导师", get: (p) => p.advisor },
  "{学院}": { label: "学院", get: (p) => p.college },
  "{专业}": { label: "专业", get: (p) => p.major },
  "{序号}": { label: "序号", get: (_p, no) => no },
};

export const DEFAULT_PATTERN = "{学号}_{姓名}_{题目}";
const MAX_STEM = 120; // 留出路径余量，避免超过 Windows 260 字符的限制

export type PlanStatus = "ok" | "same" | "missing" | "conflict" | "skip";

export interface RenamePlan {
  paper: RenamePaper;
  from: string;
  to: string;
  newName: string;
  status: PlanStatus;
  note: string;
}

const dirOf = (p: string) => p.replace(/[\\/][^\\/]*$/, "");
const baseOf = (p: string) => p.replace(/^.*[\\/]/, "");
const sep = (p: string) => (p.includes("\\") ? "\\" : "/");

/** 去掉 Windows 文件名里的非法字符，合并多余的分隔符 */
export function sanitize(stem: string): string {
  let s = stem
    .replace(/[\\/:*?"<>|\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/([_\-+\s])\1+/g, "$1") // 某一项为空时会出现「__」；题目里的「——」等标点不动
    .replace(/^[\s._\-+]+|[\s._\-+]+$/g, "");
  if (s.length > MAX_STEM) s = s.slice(0, MAX_STEM).replace(/[\s._\-+]+$/, "");
  return s;
}

export function buildStem(pattern: string, p: RenamePaper, no: string): { stem: string; missing: string[] } {
  const missing: string[] = [];
  const raw = pattern.replace(/\{[^}]+\}/g, (tok) => {
    const t = TOKENS[tok];
    if (!t) return tok;
    const v = t.get(p, no).trim();
    if (!v) missing.push(t.label);
    return v;
  });
  return { stem: sanitize(raw), missing };
}

export function planRenames(papers: RenamePaper[], pattern: string, skip: (p: RenamePaper) => boolean): RenamePlan[] {
  const width = Math.max(2, String(papers.length).length);
  const plans: RenamePlan[] = papers.map((p, i) => {
    const from = p.path;
    const base = { paper: p, from, to: from, newName: baseOf(from) };
    if (skip(p)) return { ...base, status: "skip", note: "还需要核对，暂不改名" };
    const { stem, missing } = buildStem(pattern, p, String(i + 1).padStart(width, "0"));
    if (!stem) return { ...base, status: "missing", note: "模板里的信息都是空的" };
    const newName = `${stem}.pdf`;
    const to = dirOf(from) + sep(from) + newName;
    if (to === from) return { ...base, status: "same", note: "文件名已经符合模板" };
    const note = missing.length ? `缺少${missing.join("、")}` : "";
    return { paper: p, from, to, newName, status: missing.length ? "missing" : "ok", note };
  });

  // 冲突：同一批里改成相同的名字，或者目标名字已经被另一篇（不改名的）论文占用
  const lower = (s: string) => s.toLowerCase();
  const targets = new Map<string, number>();
  for (const pl of plans) if (pl.status === "ok") targets.set(lower(pl.to), (targets.get(lower(pl.to)) ?? 0) + 1);
  const staying = new Set(plans.filter((pl) => pl.status !== "ok").map((pl) => lower(pl.from)));
  for (const pl of plans) {
    if (pl.status !== "ok") continue;
    if ((targets.get(lower(pl.to)) ?? 0) > 1) Object.assign(pl, { status: "conflict", note: "和其他论文改成了同一个名字" });
    else if (staying.has(lower(pl.to)) && lower(pl.to) !== lower(pl.from)) {
      Object.assign(pl, { status: "conflict", note: "已有同名文件" });
    }
  }
  return plans;
}
