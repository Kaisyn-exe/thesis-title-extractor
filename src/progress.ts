/**
 * 核对进度：老师改过、确认过的内容按文件夹保存在本机（WebView 的 localStorage），
 * 下次打开同一个文件夹时自动恢复。文件大小或修改时间变了（PDF 被替换）就不再套用。
 */

export const EDIT_FIELDS = ["title", "name", "studentId", "advisor", "advisorTitle", "college", "major"] as const;
export type EditField = (typeof EDIT_FIELDS)[number];
export type Fields = Record<EditField, string>;

export interface SavedItem extends Fields {
  size: number;
  mtime: number;
  method: string;
}

/** 以「相对所选文件夹的路径」为键 */
export type SavedItems = Record<string, SavedItem>;

const PREFIX = "progress:v1:";
const keyOf = (folder: string) => PREFIX + folder.replace(/\//g, "\\").replace(/\\+$/, "").toLowerCase();

export function loadProgress(folder: string): SavedItems {
  try {
    const raw = localStorage.getItem(keyOf(folder));
    return raw ? (JSON.parse(raw).items ?? {}) : {};
  } catch {
    return {};
  }
}

/** 保存失败（存储空间满等）返回 false，不抛错 */
export function saveProgress(folder: string, items: SavedItems): boolean {
  try {
    if (Object.keys(items).length) {
      localStorage.setItem(keyOf(folder), JSON.stringify({ savedAt: Date.now(), items }));
    } else {
      localStorage.removeItem(keyOf(folder));
    }
    return true;
  } catch {
    return false;
  }
}

export function pickFields(src: Fields): Fields {
  return Object.fromEntries(EDIT_FIELDS.map((k) => [k, src[k] ?? ""])) as Fields;
}

export function sameFields(a: Fields, b: Fields): boolean {
  return EDIT_FIELDS.every((k) => (a[k] ?? "") === (b[k] ?? ""));
}
