/** 检查 GitHub 上有没有新版本。只提示和给出下载链接，不自动下载安装。 */

export const REPO = "Kaisyn-exe/thesis-title-extractor";
export const RELEASES_URL = `https://github.com/${REPO}/releases/latest`;

export interface UpdateInfo {
  version: string;
  url: string;
  notes: string;
}

/** 「1.10.0」>「1.9.2」：逐段按数字比较 */
export function isNewer(latest: string, current: string): boolean {
  const a = latest.replace(/^v/i, "").split(".").map((n) => parseInt(n, 10) || 0);
  const b = current.replace(/^v/i, "").split(".").map((n) => parseInt(n, 10) || 0);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if ((a[i] ?? 0) !== (b[i] ?? 0)) return (a[i] ?? 0) > (b[i] ?? 0);
  }
  return false;
}

/** 有新版本返回其信息，没有返回 null；网络不通会抛错（调用方决定是否提示） */
export async function checkForUpdate(current: string): Promise<UpdateInfo | null> {
  const res = await fetch(`https://api.github.com/repos/${REPO}/releases/latest`, {
    headers: { Accept: "application/vnd.github+json" },
    signal: AbortSignal.timeout(10000),
  });
  if (!res.ok) throw new Error(`GitHub 返回 ${res.status}`);
  const rel = (await res.json()) as { tag_name: string; html_url: string; body?: string };
  if (!isNewer(rel.tag_name, current)) return null;
  // 只取更新说明里的条目，去掉下载提示和 Markdown 标记
  const notes = (rel.body ?? "")
    .split("\n")
    .filter((l) => /^\s*-\s/.test(l))
    .map((l) => "· " + l.replace(/^\s*-\s*/, "").replace(/\*\*|`/g, ""))
    .slice(0, 6)
    .join("\n");
  return { version: rel.tag_name.replace(/^v/i, ""), url: rel.html_url, notes };
}
