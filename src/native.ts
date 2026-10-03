/** 对 Rust 端命令的封装（见 src-tauri/src/lib.rs）。 */
import { invoke } from "@tauri-apps/api/core";

export interface PdfEntry {
  path: string;
  size: number;
  /** 修改时间（毫秒） */
  mtime: number;
}

export const collectPdfs = (paths: string[], recursive: boolean) =>
  invoke<PdfEntry[]>("collect_pdfs", { paths, recursive });

/** 重命名；目标已存在时 Rust 端会拒绝，不会覆盖 */
export const renameFile = (from: string, to: string) => invoke<void>("rename_file", { from, to });

export const isDir = (path: string) => invoke<boolean>("is_dir", { path });

export const readFile = async (path: string) =>
  new Uint8Array(await invoke<ArrayBuffer>("read_file", { path }));

export const writeFile = (path: string, data: Uint8Array) =>
  invoke<void>("write_file", data, { headers: { path: encodeURIComponent(path) } });

export const openPath = (path: string) => invoke<void>("open_path", { path });

export const revealPath = (path: string) => invoke<void>("reveal_path", { path });

export const basename = (p: string) => p.replace(/[\\/]+$/, "").replace(/^.*[\\/]/, "");
export const dirname = (p: string) => p.replace(/[\\/][^\\/]*$/, "");
export const joinPath = (dir: string, name: string) =>
  dir.replace(/[\\/]+$/, "") + (dir.includes("/") && !dir.includes("\\") ? "/" : "\\") + name;
