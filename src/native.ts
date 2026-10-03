/** 对 Rust 端命令的封装（见 src-tauri/src/lib.rs）。 */
import { invoke } from "@tauri-apps/api/core";

export const collectPdfs = (paths: string[], recursive: boolean) =>
  invoke<string[]>("collect_pdfs", { paths, recursive });

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
