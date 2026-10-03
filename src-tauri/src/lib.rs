//! 只提供文件系统相关的薄接口；PDF 解析、题目识别、导出全部在前端 TypeScript 中完成。

mod ocr;

use std::fs;
use std::path::{Path, PathBuf};

use tauri::ipc::{InvokeBody, Request, Response};
use tauri_plugin_opener::OpenerExt;

fn is_pdf(p: &Path) -> bool {
    p.extension()
        .map(|e| e.eq_ignore_ascii_case("pdf"))
        .unwrap_or(false)
}

fn walk(dir: &Path, recursive: bool, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    let mut entries: Vec<_> = fs::read_dir(dir)?.filter_map(|e| e.ok()).map(|e| e.path()).collect();
    entries.sort();
    for p in entries {
        if p.is_dir() {
            if recursive {
                // 子文件夹无权限之类的错误直接跳过
                let _ = walk(&p, recursive, out);
            }
        } else if is_pdf(&p) {
            out.push(p);
        }
    }
    Ok(())
}

/// 把若干文件夹 / 文件展开成 PDF 文件列表（文件夹按需递归）。
#[tauri::command]
fn collect_pdfs(paths: Vec<String>, recursive: bool) -> Result<Vec<String>, String> {
    let mut out = Vec::new();
    for p in paths.iter().map(PathBuf::from) {
        if p.is_dir() {
            walk(&p, recursive, &mut out).map_err(|e| format!("无法读取文件夹 {}：{e}", p.display()))?;
        } else if p.is_file() && is_pdf(&p) {
            out.push(p);
        }
    }
    Ok(out.into_iter().map(|p| p.to_string_lossy().into_owned()).collect())
}

#[tauri::command]
fn is_dir(path: String) -> bool {
    Path::new(&path).is_dir()
}

/// 以二进制形式返回文件内容（前端收到 ArrayBuffer，不经过 JSON）。
#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    fs::read(&path)
        .map(Response::new)
        .map_err(|e| format!("无法读取 {path}：{e}"))
}

/// 请求体是文件内容，目标路径放在 `path` 请求头里（URL 编码）。
#[tauri::command]
fn write_file(request: Request<'_>) -> Result<(), String> {
    let raw = request
        .headers()
        .get("path")
        .and_then(|v| v.to_str().ok())
        .ok_or("缺少 path")?;
    let path = urlencoding::decode(raw).map_err(|e| e.to_string())?.into_owned();
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("请求体必须是二进制".into());
    };
    if let Some(parent) = Path::new(&path).parent() {
        fs::create_dir_all(parent).map_err(|e| format!("无法创建文件夹：{e}"))?;
    }
    fs::write(&path, bytes).map_err(|e| {
        if e.kind() == std::io::ErrorKind::PermissionDenied {
            format!("无法写入 {path}：文件可能正在被其他程序打开，请关闭后重试")
        } else {
            format!("无法写入 {path}：{e}")
        }
    })
}

/// 对一张 PNG 做 OCR（请求体为图片字节）。在后台线程执行，不阻塞界面。
#[tauri::command]
async fn ocr_image(request: Request<'_>) -> Result<ocr::OcrOutput, String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("请求体必须是二进制".into());
    };
    let png = bytes.clone();
    tauri::async_runtime::spawn_blocking(move || ocr::recognize(&png))
        .await
        .map_err(|e| e.to_string())?
}

#[tauri::command]
fn open_path(app: tauri::AppHandle, path: String) -> Result<(), String> {
    app.opener()
        .open_path(path, None::<&str>)
        .map_err(|e| e.to_string())
}

#[tauri::command]
fn reveal_path(app: tauri::AppHandle, path: String) -> Result<(), String> {
    app.opener()
        .reveal_item_in_dir(path)
        .map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            collect_pdfs,
            is_dir,
            read_file,
            write_file,
            ocr_image,
            open_path,
            reveal_path
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
