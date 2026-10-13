//! RaabeWrite's native side: file access (atomic saves), AutoRecover files, PDF export
//! through WebView2, and single-instance handling so a double-clicked .docx opens as a new
//! tab in the running window. Everything else lives in the web frontend.

mod fonts;
mod spell;

use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use serde::{Deserialize, Serialize};
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::window::Color;
use tauri::{AppHandle, Emitter, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder};

struct StartupFiles(Mutex<Vec<String>>);

fn header(request: &Request, name: &str) -> Result<String, String> {
    let raw = request.headers().get(name).ok_or_else(|| format!("missing header {name}"))?;
    let text = raw.to_str().map_err(|e| e.to_string())?;
    Ok(percent_encoding::percent_decode_str(text).decode_utf8_lossy().into_owned())
}

fn body<'a>(request: &'a Request<'_>) -> Result<&'a [u8], String> {
    match request.body() {
        InvokeBody::Raw(data) => Ok(data),
        _ => Err("expected binary data".into()),
    }
}

/// Writes to a temporary file next to the target and renames it over the target, so a
/// crash or full disk never leaves a half-written document behind.
fn atomic_write(path: &Path, data: &[u8]) -> Result<(), String> {
    let tmp = path.with_extension("raabe-tmp");
    fs::write(&tmp, data).map_err(|e| e.to_string())?;
    if let Err(e) = fs::rename(&tmp, path) {
        let _ = fs::remove_file(&tmp);
        return Err(e.to_string());
    }
    Ok(())
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    fs::read(&path).map(Response::new).map_err(|e| e.to_string())
}

#[tauri::command]
fn write_file(request: Request) -> Result<(), String> {
    let path = header(&request, "path")?;
    atomic_write(Path::new(&path), body(&request)?)
}

#[derive(Serialize)]
struct FileInfo {
    exists: bool,
    modified: Option<u128>,
}

#[tauri::command]
fn file_info(path: String) -> FileInfo {
    match fs::metadata(&path) {
        Ok(meta) => FileInfo {
            exists: meta.is_file(),
            modified: meta
                .modified()
                .ok()
                .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                .map(|d| d.as_millis()),
        },
        Err(_) => FileInfo { exists: false, modified: None },
    }
}

// ----- AutoRecover -------------------------------------------------------------------------

#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RecoveryEntry {
    id: String,
    title: String,
    path: Option<String>,
    saved_at: u64,
}

fn recovery_dir(app: &AppHandle) -> Result<PathBuf, String> {
    // feel-lab runs use their own folder so tests never touch the user's real backups.
    let dir = if std::env::var("RAABE_FEEL_DEBUG_PORT").is_ok() {
        std::env::temp_dir().join("raabe-write-feel-recovery")
    } else {
        app.path().app_local_data_dir().map_err(|e| e.to_string())?.join("recovery")
    };
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

/// Recovery ids are UUIDs; anything else could escape the recovery folder.
fn check_id(id: &str) -> Result<(), String> {
    if !id.is_empty() && id.len() <= 64 && id.chars().all(|c| c.is_ascii_hexdigit() || c == '-') {
        Ok(())
    } else {
        Err("invalid recovery id".into())
    }
}

#[tauri::command]
fn save_recovery(app: AppHandle, request: Request) -> Result<(), String> {
    let meta = header(&request, "meta")?;
    let entry: RecoveryEntry = serde_json::from_str(&meta).map_err(|e| e.to_string())?;
    check_id(&entry.id)?;
    let dir = recovery_dir(&app)?;
    atomic_write(&dir.join(format!("{}.docx", entry.id)), body(&request)?)?;
    atomic_write(&dir.join(format!("{}.json", entry.id)), meta.as_bytes())
}

#[tauri::command]
fn list_recovery(app: AppHandle) -> Result<Vec<RecoveryEntry>, String> {
    let dir = recovery_dir(&app)?;
    let mut out = Vec::new();
    for item in fs::read_dir(&dir).map_err(|e| e.to_string())?.flatten() {
        let p = item.path();
        if p.extension().is_some_and(|e| e == "json") {
            if let Ok(entry) = fs::read_to_string(&p).map_err(|_| ()).and_then(|s| serde_json::from_str::<RecoveryEntry>(&s).map_err(|_| ())) {
                if dir.join(format!("{}.docx", entry.id)).exists() {
                    out.push(entry);
                }
            }
        }
    }
    out.sort_by_key(|e| e.saved_at);
    Ok(out)
}

#[tauri::command]
fn read_recovery(app: AppHandle, id: String) -> Result<Response, String> {
    check_id(&id)?;
    fs::read(recovery_dir(&app)?.join(format!("{id}.docx"))).map(Response::new).map_err(|e| e.to_string())
}

#[tauri::command]
fn delete_recovery(app: AppHandle, id: String) -> Result<(), String> {
    check_id(&id)?;
    let dir = recovery_dir(&app)?;
    let _ = fs::remove_file(dir.join(format!("{id}.docx")));
    let _ = fs::remove_file(dir.join(format!("{id}.json")));
    Ok(())
}

// ----- opening files from Explorer -----------------------------------------------------------

fn docx_args(args: impl IntoIterator<Item = String>, cwd: Option<&str>) -> Vec<String> {
    args.into_iter()
        .filter(|a| a.to_lowercase().ends_with(".docx"))
        .map(|a| {
            let p = PathBuf::from(&a);
            match (p.is_absolute(), cwd) {
                (false, Some(dir)) => Path::new(dir).join(p).to_string_lossy().into_owned(),
                _ => a,
            }
        })
        .collect()
}

#[tauri::command]
fn startup_files(state: tauri::State<StartupFiles>) -> Vec<String> {
    std::mem::take(&mut *state.0.lock().unwrap())
}

// ----- PDF export ---------------------------------------------------------------------------

#[tauri::command]
async fn print_to_pdf(window: WebviewWindow, path: String) -> Result<(), String> {
    #[cfg(windows)]
    {
        print_to_pdf_windows(window, path).await
    }
    #[cfg(not(windows))]
    {
        let _ = (window, path);
        Err("PDF-Export ist auf diesem System noch nicht verfügbar.".into())
    }
}

#[cfg(windows)]
async fn print_to_pdf_windows(window: WebviewWindow, path: String) -> Result<(), String> {
    use webview2_com::Microsoft::Web::WebView2::Win32::{
        COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT, ICoreWebView2_7, ICoreWebView2Environment6,
    };
    use webview2_com::PrintToPdfCompletedHandler;
    use windows::core::{HSTRING, Interface};

    let (tx, rx) = std::sync::mpsc::channel::<Result<(), String>>();
    window
        .with_webview(move |wv| {
            let start = || -> windows::core::Result<()> {
                unsafe {
                    let core: ICoreWebView2_7 = wv.controller().CoreWebView2()?.cast()?;
                    let env: ICoreWebView2Environment6 = wv.environment().cast()?;
                    let settings = env.CreatePrintSettings()?;
                    // A4 without the browser's own header/footer and margins; the pages already
                    // contain Word's margins.
                    settings.SetOrientation(COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT)?;
                    settings.SetPageWidth(8.27)?;
                    settings.SetPageHeight(11.69)?;
                    settings.SetMarginTop(0.0)?;
                    settings.SetMarginBottom(0.0)?;
                    settings.SetMarginLeft(0.0)?;
                    settings.SetMarginRight(0.0)?;
                    settings.SetShouldPrintHeaderAndFooter(false)?;
                    settings.SetShouldPrintBackgrounds(true)?;
                    let tx2 = tx.clone();
                    let handler = PrintToPdfCompletedHandler::create(Box::new(move |result, ok| {
                        let _ = tx2.send(match (result, ok) {
                            (Ok(()), true) => Ok(()),
                            (Err(e), _) => Err(e.message().to_string()),
                            _ => Err("Die PDF-Datei konnte nicht geschrieben werden.".into()),
                        });
                        Ok(())
                    }));
                    core.PrintToPdf(&HSTRING::from(path.as_str()), &settings, &handler)
                }
            };
            if let Err(e) = start() {
                let _ = tx.send(Err(e.message().to_string()));
            }
        })
        .map_err(|e| e.to_string())?;
    tauri::async_runtime::spawn_blocking(move || rx.recv().map_err(|e| e.to_string())?)
        .await
        .map_err(|e| e.to_string())?
}

// ----- app ------------------------------------------------------------------------------------

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    fonts::preload();
    spell::preload();
    tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(|app, argv, cwd| {
            let files = docx_args(argv.into_iter().skip(1), Some(&cwd));
            if let Some(w) = app.get_webview_window("main") {
                let _ = w.unminimize();
                let _ = w.set_focus();
            }
            if !files.is_empty() {
                let _ = app.emit("open-files", files);
            }
        }))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .manage(StartupFiles(Mutex::new(docx_args(args.clone(), None))))
        .setup(move |app| {
            let mut window = WebviewWindowBuilder::new(app, "main", WebviewUrl::default())
                .title("RaabeWrite")
                .inner_size(1400.0, 950.0)
                .min_inner_size(640.0, 480.0)
                .center()
                .decorations(false)
                .shadow(true)
                .background_color(Color(240, 240, 240, 255))
                .disable_drag_drop_handler();
            // feel-lab runs (see docs/FEEL.md): with RAABE_FEEL_DEBUG_PORT set, the DevTools
            // protocol is opened on that port and `--feel <fixture>` loads a test document.
            if let Ok(port) = std::env::var("RAABE_FEEL_DEBUG_PORT").map(|p| p.parse::<u16>()) {
                if let Ok(port) = port {
                    // Occlusion tracking would pause rendering whenever the window is covered.
                    // RAABE_FEEL_BROWSER_ARGS adds experimental flags (latency investigations).
                    let extra = std::env::var("RAABE_FEEL_BROWSER_ARGS").unwrap_or_default();
                    window = window
                        .data_directory(std::env::temp_dir().join(format!("raabe-write-feel-{port}")))
                        .additional_browser_args(&format!(
                            "--disable-features=msWebOOUI,msPdfOOUI,msSmartScreenProtection,CalculateNativeWinOcclusion --remote-debugging-port={port} {extra}"
                        ));
                    // Runs that need no real input stay off screen and never take the focus, so
                    // they can run while someone works at the PC.
                    if std::env::var("RAABE_FEEL_BACKGROUND").is_ok() {
                        window = window.position(-4000.0, 0.0).focused(false).skip_taskbar(true);
                    }
                    if let Some(i) = args.iter().position(|a| a == "--feel") {
                        let fixture: String = args.get(i + 1).cloned().unwrap_or_else(|| "short".into())
                            .chars().filter(|c| c.is_ascii_alphanumeric()).collect();
                        window = window.initialization_script(&format!("window.__FEEL = {fixture:?};"));
                    }
                }
            }
            window.build()?;
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            read_file, write_file, file_info, save_recovery, list_recovery, read_recovery, delete_recovery,
            startup_files, print_to_pdf, fonts::list_fonts, fonts::font_metrics, spell::spell_check, spell::spell_suggest
        ])
        .run(tauri::generate_context!())
        .expect("error while running RaabeWrite");
}
