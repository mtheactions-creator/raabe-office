//! RaabeCalc's native side. Files, AutoRecover, PDF, fonts, the clipboard, the window and
//! single-instance handling come from the shared `raabe-host` crate (D-061).

const CONFIG: raabe_host::AppConfig = raabe_host::AppConfig { title: "RaabeCalc", id: "raabe-calc", extensions: &["xlsx", "csv"] };

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    raabe_host::builder(CONFIG)
        .invoke_handler(raabe_host::handler())
        .run(tauri::generate_context!())
        .expect("error while running RaabeCalc");
}
