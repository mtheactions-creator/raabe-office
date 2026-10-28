//! RaabeCalc's native side. Files, AutoRecover, PDF, fonts, the clipboard, the window and
//! single-instance handling come from the shared `raabe-host` crate (D-061); spelling (F7) is
//! RaabeWrite's Hunspell module, included from its crate so both apps check words the same way.

#[path = "../../../raabe-write/src-tauri/src/spell.rs"]
#[allow(dead_code)]
mod spell;

const CONFIG: raabe_host::AppConfig = raabe_host::AppConfig { title: "RaabeCalc", id: "raabe-calc", extensions: &["xlsx", "csv"] };

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let host = raabe_host::handler();
    let own: Box<dyn Fn(tauri::ipc::Invoke<tauri::Wry>) -> bool + Send + Sync> = Box::new(tauri::generate_handler![spell::spell_check, spell::spell_suggest]);
    raabe_host::builder(CONFIG)
        .invoke_handler(move |invoke| if raabe_host::handles(invoke.message.command()) { host(invoke) } else { own(invoke) })
        .run(tauri::generate_context!())
        .expect("error while running RaabeCalc");
}
