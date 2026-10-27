//! RaabeWrite's native side: spelling, thesaurus and grammar. Files, AutoRecover, PDF, fonts,
//! the clipboard, the window and single-instance handling come from the shared `raabe-host`
//! crate (D-061).

mod grammar;
mod spell;
mod thesaurus;

const CONFIG: raabe_host::AppConfig = raabe_host::AppConfig { title: "RaabeWrite", id: "raabe-write", extensions: &["docx"] };

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    spell::preload();
    grammar::preload();
    let host = raabe_host::handler();
    let own: Box<dyn Fn(tauri::ipc::Invoke<tauri::Wry>) -> bool + Send + Sync> = Box::new(tauri::generate_handler![spell::spell_check, spell::spell_suggest, thesaurus::thesaurus_lookup, grammar::grammar_check]);
    raabe_host::builder(CONFIG)
        .invoke_handler(move |invoke| if raabe_host::handles(invoke.message.command()) { host(invoke) } else { own(invoke) })
        .run(tauri::generate_context!())
        .expect("error while running RaabeWrite");
}
