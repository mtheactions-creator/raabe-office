use tauri::{WebviewUrl, WebviewWindowBuilder};

/// Opens one prototype: `raabe-feel <a|b> [fixture]`, e.g. `raabe-feel b long`.
#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .setup(|app| {
            let args: Vec<String> = std::env::args().skip(1).collect();
            let engine = if args.first().map(String::as_str) == Some("a") { "a" } else { "b" };
            let fixture = args.get(1).cloned().unwrap_or_else(|| "short".into());
            let fixture_js = js_string_literal(&fixture);
            let mut builder = WebviewWindowBuilder::new(app, "main", WebviewUrl::App(format!("{engine}.html").into()))
                .title(format!("Feel-Prototyp {}", engine.to_uppercase()))
                .inner_size(1400.0, 1000.0)
                .initialization_script(format!("window.__FIXTURE = {fixture_js};"));
            // The feel-lab harness reads the page over the DevTools protocol.
            if let Ok(port) = std::env::var("RAABE_FEEL_DEBUG_PORT") {
                if let Ok(port) = port.parse::<u16>() {
                    // Own data folder per port, otherwise instances share one browser process
                    // and later ones silently ignore their flags.
                    builder = builder
                        .data_directory(std::env::temp_dir().join(format!("raabe-feel-{port}")))
                        .additional_browser_args(&format!(
                            "--disable-features=msWebOOUI,msPdfOOUI,msSmartScreenProtection --remote-debugging-port={port}"
                        ));
                }
            }
            builder.build()?;
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running prototype");
}

/// Fixture names are plain identifiers; anything else is dropped.
fn js_string_literal(s: &str) -> String {
    let escaped: String = s.chars().filter(|c| c.is_ascii_alphanumeric() || *c == '-').collect();
    format!("\"{escaped}\"")
}
