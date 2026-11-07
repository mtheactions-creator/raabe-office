//! Freihandgleichung (D-079): Windows' own handwriting recognition for mathematics, the Math Input Panel
//! (mip.exe, part of Windows), is started for the user; its "Einfügen" puts the recognised equation on the
//! clipboard as MathML, which RaabePresent reads ("MathML" or "MathML Presentation") and inserts.

#[cfg(windows)]
fn mip_path() -> Option<std::path::PathBuf> {
    let base = std::env::var_os("CommonProgramFiles")?;
    let p = std::path::Path::new(&base).join("microsoft shared").join("ink").join("mip.exe");
    p.exists().then_some(p)
}

/// Whether the Math Input Panel is there (Windows with its ink components).
#[tauri::command]
pub fn math_input_available() -> bool {
    #[cfg(windows)]
    {
        mip_path().is_some()
    }
    #[cfg(not(windows))]
    {
        false
    }
}

/// Starts the Math Input Panel.
#[tauri::command]
pub fn math_input_start() -> Result<(), String> {
    #[cfg(windows)]
    {
        let p = mip_path().ok_or("Mathematik-Eingabebereich nicht vorhanden")?;
        std::process::Command::new(p).spawn().map(|_| ()).map_err(|e| e.to_string())
    }
    #[cfg(not(windows))]
    {
        Err("not supported".into())
    }
}

/// The MathML on the clipboard (the Math Input Panel's formats), or "".
#[tauri::command]
pub fn clipboard_read_mathml() -> Result<String, String> {
    #[cfg(windows)]
    {
        use clipboard_win::{raw, register_format, Clipboard};
        let _open = Clipboard::new_attempts(10).map_err(|e| e.to_string())?;
        for name in ["MathML Presentation", "MathML", "application/mathml+xml"] {
            let Some(format) = register_format(name) else { continue };
            if !raw::is_format_avail(format.get()) {
                continue;
            }
            let mut data = Vec::new();
            raw::get_vec(format.get(), &mut data).map_err(|e| e.to_string())?;
            // UTF-8 or UTF-16 text, possibly ending in NULs.
            let text = if data.len() > 1 && data[1] == 0 {
                let units: Vec<u16> = data.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
                String::from_utf16_lossy(&units)
            } else {
                String::from_utf8_lossy(&data).into_owned()
            };
            return Ok(text.trim_end_matches('\0').to_string());
        }
        Ok(String::new())
    }
    #[cfg(not(windows))]
    {
        Ok(String::new())
    }
}
