//! The clipboard's "HTML Format" as Word uses it. The web view only hands out the copied
//! fragment, without the style sheet in which Word defines its lists (`@list l0:level1`);
//! reading the whole document keeps bullets, numbers and indents. Writing uses the same
//! envelope, so Word gets RaabeWrite's list definitions and table widths unchanged.

#[cfg(windows)]
use clipboard_win::{raw, register_format, Clipboard};

const START: &str = "<!--StartFragment-->";
const END: &str = "<!--EndFragment-->";

/// The value of a CF_HTML header line ("StartHTML:0000000105").
fn header_value(text: &str, key: &str) -> Option<usize> {
    text.lines().take(12).find_map(|l| l.strip_prefix(key)?.strip_prefix(':')?.trim().parse().ok())
}

/// The HTML document inside a CF_HTML block (from StartHTML to EndHTML).
fn document(data: &[u8]) -> String {
    let head = String::from_utf8_lossy(&data[..data.len().min(400)]).into_owned();
    let start = header_value(&head, "StartHTML").filter(|&s| s < data.len());
    let end = header_value(&head, "EndHTML").map(|e| e.min(data.len()));
    match (start, end) {
        (Some(s), Some(e)) if e > s => String::from_utf8_lossy(&data[s..e]).into_owned(),
        _ => {
            let all = String::from_utf8_lossy(data).into_owned();
            match all.find('<') { Some(i) => all[i..].to_string(), None => String::new() }
        }
    }
}

/// A CF_HTML block for an HTML document (the fragment markers added when missing).
fn envelope(html: &str) -> Vec<u8> {
    let doc = if html.contains(START) && html.contains(END) { html.to_string() } else { format!("<html><body>{START}{html}{END}</body></html>") };
    let template = "Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\nStartFragment:0000000000\r\nEndFragment:0000000000\r\n";
    let n = template.len();
    let sf = n + doc.find(START).map(|i| i + START.len()).unwrap_or(0);
    let ef = n + doc.find(END).unwrap_or(doc.len());
    let header = format!("Version:0.9\r\nStartHTML:{n:010}\r\nEndHTML:{:010}\r\nStartFragment:{sf:010}\r\nEndFragment:{ef:010}\r\n", n + doc.len());
    [header.as_bytes(), doc.as_bytes()].concat()
}

#[cfg(windows)]
#[tauri::command]
pub fn clipboard_read_html() -> Result<String, String> {
    let _open = Clipboard::new_attempts(10).map_err(|e| e.to_string())?;
    let format = register_format("HTML Format").ok_or("HTML Format")?;
    if !raw::is_format_avail(format.get()) {
        return Ok(String::new());
    }
    let mut data = Vec::new();
    raw::get_vec(format.get(), &mut data).map_err(|e| e.to_string())?;
    Ok(document(&data))
}

#[cfg(windows)]
#[tauri::command]
pub fn clipboard_write_html(html: String, text: String) -> Result<(), String> {
    let _open = Clipboard::new_attempts(10).map_err(|e| e.to_string())?;
    raw::set_string(&text).map_err(|e| e.to_string())?;
    let format = register_format("HTML Format").ok_or("HTML Format")?;
    raw::set_without_clear(format.get(), &envelope(&html)).map_err(|e| e.to_string())
}

#[cfg(not(windows))]
#[tauri::command]
pub fn clipboard_read_html() -> Result<String, String> {
    Ok(String::new())
}

#[cfg(not(windows))]
#[tauri::command]
pub fn clipboard_write_html(_html: String, _text: String) -> Result<(), String> {
    Err("not supported".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip() {
        let html = "<html><head><style>@list l0:level1 {}</style></head><body><!--StartFragment--><p>ä</p><!--EndFragment--></body></html>";
        let block = envelope(html);
        let text = String::from_utf8(block.clone()).unwrap();
        let sf = header_value(&text, "StartFragment").unwrap();
        let ef = header_value(&text, "EndFragment").unwrap();
        assert_eq!(&block[sf..ef], "<p>ä</p>".as_bytes());
        assert_eq!(document(&block), html);
    }
}
