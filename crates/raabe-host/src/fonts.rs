//! Installed fonts and their vertical metrics. Word computes line heights from the font's
//! hhea and OS/2 tables (see Fonts.lineHeight in the text engine; measured against Word's PDF
//! output for twelve fonts on 2026-10-13), so the engine needs these exact values, which a
//! browser does not expose.

use std::collections::HashMap;
use std::sync::OnceLock;

use serde::Serialize;

static DB: OnceLock<fontdb::Database> = OnceLock::new();

fn db() -> &'static fontdb::Database {
    DB.get_or_init(|| {
        let mut db = fontdb::Database::new();
        db.load_system_fonts();
        db
    })
}

static CLOUD: OnceLock<fontdb::Database> = OnceLock::new();

/// Office's cloud fonts (Aptos, Aptos Narrow, Aptos Display …), which are not Windows fonts: Office keeps them per user
/// in its font cache. They are this user's own fonts, used here only (D-089). A database of their own, because the
/// Display faces call themselves "Aptos": a face counts for its cache folder's name only.
fn cloud_db() -> &'static fontdb::Database {
    CLOUD.get_or_init(|| {
        let mut db = fontdb::Database::new();
        if let Some(dir) = cloud_dir() {
            db.load_fonts_dir(dir);
        }
        db
    })
}

/// The cache folder (the family Office files a cloud face under) of a face.
fn cloud_family(face: &fontdb::FaceInfo) -> Option<String> {
    let dir = cloud_dir()?;
    match &face.source {
        fontdb::Source::File(p) => {
            let rel = p.strip_prefix(&dir).ok()?;
            if rel.components().count() < 2 { return None; }
            rel.components().next().map(|c| c.as_os_str().to_string_lossy().into_owned())
        }
        _ => None,
    }
}

/// The regular face of a cloud family (by its folder), else its nearest to regular.
fn find_cloud_face(name: &str) -> Option<fontdb::ID> {
    let db = cloud_db();
    db.faces()
        .filter(|f| cloud_family(f).is_some_and(|n| n.eq_ignore_ascii_case(name)))
        .min_by_key(|f| (f.style != fontdb::Style::Normal, (f.weight.0 as i32 - 400).abs()))
        .map(|f| f.id)
}

/// Office's per-user cloud font cache, when there is one.
fn cloud_dir() -> Option<std::path::PathBuf> {
    let base = std::env::var_os("LOCALAPPDATA")?;
    let dir = std::path::PathBuf::from(base).join("Microsoft").join("FontCache").join("4").join("CloudFonts");
    dir.is_dir().then_some(dir)
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CloudFace {
    /// The family names of the face (typographic and Windows family, e.g. "Aptos Narrow").
    families: Vec<String>,
    weight: u16,
    italic: bool,
    /// What `cloud_font` takes to give the file's bytes.
    id: String,
}

/// The faces of Office's cloud fonts, for the page to register the families the browser does not know.
#[tauri::command]
pub fn cloud_fonts() -> Vec<CloudFace> {
    let Some(dir) = cloud_dir() else { return Vec::new() };
    cloud_db().faces()
        .filter_map(|f| {
            let family = cloud_family(f)?;
            let id = match &f.source { fontdb::Source::File(p) => p.strip_prefix(&dir).ok()?.to_string_lossy().into_owned(), _ => return None };
            Some(CloudFace { families: vec![family], weight: f.weight.0, italic: f.style != fontdb::Style::Normal, id })
        })
        .collect()
}

/// The bytes of one cloud font file (only files inside the cloud font cache).
#[tauri::command]
pub fn cloud_font(id: String) -> Result<tauri::ipc::Response, String> {
    let dir = cloud_dir().ok_or("no cloud fonts")?;
    let path = dir.join(&id);
    let canon = path.canonicalize().map_err(|e| e.to_string())?;
    if !canon.starts_with(dir.canonicalize().map_err(|e| e.to_string())?) {
        return Err("not a cloud font".into());
    }
    std::fs::read(&canon).map(tauri::ipc::Response::new).map_err(|e| e.to_string())
}

/// Starts loading the font database in the background so the first request is fast.
pub fn preload() {
    std::thread::spawn(|| {
        db();
    });
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Metrics {
    found: bool,
    units_per_em: u16,
    ascender: i16,
    descender: i16,
    line_gap: i16,
    win_ascent: u16,
    win_descent: u16,
    /// The font has a MATH table (Cambria Math): Word ignores its huge Windows values.
    math: bool,
}

#[tauri::command]
pub fn list_fonts() -> Vec<String> {
    let mut names: Vec<String> = db()
        .faces()
        .filter_map(|f| f.families.first().map(|(n, _)| n.clone()))
        .collect();
    names.sort_by_key(|n| n.to_lowercase());
    names.dedup();
    names
}

/// The regular face of a family, matched without regard to case (as Word does; fontdb's own
/// query compares names exactly and falls back to similar families).
fn find_face(db: &fontdb::Database, name: &str) -> Option<fontdb::ID> {
    let wanted = name.to_lowercase();
    let canonical = db
        .faces()
        .find_map(|f| f.families.iter().find(|(n, _)| n.to_lowercase() == wanted).map(|(n, _)| n.clone()))?;
    let query = fontdb::Query { families: &[fontdb::Family::Name(&canonical)], ..fontdb::Query::default() };
    db.query(&query)
        .filter(|id| db.face(*id).is_some_and(|f| f.families.iter().any(|(n, _)| *n == canonical)))
}

#[tauri::command]
pub fn font_metrics(families: Vec<String>) -> HashMap<String, Metrics> {
    let db = db();
    families
        .into_iter()
        .map(|name| {
            let (db, id) = match find_face(db, &name) { Some(id) => (db, Some(id)), None => (cloud_db(), find_cloud_face(&name)) };
            let metrics = id
                .and_then(|id| {
                    db.with_face_data(id, |data, index| {
                        let face = ttf_parser::Face::parse(data, index).ok()?;
                        let t = face.tables();
                        let os2 = t.os2;
                        Some(Metrics {
                            found: true,
                            units_per_em: face.units_per_em(),
                            ascender: t.hhea.ascender,
                            descender: t.hhea.descender,
                            line_gap: t.hhea.line_gap,
                            win_ascent: os2.map_or(0, |o| o.windows_ascender() as u16),
                            win_descent: os2.map_or(0, |o| o.windows_descender().unsigned_abs()),
                            math: face.raw_face().table(ttf_parser::Tag::from_bytes(b"MATH")).is_some(),
                        })
                    })
                    .flatten()
                })
                .unwrap_or(Metrics {
                    found: false,
                    units_per_em: 0,
                    ascender: 0,
                    descender: 0,
                    line_gap: 0,
                    win_ascent: 0,
                    win_descent: 0,
                    math: false,
                });
            (name, metrics)
        })
        .collect()
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FontChars {
    found: bool,
    /// A symbol font (cmap 3,0: Symbol, Wingdings): its characters are U+F020-U+F0FF.
    symbol: bool,
    /// Ranges of code points the font has a glyph for, as [first, last] pairs.
    ranges: Vec<[u32; 2]>,
}

/// Einfügen > Symbol: the characters a font has (its cmap), for the dialog's grid.
#[tauri::command]
pub fn font_chars(family: String) -> FontChars {
    let db = db();
    let none = FontChars { found: false, symbol: false, ranges: Vec::new() };
    let Some(id) = find_face(db, &family) else { return none };
    db.with_face_data(id, |data, index| {
        let face = ttf_parser::Face::parse(data, index).ok()?;
        let cmap = face.tables().cmap?;
        let mut codes: Vec<u32> = Vec::new();
        let mut symbol = false;
        for sub in cmap.subtables {
            if sub.platform_id == ttf_parser::PlatformId::Windows && sub.encoding_id == 0 {
                symbol = true;
            }
            if sub.is_unicode() || (sub.platform_id == ttf_parser::PlatformId::Windows && sub.encoding_id == 0) {
                sub.codepoints(|c| {
                    if sub.glyph_index(c).is_some_and(|g| g.0 != 0) {
                        codes.push(c);
                    }
                });
            }
        }
        codes.sort_unstable();
        codes.dedup();
        let mut ranges: Vec<[u32; 2]> = Vec::new();
        for c in codes {
            match ranges.last_mut() {
                Some(r) if r[1] + 1 == c => r[1] = c,
                _ => ranges.push([c, c]),
            }
        }
        Some(FontChars { found: true, symbol, ranges })
    })
    .flatten()
    .unwrap_or(none)
}
