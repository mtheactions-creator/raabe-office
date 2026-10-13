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
            let metrics = find_face(db, &name)
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
