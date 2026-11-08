//! OfficeLite - Native Rust Desktop Application
//! Standalone, zero-localhost Word & Excel Editor.

mod macos;

use std::borrow::Cow;
use std::fs;
use std::io::Write;
use std::os::unix::fs::PermissionsExt;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use tao::{
    dpi::LogicalSize,
    event::{Event, StartCause, WindowEvent},
    event_loop::{ControlFlow, EventLoopBuilder},
    window::WindowBuilder,
};
use wry::{http::Response, WebViewBuilder};

// Embed static web assets directly into the compiled native binary
const INDEX_HTML: &str = include_str!("../index.html");
const STYLES_CSS: &str = include_str!("../styles.css");
const APP_JS: &str = include_str!("../app.js");
const FORMULA_JS: &str = include_str!("../formula.js");
const DOCX_JS: &str = include_str!("../docx.js");
const VENDOR_JSZIP: &[u8] = include_bytes!("../vendor/jszip.min.js");
const VENDOR_MAMMOTH: &[u8] = include_bytes!("../vendor/mammoth.browser.min.js");
const VENDOR_XLSX: &[u8] = include_bytes!("../vendor/xlsx.full.min.js");
const FONT_GEIST: &[u8] = include_bytes!("../vendor/fonts/Geist-Variable.woff2");
const FONT_GEIST_MONO: &[u8] = include_bytes!("../vendor/fonts/GeistMono-Variable.woff2");
const SAMPLE_DOCX: &[u8] = include_bytes!("../samples/Sample_Document.docx");
const SAMPLE_XLSX: &[u8] = include_bytes!("../samples/Sample_Budget.xlsx");

/// Messages from the web view's IPC handler back to the event loop.
enum UserEvent {
    Print { title: String },
    /// The page has loaded and can take File menu commands and opened files.
    PageReady,
    FileCommand(&'static str),
}

/// Files that Finder asked the app to open. The page fetches one by its index, so it never
/// sees or picks a path itself.
type OpenedFiles = Arc<Mutex<Vec<PathBuf>>>;

fn handle_opened_file(query: Option<&str>, opened: &OpenedFiles) -> Response<Cow<'static, [u8]>> {
    let path = query
        .and_then(|q| q.split('&').find_map(|pair| pair.strip_prefix("id=")))
        .and_then(|id| id.parse::<usize>().ok())
        .and_then(|id| opened.lock().ok()?.get(id).cloned());
    let Some(path) = path else {
        return json_response(404, serde_json::json!({ "error": "This file is no longer available" }));
    };
    match fs::read(&path) {
        Ok(bytes) => Response::builder()
            .header("Content-Type", "application/octet-stream")
            .body(cow_vec(bytes))
            .unwrap(),
        Err(e) => json_response(
            500,
            serde_json::json!({ "error": format!("{} ({})", e, path.display()) }),
        ),
    }
}

/// Script that hands an opened file to the page.
fn open_file_script(id: usize, path: &Path) -> String {
    let name = path
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default();
    format!(
        "window.officeLiteOpenFile && window.officeLiteOpenFile({}, {})",
        id,
        serde_json::Value::String(name)
    )
}

fn cow_slice<'a>(s: &'a [u8]) -> Cow<'a, [u8]> {
    Cow::Borrowed(s)
}

fn cow_vec(v: Vec<u8>) -> Cow<'static, [u8]> {
    Cow::Owned(v)
}

fn json_response(status: u16, value: serde_json::Value) -> Response<Cow<'static, [u8]>> {
    Response::builder()
        .status(status)
        .header("Content-Type", "application/json")
        .body(cow_vec(value.to_string().into_bytes()))
        .unwrap()
}

fn percent_decode(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        let hex = (bytes[i] == b'%')
            .then(|| input.get(i + 1..i + 3))
            .flatten()
            .and_then(|h| u8::from_str_radix(h, 16).ok());
        match hex {
            Some(byte) => {
                out.push(byte);
                i += 3;
            }
            None => {
                out.push(bytes[i]);
                i += 1;
            }
        }
    }
    String::from_utf8_lossy(&out).into_owned()
}

/// Writes through a temp file so a failed save never damages the file being replaced.
fn write_atomically(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let dir = path.parent().unwrap_or_else(|| Path::new("."));
    let permissions = fs::metadata(path)
        .map(|meta| meta.permissions())
        .unwrap_or_else(|_| fs::Permissions::from_mode(0o644));

    let mut temp = tempfile::NamedTempFile::new_in(dir)?;
    temp.write_all(bytes)?;
    temp.as_file().sync_all()?;
    temp.as_file().set_permissions(permissions)?;
    temp.persist(path).map_err(|e| e.error)?;
    Ok(())
}

/// Shows the Save panel, then writes the request body to the location the person chose.
/// The page only suggests a name: it can never pick the path.
fn handle_save_file(query: Option<&str>, body: &[u8]) -> Response<Cow<'static, [u8]>> {
    let default_name = query
        .and_then(|q| q.split('&').find_map(|pair| pair.strip_prefix("name=")))
        .map(percent_decode)
        .unwrap_or_default();
    let default_name = Path::new(&default_name)
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("Untitled")
        .to_string();

    let chosen = match macos::prompt_save_path(&default_name) {
        Ok(Some(path)) => path,
        Ok(None) => return json_response(200, serde_json::json!({ "status": "cancelled" })),
        Err(message) => return json_response(500, serde_json::json!({ "error": message })),
    };

    // The panel lets people delete the extension, which would make the file hard to open.
    let chosen = match (chosen.extension(), Path::new(&default_name).extension()) {
        (None, Some(ext)) => chosen.with_extension(ext),
        _ => chosen,
    };

    match write_atomically(&chosen, body) {
        Ok(()) => json_response(
            200,
            serde_json::json!({
                "status": "saved",
                "name": chosen.file_name().map(|n| n.to_string_lossy().into_owned()),
                "path": chosen.to_string_lossy(),
            }),
        ),
        Err(e) => json_response(
            500,
            serde_json::json!({ "error": format!("{} ({})", e, chosen.display()) }),
        ),
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {

    let event_loop = EventLoopBuilder::<UserEvent>::with_user_event().build();
    let ipc_proxy = event_loop.create_proxy();
    let menu_proxy = event_loop.create_proxy();
    let opened_files: OpenedFiles = Arc::new(Mutex::new(Vec::new()));
    let protocol_opened_files = opened_files.clone();
    let window = WindowBuilder::new()
        .with_title("OfficeLite")
        .with_inner_size(LogicalSize::new(1280.0, 860.0))
        .with_min_inner_size(LogicalSize::new(960.0, 600.0))
        .build(&event_loop)?;

    let webview = WebViewBuilder::new()
        .with_ipc_handler(move |request| {
            let Ok(message) = serde_json::from_str::<serde_json::Value>(request.body()) else {
                return;
            };
            if message["cmd"] == "print" {
                let title = message["title"].as_str().unwrap_or("OfficeLite").to_string();
                let _ = ipc_proxy.send_event(UserEvent::Print { title });
            } else if message["cmd"] == "ready" {
                let _ = ipc_proxy.send_event(UserEvent::PageReady);
            }
        })
        .with_custom_protocol("officelite".into(), move |_webview_id, req| {
            let path = req.uri().path();
            let method = req.method().as_str();

            // API: Save a file through the native Save panel
            if path == "/api/save-file" && method == "POST" {
                return handle_save_file(req.uri().query(), req.body());
            }

            // API: Read a file that Finder asked the app to open
            if path == "/api/opened-file" && method == "GET" {
                return handle_opened_file(req.uri().query(), &protocol_opened_files);
            }

            // Embedded Static Assets
            match path {
                "/" | "/index.html" => Response::builder()
                    .header("Content-Type", "text/html; charset=utf-8")
                    .body(cow_slice(INDEX_HTML.as_bytes()))
                    .unwrap(),
                "/styles.css" => Response::builder()
                    .header("Content-Type", "text/css; charset=utf-8")
                    .body(cow_slice(STYLES_CSS.as_bytes()))
                    .unwrap(),
                "/app.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(APP_JS.as_bytes()))
                    .unwrap(),
                "/formula.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(FORMULA_JS.as_bytes()))
                    .unwrap(),
                "/docx.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(DOCX_JS.as_bytes()))
                    .unwrap(),
                "/vendor/jszip.min.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(VENDOR_JSZIP))
                    .unwrap(),
                "/vendor/mammoth.browser.min.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(VENDOR_MAMMOTH))
                    .unwrap(),
                "/vendor/xlsx.full.min.js" => Response::builder()
                    .header("Content-Type", "application/javascript; charset=utf-8")
                    .body(cow_slice(VENDOR_XLSX))
                    .unwrap(),
                "/vendor/fonts/Geist-Variable.woff2" => Response::builder()
                    .header("Content-Type", "font/woff2")
                    .body(cow_slice(FONT_GEIST))
                    .unwrap(),
                "/vendor/fonts/GeistMono-Variable.woff2" => Response::builder()
                    .header("Content-Type", "font/woff2")
                    .body(cow_slice(FONT_GEIST_MONO))
                    .unwrap(),
                "/samples/Sample_Document.docx" => Response::builder()
                    .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    .body(cow_slice(SAMPLE_DOCX))
                    .unwrap(),
                "/samples/Sample_Budget.xlsx" => Response::builder()
                    .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    .body(cow_slice(SAMPLE_XLSX))
                    .unwrap(),
                _ => Response::builder()
                    .status(404)
                    .header("Content-Type", "text/plain")
                    .body(cow_slice(b"404 Not Found"))
                    .unwrap(),
            }
        })
        .with_url("officelite://localhost/index.html")
        .build(&window)?;

    // Files opened before the page finished loading wait here.
    let mut page_ready = false;
    let mut waiting_file: Option<usize> = None;

    event_loop.run(move |event, _, control_flow| {
        *control_flow = ControlFlow::Wait;

        match event {
            Event::NewEvents(StartCause::Init) => {
                let proxy = menu_proxy.clone();
                macos::install_main_menu(move |command| {
                    let _ = proxy.send_event(UserEvent::FileCommand(command));
                });
            }
            Event::Opened { urls } => {
                // OfficeLite edits one file at a time, so it opens the first one.
                let Some(path) = urls.iter().find_map(|url| url.to_file_path().ok()) else {
                    return;
                };
                let id = {
                    let mut files = opened_files.lock().unwrap();
                    files.push(path.clone());
                    files.len() - 1
                };
                if page_ready {
                    let _ = webview.evaluate_script(&open_file_script(id, &path));
                } else {
                    waiting_file = Some(id);
                }
                window.set_focus();
            }
            Event::UserEvent(UserEvent::PageReady) => {
                page_ready = true;
                if let Some(id) = waiting_file.take() {
                    let path = opened_files.lock().unwrap()[id].clone();
                    let _ = webview.evaluate_script(&open_file_script(id, &path));
                }
            }
            Event::UserEvent(UserEvent::FileCommand(command)) => {
                let script = format!(
                    "window.officeLiteCommand && window.officeLiteCommand({})",
                    serde_json::Value::String(command.to_string())
                );
                let _ = webview.evaluate_script(&script);
            }
            Event::UserEvent(UserEvent::Print { title }) => {
                macos::print_webview(&webview, &title);
            }
            Event::WindowEvent {
                event: WindowEvent::CloseRequested,
                ..
            } => {
                *control_flow = ControlFlow::Exit;
            }
            _ => (),
        }
    });
}
