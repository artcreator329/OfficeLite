//! OfficeLite - Native Rust Desktop Application
//! Standalone, zero-localhost Word & Excel Editor.

mod macos;

use std::borrow::Cow;
use std::fs;
use std::io::Write;
use std::os::unix::fs::PermissionsExt;
use std::path::{Path, PathBuf};
use std::process::Command;

use serde::{Deserialize, Serialize};
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
}

fn cow_slice<'a>(s: &'a [u8]) -> Cow<'a, [u8]> {
    Cow::Borrowed(s)
}

fn cow_vec(v: Vec<u8>) -> Cow<'static, [u8]> {
    Cow::Owned(v)
}

#[derive(Serialize)]
struct StatusResponse {
    status: &'static str,
    pandoc_available: bool,
    pandoc_path: Option<String>,
    pandoc_version: Option<String>,
    engine: &'static str,
}

#[derive(Deserialize)]
struct ExportDocxRequest {
    html: String,
    filename: Option<String>,
}

#[derive(Serialize)]
struct ImportDocxResponse {
    html: String,
}

fn find_pandoc() -> Option<(PathBuf, String)> {
    let candidate_paths = [
        PathBuf::from("/opt/homebrew/bin/pandoc"),
        PathBuf::from("/usr/local/bin/pandoc"),
        PathBuf::from("/usr/bin/pandoc"),
    ];

    for path in &candidate_paths {
        if path.exists() {
            if let Ok(output) = Command::new(path).arg("--version").output() {
                if output.status.success() {
                    let text = String::from_utf8_lossy(&output.stdout);
                    let first_line = text.lines().next().unwrap_or("pandoc").to_string();
                    return Some((path.clone(), first_line));
                }
            }
        }
    }

    // Try finding via `which pandoc`
    if let Ok(output) = Command::new("which").arg("pandoc").output() {
        if output.status.success() {
            let path_str = String::from_utf8_lossy(&output.stdout).trim().to_string();
            let path = PathBuf::from(path_str);
            if path.exists() {
                if let Ok(ver_out) = Command::new(&path).arg("--version").output() {
                    let text = String::from_utf8_lossy(&ver_out.stdout);
                    let first_line = text.lines().next().unwrap_or("pandoc").to_string();
                    return Some((path, first_line));
                }
            }
        }
    }

    None
}

fn handle_export_docx(body: &[u8], pandoc_info: &Option<(PathBuf, String)>) -> Response<Cow<'static, [u8]>> {
    let pandoc_bin = match pandoc_info {
        Some((path, _)) => path,
        None => {
            return Response::builder()
                .status(501)
                .header("Content-Type", "application/json")
                .body(cow_slice(b"{\"error\":\"Pandoc not installed\"}"))
                .unwrap();
        }
    };

    let req_data: ExportDocxRequest = match serde_json::from_slice(body) {
        Ok(d) => d,
        Err(e) => {
            let err_json = format!("{{\"error\":\"Invalid JSON: {}\"}}", e);
            return Response::builder()
                .status(400)
                .header("Content-Type", "application/json")
                .body(cow_vec(err_json.into_bytes()))
                .unwrap();
        }
    };

    let filename = req_data.filename.unwrap_or_else(|| "document.docx".to_string());
    let mut temp_html = tempfile::Builder::new().suffix(".html").tempfile().unwrap();
    if let Err(e) = temp_html.write_all(req_data.html.as_bytes()) {
        let err_json = format!("{{\"error\":\"Temp file error: {}\"}}", e);
        return Response::builder()
            .status(500)
            .header("Content-Type", "application/json")
            .body(cow_vec(err_json.into_bytes()))
            .unwrap();
    }
    let html_path = temp_html.path().to_path_buf();
    let docx_path = html_path.with_extension("docx");

    let cmd_status = Command::new(pandoc_bin)
        .arg("-f")
        .arg("html")
        .arg("-t")
        .arg("docx")
        .arg("-o")
        .arg(&docx_path)
        .arg(&html_path)
        .status();

    match cmd_status {
        Ok(status) if status.success() => {
            if let Ok(docx_bytes) = fs::read(&docx_path) {
                let _ = fs::remove_file(&docx_path);
                Response::builder()
                    .status(200)
                    .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    .header("Content-Disposition", format!("attachment; filename=\"{}\"", filename))
                    .body(cow_vec(docx_bytes))
                    .unwrap()
            } else {
                Response::builder()
                    .status(500)
                    .header("Content-Type", "application/json")
                    .body(cow_slice(b"{\"error\":\"Failed to read generated docx\"}"))
                    .unwrap()
            }
        }
        Ok(status) => {
            let err_json = format!("{{\"error\":\"Pandoc failed with exit code: {}\"}}", status);
            Response::builder()
                .status(500)
                .header("Content-Type", "application/json")
                .body(cow_vec(err_json.into_bytes()))
                .unwrap()
        }
        Err(e) => {
            let err_json = format!("{{\"error\":\"Pandoc execution error: {}\"}}", e);
            Response::builder()
                .status(500)
                .header("Content-Type", "application/json")
                .body(cow_vec(err_json.into_bytes()))
                .unwrap()
        }
    }
}

fn handle_import_docx(body: &[u8], pandoc_info: &Option<(PathBuf, String)>) -> Response<Cow<'static, [u8]>> {
    let pandoc_bin = match pandoc_info {
        Some((path, _)) => path,
        None => {
            return Response::builder()
                .status(501)
                .header("Content-Type", "application/json")
                .body(cow_slice(b"{\"error\":\"Pandoc not installed\"}"))
                .unwrap();
        }
    };

    let mut temp_docx = tempfile::Builder::new().suffix(".docx").tempfile().unwrap();
    if let Err(e) = temp_docx.write_all(body) {
        let err_json = format!("{{\"error\":\"Temp file error: {}\"}}", e);
        return Response::builder()
            .status(500)
            .header("Content-Type", "application/json")
            .body(cow_vec(err_json.into_bytes()))
            .unwrap();
    }
    let docx_path = temp_docx.path().to_path_buf();

    let output = Command::new(pandoc_bin)
        .arg("-f")
        .arg("docx")
        .arg("-t")
        .arg("html")
        .arg(&docx_path)
        .output();

    match output {
        Ok(out) if out.status.success() => {
            let html_str = String::from_utf8_lossy(&out.stdout).to_string();
            let json_resp = serde_json::to_string(&ImportDocxResponse { html: html_str }).unwrap();
            Response::builder()
                .status(200)
                .header("Content-Type", "application/json")
                .body(cow_vec(json_resp.into_bytes()))
                .unwrap()
        }
        Ok(out) => {
            let err_text = String::from_utf8_lossy(&out.stderr);
            let err_json = format!("{{\"error\":\"Pandoc failed: {}\"}}", err_text);
            Response::builder()
                .status(500)
                .header("Content-Type", "application/json")
                .body(cow_vec(err_json.into_bytes()))
                .unwrap()
        }
        Err(e) => {
            let err_json = format!("{{\"error\":\"Pandoc execution error: {}\"}}", e);
            Response::builder()
                .status(500)
                .header("Content-Type", "application/json")
                .body(cow_vec(err_json.into_bytes()))
                .unwrap()
        }
    }
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
    let pandoc_info = find_pandoc();
    let pandoc_clone_for_protocol = pandoc_info.clone();

    let event_loop = EventLoopBuilder::<UserEvent>::with_user_event().build();
    let ipc_proxy = event_loop.create_proxy();
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
            }
        })
        .with_custom_protocol("officelite".into(), move |_webview_id, req| {
            let path = req.uri().path();
            let method = req.method().as_str();

            // API: Status
            if path == "/api/status" {
                let resp = StatusResponse {
                    status: "online",
                    pandoc_available: pandoc_clone_for_protocol.is_some(),
                    pandoc_path: pandoc_clone_for_protocol.as_ref().map(|(p, _)| p.to_string_lossy().to_string()),
                    pandoc_version: pandoc_clone_for_protocol.as_ref().map(|(_, v)| v.clone()),
                    engine: "Native Rust (WebKit / Wry - Zero Localhost)",
                };
                let json = serde_json::to_string(&resp).unwrap();
                return Response::builder()
                    .header("Content-Type", "application/json")
                    .header("Access-Control-Allow-Origin", "*")
                    .body(cow_vec(json.into_bytes()))
                    .unwrap();
            }

            // API: Export DOCX
            if path == "/api/export-docx" && method == "POST" {
                return handle_export_docx(req.body(), &pandoc_clone_for_protocol);
            }

            // API: Import DOCX
            if path == "/api/import-docx" && method == "POST" {
                return handle_import_docx(req.body(), &pandoc_clone_for_protocol);
            }

            // API: Save a file through the native Save panel
            if path == "/api/save-file" && method == "POST" {
                return handle_save_file(req.uri().query(), req.body());
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

    event_loop.run(move |event, _, control_flow| {
        *control_flow = ControlFlow::Wait;

        match event {
            Event::NewEvents(StartCause::Init) => {
                macos::install_main_menu();
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
