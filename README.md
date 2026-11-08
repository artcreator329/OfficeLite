# OfficeLite 🦀📄📊
### Native Rust Desktop Word (.docx) & Excel (.xlsx) Viewer & Editor
**Zero Localhost &bull; Zero Network Ports &bull; 100% Native macOS &bull; 3.2 MB Binary**

**OfficeLite** is a blazingly fast, standalone desktop app built in **Rust** (powered by Tao + Wry / macOS WebKit) so you never need to pay for a Microsoft 365 / MS Office subscription.

![OfficeLite editing a Word document](docs/screenshot-document.jpg)

![OfficeLite editing an Excel spreadsheet](docs/screenshot-spreadsheet.jpg)

---

## ⚡ Why Rust Instead of Electron or Localhost?

| Metric | 🐌 Electron App | 🌐 Python Localhost | 🦀 **OfficeLite (Rust)** |
| :--- | :--- | :--- | :--- |
| **Localhost Server** | Often required | Required (`:8765`) | **None (Zero Ports)** |
| **Binary Size** | ~200 MB+ | ~50 MB+ runtime | **3.2 MB (Total)** |
| **RAM Usage** | ~300 MB - 600 MB | ~80 MB | **~35 MB** |
| **Startup Time** | 2.5 - 4.0 seconds | 1.5 seconds | **< 0.2 seconds** |
| **Engine** | Bundled Chromium | External Browser | **Native Apple WebKit** |
| **Network Exposure**| Local ports | Port 8765 | **Zero Open Sockets** |

---

## 🚀 Build From Source

There is no ready-made download yet, so build the app yourself. You need a Mac running macOS 11 or newer, with the [Rust toolchain](https://rustup.rs) installed.

```bash
git clone https://github.com/artcreator329/OfficeLite.git
cd OfficeLite
```

Build a double clickable app at `target/release/OfficeLite.app`, which you can drag into `/Applications`. Once it is there, Finder lists OfficeLite under **Open With** for `.docx`, `.xlsx`, `.xls`, `.csv`, `.txt`, and `.html` files:

```bash
./scripts/make-app.sh
```

Or run it directly:

```bash
cargo run --release
```

---

## ✨ Features

### 📄 Word Document Editor (.docx, .doc, .html, .txt)
- **View & Edit .docx Files**: Open a Word document from Finder, the File menu, or by dragging it into the window.
- **Rich Typography & Formatting**: Headings (H1, H2, H3), Paragraphs, Quotes, Code blocks.
- **Font & Size Selection**: Calibri, Arial, Times New Roman, Georgia, SF Pro, Geist, etc.
- **Styling**: Bold, Italic, Underline, Strikethrough, Text Color, Highlight Color.
- **Lists & Alignment**: Bulleted lists, Numbered lists, Indent/Outdent, Left/Center/Right/Justify.
- **Insertables**:
  - Interactive **Tables** (customizable rows and columns).
  - **Images** (file upload, clipboard paste, or drag & drop).
  - **Hyperlinks** and dividers.
- **Document Stats**: Live word count, character count, and estimated reading time.
- **Exporting**:
  - **Save as .docx** (Cmd+S): Choose the location in the native macOS Save panel. Keeps text color, highlight, fonts, sizes, alignment, lists, tables, images, and links.
  - **PDF / Print** (Cmd+P): The native macOS print panel, including Save as PDF.

### 📊 Excel Spreadsheet Editor (.xlsx, .xls, .csv)
- **View & Edit .xlsx Workbooks**: Open multi-sheet Excel files or CSVs from Finder, the File menu, or by dragging them into the window.
- **Formula Engine**:
  - Math and stats: `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `MEDIAN`, `ROUND`, `MOD`, and more.
  - Logic and lookup: `IF`, `IFERROR`, `AND`, `OR`, `SUMIF`, `COUNTIF`, `VLOOKUP`, `INDEX`, `MATCH`.
  - Text and dates: `CONCAT`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `DATE`, `TODAY`.
  - Arithmetic, `$` absolute references, and references to other sheets: `=A1 * 1.08`, `=SUM($B$2:$B$9)`, `=Budget!F6`.
  - Click a cell while typing a formula to insert its address.
  - Functions it does not know keep the value Excel saved in the file.
- **Editing**: Drag or Shift+arrow to select ranges, then see their sum and average in the status bar. Undo, fill handle, right click menu, and resizable columns.
- **Cell Formatting**: Bold, Italic, Text Color, Background Fill, Text Alignment, and Merged Cells.
- **Number Formats**: Currency (`$1,234.50`), Percentage (`12.5%`), Decimals, Integers, Dates, and Plain Text.
- **Grid Manipulation**: Insert and delete rows and columns. Formulas update to match.
- **Multi-Sheet Management**: Switch sheets, add new sheets (`+`), double-click to rename, or delete sheets.
- **Clipboard**: Cut, copy, and paste cells. Tab separated rows from Excel, Numbers, or a web table fill the grid.
- **Exporting**: Save as `.xlsx` (Cmd+S) through the native macOS Save panel. Bold, italic, colors, fills, alignment, number formats, and merged cells are kept.

### 🍎 Mac Menus and Shortcuts
- **File menu**: New (Cmd+N), Open (Cmd+O), Save (Cmd+S), and Print (Cmd+P).
- OfficeLite asks before replacing a file that has unsaved changes, including after the app restarts.

---

## 🔒 100% Private & Offline
- **Zero Localhost**: Communicates using an in-memory custom protocol (`officelite://`). It does not open any network sockets or listen on any ports.
- **Self-Contained**: All HTML, CSS, JavaScript, and libraries (SheetJS, Mammoth, JSZip) are compiled directly into the Mach-O binary.
- **Auto-Save Protection**: Real-time autosave to local application storage.

---

## ⚠️ Known Limits
- `.xlsx` files keep the formatting listed above. Borders, charts, and theme colors are not preserved.
- Spreadsheets have no sorting, filtering, frozen panes, or find and replace yet.
- `.docx` headers, footers, and comments are not kept. OfficeLite says so when it opens a file that has them. Footnotes move to the end of the document, and text boxes and equations become plain text.
- The old `.doc` format cannot be opened. Save it as `.docx` in Word first.
- The page on screen is always Letter size, but a `.docx` file keeps its own page size and margins when saved.
- Save asks for a location each time.
- OfficeLite edits one document and one spreadsheet at a time. Opening several files from Finder at once opens only the first.

## 📜 License
[ISC](LICENSE). Bundled components keep their own licenses, listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

OfficeLite is an independent project and is not affiliated with or endorsed by Microsoft.
