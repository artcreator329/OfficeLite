# OfficeLite 🦀📄📊
### Free Word (.docx) & Excel (.xlsx) Editor for Mac, Windows, and Linux
**Native Mac App &bull; Browser Version for Windows and Linux &bull; Private and Offline &bull; 3.2 MB Mac App**

**OfficeLite** is a fast, free editor for Word documents and Excel spreadsheets, so you never need to pay for a Microsoft 365 / MS Office subscription. On a Mac it runs as a small native app built in **Rust** (powered by Tao + Wry / macOS WebKit). On Windows and Linux it runs in your browser.

![OfficeLite editing a Word document](docs/screenshot-document.jpg)

![OfficeLite editing an Excel spreadsheet](docs/screenshot-spreadsheet.jpg)

---

## ⚡ Why a Native Rust App on the Mac?

| Metric | 🐌 Electron App | 🌐 Python Localhost | 🦀 **OfficeLite (Rust)** |
| :--- | :--- | :--- | :--- |
| **Localhost Server** | Often required | Required (`:8765`) | **None (Zero Ports)** |
| **Binary Size** | ~200 MB+ | ~50 MB+ runtime | **3.2 MB (Total)** |
| **RAM Usage** | ~300 MB - 600 MB | ~80 MB | **~35 MB** |
| **Startup Time** | 2.5 - 4.0 seconds | 1.5 seconds | **< 0.2 seconds** |
| **Engine** | Bundled Chromium | External Browser | **Native Apple WebKit** |
| **Network Exposure**| Local ports | Port 8765 | **Zero Open Sockets** |

---

## 🚀 Run It

| Your computer | How to run OfficeLite |
| :--- | :--- |
| **Windows** | In your browser (Option 1) |
| **Linux** | In your browser (Option 1) |
| **Mac** | As a native app (Option 2), or in your browser (Option 1) |

There is no ready-made download yet. First, get a copy of this repository:

```bash
git clone https://github.com/artcreator329/OfficeLite.git
cd OfficeLite
```

No Git? On this GitHub page, click **Code > Download ZIP**, then unzip it.

### Option 1: In your browser (Windows, Linux, or Mac)

You need [Python 3](https://www.python.org/downloads/) and nothing else. On Windows, tick **Add python.exe to PATH** in the Python installer.

| Your computer | Start it by |
| :--- | :--- |
| **Windows** | Double clicking `start.bat`, or running `py server.py` in Command Prompt |
| **Mac** | Double clicking `start.command`, or running `python3 server.py` in Terminal |
| **Linux** | Running `python3 server.py` in a terminal |

OfficeLite opens in your default browser at `http://127.0.0.1:8765/`. If that port is busy, it picks the next free one and prints the address. Keep the window open while you work, and press **Ctrl+C** in it to stop.

Tested in Chrome and Safari. Edge uses the same engine as Chrome. Firefox has not been tested.

How the browser version differs from the Mac app:
- **Save** downloads the file to your Downloads folder instead of asking where to put it.
- **Print** uses your browser's print dialog, which can also save a PDF.
- To open a file, use the **Open** button or drag the file into the window. There is no File menu and no Open With.
- Your work autosaves in that browser only.

The server only listens on your own computer (`127.0.0.1`), so no other device can reach it, and your files never leave your computer.

### Option 2: The Mac app

You need macOS 11 or newer and the [Rust toolchain](https://rustup.rs). The first build downloads the Rust dependencies and takes a few minutes.

Build a double clickable app at `target/release/OfficeLite.app`:

```bash
./scripts/make-app.sh
```

Drag it into `/Applications`. Finder then lists OfficeLite under **Open With** for `.docx`, `.xlsx`, `.xls`, `.csv`, `.txt`, and `.html` files.

The first time you open it, macOS may block it because it is not signed with an Apple Developer ID. Go to **System Settings > Privacy & Security** and click **Open Anyway**.

To run it without making the app bundle:

```bash
cargo run --release
```

The Mac app does not build on Windows or Linux. Use Option 1 there.

### Run the tests

The unit tests cover the formula engine and the Word file reader and writer. They need [Node.js](https://nodejs.org) 18 or newer, and nothing to install:

```bash
npm test
```

To test Word files in a real browser, start Option 1, then open `http://127.0.0.1:8765/tests/docx-roundtrip.html`. It saves a sample document as `.docx`, opens it again, and shows how many checks passed.

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
- **No Network Ports on the Mac**: The Mac app talks to its page through an in-memory custom protocol (`officelite://`). It does not open any network sockets or listen on any ports.
- **Local Only in the Browser**: The browser version's server listens only on `127.0.0.1`, so nothing outside your computer can reach it.
- **Self-Contained**: All HTML, CSS, JavaScript, and libraries (SheetJS, Mammoth, JSZip) ship with OfficeLite, and the Mac app compiles them into its binary. Nothing is loaded from the internet.
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
