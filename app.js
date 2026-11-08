/**
 * OfficeLite - Core Application Logic
 * Comprehensive DOCX & XLSX Viewer, Editor, and Generator.
 */

(function () {
  'use strict';

  const IS_NATIVE_APP = window.location.protocol === 'officelite:';

  // --- App State ---
  const state = {
    mode: 'doc', // 'doc' | 'sheet'
    docTitle: 'Untitled Document',
    docSection: null, // Page size and margins from the .docx file that was opened
    sheetTitle: 'Untitled Spreadsheet',
    docModified: false,
    sheetModified: false,

    // Spreadsheet State
    sheet: {
      activeSheet: 'Sheet1',
      sheetNames: ['Sheet1'],
      sheets: {
        'Sheet1': {
          cells: {}, // key: "A1" -> { v: "100", f: "=SUM(...)", s: { bold: true, align: 'left', format: 'general', color: '', bg: '' } }
          rowCount: 40,
          colCount: 20
        }
      },
      selectedCell: { r: 0, c: 0 }, // The active cell. Row and column indexes start at 0.
      rangeEnd: { r: 0, c: 0 } // The opposite corner of the selection
    }
  };

  // --- DOM Elements ---
  const el = {
    brandLogo: document.getElementById('brandLogo'),
    docNameInput: document.getElementById('docNameInput'),
    saveStatus: document.getElementById('saveStatus'),
    saveStatusText: document.getElementById('saveStatusText'),
    btnModeDoc: document.getElementById('btnModeDoc'),
    btnModeSheet: document.getElementById('btnModeSheet'),
    btnNew: document.getElementById('btnNew'),
    btnOpenFile: document.getElementById('btnOpenFile'),
    fileInput: document.getElementById('fileInput'),
    btnSaveMain: document.getElementById('btnSaveMain'),
    saveBtnLabel: document.getElementById('saveBtnLabel'),
    btnPrint: document.getElementById('btnPrint'),
    btnSamples: document.getElementById('btnSamples'),
    btnHelp: document.getElementById('btnHelp'),

    // Ribbon toolbars
    toolbarDoc: document.getElementById('toolbarDoc'),
    toolbarSheet: document.getElementById('toolbarSheet'),
    formulaBarContainer: document.getElementById('formulaBarContainer'),
    cellAddress: document.getElementById('cellAddress'),
    formulaInput: document.getElementById('formulaInput'),

    // Workspaces
    docWorkspace: document.getElementById('docWorkspace'),
    documentPage: document.getElementById('documentPage'),
    sheetWorkspace: document.getElementById('sheetWorkspace'),
    gridContainer: document.getElementById('gridContainer'),
    spreadsheetGrid: document.getElementById('spreadsheetGrid'),
    tabsList: document.getElementById('tabsList'),
    btnAddSheet: document.getElementById('btnAddSheet'),

    // Stats
    statWordCount: document.getElementById('statWordCount'),
    statCharCount: document.getElementById('statCharCount'),
    statReadingTime: document.getElementById('statReadingTime'),
    sheetStats: document.getElementById('sheetStats'),
    docStats: document.getElementById('docStats'),
    sheetStatCell: document.getElementById('sheetStatCell'),
    sheetStatSum: document.getElementById('sheetStatSum'),
    sheetStatCount: document.getElementById('sheetStatCount'),

    // Modals
    modalTable: document.getElementById('modalTable'),
    modalLink: document.getElementById('modalLink'),
    modalHelp: document.getElementById('modalHelp'),
    modalSamples: document.getElementById('modalSamples'),
    modalConfirm: document.getElementById('modalConfirm'),
    toastContainer: document.getElementById('toastContainer'),
    dragOverlay: document.getElementById('dragOverlay')
  };

  // --- Helpers: Spreadsheet Coordinates ---
  function indexToCol(index) {
    let col = '';
    let temp = index;
    while (temp >= 0) {
      col = String.fromCharCode((temp % 26) + 65) + col;
      temp = Math.floor(temp / 26) - 1;
    }
    return col;
  }

  function colToIndex(colStr) {
    let index = 0;
    const str = colStr.toUpperCase();
    for (let i = 0; i < str.length; i++) {
      index = index * 26 + (str.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  function coordToAddress(r, c) {
    return `${indexToCol(c)}${r + 1}`;
  }

  function addressToCoord(addr) {
    const match = addr.match(/^([A-Z]+)([0-9]+)$/i);
    if (!match) return null;
    return {
      c: colToIndex(match[1]),
      r: parseInt(match[2], 10) - 1
    };
  }

  // --- Toast Notifications ---
  function escapeXml(unsafe) {
    const entities = { '<': '&lt;', '>': '&gt;', '&': '&amp;', '\'': '&apos;', '"': '&quot;' };
    return String(unsafe ?? '').replace(/[<>&'"]/g, (c) => entities[c]);
  }

  function icon(name) {
    return `<svg class="icon" aria-hidden="true"><use href="#i-${name}"/></svg>`;
  }

  const TOAST_ICONS = { info: 'info', success: 'check-circle', warning: 'warning-circle', error: 'warning-circle' };
  const MAX_TOASTS = 3;

  function showToast(message, kind = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${kind}`;
    toast.setAttribute('role', kind === 'error' ? 'alert' : 'status');
    toast.innerHTML = `${icon(TOAST_ICONS[kind])}<span></span>`;
    toast.querySelector('span').textContent = message;
    el.toastContainer.appendChild(toast);
    while (el.toastContainer.children.length > MAX_TOASTS) {
      el.toastContainer.firstElementChild.remove();
    }
    setTimeout(() => {
      toast.classList.add('is-leaving');
      setTimeout(() => toast.remove(), 200);
    }, kind === 'error' || kind === 'warning' ? 6000 : 2800);
  }

  // --- Modals ---
  // Replaces window.alert/confirm/prompt, which the native WebKit shell does not implement.
  let modalReturnFocus = null;
  let savedDocRange = null;
  let pendingConfirm = null;

  function openModal(modal) {
    const sel = window.getSelection();
    savedDocRange = sel.rangeCount && el.documentPage.contains(sel.anchorNode)
      ? sel.getRangeAt(0).cloneRange()
      : null;
    modalReturnFocus = document.activeElement;
    modal.classList.add('active');
    const target = modal.querySelector('[data-autofocus]') || modal.querySelector('button');
    if (target) {
      target.focus();
      if (target.select) target.select();
    }
  }

  function closeModal(modal) {
    modal.classList.remove('active');
    if (modal === el.modalConfirm && pendingConfirm) {
      pendingConfirm(false);
      pendingConfirm = null;
    }
    if (modalReturnFocus && document.contains(modalReturnFocus)) modalReturnFocus.focus();
    modalReturnFocus = null;
  }

  function restoreDocSelection() {
    if (!savedDocRange) return;
    el.documentPage.focus();
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(savedDocRange);
    savedDocRange = null;
  }

  function confirmDialog({ title, message, confirmLabel }) {
    document.getElementById('modalConfirmTitle').textContent = title;
    document.getElementById('modalConfirmMessage').textContent = message;
    document.getElementById('btnConfirmOk').textContent = confirmLabel;
    return new Promise((resolve) => {
      pendingConfirm = resolve;
      openModal(el.modalConfirm);
    });
  }

  function initModals() {
    document.getElementById('btnConfirmCancel').onclick = () => closeModal(el.modalConfirm);
    document.getElementById('btnConfirmOk').onclick = () => {
      const resolve = pendingConfirm;
      pendingConfirm = null;
      closeModal(el.modalConfirm);
      if (resolve) resolve(true);
    };

    document.querySelectorAll('.modal-overlay').forEach((modal) => {
      modal.addEventListener('mousedown', (e) => {
        if (e.target === modal) closeModal(modal);
      });
    });

    document.addEventListener('keydown', (e) => {
      const modal = document.querySelector('.modal-overlay.active');
      if (!modal) return;

      if (e.key === 'Escape') {
        e.preventDefault();
        closeModal(modal);
      } else if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
        const confirmBtn = modal.querySelector('[data-modal-confirm]');
        if (confirmBtn) {
          e.preventDefault();
          confirmBtn.click();
        }
      } else if (e.key === 'Tab') {
        const focusable = [...modal.querySelectorAll('button, input')];
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      }
    });
  }

  // --- Keep toolbars from stealing focus or the text selection on click ---
  function initToolbars() {
    document.querySelectorAll('.ribbon-toolbar').forEach((toolbar) => {
      toolbar.addEventListener('mousedown', (e) => {
        if (!e.target.closest('select, input, label')) e.preventDefault();
      });
    });

    document.querySelectorAll('.color-input-hidden').forEach((input) => {
      const setSwatch = () => input.parentElement.style.setProperty('--swatch', input.value);
      setSwatch();
      input.addEventListener('input', setSwatch);
    });
  }

  function setPressed(id, pressed) {
    document.getElementById(id).setAttribute('aria-pressed', pressed ? 'true' : 'false');
  }

  // AppKit has no Format menu here, so these need handling in the page.
  function handleFormatShortcut(e, key) {
    if (state.mode === 'doc' && document.activeElement === el.documentPage) {
      e.preventDefault();
      execDocCmd({ b: 'bold', i: 'italic', u: 'underline' }[key]);
    } else if (gridHasFocus() && key !== 'u') {
      e.preventDefault();
      toggleCellFormat(key === 'b' ? 'bold' : 'italic');
    }
  }

  // --- App shortcuts ---
  function initShortcuts() {
    document.addEventListener('keydown', (e) => {
      if (!e.metaKey && !e.ctrlKey) return;
      if (document.querySelector('.modal-overlay.active')) return;
      const key = e.key.toLowerCase();
      if (key === 's') {
        e.preventDefault();
        el.btnSaveMain.click();
      } else if (key === 'o') {
        e.preventDefault();
        el.btnOpenFile.click();
      } else if (key === 'p') {
        e.preventDefault();
        el.btnPrint.click();
      } else if (!e.shiftKey && !e.altKey && ['b', 'i', 'u'].includes(key)) {
        handleFormatShortcut(e, key);
      }
    });
  }

  // ==========================================
  // --- DOCUMENT (WORD) EDITOR LOGIC ---
  // ==========================================

  function execDocCmd(cmd, value = null) {
    el.documentPage.focus();
    document.execCommand(cmd, false, value);
    updateDocStats();
    markDocModified();
    syncDocToolbar();
  }

  function updateDocStats() {
    const text = el.documentPage.innerText || '';
    const words = text.trim().length ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    const readingTime = Math.ceil(words / 200);

    el.statWordCount.textContent = `${words.toLocaleString()} ${words === 1 ? 'word' : 'words'}`;
    el.statCharCount.textContent = `${chars.toLocaleString()} ${chars === 1 ? 'character' : 'characters'}`;
    el.statReadingTime.textContent = `${readingTime} min read`;

    const isEmpty = !text.trim() && !el.documentPage.querySelector('img, table, hr');
    el.documentPage.classList.toggle('is-empty', isEmpty);
  }

  function updateSaveStatus() {
    const modified = state.mode === 'doc' ? state.docModified : state.sheetModified;
    el.saveStatus.dataset.state = modified ? 'unsaved' : 'saved';
    el.saveStatusText.textContent = modified ? 'Unsaved changes' : 'Saved';
  }

  function markDocModified() {
    state.docModified = true;
    updateSaveStatus();
  }

  function markSheetModified() {
    state.sheetModified = true;
    updateSaveStatus();
  }

  // Marks the file in the active editor as saved.
  function markSaved() {
    if (state.mode === 'doc') state.docModified = false;
    else state.sheetModified = false;
    updateSaveStatus();
  }

  // Reflect the formatting at the cursor in the document toolbar.
  function syncDocToolbar() {
    if (state.mode !== 'doc') return;
    const sel = window.getSelection();
    if (!sel.rangeCount || !el.documentPage.contains(sel.anchorNode)) return;

    const commands = {
      docBold: 'bold',
      docItalic: 'italic',
      docUnderline: 'underline',
      docStrike: 'strikeThrough',
      docAlignLeft: 'justifyLeft',
      docAlignCenter: 'justifyCenter',
      docAlignRight: 'justifyRight',
      docAlignJustify: 'justifyFull',
      docBulletList: 'insertUnorderedList',
      docNumberList: 'insertOrderedList'
    };
    Object.entries(commands).forEach(([id, cmd]) => setPressed(id, document.queryCommandState(cmd)));

    const block = document.queryCommandValue('formatBlock').toLowerCase();
    const formatSelect = document.getElementById('docFormatSelect');
    formatSelect.value = [...formatSelect.options].some((o) => o.value === block) ? block : 'p';

    const selection = window.getSelection();
    const anchor = selection.anchorNode && (selection.anchorNode.nodeType === 1 ? selection.anchorNode : selection.anchorNode.parentElement);
    if (anchor && el.documentPage.contains(anchor)) {
      const computed = getComputedStyle(anchor);

      // Blank when the text is a size or font the menus do not list.
      const points = parseFloat(computed.fontSize) * 0.75;
      const sizes = window.OfficeLiteDocx.FONT_SIZE_PT;
      const sizeSelect = document.getElementById('docFontSize');
      sizeSelect.value = Object.keys(sizes).find((key) => Math.abs(sizes[key] - points) < 0.1) || '';

      const family = (value) => value.split(',')[0].replace(/['"]/g, '').trim().toLowerCase();
      const fontSelect = document.getElementById('docFontFamily');
      const current = family(computed.fontFamily);
      const match = [...fontSelect.options].find((option) => family(option.value) === current);
      fontSelect.value = match ? match.value : '';
    }
  }

  // Setup Document Ribbon Toolbar
  function initDocToolbar() {
    document.getElementById('docUndo').onclick = () => execDocCmd('undo');
    document.getElementById('docRedo').onclick = () => execDocCmd('redo');

    document.getElementById('docFormatSelect').onchange = (e) => {
      const val = e.target.value;
      if (val === 'p' || val.startsWith('h')) {
        execDocCmd('formatBlock', `<${val}>`);
      } else if (val === 'blockquote') {
        execDocCmd('formatBlock', '<blockquote>');
      } else if (val === 'pre') {
        execDocCmd('formatBlock', '<pre>');
      }
    };

    document.getElementById('docFontFamily').onchange = (e) => {
      execDocCmd('fontName', e.target.value);
    };

    document.getElementById('docFontSize').onchange = (e) => {
      execDocCmd('fontSize', e.target.value);
    };

    document.getElementById('docBold').onclick = () => execDocCmd('bold');
    document.getElementById('docItalic').onclick = () => execDocCmd('italic');
    document.getElementById('docUnderline').onclick = () => execDocCmd('underline');
    document.getElementById('docStrike').onclick = () => execDocCmd('strikeThrough');

    document.getElementById('docTextColor').oninput = (e) => {
      execDocCmd('foreColor', e.target.value);
    };

    document.getElementById('docHighlightColor').oninput = (e) => {
      execDocCmd('hiliteColor', e.target.value);
    };

    document.getElementById('docAlignLeft').onclick = () => execDocCmd('justifyLeft');
    document.getElementById('docAlignCenter').onclick = () => execDocCmd('justifyCenter');
    document.getElementById('docAlignRight').onclick = () => execDocCmd('justifyRight');
    document.getElementById('docAlignJustify').onclick = () => execDocCmd('justifyFull');

    document.getElementById('docBulletList').onclick = () => execDocCmd('insertUnorderedList');
    document.getElementById('docNumberList').onclick = () => execDocCmd('insertOrderedList');
    document.getElementById('docIndent').onclick = () => execDocCmd('indent');
    document.getElementById('docOutdent').onclick = () => execDocCmd('outdent');
    document.getElementById('docInsertHr').onclick = () => execDocCmd('insertHorizontalRule');
    document.getElementById('docClearFormat').onclick = () => execDocCmd('removeFormat');

    // Table Modal
    document.getElementById('docInsertTable').onclick = () => openModal(el.modalTable);
    document.getElementById('modalTableClose').onclick = () => closeModal(el.modalTable);
    document.getElementById('btnCancelTable').onclick = () => closeModal(el.modalTable);
    document.getElementById('btnConfirmTable').onclick = () => {
      const rows = parseInt(document.getElementById('tableRows').value, 10) || 3;
      const cols = parseInt(document.getElementById('tableCols').value, 10) || 3;
      closeModal(el.modalTable);
      restoreDocSelection();
      insertTable(rows, cols);
    };

    // Image Upload
    const imageInput = document.getElementById('imageInput');
    document.getElementById('docInsertImage').onclick = () => imageInput.click();
    imageInput.onchange = (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (re) => {
          execDocCmd('insertImage', re.target.result);
        };
        reader.readAsDataURL(file);
      }
      imageInput.value = '';
    };

    // Link Modal
    const linkUrl = document.getElementById('linkUrl');
    const linkText = document.getElementById('linkText');
    linkUrl.addEventListener('input', () => linkUrl.removeAttribute('aria-invalid'));

    document.getElementById('docInsertLink').onclick = () => {
      linkText.value = window.getSelection().toString();
      linkUrl.value = '';
      linkUrl.removeAttribute('aria-invalid');
      openModal(el.modalLink);
    };
    document.getElementById('modalLinkClose').onclick = () => closeModal(el.modalLink);
    document.getElementById('btnCancelLink').onclick = () => closeModal(el.modalLink);
    document.getElementById('btnConfirmLink').onclick = () => {
      const url = linkUrl.value.trim();
      const text = linkText.value.trim();
      if (!url) {
        linkUrl.setAttribute('aria-invalid', 'true');
        linkUrl.focus();
        return;
      }
      closeModal(el.modalLink);
      restoreDocSelection();
      if (window.getSelection().toString()) {
        execDocCmd('createLink', url);
      } else {
        execDocCmd('insertHTML', `<a href="${escapeXml(url)}" target="_blank">${escapeXml(text || url)}</a>`);
      }
    };

    document.addEventListener('selectionchange', syncDocToolbar);

    // Listen to changes in document
    el.documentPage.addEventListener('input', () => {
      updateDocStats();
      markDocModified();
    });

    // Paste handler to clean up Word/Office HTML
    el.documentPage.addEventListener('paste', (e) => {
      // Allow rich HTML or clean text
      setTimeout(() => {
        updateDocStats();
        markDocModified();
      }, 50);
    });
  }

  function insertTable(rows, cols) {
    let html = '<table border="1" style="width:100%; border-collapse:collapse; margin:14px 0;"><thead><tr>';
    for (let c = 0; c < cols; c++) {
      html += `<th style="border:1px solid #cbd5e1; padding:8px 12px; background:#f8fafc;">Header ${c + 1}</th>`;
    }
    html += '</tr></thead><tbody>';
    // The header row is one of the rows asked for.
    for (let r = 1; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        html += '<td style="border:1px solid #cbd5e1; padding:8px 12px;">Cell</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    execDocCmd('insertHTML', html);
  }

  // --- DOCX IMPORT (Mammoth or Server) ---
  async function loadDocxFile(file) {
    showToast('Opening document');
    const arrayBuffer = await file.arrayBuffer();
    const isZip = new Uint8Array(arrayBuffer, 0, Math.min(2, arrayBuffer.byteLength)).join() === '80,75';
    if (!isZip) {
      showToast(/\.doc$/i.test(file.name)
        ? 'OfficeLite cannot open the old .doc format. Save it as .docx in Word first.'
        : 'This file is not a valid .docx document.', 'error');
      return;
    }

    let html;
    let section = null;
    let warnings = [];
    try {
      ({ html, section, warnings } = await window.OfficeLiteDocx.docxToHtml(arrayBuffer));
    } catch (err) {
      // A plainer reader is better than no document, as long as the user is told.
      try {
        html = (await window.mammoth.convertToHtml({ arrayBuffer })).value;
        warnings = ['some formatting'];
      } catch (fallbackErr) {
        showToast('Could not open this .docx file: ' + err.message, 'error');
        return;
      }
    }

    el.documentPage.innerHTML = html || '';
    state.docSection = section;
    state.docTitle = file.name.replace(/\.[^/.]+$/, '');
    el.docNameInput.value = state.docTitle;
    updateDocStats();
    markSaved();
    if (warnings.length) showToast(`Document opened. OfficeLite cannot keep ${warnings.join(', ')}.`, 'warning');
    else showToast('Document opened', 'success');
  }

  // --- DOCX EXPORT ---
  async function exportDocx() {
    const filename = (state.docTitle || 'document').replace(/\.[^/.]+$/, '') + '.docx';
    try {
      const { blob, warnings } = await window.OfficeLiteDocx.htmlToDocx(el.documentPage, { section: state.docSection });
      await saveBlob(blob, filename);
      if (warnings.length) showToast(`${warnings.join(', ')}.`, 'warning');
    } catch (err) {
      showToast('Could not save this .docx file: ' + err.message, 'error');
    }
  }

  async function saveBlob(blob, filename) {
    if (!IS_NATIVE_APP) {
      downloadBlob(blob, filename);
      markSaved();
      showToast(`Downloaded ${filename}`, 'success');
      return;
    }

    const res = await fetch(`/api/save-file?name=${encodeURIComponent(filename)}`, {
      method: 'POST',
      body: await blob.arrayBuffer()
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'The file could not be written.');
    if (result.status !== 'saved') return;

    // The Save panel may have renamed the file.
    const title = result.name.replace(/\.[^/.]+$/, '');
    if (state.mode === 'doc') state.docTitle = title;
    else state.sheetTitle = title;
    el.docNameInput.value = title;

    markSaved();
    showToast(`Saved ${result.name}`, 'success');
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 200);
  }

  // ==========================================
  // --- SPREADSHEET (EXCEL) LOGIC ---
  // ==========================================

  const {
    Engine, parseInput, formatValue, shiftRelative, adjustReferences,
    renameSheetReferences, removeSheetReferences
  } = window.OfficeLiteFormula;

  // These match the sizes in styles.css. Fixed row heights let the grid draw only the visible rows.
  const ROW_HEIGHT = 24;
  const HEADER_HEIGHT = 24;
  const ROW_HEADER_WIDTH = 48;
  const DEFAULT_COL_WIDTH = 96;
  const MIN_COL_WIDTH = 32;
  const MAX_AUTO_COL_WIDTH = 320;
  const MIN_ROWS = 1000;
  const MIN_COLS = 26;
  const ROW_OVERSCAN = 20;
  const MAX_UNDO_STEPS = 100;
  const MAX_UNDO_BYTES = 64 * 1024 * 1024;
  const ERROR_VALUE = /^#[A-Z0-9/]+[!?A]$/;

  const formulas = new Engine((name) => {
    const match = state.sheet.sheetNames.find((n) => n.toLowerCase() === name.toLowerCase());
    return match ? state.sheet.sheets[match] : null;
  });

  let colOffsets = [0];
  let renderedRows = null;
  let cellEditor = null; // { input, r, c, typed }
  let formulaEdit = null; // { r, c, original }
  let refPick = null; // { input, start, end, anchor }
  let dragState = null;
  let gridClipboard = null; // { text, sheetName, range, matrix, cut }
  let undoStack = [];
  let redoStack = [];

  function getActiveSheetData() {
    return state.sheet.sheets[state.sheet.activeSheet] || null;
  }

  function normalizeSheet(sheet) {
    sheet.rowCount = Math.max(sheet.rowCount || 0, MIN_ROWS);
    sheet.colCount = Math.max(sheet.colCount || 0, MIN_COLS);
    if (!sheet.colWidths) sheet.colWidths = {};
  }

  function colWidth(sheet, c) {
    return sheet.colWidths[c] || DEFAULT_COL_WIDTH;
  }

  // --- Merged cells ---
  // A merge is stored as { r1, c1, r2, c2 }. Its value and style live in the upper left cell.
  function mergeAt(sheet, r, c) {
    if (!sheet || !sheet.merges) return null;
    return sheet.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;
  }

  function mergeRect(sheet, r, c) {
    const m = mergeAt(sheet, r, c);
    return m ? rectFor(m.r1, m.c1, m.r2, m.c2) : rectFor(r, c);
  }

  // A selection that touches a merged cell grows to cover all of it, as in Excel.
  function expandToMerges(sheet, range) {
    if (!sheet || !sheet.merges || !sheet.merges.length) return range;
    const out = { ...range };
    let grown = true;
    while (grown) {
      grown = false;
      for (const m of sheet.merges) {
        if (m.r1 > out.r2 || m.r2 < out.r1 || m.c1 > out.c2 || m.c2 < out.c1) continue;
        if (m.r1 < out.r1) { out.r1 = m.r1; grown = true; }
        if (m.c1 < out.c1) { out.c1 = m.c1; grown = true; }
        if (m.r2 > out.r2) { out.r2 = m.r2; grown = true; }
        if (m.c2 > out.c2) { out.c2 = m.c2; grown = true; }
      }
    }
    return out;
  }

  // True when the selection is one cell, or exactly one merged cell.
  function isSingleSelection(range) {
    const a = state.sheet.selectedCell;
    const m = mergeAt(getActiveSheetData(), a.r, a.c);
    const own = m || { r1: a.r, c1: a.c, r2: a.r, c2: a.c };
    return range.r1 === own.r1 && range.c1 === own.c1 && range.r2 === own.r2 && range.c2 === own.c2;
  }

  // The next cell in a direction, stepping over the whole of a merged cell.
  function stepFrom(r, c, dr, dc) {
    const m = mergeAt(getActiveSheetData(), r, c);
    if (m) {
      if (dr) r = dr > 0 ? m.r2 : m.r1;
      if (dc) c = dc > 0 ? m.c2 : m.c1;
    }
    return { r: r + dr, c: c + dc };
  }

  // Moves a merge when rows or columns are inserted or deleted. Returns null when it is gone.
  function shiftMerge(m, axis, at, count) {
    const [lo, hi] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
    let start = m[lo];
    let end = m[hi];
    if (count > 0) {
      if (start >= at) start += count;
      if (end >= at) end += count;
    } else {
      const removed = -count;
      const last = at + removed - 1;
      if (start > last) start -= removed;
      else if (start >= at) start = at;
      if (end > last) end -= removed;
      else if (end >= at) end = at - 1;
      if (end < start) return null;
    }
    const out = { ...m, [lo]: start, [hi]: end };
    return out.r1 === out.r2 && out.c1 === out.c2 ? null : out;
  }

  // --- Selection geometry ---
  function getRange() {
    const a = state.sheet.selectedCell;
    const e = state.sheet.rangeEnd || a;
    return expandToMerges(getActiveSheetData(), {
      r1: Math.min(a.r, e.r), c1: Math.min(a.c, e.c),
      r2: Math.max(a.r, e.r), c2: Math.max(a.c, e.c)
    });
  }

  function rangeLabel(range) {
    const start = coordToAddress(range.r1, range.c1);
    if (range.r1 === range.r2 && range.c1 === range.c2) return start;
    return `${start}:${coordToAddress(range.r2, range.c2)}`;
  }

  function rectFor(r1, c1, r2 = r1, c2 = c1) {
    return {
      left: ROW_HEADER_WIDTH + colOffsets[c1],
      top: HEADER_HEIGHT + r1 * ROW_HEIGHT,
      width: colOffsets[c2 + 1] - colOffsets[c1],
      height: (r2 - r1 + 1) * ROW_HEIGHT
    };
  }

  function placeBox(box, rect, visible = true) {
    box.hidden = !visible;
    if (!visible) return;
    box.style.transform = `translate(${rect.left}px, ${rect.top}px)`;
    box.style.width = `${rect.width}px`;
    box.style.height = `${rect.height}px`;
  }

  function cellAtPoint(clientX, clientY) {
    const sheet = getActiveSheetData();
    const bounds = el.gridContainer.getBoundingClientRect();
    const x = clientX - bounds.left + el.gridContainer.scrollLeft - ROW_HEADER_WIDTH;
    const y = clientY - bounds.top + el.gridContainer.scrollTop - HEADER_HEIGHT;

    let low = 0;
    let high = sheet.colCount - 1;
    while (low < high) {
      const mid = Math.ceil((low + high) / 2);
      if (colOffsets[mid] <= x) low = mid;
      else high = mid - 1;
    }
    return { r: Math.max(0, Math.min(sheet.rowCount - 1, Math.floor(y / ROW_HEIGHT))), c: low };
  }

  function scrollCellIntoView(r, c) {
    const box = el.gridContainer;
    const rect = rectFor(r, c);
    if (rect.top < box.scrollTop + HEADER_HEIGHT) box.scrollTop = rect.top - HEADER_HEIGHT;
    else if (rect.top + rect.height > box.scrollTop + box.clientHeight) box.scrollTop = rect.top + rect.height - box.clientHeight;
    if (rect.left < box.scrollLeft + ROW_HEADER_WIDTH) box.scrollLeft = rect.left - ROW_HEADER_WIDTH;
    else if (rect.left + rect.width > box.scrollLeft + box.clientWidth) box.scrollLeft = rect.left + rect.width - box.clientWidth;
    renderVisibleRows();
  }

  // --- Rendering ---
  function displayValue(sheet, addr, cell) {
    return cell.f ? formulas.value(sheet, addr) : (cell.v ?? '');
  }

  function isBlankForSpill(cell) {
    return !cell || ((cell.v === undefined || cell.v === '') && !cell.f && !(cell.s && cell.s.bg));
  }

  function cellLook(sheet, addr, cell) {
    const s = cell.s || {};
    const value = displayValue(sheet, addr, cell);
    const text = String(formatValue(value, s.format || 'general'));
    const classes = ['grid-cell'];

    let style = '';
    if (s.bold) style += 'font-weight:600;';
    if (s.italic) style += 'font-style:italic;';
    if (s.color) style += `color:${s.color};`;
    if (s.bg) style += `background-color:${s.bg};`;

    let align = s.align;
    if (!align) {
      if (typeof value === 'number') align = 'right';
      else if (typeof value === 'boolean' || (cell.f && ERROR_VALUE.test(text))) align = 'center';
    }
    if (align) style += `text-align:${align};`;
    if (cell.f && ERROR_VALUE.test(text)) classes.push('is-error');
    return { value, text, align, style, classes };
  }

  function cellHtml(sheet, r, c) {
    const addr = coordToAddress(r, c);
    const cell = sheet.cells[addr];
    if (!cell) return `<td class="grid-cell" data-row="${r}" data-col="${c}" data-addr="${addr}"></td>`;

    const { value, text, align, style, classes } = cellLook(sheet, addr, cell);

    // renderMerges draws merged cells on screen. The text here is only for printing.
    if (mergeAt(sheet, r, c)) {
      return `<td class="grid-cell merge-anchor" data-row="${r}" data-col="${c}" data-addr="${addr}" style="${style}">` +
        `<div class="cell-content">${escapeXml(text)}</div></td>`;
    }

    // Long text runs into empty cells on its right, as it does in Excel.
    const isLeftText = typeof value === 'string' && (!align || align === 'left');
    if (isLeftText && text.length * 7 > colWidth(sheet, c) - 16 && isBlankForSpill(sheet.cells[coordToAddress(r, c + 1)])) {
      classes.push('spills');
    }

    return `<td class="${classes.join(' ')}" data-row="${r}" data-col="${c}" data-addr="${addr}" style="${style}">` +
      `<div class="cell-content">${escapeXml(text)}</div></td>`;
  }

  function renderRowsInto(sheet, first, last) {
    const spacer = (rows) => `<tr class="grid-spacer"><td colspan="${sheet.colCount + 1}" style="height:${rows * ROW_HEIGHT}px"></td></tr>`;
    let html = first > 0 ? spacer(first) : '';
    for (let r = first; r <= last; r++) {
      html += `<tr data-r="${r}"><th class="row-header" data-row="${r}">${r + 1}</th>`;
      for (let c = 0; c < sheet.colCount; c++) html += cellHtml(sheet, r, c);
      html += '</tr>';
    }
    if (last < sheet.rowCount - 1) html += spacer(sheet.rowCount - 1 - last);

    document.getElementById('gridBody').innerHTML = html;
    renderedRows = { first, last };
  }

  function renderVisibleRows(force = false) {
    const sheet = getActiveSheetData();
    const body = document.getElementById('gridBody');
    if (!sheet || !body) return;

    const box = el.gridContainer;
    const viewport = box.clientHeight || 800;
    const firstVisible = Math.floor(Math.max(0, box.scrollTop - HEADER_HEIGHT) / ROW_HEIGHT);
    const lastVisible = Math.min(sheet.rowCount - 1, Math.ceil((box.scrollTop + viewport) / ROW_HEIGHT));
    if (!force && renderedRows && firstVisible >= renderedRows.first && lastVisible <= renderedRows.last) return;

    renderRowsInto(sheet, Math.max(0, firstVisible - ROW_OVERSCAN), Math.min(sheet.rowCount - 1, lastVisible + ROW_OVERSCAN));
    paintCellTints();
  }

  function renderGrid() {
    const sheet = getActiveSheetData();
    if (!sheet) return;

    normalizeSheet(sheet);
    formulas.reset();

    colOffsets = [0];
    let cols = `<col style="width:${ROW_HEADER_WIDTH}px">`;
    let headers = '<th class="grid-corner" title="Select all"></th>';
    for (let c = 0; c < sheet.colCount; c++) {
      const width = colWidth(sheet, c);
      colOffsets.push(colOffsets[c] + width);
      cols += `<col data-col="${c}" style="width:${width}px">`;
      headers += `<th class="col-header" data-col="${c}">${indexToCol(c)}<span class="col-resize" title="Drag to resize. Double click to fit."></span></th>`;
    }

    el.spreadsheetGrid.style.width = `${ROW_HEADER_WIDTH + colOffsets[sheet.colCount]}px`;
    el.spreadsheetGrid.innerHTML = `<colgroup>${cols}</colgroup><thead><tr>${headers}</tr></thead><tbody id="gridBody"></tbody>`;

    const active = state.sheet.selectedCell;
    active.r = Math.min(active.r, sheet.rowCount - 1);
    active.c = Math.min(active.c, sheet.colCount - 1);
    if (state.sheet.rangeEnd) {
      state.sheet.rangeEnd.r = Math.min(state.sheet.rangeEnd.r, sheet.rowCount - 1);
      state.sheet.rangeEnd.c = Math.min(state.sheet.rangeEnd.c, sheet.colCount - 1);
    }

    renderVisibleRows(true);
    renderMerges(sheet);
    paintSelection();
    updateFormulaBar();
    updateSheetStats();
    renderSheetTabs();
    syncUndoButtons();
  }

  // Merged cells are drawn over the grid, so drawing only the visible rows cannot split them.
  function renderMerges(sheet) {
    el.gridContainer.querySelectorAll('.merge-cell').forEach((node) => node.remove());
    const before = document.getElementById('selectionRange');
    for (const m of sheet.merges || []) {
      if (m.r1 >= sheet.rowCount || m.c1 >= sheet.colCount) continue;
      const box = document.createElement('div');
      box.className = 'merge-cell';
      const addr = coordToAddress(m.r1, m.c1);
      const cell = sheet.cells[addr];
      if (cell) {
        const look = cellLook(sheet, addr, cell);
        box.className = ['merge-cell', ...look.classes.filter((name) => name !== 'grid-cell')].join(' ');
        box.style.cssText = look.style;
        const content = document.createElement('div');
        content.className = 'cell-content';
        content.textContent = look.text;
        box.appendChild(content);
      }
      box.dataset.merge = `${m.r1},${m.c1},${m.r2},${m.c2}`;
      placeBox(box, rectFor(m.r1, m.c1, Math.min(m.r2, sheet.rowCount - 1), Math.min(m.c2, sheet.colCount - 1)));
      el.gridContainer.insertBefore(box, before);
    }
  }

  function paintCellTints() {
    if (!renderedRows) return;
    const body = document.getElementById('gridBody');
    body.querySelectorAll('.in-range').forEach((td) => td.classList.remove('in-range'));
    body.querySelectorAll('.is-active').forEach((th) => th.classList.remove('is-active'));

    const range = getRange();
    const active = state.sheet.selectedCell;
    const single = isSingleSelection(range);

    el.gridContainer.querySelectorAll('.merge-cell').forEach((box) => {
      const [r1, c1, r2, c2] = box.dataset.merge.split(',').map(Number);
      const inside = !single && r1 >= range.r1 && r2 <= range.r2 && c1 >= range.c1 && c2 <= range.c2;
      const isActive = r1 === active.r && c1 === active.c;
      box.classList.toggle('in-range', inside && !isActive);
    });

    for (const row of body.children) {
      if (row.dataset.r === undefined) continue;
      const r = Number(row.dataset.r);
      if (r < range.r1 || r > range.r2) continue;
      row.firstElementChild.classList.add('is-active');
      if (single) continue;
      for (let c = range.c1; c <= range.c2; c++) {
        if (r !== active.r || c !== active.c) row.children[c + 1].classList.add('in-range');
      }
    }
  }

  function paintSelection() {
    const range = getRange();
    const active = state.sheet.selectedCell;
    const single = isSingleSelection(range);

    placeBox(document.getElementById('selectionRange'), rectFor(range.r1, range.c1, range.r2, range.c2), !single);
    placeBox(document.getElementById('selectionActive'), mergeRect(getActiveSheetData(), active.r, active.c));

    const corner = rectFor(range.r2, range.c2);
    const handle = document.getElementById('fillHandle');
    handle.style.transform = `translate(${corner.left + corner.width - 4}px, ${corner.top + corner.height - 4}px)`;

    paintCopyBox();
    paintCellTints();

    el.spreadsheetGrid.querySelectorAll('.col-header.is-active').forEach((th) => th.classList.remove('is-active'));
    const headers = el.spreadsheetGrid.querySelectorAll('.col-header');
    for (let c = range.c1; c <= range.c2; c++) headers[c].classList.add('is-active');

    el.cellAddress.value = single ? coordToAddress(active.r, active.c) : rangeLabel(range);
  }

  function paintCopyBox() {
    const box = document.getElementById('copyBox');
    if (refPick) {
      const { anchor, end } = refPick;
      placeBox(box, rectFor(Math.min(anchor.r, end.r), Math.min(anchor.c, end.c), Math.max(anchor.r, end.r), Math.max(anchor.c, end.c)));
    } else if (dragState && dragState.type === 'fill' && dragState.target) {
      const t = dragState.target;
      placeBox(box, rectFor(t.r1, t.c1, t.r2, t.c2));
    } else if (gridClipboard && gridClipboard.sheetName === state.sheet.activeSheet) {
      const g = gridClipboard.range;
      placeBox(box, rectFor(g.r1, g.c1, g.r2, g.c2));
    } else {
      box.hidden = true;
    }
  }

  function selectRange(anchorR, anchorC, endR, endC) {
    state.sheet.selectedCell = { r: anchorR, c: anchorC };
    state.sheet.rangeEnd = { r: endR, c: endC };
    paintSelection();
    updateFormulaBar();
    updateSheetStats();
  }

  function selectCell(r, c, { extend = false, scroll = true } = {}) {
    const sheet = getActiveSheetData();
    let row = Math.max(0, Math.min(sheet.rowCount - 1, r));
    let col = Math.max(0, Math.min(sheet.colCount - 1, c));
    const merge = !extend && mergeAt(sheet, row, col);
    if (merge) {
      row = merge.r1;
      col = merge.c1;
    }
    if (extend) {
      const a = state.sheet.selectedCell;
      selectRange(a.r, a.c, row, col);
    } else {
      selectRange(row, col, row, col);
    }
    if (scroll) scrollCellIntoView(row, col);
  }

  function rawInput(cell) {
    if (!cell) return '';
    return String(cell.f ?? cell.v ?? '');
  }

  function updateFormulaBar() {
    if (formulaEdit || cellEditor) return;
    const sheet = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const cell = (sheet && sheet.cells[coordToAddress(r, c)]) || null;
    el.formulaInput.value = rawInput(cell);
    syncSheetToolbar(cell || {});
  }

  // Reflect the active cell in the toolbar.
  function syncSheetToolbar(cell) {
    const style = cell.s || {};
    setPressed('sheetBold', !!style.bold);
    setPressed('sheetItalic', !!style.italic);
    setPressed('sheetAlignLeft', style.align === 'left');
    setPressed('sheetAlignCenter', style.align === 'center');
    setPressed('sheetAlignRight', style.align === 'right');
    document.getElementById('sheetNumberFormat').value = style.format || 'general';
    const { r, c } = state.sheet.selectedCell;
    setPressed('sheetMerge', !!mergeAt(getActiveSheetData(), r, c));
  }

  function updateSheetStats() {
    const sheet = getActiveSheetData();
    if (!sheet) return;
    const range = getRange();
    const single = range.r1 === range.r2 && range.c1 === range.c2;

    if (single) {
      const addr = coordToAddress(range.r1, range.c1);
      const cell = sheet.cells[addr];
      const value = cell ? displayValue(sheet, addr, cell) : '';
      const filled = Object.values(sheet.cells).filter((item) => rawInput(item) !== '').length;
      el.sheetStatCell.textContent = `Cell: ${addr}`;
      el.sheetStatSum.textContent = `Value: ${value === '' ? '(empty)' : formatValue(value, (cell.s && cell.s.format) || 'general')}`;
      el.sheetStatCount.textContent = `Cells filled: ${filled}`;
      return;
    }

    let sum = 0;
    let numbers = 0;
    let count = 0;
    forEachExistingCell(sheet, range, (addr, cell) => {
      const value = displayValue(sheet, addr, cell);
      if (value === '') return;
      count++;
      if (typeof value === 'number') {
        sum += value;
        numbers++;
      }
    });
    const pretty = (n) => n.toLocaleString('en-US', { maximumFractionDigits: 2 });
    el.sheetStatCell.textContent = numbers ? `Average: ${pretty(sum / numbers)}` : 'Average: none';
    el.sheetStatSum.textContent = `Sum: ${pretty(sum)}`;
    el.sheetStatCount.textContent = `Count: ${count}`;
  }

  function forEachExistingCell(sheet, range, callback) {
    const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
    const keys = Object.keys(sheet.cells);
    if (area > keys.length) {
      keys.forEach((addr) => {
        const { r, c } = addressToCoord(addr);
        if (r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2) callback(addr, sheet.cells[addr], r, c);
      });
      return;
    }
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const addr = coordToAddress(r, c);
        if (sheet.cells[addr]) callback(addr, sheet.cells[addr], r, c);
      }
    }
  }

  // --- Undo ---
  function snapshot() {
    const { sheetNames, sheets, activeSheet, selectedCell, rangeEnd } = state.sheet;
    return JSON.stringify({ sheetNames, sheets, activeSheet, selectedCell, rangeEnd });
  }

  function pushUndo() {
    undoStack.push(snapshot());
    redoStack = [];
    let bytes = undoStack.reduce((total, entry) => total + entry.length, 0);
    while (undoStack.length > 1 && (undoStack.length > MAX_UNDO_STEPS || bytes > MAX_UNDO_BYTES)) {
      bytes -= undoStack.shift().length;
    }
  }

  function restoreSnapshot(json) {
    Object.assign(state.sheet, JSON.parse(json));
    gridClipboard = null;
    markSheetModified();
    renderGrid();
    const { r, c } = state.sheet.selectedCell;
    scrollCellIntoView(r, c);
  }

  function undoSheet() {
    commitPendingEdit();
    if (!undoStack.length) return;
    redoStack.push(snapshot());
    restoreSnapshot(undoStack.pop());
  }

  function redoSheet() {
    commitPendingEdit();
    if (!redoStack.length) return;
    undoStack.push(snapshot());
    restoreSnapshot(redoStack.pop());
  }

  function syncUndoButtons() {
    document.getElementById('sheetUndo').disabled = !undoStack.length;
    document.getElementById('sheetRedo').disabled = !redoStack.length;
  }

  // Forgets edit history and view state when a different workbook is loaded.
  function resetSheetView() {
    undoStack = [];
    redoStack = [];
    gridClipboard = null;
    state.sheet.selectedCell = { r: 0, c: 0 };
    state.sheet.rangeEnd = { r: 0, c: 0 };
    el.gridContainer.scrollTop = 0;
    el.gridContainer.scrollLeft = 0;
  }

  /** Runs one undoable change to the workbook and redraws. */
  function mutate(change, { keepClipboard = false } = {}) {
    pushUndo();
    change(getActiveSheetData());
    if (!keepClipboard) gridClipboard = null;
    markSheetModified();
    renderGrid();
  }

  // --- Cell values ---
  function writeCellInput(sheet, r, c, input) {
    const addr = coordToAddress(r, c);
    const text = input === null || input === undefined ? '' : String(input);
    const existing = sheet.cells[addr];

    if (text.trim() === '') {
      if (!existing) return;
      if (existing.s && Object.keys(existing.s).length) {
        delete existing.f;
        delete existing.v;
      } else {
        delete sheet.cells[addr];
      }
      return;
    }

    const cell = existing || (sheet.cells[addr] = { s: {} });
    if (!cell.s) cell.s = {};

    if (text.length > 1 && text.startsWith('=')) {
      cell.f = text;
      delete cell.v;
      return;
    }

    delete cell.f;
    if (cell.s.format === 'text') {
      cell.v = text;
      return;
    }
    const parsed = parseInput(text);
    cell.v = parsed.value;
    if (parsed.format && (!cell.s.format || cell.s.format === 'general')) cell.s.format = parsed.format;
  }

  function setCellVal(r, c, inputVal) {
    const sheet = getActiveSheetData();
    if (!sheet || rawInput(sheet.cells[coordToAddress(r, c)]) === String(inputVal ?? '')) {
      renderGrid();
      return;
    }
    mutate((current) => writeCellInput(current, r, c, inputVal));
  }

  function clearSelection() {
    const sheet = getActiveSheetData();
    const range = getRange();
    let hasContent = false;
    forEachExistingCell(sheet, range, (addr, cell) => {
      if (rawInput(cell) !== '') hasContent = true;
    });
    if (!hasContent) return;
    mutate((current) => {
      const targets = [];
      forEachExistingCell(current, range, (addr, cell, r, c) => targets.push([r, c]));
      targets.forEach(([r, c]) => writeCellInput(current, r, c, ''));
    });
  }

  // --- Print ---
  // Draws every row with data and keeps empty rows and columns off the printed page.
  function markPrintRange() {
    const sheet = getActiveSheetData();
    let lastRow = -1;
    let lastCol = -1;
    Object.keys(sheet.cells).forEach((addr) => {
      const { r, c } = addressToCoord(addr);
      if (rawInput(sheet.cells[addr]) === '' && !(sheet.cells[addr].s && sheet.cells[addr].s.bg)) return;
      lastRow = Math.max(lastRow, r);
      lastCol = Math.max(lastCol, c);
    });

    const visibleLast = renderedRows ? renderedRows.last : 0;
    renderRowsInto(sheet, 0, Math.max(lastRow, visibleLast));
    paintCellTints();

    el.spreadsheetGrid.querySelectorAll('#gridBody tr').forEach((row) => {
      row.classList.toggle('print-hidden', row.dataset.r === undefined || Number(row.dataset.r) > lastRow);
    });
    el.spreadsheetGrid.querySelectorAll('[data-col]').forEach((cell) => {
      cell.classList.toggle('print-hidden', Number(cell.dataset.col) > lastCol);
    });
  }

  // --- Editing: the cell editor and formula bar ---
  function activeEditInput() {
    if (cellEditor) return cellEditor.input;
    return document.activeElement === el.formulaInput ? el.formulaInput : null;
  }

  // Clicking a cell inserts its address when the caret sits where a reference can go.
  function canPickReference(input) {
    if (!input.value.startsWith('=')) return false;
    if (refPick && refPick.input === input) return true;
    const before = input.value.slice(0, input.selectionStart);
    return /[=(,;+\-*/^&<>:]\s*$/.test(before);
  }

  function writePickedReference(endCell) {
    const { input, start, anchor } = refPick;
    const label = rangeLabel({
      r1: Math.min(anchor.r, endCell.r), c1: Math.min(anchor.c, endCell.c),
      r2: Math.max(anchor.r, endCell.r), c2: Math.max(anchor.c, endCell.c)
    });
    const after = refPick.end === undefined ? input.selectionStart : refPick.textEnd;
    input.value = input.value.slice(0, start) + label + input.value.slice(after);
    refPick.textEnd = start + label.length;
    refPick.end = endCell;
    input.setSelectionRange(refPick.textEnd, refPick.textEnd);
    if (input !== el.formulaInput) el.formulaInput.value = input.value;
    paintCopyBox();
  }

  function clearRefPick() {
    if (!refPick) return;
    refPick = null;
    paintCopyBox();
  }

  function startCellEdit(initialText = null) {
    commitPendingEdit();
    const sheet = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    scrollCellIntoView(r, c);

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-editor';
    input.spellcheck = false;
    input.setAttribute('aria-label', `Edit ${coordToAddress(r, c)}`);
    input.value = initialText !== null ? initialText : rawInput(sheet.cells[coordToAddress(r, c)]);

    const rect = mergeRect(sheet, r, c);
    input.style.transform = `translate(${rect.left}px, ${rect.top}px)`;
    input.style.height = `${rect.height}px`;
    const fit = () => {
      input.style.width = `${rect.width}px`;
      input.style.width = `${Math.max(rect.width, Math.min(input.scrollWidth + 16, 640))}px`;
    };

    cellEditor = { input, r, c, typed: initialText !== null };
    el.gridContainer.appendChild(input);
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    el.formulaInput.value = input.value;
    fit();

    input.addEventListener('input', () => {
      refPick = null;
      el.formulaInput.value = input.value;
      fit();
      paintCopyBox();
    });
    // Switching to another app also blurs the input, but that must not end the edit.
    input.addEventListener('blur', () => {
      if (document.hasFocus()) finishCellEdit(true);
    });
    input.addEventListener('keydown', (e) => {
      const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (e.key === 'Enter') {
        e.preventDefault();
        finishCellEdit(true);
        const to = stepFrom(r, c, e.shiftKey ? -1 : 1, 0);
        selectCell(to.r, to.c);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        finishCellEdit(true);
        const to = stepFrom(r, c, 0, e.shiftKey ? -1 : 1);
        selectCell(to.r, to.c);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finishCellEdit(false);
      } else if (moves[e.key] && cellEditor && cellEditor.typed && !input.value.startsWith('=')) {
        e.preventDefault();
        finishCellEdit(true);
        const to = stepFrom(r, c, moves[e.key][0], moves[e.key][1]);
        selectCell(to.r, to.c);
      }
    });
  }

  function finishCellEdit(commit) {
    if (!cellEditor) return;
    const { input, r, c } = cellEditor;
    cellEditor = null;
    const value = input.value;
    input.remove();
    clearRefPick();
    if (commit) setCellVal(r, c, value);
    else updateFormulaBar();
  }

  function commitFormulaBar() {
    if (!formulaEdit) return;
    const { r, c, original } = formulaEdit;
    formulaEdit = null;
    clearRefPick();
    if (el.formulaInput.value !== original) setCellVal(r, c, el.formulaInput.value);
    else updateFormulaBar();
  }

  function commitPendingEdit() {
    finishCellEdit(true);
    commitFormulaBar();
  }

  // --- Applying formatting to the selection ---
  function setCellFormatProperty(prop, value) {
    const range = getRange();
    mutate((sheet) => {
      for (let r = range.r1; r <= range.r2; r++) {
        for (let c = range.c1; c <= range.c2; c++) {
          const addr = coordToAddress(r, c);
          const cell = sheet.cells[addr] || (sheet.cells[addr] = { s: {} });
          if (!cell.s) cell.s = {};
          if (value === false || value === '') delete cell.s[prop];
          else cell.s[prop] = value;
        }
      }
    });
  }

  function toggleCellFormat(prop) {
    const sheet = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const cell = sheet.cells[coordToAddress(r, c)];
    setCellFormatProperty(prop, !(cell && cell.s && cell.s[prop]));
  }

  // --- Rows and columns ---
  function changeStructure(axis, at, count) {
    const owner = state.sheet.activeSheet;
    mutate((sheet) => {
      const moved = {};
      const key = axis === 'row' ? 'r' : 'c';
      const removed = -count;
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        const coord = addressToCoord(addr);
        if (count > 0) {
          if (coord[key] >= at) coord[key] += count;
        } else if (coord[key] >= at + removed) {
          coord[key] -= removed;
        } else if (coord[key] >= at) {
          continue;
        }
        moved[coordToAddress(coord.r, coord.c)] = cell;
      }
      sheet.cells = moved;
      if (sheet.merges) {
        sheet.merges = sheet.merges.map((m) => shiftMerge(m, axis, at, count)).filter(Boolean);
      }

      if (axis === 'row') {
        sheet.rowCount += count;
      } else {
        sheet.colCount += count;
        const widths = {};
        for (const [index, width] of Object.entries(sheet.colWidths)) {
          const c = Number(index);
          if (count > 0) widths[c >= at ? c + count : c] = width;
          else if (c >= at + removed) widths[c - removed] = width;
          else if (c < at) widths[c] = width;
        }
        sheet.colWidths = widths;
      }

      // Formulas on every sheet can point at what just moved.
      state.sheet.sheetNames.forEach((name) => {
        Object.values(state.sheet.sheets[name].cells).forEach((cell) => {
          if (cell.f) cell.f = adjustReferences(cell.f, name, owner, axis, at, count);
        });
      });
    });
  }

  function insertRows(position) {
    const range = getRange();
    const count = range.r2 - range.r1 + 1;
    changeStructure('row', position === 'above' ? range.r1 : range.r2 + 1, count);
  }

  function insertCols(position) {
    const range = getRange();
    const count = range.c2 - range.c1 + 1;
    changeStructure('col', position === 'left' ? range.c1 : range.c2 + 1, count);
  }

  function deleteRows() {
    const range = getRange();
    changeStructure('row', range.r1, -(range.r2 - range.r1 + 1));
    selectCell(range.r1, state.sheet.selectedCell.c);
  }

  function deleteCols() {
    const range = getRange();
    changeStructure('col', range.c1, -(range.c2 - range.c1 + 1));
    selectCell(state.sheet.selectedCell.r, range.c1);
  }

  function mergesInside(sheet, range) {
    return (sheet.merges || []).filter((m) =>
      m.r1 >= range.r1 && m.r2 <= range.r2 && m.c1 >= range.c1 && m.c2 <= range.c2);
  }

  /** Merges the selected cells, or unmerges them when the selection holds a merged cell. */
  function toggleMerge() {
    commitPendingEdit();
    const sheet = getActiveSheetData();
    const range = getRange();
    const existing = mergesInside(sheet, range);

    if (existing.length) {
      mutate((current) => {
        current.merges = current.merges.filter((m) => !existing.includes(m));
      });
      selectRange(range.r1, range.c1, range.r2, range.c2);
      return;
    }
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      showToast('Select two or more cells to merge them.');
      return;
    }

    let otherValues = false;
    forEachExistingCell(sheet, range, (addr, cell, r, c) => {
      if ((r !== range.r1 || c !== range.c1) && rawInput(cell) !== '') otherValues = true;
    });

    const merge = () => {
      mutate((current) => {
        forEachExistingCell(current, range, (addr, cell, r, c) => {
          if (r !== range.r1 || c !== range.c1) delete current.cells[addr];
        });
        current.merges = (current.merges || []).concat([{ ...range }]);
      });
      selectCell(range.r1, range.c1, { scroll: false });
    };

    if (!otherValues) {
      merge();
      return;
    }
    confirmDialog({
      title: 'Merge these cells?',
      message: 'The merged cell keeps only the value in the upper left cell. The other values are removed.',
      confirmLabel: 'Merge cells'
    }).then((confirmed) => {
      if (confirmed) merge();
    });
  }

  function fittedColWidth(sheet, c) {
    let longest = 0;
    for (const [addr, cell] of Object.entries(sheet.cells)) {
      const coord = addressToCoord(addr);
      if (coord.c !== c) continue;
      const merge = mergeAt(sheet, coord.r, coord.c);
      if (merge && merge.c2 > merge.c1) continue;
      const text = String(formatValue(displayValue(sheet, addr, cell), (cell.s && cell.s.format) || 'general'));
      const weight = cell.s && cell.s.bold ? 8.5 : 8;
      longest = Math.max(longest, text.length * weight);
    }
    return Math.max(DEFAULT_COL_WIDTH, Math.min(MAX_AUTO_COL_WIDTH, Math.ceil(longest + 20)));
  }

  function autoSizeColumns(sheet) {
    normalizeSheet(sheet);
    const used = new Set(Object.keys(sheet.cells).map((addr) => addressToCoord(addr).c));
    used.forEach((c) => {
      if (!sheet.colWidths[c]) sheet.colWidths[c] = fittedColWidth(sheet, c);
    });
  }

  // --- Spreadsheet clipboard ---
  function gridHasFocus() {
    if (state.mode !== 'sheet' || cellEditor || formulaEdit) return false;
    if (document.querySelector('.modal-overlay.active')) return false;
    const active = document.activeElement;
    return !(active && active.closest('input, select, textarea, [contenteditable="true"]'));
  }

  function copySelection(clipboardData, cut) {
    const sheet = getActiveSheetData();
    const range = getRange();
    const lines = [];
    const matrix = [];
    for (let r = range.r1; r <= range.r2; r++) {
      const texts = [];
      const row = [];
      for (let c = range.c1; c <= range.c2; c++) {
        const addr = coordToAddress(r, c);
        const cell = sheet.cells[addr];
        row.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
        texts.push(cell ? String(formatValue(displayValue(sheet, addr, cell), (cell.s && cell.s.format) || 'general')) : '');
      }
      lines.push(texts.join('\t'));
      matrix.push(row);
    }
    const text = lines.join('\n');
    clipboardData.setData('text/plain', text);
    gridClipboard = { text, matrix, range, cut, sheetName: state.sheet.activeSheet };
    paintCopyBox();
  }

  function pasteIntoGrid(clipboardData) {
    const text = clipboardData.getData('text/plain');
    if (!text) return;
    const target = getRange();
    const source = gridClipboard && gridClipboard.text === text ? gridClipboard : null;

    mutate((sheet) => {
      if (source) {
        const rows = source.matrix.length;
        const cols = source.matrix[0].length;
        const oneCell = rows === 1 && cols === 1;
        const lastR = oneCell ? target.r2 : target.r1 + rows - 1;
        const lastC = oneCell ? target.c2 : target.c1 + cols - 1;

        if (source.cut) {
          const from = state.sheet.sheets[source.sheetName];
          if (from) forEachExistingCell(from, source.range, (addr) => delete from.cells[addr]);
        }

        for (let r = target.r1; r <= lastR; r++) {
          for (let c = target.c1; c <= lastC; c++) {
            const addr = coordToAddress(r, c);
            const original = source.matrix[oneCell ? 0 : r - target.r1][oneCell ? 0 : c - target.c1];
            if (!original) {
              delete sheet.cells[addr];
              continue;
            }
            const cell = JSON.parse(JSON.stringify(original));
            if (cell.f && !source.cut) {
              const fromR = source.range.r1 + (oneCell ? 0 : r - target.r1);
              const fromC = source.range.c1 + (oneCell ? 0 : c - target.c1);
              cell.f = shiftRelative(cell.f, r - fromR, c - fromC);
            }
            sheet.cells[addr] = cell;
          }
        }
        sheet.rowCount = Math.max(sheet.rowCount, lastR + 1);
        sheet.colCount = Math.max(sheet.colCount, lastC + 1);
        state.sheet.selectedCell = { r: target.r1, c: target.c1 };
        state.sheet.rangeEnd = { r: lastR, c: lastC };
      } else {
        // Rows and columns from Excel, Numbers, or a web table arrive as tab separated text.
        const rows = text.replace(/\r?\n$/, '').split(/\r?\n/).map((row) => row.split('\t'));
        const width = Math.max(...rows.map((row) => row.length));
        rows.forEach((row, rowOffset) => {
          row.forEach((value, colOffset) => writeCellInput(sheet, target.r1 + rowOffset, target.c1 + colOffset, value));
        });
        sheet.rowCount = Math.max(sheet.rowCount, target.r1 + rows.length);
        sheet.colCount = Math.max(sheet.colCount, target.c1 + width);
        state.sheet.selectedCell = { r: target.r1, c: target.c1 };
        state.sheet.rangeEnd = { r: target.r1 + rows.length - 1, c: target.c1 + width - 1 };
      }
    }, { keepClipboard: !!source && !source.cut });
  }

  function initGridClipboard() {
    // WebKit disables Cut, Copy, and Paste when nothing is selected unless these are cancelled.
    ['beforecut', 'beforecopy', 'beforepaste'].forEach((type) => {
      document.addEventListener(type, (e) => {
        if (gridHasFocus()) e.preventDefault();
      });
    });

    document.addEventListener('copy', (e) => {
      if (!gridHasFocus()) return;
      e.preventDefault();
      copySelection(e.clipboardData, false);
    });

    document.addEventListener('cut', (e) => {
      if (!gridHasFocus()) return;
      e.preventDefault();
      copySelection(e.clipboardData, true);
    });

    document.addEventListener('paste', (e) => {
      if (!gridHasFocus()) return;
      e.preventDefault();
      pasteIntoGrid(e.clipboardData);
    });
  }

  // --- Fill handle ---
  function numericSeriesStep(values) {
    if (values.length < 2 || values.some((v) => typeof v !== 'number')) return null;
    const step = values[1] - values[0];
    return values.every((v, i) => i === 0 || Math.abs(v - values[i - 1] - step) < 1e-9) ? step : null;
  }

  function applyFill(source, target) {
    mutate((sheet) => {
      const vertical = target.c1 === source.c1 && target.c2 === source.c2;
      const sourceCell = (r, c) => sheet.cells[coordToAddress(r, c)];
      const lanes = vertical ? [source.c1, source.c2] : [source.r1, source.r2];

      for (let lane = lanes[0]; lane <= lanes[1]; lane++) {
        const length = vertical ? source.r2 - source.r1 + 1 : source.c2 - source.c1 + 1;
        const start = vertical ? source.r1 : source.c1;
        const cellAt = (i) => (vertical ? sourceCell(start + i, lane) : sourceCell(lane, start + i));
        const values = [];
        for (let i = 0; i < length; i++) {
          const cell = cellAt(i);
          values.push(cell && !cell.f ? cell.v : undefined);
        }
        const step = numericSeriesStep(values);

        const from = vertical ? target.r1 : target.c1;
        const to = vertical ? target.r2 : target.c2;
        const snapshotCells = values.map((_, i) => {
          const cell = cellAt(i);
          return cell ? JSON.parse(JSON.stringify(cell)) : null;
        });

        for (let pos = from; pos <= to; pos++) {
          if (pos >= start && pos < start + length) continue;
          const offset = pos - start;
          const index = ((offset % length) + length) % length;
          const addr = vertical ? coordToAddress(pos, lane) : coordToAddress(lane, pos);
          const template = snapshotCells[index];
          if (!template) {
            delete sheet.cells[addr];
            continue;
          }
          const cell = JSON.parse(JSON.stringify(template));
          if (cell.f) {
            const distance = pos - (start + index);
            cell.f = vertical ? shiftRelative(cell.f, distance, 0) : shiftRelative(cell.f, 0, distance);
          } else if (step !== null) {
            cell.v = values[0] + step * offset;
          }
          sheet.cells[addr] = cell;
        }
      }
      state.sheet.selectedCell = { r: target.r1, c: target.c1 };
      state.sheet.rangeEnd = { r: target.r2, c: target.c2 };
    });
  }

  // --- AutoSum ---
  function runAutoSum() {
    commitPendingEdit();
    const sheet = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    formulas.reset();
    const isNumber = (row, col) => {
      const addr = coordToAddress(row, col);
      const cell = sheet.cells[addr];
      return !!cell && typeof displayValue(sheet, addr, cell) === 'number';
    };

    let top = r;
    while (top > 0 && isNumber(top - 1, c)) top--;
    if (top < r) {
      setCellVal(r, c, `=SUM(${coordToAddress(top, c)}:${coordToAddress(r - 1, c)})`);
      return;
    }
    let left = c;
    while (left > 0 && isNumber(r, left - 1)) left--;
    if (left < c) {
      setCellVal(r, c, `=SUM(${coordToAddress(r, left)}:${coordToAddress(r, c - 1)})`);
      return;
    }
    startCellEdit('=SUM(');
  }

  // --- Sheet tabs ---
  function renderSheetTabs() {
    el.tabsList.innerHTML = '';
    state.sheet.sheetNames.forEach(name => {
      const isActive = name === state.sheet.activeSheet;
      const tab = document.createElement('div');
      tab.className = `sheet-tab ${isActive ? 'active' : ''}`;
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
      tab.tabIndex = 0;
      tab.title = 'Double click to rename';
      tab.innerHTML = `<span class="sheet-tab-name">${escapeXml(name)}</span>` +
        (state.sheet.sheetNames.length > 1
          ? `<button class="sheet-tab-close" tabindex="-1" aria-label="Delete ${escapeXml(name)}" title="Delete sheet">${icon('x')}</button>`
          : '');

      tab.onclick = (e) => {
        if (e.target.closest('.sheet-tab-close')) {
          e.stopPropagation();
          deleteSheet(name);
          return;
        }
        if (e.target.closest('.sheet-tab-input')) return;
        // Rebuilding the active tab here would swallow the double click that renames it.
        if (name !== state.sheet.activeSheet) switchSheet(name);
      };

      tab.onkeydown = (e) => {
        if (e.target !== tab) return;
        if (e.key === 'Enter') tab.click();
        else if (e.key === 'F2') startTabRename(tab, name);
      };

      el.tabsList.appendChild(tab);
    });
  }

  function switchSheet(name) {
    commitPendingEdit();
    const current = getActiveSheetData();
    if (current) {
      current.view = {
        selectedCell: state.sheet.selectedCell,
        rangeEnd: state.sheet.rangeEnd,
        scrollTop: el.gridContainer.scrollTop,
        scrollLeft: el.gridContainer.scrollLeft
      };
    }
    state.sheet.activeSheet = name;
    const view = getActiveSheetData().view || {};
    state.sheet.selectedCell = view.selectedCell || { r: 0, c: 0 };
    state.sheet.rangeEnd = view.rangeEnd || { ...state.sheet.selectedCell };
    el.gridContainer.scrollTop = view.scrollTop || 0;
    el.gridContainer.scrollLeft = view.scrollLeft || 0;
    renderGrid();
  }

  function startTabRename(tab, name) {
    const label = tab.querySelector('.sheet-tab-name');
    if (!label) return;

    const input = document.createElement('input');
    input.className = 'sheet-tab-input';
    input.value = name;
    input.setAttribute('aria-label', 'Sheet name');
    input.spellcheck = false;
    label.replaceWith(input);
    input.focus();
    input.select();

    let finished = false;
    const finish = (commit) => {
      if (finished) return;
      finished = true;
      const newName = input.value.trim();
      if (commit && newName && newName !== name) {
        renameSheet(name, newName);
      } else {
        renderSheetTabs();
      }
    };

    input.addEventListener('blur', () => finish(true));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') finish(true);
      else if (e.key === 'Escape') finish(false);
    });
  }

  function addSheet(name = null) {
    commitPendingEdit();
    let newName = name;
    if (!newName) {
      let i = state.sheet.sheetNames.length + 1;
      while (state.sheet.sheetNames.includes(`Sheet${i}`)) i++;
      newName = `Sheet${i}`;
    }
    pushUndo();
    state.sheet.sheetNames.push(newName);
    state.sheet.sheets[newName] = { cells: {}, rowCount: MIN_ROWS, colCount: MIN_COLS, colWidths: {} };
    markSheetModified();
    switchSheet(newName);
  }

  function renameSheet(oldName, newName) {
    const taken = state.sheet.sheetNames.some((n) => n.toLowerCase() === newName.toLowerCase() && n !== oldName);
    if (taken) {
      showToast(`A sheet named ${newName} already exists`, 'error');
      renderSheetTabs();
      return;
    }
    mutate(() => {
      const idx = state.sheet.sheetNames.indexOf(oldName);
      state.sheet.sheetNames[idx] = newName;
      state.sheet.sheets[newName] = state.sheet.sheets[oldName];
      delete state.sheet.sheets[oldName];
      if (state.sheet.activeSheet === oldName) state.sheet.activeSheet = newName;
      Object.values(state.sheet.sheets).forEach((sheet) => {
        Object.values(sheet.cells).forEach((cell) => {
          if (cell.f) cell.f = renameSheetReferences(cell.f, oldName, newName);
        });
      });
    });
  }

  function deleteSheet(name) {
    if (state.sheet.sheetNames.length <= 1) return;
    confirmDialog({
      title: `Delete ${name}?`,
      message: 'Everything on this sheet will be removed. Formulas that use it will show #REF!.',
      confirmLabel: 'Delete sheet'
    }).then((confirmed) => {
      if (!confirmed) return;
      commitPendingEdit();
      pushUndo();
      state.sheet.sheetNames = state.sheet.sheetNames.filter(n => n !== name);
      delete state.sheet.sheets[name];
      Object.values(state.sheet.sheets).forEach((sheet) => {
        Object.values(sheet.cells).forEach((cell) => {
          if (cell.f) cell.f = removeSheetReferences(cell.f, name);
        });
      });
      markSheetModified();
      if (state.sheet.activeSheet === name) {
        state.sheet.activeSheet = null;
        switchSheet(state.sheet.sheetNames[0]);
      } else {
        renderGrid();
      }
    });
  }

  // --- Grid mouse handling ---
  function beginDrag(state_, onMove, onEnd) {
    dragState = state_;
    const move = (e) => {
      const bounds = el.gridContainer.getBoundingClientRect();
      if (e.clientY > bounds.bottom - 16) el.gridContainer.scrollTop += ROW_HEIGHT;
      else if (e.clientY < bounds.top + HEADER_HEIGHT + 8) el.gridContainer.scrollTop -= ROW_HEIGHT;
      if (e.clientX > bounds.right - 16) el.gridContainer.scrollLeft += DEFAULT_COL_WIDTH / 2;
      else if (e.clientX < bounds.left + ROW_HEADER_WIDTH + 8) el.gridContainer.scrollLeft -= DEFAULT_COL_WIDTH / 2;
      renderVisibleRows();
      onMove(e);
    };
    const up = (e) => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      const finished = dragState;
      dragState = null;
      if (onEnd) onEnd(e, finished);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  function startColResize(e, c) {
    e.preventDefault();
    const sheet = getActiveSheetData();
    const startX = e.clientX;
    const startWidth = colWidth(sheet, c);
    let width = startWidth;
    const col = el.spreadsheetGrid.querySelector(`col[data-col="${c}"]`);
    const tableWidth = parseFloat(el.spreadsheetGrid.style.width);

    beginDrag({ type: 'resize' }, (move) => {
      width = Math.max(MIN_COL_WIDTH, startWidth + move.clientX - startX);
      col.style.width = `${width}px`;
      el.spreadsheetGrid.style.width = `${tableWidth + width - startWidth}px`;
    }, () => {
      if (width === startWidth) return;
      mutate((current) => {
        current.colWidths[c] = width;
      }, { keepClipboard: true });
    });
  }

  function startFillDrag(e) {
    e.preventDefault();
    e.stopPropagation();
    commitPendingEdit();
    const source = getRange();

    beginDrag({ type: 'fill', target: null }, (move) => {
      const { r, c } = cellAtPoint(move.clientX, move.clientY);
      const below = r - source.r2;
      const above = source.r1 - r;
      const right = c - source.c2;
      const left = source.c1 - c;
      const furthest = Math.max(below, above, right, left);
      let target = null;
      if (furthest > 0) {
        if (furthest === below) target = { ...source, r2: r };
        else if (furthest === above) target = { ...source, r1: r };
        else if (furthest === right) target = { ...source, c2: c };
        else target = { ...source, c1: c };
      }
      dragState.target = target;
      paintCopyBox();
    }, (up, finished) => {
      paintCopyBox();
      if (finished.target) applyFill(source, finished.target);
    });
  }

  function onGridMouseDown(e) {
    if (e.target.closest('.cell-editor')) return;
    closeGridMenu();
    const sheet = getActiveSheetData();

    const resize = e.target.closest('.col-resize');
    if (resize) {
      commitPendingEdit();
      startColResize(e, Number(resize.parentElement.dataset.col));
      return;
    }
    if (e.target.closest('.fill-handle')) {
      startFillDrag(e);
      return;
    }

    const td = e.target.closest('td.grid-cell');
    const colHeader = e.target.closest('.col-header');
    const rowHeader = e.target.closest('.row-header');
    const corner = e.target.closest('.grid-corner');
    if (!td && !colHeader && !rowHeader && !corner) return;

    // Right clicking inside the selection keeps it, so the menu acts on all of it.
    if (e.button === 2) {
      commitPendingEdit();
      const range = getRange();
      if (td) {
        const r = Number(td.dataset.row);
        const c = Number(td.dataset.col);
        if (r < range.r1 || r > range.r2 || c < range.c1 || c > range.c2) selectCell(r, c, { scroll: false });
      } else if (colHeader) {
        const c = Number(colHeader.dataset.col);
        if (c < range.c1 || c > range.c2) selectRange(0, c, sheet.rowCount - 1, c);
      } else if (rowHeader) {
        const r = Number(rowHeader.dataset.row);
        if (r < range.r1 || r > range.r2) selectRange(r, 0, r, sheet.colCount - 1);
      }
      return;
    }
    if (e.button !== 0) return;

    if (td) {
      const input = activeEditInput();
      if (input && canPickReference(input)) {
        e.preventDefault();
        const cell = { r: Number(td.dataset.row), c: Number(td.dataset.col) };
        const continuing = refPick && refPick.input === input;
        refPick = {
          input,
          start: continuing ? refPick.start : input.selectionStart,
          textEnd: continuing ? refPick.textEnd : input.selectionStart,
          end: continuing ? refPick.end : undefined,
          anchor: cell
        };
        if (refPick.end === undefined) refPick.end = null;
        writePickedReference(cell);
        beginDrag({ type: 'pick' }, (move) => writePickedReference(cellAtPoint(move.clientX, move.clientY)));
        return;
      }
    }

    const r = td ? Number(td.dataset.row) : rowHeader ? Number(rowHeader.dataset.row) : 0;
    const c = td ? Number(td.dataset.col) : colHeader ? Number(colHeader.dataset.col) : 0;

    commitPendingEdit();
    if (document.activeElement && document.activeElement.matches('input, select, button')) document.activeElement.blur();
    e.preventDefault();

    const lastRow = sheet.rowCount - 1;
    const lastCol = sheet.colCount - 1;

    if (corner) {
      selectRange(0, 0, lastRow, lastCol);
    } else if (colHeader) {
      const anchor = e.shiftKey ? state.sheet.selectedCell.c : c;
      selectRange(0, anchor, lastRow, c);
      beginDrag({ type: 'cols' }, (move) => selectRange(0, anchor, lastRow, cellAtPoint(move.clientX, move.clientY).c));
    } else if (rowHeader) {
      const anchor = e.shiftKey ? state.sheet.selectedCell.r : r;
      selectRange(anchor, 0, r, lastCol);
      beginDrag({ type: 'rows' }, (move) => selectRange(anchor, 0, cellAtPoint(move.clientX, move.clientY).r, lastCol));
    } else {
      selectCell(r, c, { extend: e.shiftKey, scroll: false });
      beginDrag({ type: 'cells' }, (move) => {
        const end = cellAtPoint(move.clientX, move.clientY);
        const a = state.sheet.selectedCell;
        const current = state.sheet.rangeEnd;
        if (!current || current.r !== end.r || current.c !== end.c) selectRange(a.r, a.c, end.r, end.c);
      });
    }
  }

  // --- Right click menu ---
  function closeGridMenu() {
    document.getElementById('gridMenu').hidden = true;
  }

  function openGridMenu(x, y) {
    const menu = document.getElementById('gridMenu');
    const range = getRange();
    const mergeItem = menu.querySelector('[data-action="merge"]');
    const hasMerge = mergesInside(getActiveSheetData(), range).length > 0;
    mergeItem.textContent = hasMerge ? 'Unmerge cells' : 'Merge cells';
    mergeItem.disabled = !hasMerge && range.r1 === range.r2 && range.c1 === range.c2;
    menu.hidden = false;
    menu.style.left = `${Math.min(x, window.innerWidth - menu.offsetWidth - 8)}px`;
    menu.style.top = `${Math.min(y, window.innerHeight - menu.offsetHeight - 8)}px`;
  }

  function initGridMenu() {
    const actions = {
      insertRowAbove: () => insertRows('above'),
      insertRowBelow: () => insertRows('below'),
      deleteRow: deleteRows,
      insertColLeft: () => insertCols('left'),
      insertColRight: () => insertCols('right'),
      deleteCol: deleteCols,
      merge: toggleMerge,
      clear: clearSelection
    };
    const menu = document.getElementById('gridMenu');
    menu.addEventListener('mousedown', (e) => e.preventDefault());
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('[data-action]');
      if (!item) return;
      closeGridMenu();
      actions[item.dataset.action]();
    });
    el.gridContainer.addEventListener('contextmenu', (e) => {
      if (e.target.closest('.cell-editor')) return;
      e.preventDefault();
      openGridMenu(e.clientX, e.clientY);
    });
    document.addEventListener('mousedown', (e) => {
      if (!e.target.closest('#gridMenu') && !e.target.closest('#gridContainer')) closeGridMenu();
    });
  }

  // --- Keyboard ---
  function jumpTarget(r, c, dr, dc) {
    const sheet = getActiveSheetData();
    const filled = (row, col) => rawInput(sheet.cells[coordToAddress(row, col)]) !== '';
    const inBounds = (row, col) => row >= 0 && col >= 0 && row < sheet.rowCount && col < sheet.colCount;

    let row = r + dr;
    let col = c + dc;
    if (!inBounds(row, col)) return { r, c };

    if (filled(r, c) && filled(row, col)) {
      while (inBounds(row + dr, col + dc) && filled(row + dr, col + dc)) {
        row += dr;
        col += dc;
      }
      return { r: row, c: col };
    }
    while (inBounds(row, col) && !filled(row, col)) {
      row += dr;
      col += dc;
    }
    if (inBounds(row, col)) return { r: row, c: col };
    return { r: dr ? (dr > 0 ? sheet.rowCount - 1 : 0) : r, c: dc ? (dc > 0 ? sheet.colCount - 1 : 0) : c };
  }

  function onGridKeyDown(e) {
    if (state.mode !== 'sheet' || cellEditor || e.isComposing) return;
    if (e.target.closest && e.target.closest('input, select, button, [role="tab"], .modal-overlay')) return;
    if (document.querySelector('.modal-overlay.active')) return;

    const sheet = getActiveSheetData();
    const a = state.sheet.selectedCell;
    const end = state.sheet.rangeEnd || a;
    const command = e.metaKey || e.ctrlKey;
    const key = e.key.toLowerCase();

    if (command) {
      if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) redoSheet();
        else undoSheet();
      } else if (key === 'y') {
        e.preventDefault();
        redoSheet();
      } else if (key === 'a') {
        e.preventDefault();
        selectRange(0, 0, sheet.rowCount - 1, sheet.colCount - 1);
      }
    }

    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[e.key]) {
      e.preventDefault();
      const [dr, dc] = arrows[e.key];
      const from = e.shiftKey ? end : a;
      const to = command ? jumpTarget(from.r, from.c, dr, dc) : stepFrom(from.r, from.c, dr, dc);
      selectCell(to.r, to.c, { extend: e.shiftKey });
      return;
    }

    const page = Math.max(1, Math.floor(el.gridContainer.clientHeight / ROW_HEIGHT) - 2);
    switch (e.key) {
      case 'Tab': {
        e.preventDefault();
        const to = stepFrom(a.r, a.c, 0, e.shiftKey ? -1 : 1);
        selectCell(to.r, to.c);
        break;
      }
      case 'Enter': {
        e.preventDefault();
        const to = stepFrom(a.r, a.c, e.shiftKey ? -1 : 1, 0);
        selectCell(to.r, to.c);
        break;
      }
      case 'Home':
        e.preventDefault();
        selectCell(command ? 0 : a.r, 0, { extend: e.shiftKey });
        break;
      case 'PageDown':
        e.preventDefault();
        selectCell(a.r + page, a.c, { extend: e.shiftKey });
        break;
      case 'PageUp':
        e.preventDefault();
        selectCell(a.r - page, a.c, { extend: e.shiftKey });
        break;
      case 'F2':
        e.preventDefault();
        startCellEdit();
        break;
      case 'Escape':
        if (gridClipboard) {
          gridClipboard = null;
          paintCopyBox();
        }
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        clearSelection();
        break;
      default:
        if (e.key.length === 1 && !command && !e.altKey) {
          // Otherwise the browser also types this key into the editor that just took focus.
          e.preventDefault();
          selectCell(a.r, a.c);
          startCellEdit(e.key);
        }
    }
  }

  // --- Spreadsheet ribbon and grid events ---
  function initSheetEvents() {
    el.btnAddSheet.onclick = () => addSheet();

    // Delegated, because selecting another tab rebuilds the tab elements mid double click.
    el.tabsList.addEventListener('dblclick', (e) => {
      if (e.target.closest('.sheet-tab-close, .sheet-tab-input')) return;
      const tab = e.target.closest('.sheet-tab') || el.tabsList.querySelector('.sheet-tab.active');
      if (tab) {
        const name = tab.querySelector('.sheet-tab-name');
        if (name) startTabRename(tab, name.textContent);
      }
    });

    // Formula bar
    const beginFormulaEdit = () => {
      if (formulaEdit || cellEditor) return;
      const { r, c } = state.sheet.selectedCell;
      formulaEdit = { r, c, original: rawInput(getActiveSheetData().cells[coordToAddress(r, c)]) };
    };
    el.formulaInput.addEventListener('focus', beginFormulaEdit);
    el.formulaInput.addEventListener('input', () => {
      beginFormulaEdit();
      refPick = null;
      paintCopyBox();
    });
    el.formulaInput.addEventListener('blur', () => {
      if (document.hasFocus()) commitFormulaBar();
    });
    el.formulaInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        const { r, c } = formulaEdit || state.sheet.selectedCell;
        commitFormulaBar();
        el.formulaInput.blur();
        if (e.key === 'Enter') selectCell(r + (e.shiftKey ? -1 : 1), c);
        else selectCell(r, c + (e.shiftKey ? -1 : 1));
      } else if (e.key === 'Escape' && formulaEdit) {
        e.preventDefault();
        el.formulaInput.value = formulaEdit.original;
        formulaEdit = null;
        clearRefPick();
        el.formulaInput.blur();
      }
    });

    // Name box: type an address or range to go there
    el.cellAddress.addEventListener('focus', () => el.cellAddress.select());
    el.cellAddress.addEventListener('blur', () => paintSelection());
    el.cellAddress.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        el.cellAddress.blur();
      } else if (e.key === 'Enter') {
        e.preventDefault();
        const parts = el.cellAddress.value.trim().split(':').map((part) => addressToCoord(part.replace(/\$/g, '')));
        if (parts.length <= 2 && parts.every(Boolean)) {
          const sheet = getActiveSheetData();
          const fit = (coord) => ({
            r: Math.min(coord.r, sheet.rowCount - 1),
            c: Math.min(coord.c, sheet.colCount - 1)
          });
          const a = fit(parts[0]);
          const b = fit(parts[1] || parts[0]);
          selectRange(a.r, a.c, b.r, b.c);
          scrollCellIntoView(b.r, b.c);
          scrollCellIntoView(a.r, a.c);
        } else {
          showToast('Enter a cell or range, such as B20 or A1:C5', 'error');
        }
        el.cellAddress.blur();
      }
    });

    // Grid
    el.gridContainer.addEventListener('mousedown', onGridMouseDown);
    el.gridContainer.addEventListener('dblclick', (e) => {
      if (e.target.closest('.cell-editor')) return;
      const resize = e.target.closest('.col-resize');
      if (resize) {
        const c = Number(resize.parentElement.dataset.col);
        mutate((sheet) => {
          sheet.colWidths[c] = fittedColWidth(sheet, c);
        }, { keepClipboard: true });
        return;
      }
      if (e.target.closest('td.grid-cell')) startCellEdit();
    });

    el.gridContainer.addEventListener('scroll', () => {
      closeGridMenu();
      renderVisibleRows();
    });
    window.addEventListener('resize', () => renderVisibleRows());

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeGridMenu();
      onGridKeyDown(e);
    });

    initGridMenu();

    // Toolbar
    document.getElementById('sheetUndo').onclick = undoSheet;
    document.getElementById('sheetRedo').onclick = redoSheet;

    const toolbarAction = (id, action) => {
      document.getElementById(id).onclick = () => {
        commitPendingEdit();
        action();
      };
    };
    toolbarAction('sheetBold', () => toggleCellFormat('bold'));
    toolbarAction('sheetItalic', () => toggleCellFormat('italic'));
    toolbarAction('sheetAlignLeft', () => setCellFormatProperty('align', 'left'));
    toolbarAction('sheetAlignCenter', () => setCellFormatProperty('align', 'center'));
    toolbarAction('sheetAlignRight', () => setCellFormatProperty('align', 'right'));
    toolbarAction('sheetInsertRowAbove', () => insertRows('above'));
    toolbarAction('sheetInsertRowBelow', () => insertRows('below'));
    toolbarAction('sheetDeleteRow', deleteRows);
    toolbarAction('sheetInsertColLeft', () => insertCols('left'));
    toolbarAction('sheetInsertColRight', () => insertCols('right'));
    toolbarAction('sheetDeleteCol', deleteCols);
    document.getElementById('sheetAutoSum').onclick = runAutoSum;
    document.getElementById('sheetMerge').onclick = toggleMerge;

    document.getElementById('sheetTextColor').onchange = (e) => setCellFormatProperty('color', e.target.value);
    document.getElementById('sheetFillColor').onchange = (e) => setCellFormatProperty('bg', e.target.value);
    document.getElementById('sheetNumberFormat').onchange = (e) => {
      setCellFormatProperty('format', e.target.value === 'general' ? '' : e.target.value);
    };
  }

  // --- XLSX IMPORT & EXPORT via SheetJS ---
  async function loadSpreadsheetFile(file) {
    showToast('Opening spreadsheet');
    const arrayBuffer = await file.arrayBuffer();
    try {
      if (!window.XLSX) throw new Error('SheetJS library not loaded');
      const wb = window.XLSX.read(arrayBuffer, { type: 'array', cellFormula: true, cellStyles: true });
      const styles = await readCellStyles(arrayBuffer);

      state.sheet.sheetNames = wb.SheetNames;
      state.sheet.sheets = {};

      wb.SheetNames.forEach(sheetName => {
        const ws = wb.Sheets[sheetName];
        const range = window.XLSX.utils.decode_range(ws['!ref'] || 'A1:M30');
        const cells = {};

        for (const [cellAddr, cellObj] of Object.entries(ws)) {
          if (cellAddr.startsWith('!')) continue;
          // Error cells store a numeric code in v and the text, such as #N/A, in w.
          const isError = cellObj.t === 'e';
          cells[cellAddr] = {
            v: isError ? cellObj.w : (cellObj.v !== undefined ? cellObj.v : ''),
            f: cellObj.f ? `=${cellObj.f}` : undefined,
            s: (styles[sheetName] && styles[sheetName][cellAddr]) || {}
          };
        }

        // Filled or formatted cells with no value, such as the blanks in a header band.
        Object.entries(styles[sheetName] || {}).forEach(([cellAddr, style]) => {
          const coord = addressToCoord(cellAddr);
          const inRange = coord && coord.r <= range.e.r && coord.c <= range.e.c;
          if (inRange && !cells[cellAddr]) cells[cellAddr] = { v: '', s: style };
        });

        const colWidths = {};
        (ws['!cols'] || []).forEach((col, index) => {
          if (!col) return;
          const px = col.wpx || (col.wch ? col.wch * 7 + 5 : 0);
          if (px) colWidths[index] = Math.max(MIN_COL_WIDTH, Math.round(px));
        });

        const merges = (ws['!merges'] || [])
          .map((m) => ({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }))
          .filter((m) => m.r2 > m.r1 || m.c2 > m.c1);

        state.sheet.sheets[sheetName] = {
          cells: cells,
          colWidths: colWidths,
          merges: merges,
          rowCount: Math.max(range.e.r + 100, MIN_ROWS),
          colCount: Math.max(range.e.c + 5, MIN_COLS)
        };
      });

      state.sheet.activeSheet = wb.SheetNames[0];
      resetSheetView();
      wb.SheetNames.forEach((sheetName) => autoSizeColumns(state.sheet.sheets[sheetName]));
      const title = file.name.replace(/\.[^/.]+$/, '');
      state.sheetTitle = title;
      el.docNameInput.value = title;

      switchMode('sheet');
      renderGrid();
      markSaved();
      showToast(`Opened ${wb.SheetNames.length} ${wb.SheetNames.length === 1 ? 'sheet' : 'sheets'}`, 'success');
    } catch (err) {
      showToast('Could not open this spreadsheet: ' + err.message, 'error');
    }
  }

  // --- XLSX Cell Styles ---
  // SheetJS Community reads and writes values and formulas but not cell styles,
  // so styles are read from, and written into, the .xlsx package directly.
  const XLSX_NUMBER_FORMATS = {
    general: { id: 0 },
    currency: { id: 164, code: '"$"#,##0.00' },
    percent: { id: 165, code: '0.0%' },
    decimal: { id: 4 },
    integer: { id: 3 },
    text: { id: 49 },
    date: { id: 14 }
  };

  const BUILT_IN_FORMATS = {
    2: 'decimal', 3: 'integer', 4: 'decimal', 9: 'percent', 10: 'percent', 49: 'text',
    14: 'date', 15: 'date', 16: 'date', 17: 'date', 22: 'date'
  };

  function toArgb(color) {
    const match = /^#?([0-9a-f]{6})$/i.exec(color || '');
    return match ? `FF${match[1].toUpperCase()}` : null;
  }

  function fromArgb(element) {
    const rgb = element && element.getAttribute('rgb');
    return rgb && rgb.length === 8 ? `#${rgb.slice(2).toLowerCase()}` : '';
  }

  function xmlAttr(value) {
    return String(value).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  }

  // Assigns each distinct font, fill, and cell format an index, reusing identical ones.
  function createStyleTable() {
    const intern = (list, xml) => {
      const found = list.indexOf(xml);
      return found >= 0 ? found : list.push(xml) - 1;
    };
    const fontXml = (bold, italic, argb) => '<font>' +
      `${bold ? '<b/>' : ''}${italic ? '<i/>' : ''}<sz val="11"/>` +
      `${argb ? `<color rgb="${argb}"/>` : ''}<name val="Calibri"/></font>`;

    const fonts = [fontXml(false, false, null)];
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];

    return {
      indexFor(style) {
        if (!style) return 0;
        const format = XLSX_NUMBER_FORMATS[style.format] || XLSX_NUMBER_FORMATS.general;
        const font = intern(fonts, fontXml(style.bold, style.italic, toArgb(style.color)));
        const bg = toArgb(style.bg);
        const fill = bg
          ? intern(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${bg}"/><bgColor indexed="64"/></patternFill></fill>`)
          : 0;
        const align = ['left', 'center', 'right'].includes(style.align) ? style.align : null;
        if (!format.id && !font && !fill && !align) return 0;

        return intern(xfs, `<xf numFmtId="${format.id}" fontId="${font}" fillId="${fill}" borderId="0" xfId="0"` +
          `${format.id ? ' applyNumberFormat="1"' : ''}${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}` +
          (align ? ` applyAlignment="1"><alignment horizontal="${align}"/></xf>` : '/>'));
      },
      toXml() {
        const customFormats = Object.values(XLSX_NUMBER_FORMATS).filter((format) => format.code);
        const numFmts = customFormats
          .map((format) => `<numFmt numFmtId="${format.id}" formatCode="${xmlAttr(format.code)}"/>`)
          .join('');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          `<numFmts count="${customFormats.length}">${numFmts}</numFmts>` +
          `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
          `<fills count="${fills.length}">${fills.join('')}</fills>` +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          '</styleSheet>';
      }
    };
  }

  async function writeCellStyles(xlsxBytes, sheets) {
    const zip = await window.JSZip.loadAsync(xlsxBytes);
    const table = createStyleTable();

    for (let i = 0; i < sheets.length; i++) {
      const path = `xl/worksheets/sheet${i + 1}.xml`;
      const xml = await zip.file(path).async('string');
      zip.file(path, xml.replace(/<c r="([A-Z]+[0-9]+)"/g, (tag, addr) => {
        const cell = sheets[i].cells[addr];
        const index = table.indexFor(cell && cell.s);
        return index ? `${tag} s="${index}"` : tag;
      }), { createFolders: false });
    }

    zip.file('xl/styles.xml', table.toXml(), { createFolders: false });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  function numberFormatName(numFmtId, customCodes) {
    if (BUILT_IN_FORMATS[numFmtId]) return BUILT_IN_FORMATS[numFmtId];
    const code = customCodes[numFmtId] || '';
    // Ignore quoted text, locale tags, and escapes. Date codes must not read as currency: [$-409]d-mmm
    const tokens = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    if (/[dy]/i.test(tokens) || (/m/i.test(tokens) && !/[hs]/i.test(tokens))) return 'date';
    if (/[hs]/i.test(tokens)) return null;
    if (tokens.includes('%')) return 'percent';
    if (code.replace(/\[\$-[^\]]*\]/g, '').includes('$')) return 'currency';
    return null;
  }

  // Returns { sheetName: { A1: style } }. Any file this cannot read, such as .csv or .xls, has no styles.
  async function readCellStyles(arrayBuffer) {
    const result = {};
    try {
      const zip = await window.JSZip.loadAsync(arrayBuffer);
      const parse = async (path) => {
        const file = zip.file(path);
        return file ? new DOMParser().parseFromString(await file.async('string'), 'application/xml') : null;
      };
      const byTag = (node, tag) => [...node.getElementsByTagNameNS('*', tag)];

      const [stylesDoc, workbookDoc, relsDoc] = await Promise.all([
        parse('xl/styles.xml'), parse('xl/workbook.xml'), parse('xl/_rels/workbook.xml.rels')
      ]);
      if (!stylesDoc || !workbookDoc || !relsDoc) return result;

      const customCodes = {};
      byTag(stylesDoc, 'numFmt').forEach((numFmt) => {
        customCodes[numFmt.getAttribute('numFmtId')] = numFmt.getAttribute('formatCode');
      });

      const isOn = (element) => !!element && element.getAttribute('val') !== '0';
      const fonts = byTag(byTag(stylesDoc, 'fonts')[0], 'font').map((font) => ({
        bold: isOn(byTag(font, 'b')[0]),
        italic: isOn(byTag(font, 'i')[0]),
        color: fromArgb(byTag(font, 'color')[0])
      }));
      const fills = byTag(byTag(stylesDoc, 'fills')[0], 'fill').map((fill) => {
        const pattern = byTag(fill, 'patternFill')[0];
        return pattern && pattern.getAttribute('patternType') === 'solid' ? fromArgb(byTag(fill, 'fgColor')[0]) : '';
      });

      const cellXfs = byTag(byTag(stylesDoc, 'cellXfs')[0], 'xf').map((xf) => {
        const style = {};
        const font = fonts[Number(xf.getAttribute('fontId'))] || {};
        if (font.bold) style.bold = true;
        if (font.italic) style.italic = true;
        if (font.color) style.color = font.color;
        const bg = fills[Number(xf.getAttribute('fillId'))];
        if (bg) style.bg = bg;
        const format = numberFormatName(xf.getAttribute('numFmtId'), customCodes);
        if (format) style.format = format;
        const alignment = byTag(xf, 'alignment')[0];
        const align = alignment && alignment.getAttribute('horizontal');
        if (['left', 'center', 'right'].includes(align)) style.align = align;
        return style;
      });

      const targets = {};
      byTag(relsDoc, 'Relationship').forEach((rel) => {
        targets[rel.getAttribute('Id')] = rel.getAttribute('Target');
      });

      for (const sheet of byTag(workbookDoc, 'sheet')) {
        const idAttr = [...sheet.attributes].find((attr) => attr.localName === 'id');
        const target = idAttr && targets[idAttr.value];
        if (!target) continue;
        const sheetDoc = await parse(target.startsWith('/') ? target.slice(1) : `xl/${target}`);
        if (!sheetDoc) continue;

        const cells = {};
        byTag(sheetDoc, 'c').forEach((c) => {
          const style = cellXfs[Number(c.getAttribute('s'))];
          if (style && Object.keys(style).length) cells[c.getAttribute('r')] = { ...style };
        });
        result[sheet.getAttribute('name')] = cells;
      }
    } catch (e) {
      console.warn('Could not read cell styles', e);
    }
    return result;
  }

  // Fits each column to its longest formatted value. Excel shows #### for numbers that do not fit.
  function columnWidths(sheetData, maxCol) {
    const widths = Array.from({ length: maxCol + 1 }, () => ({ wch: 10 }));
    for (const [addr, cell] of Object.entries(sheetData.cells)) {
      const text = String(formatValue(displayValue(sheetData, addr, cell), (cell.s && cell.s.format) || 'general'));
      const column = widths[addressToCoord(addr).c];
      column.wch = Math.max(column.wch, Math.min(text.length + 2, 60));
    }
    // Columns someone sized by hand keep that size.
    Object.entries(sheetData.colWidths || {}).forEach(([c, px]) => {
      if (widths[c]) widths[c].wch = Math.max(1, (px - 5) / 7);
    });
    return widths;
  }

  async function exportXlsx() {
    try {
      commitPendingEdit();
      formulas.reset();
      if (!window.XLSX) throw new Error('SheetJS not loaded');
      const wb = window.XLSX.utils.book_new();

      state.sheet.sheetNames.forEach(sheetName => {
        const sheetData = state.sheet.sheets[sheetName];
        const ws = {};
        let maxR = 0, maxC = 0;

        for (const [addr, cell] of Object.entries(sheetData.cells)) {
          const coord = addressToCoord(addr);
          if (coord.r > maxR) maxR = coord.r;
          if (coord.c > maxC) maxC = coord.c;

          const cellObj = {};
          const excelType = (value) => ({ number: 'n', boolean: 'b' }[typeof value] || 's');
          if (cell.f) {
            cellObj.f = cell.f.replace(/^=/, '');
            const val = formulas.value(sheetData, addr);
            cellObj.v = val;
            cellObj.t = excelType(val);
          } else {
            cellObj.v = cell.v !== undefined ? cell.v : '';
            cellObj.t = excelType(cellObj.v);
          }
          ws[addr] = cellObj;
        }

        const merges = sheetData.merges || [];
        merges.forEach((m) => {
          maxR = Math.max(maxR, m.r2);
          maxC = Math.max(maxC, m.c2);
        });
        if (merges.length) {
          ws['!merges'] = merges.map((m) => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
        }
        ws['!ref'] = `A1:${coordToAddress(maxR, maxC)}`;
        ws['!cols'] = columnWidths(sheetData, maxC);
        window.XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });

      const filename = (state.sheetTitle || 'spreadsheet').replace(/\.[^/.]+$/, '') + '.xlsx';
      const outBuf = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const sheets = state.sheet.sheetNames.map((name) => state.sheet.sheets[name]);
      const styled = await writeCellStyles(outBuf, sheets);
      const blob = new Blob([styled], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      await saveBlob(blob, filename);
    } catch (err) {
      showToast('Could not save this spreadsheet: ' + err.message, 'error');
    }
  }

  // ==========================================
  // --- MODE SWITCHING & GENERAL APP UI ---
  // ==========================================

  function switchMode(newMode) {
    state.mode = newMode;
    document.body.dataset.mode = newMode;
    el.btnModeDoc.setAttribute('aria-pressed', newMode === 'doc' ? 'true' : 'false');
    el.btnModeSheet.setAttribute('aria-pressed', newMode === 'sheet' ? 'true' : 'false');
    if (newMode === 'doc') {
      el.saveBtnLabel.textContent = 'Save .docx';
      el.docNameInput.value = state.docTitle;
    } else {
      el.saveBtnLabel.textContent = 'Save .xlsx';
      el.docNameInput.value = state.sheetTitle;
      renderGrid();
    }
    updateSaveStatus();
  }

  function initModeSwitcher() {
    el.btnModeDoc.onclick = () => switchMode('doc');
    el.btnModeSheet.onclick = () => switchMode('sheet');

    el.docNameInput.oninput = (e) => {
      if (state.mode === 'doc') {
        state.docTitle = e.target.value;
      } else {
        state.sheetTitle = e.target.value;
      }
    };

    // Save Button
    el.btnSaveMain.onclick = () => {
      if (state.mode === 'doc') exportDocx();
      else exportXlsx();
    };

    // New File Button
    el.btnNew.onclick = () => {
      const noun = state.mode === 'doc' ? 'document' : 'spreadsheet';
      const modified = state.mode === 'doc' ? state.docModified : state.sheetModified;
      const ask = modified
        ? confirmDialog({
          title: `Create a blank ${noun}?`,
          message: `This replaces the current ${noun}. Changes you have not saved will be lost.`,
          confirmLabel: `Create blank ${noun}`
        })
        : Promise.resolve(true);
      ask.then((confirmed) => {
        if (!confirmed) return;
        if (state.mode === 'doc') {
          el.documentPage.innerHTML = '';
          state.docSection = null;
          state.docTitle = 'Untitled Document';
          el.docNameInput.value = state.docTitle;
          updateDocStats();
          el.documentPage.focus();
        } else {
          state.sheet = {
            activeSheet: 'Sheet1',
            sheetNames: ['Sheet1'],
            sheets: {
              'Sheet1': { cells: {}, rowCount: MIN_ROWS, colCount: MIN_COLS, colWidths: {} }
            },
            selectedCell: { r: 0, c: 0 },
            rangeEnd: { r: 0, c: 0 }
          };
          resetSheetView();
          state.sheetTitle = 'Untitled Spreadsheet';
          el.docNameInput.value = state.sheetTitle;
          renderGrid();
        }
        markSaved();
      });
    };

    // Print Button
    el.btnPrint.onclick = () => {
      if (state.mode === 'sheet') markPrintRange();
      if (IS_NATIVE_APP && window.ipc) {
        // WKWebView ignores window.print(), so the app runs the print operation itself.
        const title = state.mode === 'doc' ? state.docTitle : state.sheetTitle;
        window.ipc.postMessage(JSON.stringify({ cmd: 'print', title }));
      } else {
        window.print();
      }
    };

    // Help Modal
    el.btnHelp.onclick = () => openModal(el.modalHelp);
    document.getElementById('modalHelpClose').onclick = () => closeModal(el.modalHelp);
    document.getElementById('btnHelpDone').onclick = () => closeModal(el.modalHelp);

    // Samples Modal
    el.btnSamples.onclick = () => openModal(el.modalSamples);
    document.getElementById('modalSamplesClose').onclick = () => closeModal(el.modalSamples);
    document.getElementById('btnCancelSamples').onclick = () => closeModal(el.modalSamples);

    const openSample = (noun, isModified, load) => {
      closeModal(el.modalSamples);
      if (!isModified) {
        load();
        return;
      }
      confirmDialog({
        title: 'Open the sample?',
        message: `This replaces your current ${noun}. Changes you have not saved will be lost.`,
        confirmLabel: 'Open sample'
      }).then((confirmed) => {
        if (confirmed) load();
      });
    };

    document.getElementById('btnLoadDocSample').onclick = () => {
      openSample('document', state.docModified, loadSampleDocument);
    };

    document.getElementById('btnLoadSheetSample').onclick = () => {
      openSample('spreadsheet', state.sheetModified, loadSampleSpreadsheet);
    };

    // Open File picker
    el.btnOpenFile.onclick = () => el.fileInput.click();
    el.fileInput.onchange = (e) => {
      const file = e.target.files[0];
      if (file) handleIncomingFile(file);
      el.fileInput.value = '';
    };

    // Drag and Drop files
    window.addEventListener('dragover', (e) => {
      e.preventDefault();
      el.dragOverlay.classList.add('active');
    });

    window.addEventListener('dragleave', (e) => {
      if (e.relatedTarget === null) {
        el.dragOverlay.classList.remove('active');
      }
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      el.dragOverlay.classList.remove('active');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleIncomingFile(e.dataTransfer.files[0]);
      }
    });
  }

  const DOC_EXTENSIONS = ['docx', 'doc', 'html', 'htm', 'txt'];
  const SHEET_EXTENSIONS = ['xlsx', 'xls', 'csv'];

  // Opened HTML can come from anywhere, so scripts, frames, and event handlers are removed.
  function cleanImportedHtml(html) {
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    parsed.querySelectorAll('script, iframe, frame, frameset, object, embed, link, meta, base, form, noscript, template')
      .forEach((node) => node.remove());
    parsed.body.querySelectorAll('*').forEach((node) => {
      for (const attr of [...node.attributes]) {
        const name = attr.name.toLowerCase();
        const value = attr.value.replace(/[\u0000-\u0020]/g, '').toLowerCase();
        const unsafeUrl = ['href', 'src', 'xlink:href', 'action', 'formaction'].includes(name) &&
          (value.startsWith('javascript:') || value.startsWith('vbscript:') ||
            (value.startsWith('data:') && !value.startsWith('data:image/')));
        if (name.startsWith('on') || name === 'srcdoc' || unsafeUrl) node.removeAttribute(attr.name);
      }
    });
    return parsed.body.innerHTML;
  }

  /** Opens a file, first asking before it replaces work that has not been saved. */
  function handleIncomingFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const isDoc = DOC_EXTENSIONS.includes(ext);
    if (!isDoc && !SHEET_EXTENSIONS.includes(ext)) {
      showToast(`OfficeLite cannot open .${ext} files. Try a .docx, .xlsx, or .csv file.`, 'error');
      return;
    }
    const modified = isDoc ? state.docModified : state.sheetModified;
    if (!modified) {
      openIncomingFile(file, ext);
      return;
    }
    const noun = isDoc ? 'document' : 'spreadsheet';
    confirmDialog({
      title: `Open ${file.name}?`,
      message: `This replaces your current ${noun}. Changes you have not saved will be lost.`,
      confirmLabel: 'Open file'
    }).then((confirmed) => {
      if (confirmed) openIncomingFile(file, ext);
    });
  }

  function openIncomingFile(file, ext) {
    if (ext === 'docx' || ext === 'doc') {
      switchMode('doc');
      loadDocxFile(file);
    } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
      switchMode('sheet');
      loadSpreadsheetFile(file);
    } else if (ext === 'html' || ext === 'htm' || ext === 'txt') {
      const reader = new FileReader();
      reader.onload = (re) => {
        switchMode('doc');
        state.docSection = null;
        if (ext !== 'txt') el.documentPage.innerHTML = cleanImportedHtml(re.target.result);
        else el.documentPage.innerText = re.target.result;
        state.docTitle = file.name.replace(/\.[^/.]+$/, '');
        el.docNameInput.value = state.docTitle;
        updateDocStats();
        markSaved();
        showToast(`Opened ${file.name}`, 'success');
      };
      reader.readAsText(file);
    }
  }

  // --- Native app: File menu commands and files opened from Finder ---
  function initNativeBridge() {
    if (!IS_NATIVE_APP) return;

    window.officeLiteCommand = (command) => {
      if (document.querySelector('.modal-overlay.active')) return;
      const button = { new: el.btnNew, open: el.btnOpenFile, save: el.btnSaveMain, print: el.btnPrint }[command];
      if (button) button.click();
    };

    window.officeLiteOpenFile = async (id, name) => {
      try {
        const response = await fetch(`/api/opened-file?id=${encodeURIComponent(id)}`);
        if (!response.ok) {
          const detail = await response.json().catch(() => ({}));
          throw new Error(detail.error || `status ${response.status}`);
        }
        const blob = await response.blob();
        // A dialog may be open, such as the one that asks before replacing work.
        if (document.querySelector('.modal-overlay.active')) {
          document.querySelectorAll('.modal-overlay.active').forEach((modal) => closeModal(modal));
        }
        handleIncomingFile(new File([blob], name));
      } catch (err) {
        showToast(`Could not open ${name}: ${err.message}`, 'error');
      }
    };

    if (window.ipc) window.ipc.postMessage(JSON.stringify({ cmd: 'ready' }));
  }

  // --- Sample Document & Spreadsheet Loaders ---
  function loadSampleDocument() {
    switchMode('doc');
    state.docTitle = 'Sample Project Proposal';
    state.docSection = null;
    el.docNameInput.value = state.docTitle;

    el.documentPage.innerHTML = `
      <h1>Project Proposal &amp; Operational Plan</h1>
      <p>This is a complete demonstration of the <strong>OfficeLite</strong> document editor. You can view, format, add tables, insert images, and save back to <strong>.docx</strong> without paying for Microsoft Office.</p>
      
      <h2>1. Key Strategic Objectives</h2>
      <ul>
        <li>Full document editing without paying <strong>$100/yr</strong> for Microsoft 365.</li>
        <li>Native OpenXML format compatibility with Word, Apple Pages, Google Docs, and LibreOffice.</li>
        <li>100% private, local processing on your Mac with zero cloud exposure.</li>
      </ul>

      <h2>2. Department Estimates &amp; Targets</h2>
      <table border="1" cellpadding="8" style="width:100%; border-collapse:collapse; margin:16px 0;">
        <thead>
          <tr style="background-color:#f1f5f9;">
            <th style="border:1px solid #cbd5e1; padding:8px 12px; text-align:left;">Department</th>
            <th style="border:1px solid #cbd5e1; padding:8px 12px; text-align:left;">Q1 Budget</th>
            <th style="border:1px solid #cbd5e1; padding:8px 12px; text-align:left;">Q2 Budget</th>
            <th style="border:1px solid #cbd5e1; padding:8px 12px; text-align:left;">Status</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">Engineering &amp; Dev</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$45,000</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$52,000</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px; color:#10b981; font-weight:bold;">On Track</td>
          </tr>
          <tr>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">Design &amp; Marketing</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$28,000</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$34,000</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px; color:#10b981; font-weight:bold;">Approved</td>
          </tr>
          <tr>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">Operations &amp; Admin</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$16,500</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px;">$18,000</td>
            <td style="border:1px solid #cbd5e1; padding:8px 12px; color:#f59e0b; font-weight:bold;">Under Review</td>
          </tr>
        </tbody>
      </table>

      <blockquote>
        "OfficeLite allows you to read, write, and export MS Office files completely offline, fast, and free forever."
      </blockquote>

      <p>Click anywhere to begin editing this document, or click <strong>Save .docx</strong> above to download the file.</p>
    `;
    updateDocStats();
    markSaved();
    showToast('Opened sample document', 'success');
  }

  function loadSampleSpreadsheet() {
    switchMode('sheet');
    state.sheetTitle = 'Annual Financial Budget';
    el.docNameInput.value = state.sheetTitle;

    state.sheet.sheetNames = ['Annual Budget', 'Staffing Plan'];
    state.sheet.sheets = {
      'Annual Budget': {
        rowCount: 30,
        colCount: 15,
        cells: {
          'A1': { v: 'Expense Category', s: { bold: true, bg: '#f1f5f9' } },
          'B1': { v: 'Q1 Budget', s: { bold: true, bg: '#f1f5f9', align: 'right' } },
          'C1': { v: 'Q2 Budget', s: { bold: true, bg: '#f1f5f9', align: 'right' } },
          'D1': { v: 'Q3 Budget', s: { bold: true, bg: '#f1f5f9', align: 'right' } },
          'E1': { v: 'Q4 Budget', s: { bold: true, bg: '#f1f5f9', align: 'right' } },
          'F1': { v: 'Annual Total', s: { bold: true, bg: '#e2e8f0', align: 'right' } },

          'A2': { v: 'Cloud Hosting & Servers', s: {} },
          'B2': { v: 1200, s: { format: 'currency', align: 'right' } },
          'C2': { v: 1350, s: { format: 'currency', align: 'right' } },
          'D2': { v: 1400, s: { format: 'currency', align: 'right' } },
          'E2': { v: 1550, s: { format: 'currency', align: 'right' } },
          'F2': { f: '=SUM(B2:E2)', s: { format: 'currency', align: 'right', bold: true } },

          'A3': { v: 'Software & Tools', s: {} },
          'B3': { v: 850, s: { format: 'currency', align: 'right' } },
          'C3': { v: 850, s: { format: 'currency', align: 'right' } },
          'D3': { v: 920, s: { format: 'currency', align: 'right' } },
          'E3': { v: 950, s: { format: 'currency', align: 'right' } },
          'F3': { f: '=SUM(B3:E3)', s: { format: 'currency', align: 'right', bold: true } },

          'A4': { v: 'Marketing & Ads', s: {} },
          'B4': { v: 2500, s: { format: 'currency', align: 'right' } },
          'C4': { v: 3200, s: { format: 'currency', align: 'right' } },
          'D4': { v: 4000, s: { format: 'currency', align: 'right' } },
          'E4': { v: 4800, s: { format: 'currency', align: 'right' } },
          'F4': { f: '=SUM(B4:E4)', s: { format: 'currency', align: 'right', bold: true } },

          'A5': { v: 'Office Supplies & Misc', s: {} },
          'B5': { v: 400, s: { format: 'currency', align: 'right' } },
          'C5': { v: 380, s: { format: 'currency', align: 'right' } },
          'D5': { v: 420, s: { format: 'currency', align: 'right' } },
          'E5': { v: 450, s: { format: 'currency', align: 'right' } },
          'F5': { f: '=SUM(B5:E5)', s: { format: 'currency', align: 'right', bold: true } },

          'A6': { v: 'TOTAL EXPENDITURE', s: { bold: true, bg: '#dcfce7' } },
          'B6': { f: '=SUM(B2:B5)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7' } },
          'C6': { f: '=SUM(C2:C5)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7' } },
          'D6': { f: '=SUM(D2:D5)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7' } },
          'E6': { f: '=SUM(E2:E5)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7' } },
          'F6': { f: '=SUM(F2:F5)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7', color: '#166534' } }
        }
      },
      'Staffing Plan': {
        rowCount: 20,
        colCount: 10,
        cells: {
          'A1': { v: 'Emp ID', s: { bold: true, bg: '#f1f5f9' } },
          'B1': { v: 'Name', s: { bold: true, bg: '#f1f5f9' } },
          'C1': { v: 'Department', s: { bold: true, bg: '#f1f5f9' } },
          'D1': { v: 'Monthly Salary', s: { bold: true, bg: '#f1f5f9', align: 'right' } },

          'A2': { v: 'EMP-01', s: {} },
          'B2': { v: 'Alice Johnson', s: {} },
          'C2': { v: 'Engineering', s: {} },
          'D2': { v: 9500, s: { format: 'currency', align: 'right' } },

          'A3': { v: 'EMP-02', s: {} },
          'B3': { v: 'Brian Miller', s: {} },
          'C3': { v: 'Design', s: {} },
          'D3': { v: 7800, s: { format: 'currency', align: 'right' } },

          'A4': { v: 'EMP-03', s: {} },
          'B4': { v: 'Chloe Davis', s: {} },
          'C4': { v: 'Marketing', s: {} },
          'D4': { v: 7200, s: { format: 'currency', align: 'right' } },

          'A5': { v: 'TOTAL PAYROLL', s: { bold: true, bg: '#dcfce7' } },
          'B5': { v: '', s: { bg: '#dcfce7' } },
          'C5': { v: '', s: { bg: '#dcfce7' } },
          'D5': { f: '=SUM(D2:D4)', s: { format: 'currency', align: 'right', bold: true, bg: '#dcfce7', color: '#166534' } }
        }
      }
    };
    state.sheet.activeSheet = 'Annual Budget';
    resetSheetView();
    state.sheet.sheetNames.forEach((name) => autoSizeColumns(state.sheet.sheets[name]));
    renderGrid();
    markSaved();
    showToast('Opened sample budget', 'success');
  }

  // --- Auto-Save to LocalStorage ---
  function initAutoSave() {
    setInterval(() => {
      try {
        const payload = {
          mode: state.mode,
          docTitle: state.docTitle,
          sheetTitle: state.sheetTitle,
          docHtml: el.documentPage.innerHTML,
          docSection: state.docSection,
          sheet: state.sheet,
          docModified: state.docModified,
          sheetModified: state.sheetModified
        };
        localStorage.setItem('officelite_autosave', JSON.stringify(payload));
      } catch (e) {
        // storage quota exceeded or disabled
      }
    }, 4000);

    // Check if previous session exists. Returns which restored files still have unsaved changes.
    const restored = { doc: false, sheet: false };
    try {
      const saved = localStorage.getItem('officelite_autosave');
      if (saved) {
        const data = JSON.parse(saved);
        if (data.docHtml && data.docHtml.length > 50) {
          el.documentPage.innerHTML = data.docHtml;
          state.docTitle = data.docTitle || 'Untitled Document';
          state.docSection = data.docSection || null;
          el.docNameInput.value = state.docTitle;
          updateDocStats();
          restored.doc = !!data.docModified;
        }
        if (data.sheet && data.sheet.sheetNames) {
          restored.sheet = !!data.sheetModified;
          state.sheet = data.sheet;
          delete state.sheet.editingCell;
          if (!state.sheet.rangeEnd) state.sheet.rangeEnd = { ...state.sheet.selectedCell };
          state.sheetTitle = data.sheetTitle || 'Untitled Spreadsheet';
        }
      } else {
        // Load default starting document
        loadSampleDocument();
      }
    } catch (e) {
      loadSampleDocument();
    }
    return restored;
  }

  // --- App Initialization ---
  function init() {
    initModals();
    initToolbars();
    initShortcuts();
    initGridClipboard();
    initDocToolbar();
    initSheetEvents();
    initModeSwitcher();
    const restored = initAutoSave();
    updateDocStats();
    state.docModified = restored.doc;
    state.sheetModified = restored.sheet;
    updateSaveStatus();
    initNativeBridge();
  }

  // Start app when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
