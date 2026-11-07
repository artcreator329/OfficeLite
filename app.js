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
    sheetTitle: 'Untitled Spreadsheet',
    serverOnline: false,
    pandocOnline: false,
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
      selectedCell: { r: 1, c: 0 }, // Row index (0-based: row 1 is idx 0), Col index (0-based: col A is idx 0)
      editingCell: null
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
    engineBadge: document.getElementById('engineBadge'),
    engineBadgeText: document.getElementById('engineBadgeText'),

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
  function icon(name) {
    return `<svg class="icon" aria-hidden="true"><use href="#i-${name}"/></svg>`;
  }

  const TOAST_ICONS = { info: 'info', success: 'check-circle', error: 'warning-circle' };
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
    }, kind === 'error' ? 6000 : 2800);
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

  // --- Server & Pandoc Health Check ---
  async function checkServerStatus() {
    try {
      const res = await fetch('/api/status', { method: 'GET' });
      if (res.ok) {
        const data = await res.json();
        state.serverOnline = true;
        state.pandocOnline = data.pandoc_available;
        if (data.pandoc_available) {
          el.engineBadgeText.textContent = 'Pandoc converter';
          el.engineBadge.title = `Converting .docx files with ${data.pandoc_version}`;
        } else {
          el.engineBadgeText.textContent = 'Built in converter';
          el.engineBadge.title = 'Install Pandoc for higher fidelity .docx conversion';
        }
        return;
      }
    } catch (e) {
      // Running standalone without python server
    }
    state.serverOnline = false;
    state.pandocOnline = false;
    el.engineBadgeText.textContent = 'Built in converter';
    el.engineBadge.title = 'Converting .docx files inside the app';
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

    const size = document.queryCommandValue('fontSize');
    const sizeSelect = document.getElementById('docFontSize');
    if ([...sizeSelect.options].some((o) => o.value === size)) sizeSelect.value = size;
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
    for (let r = 0; r < rows; r++) {
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

    // Set doc title from file name
    const title = file.name.replace(/\.[^/.]+$/, '');
    state.docTitle = title;
    el.docNameInput.value = title;

    // Try Pandoc server first if available for high-fidelity HTML conversion
    if (state.pandocOnline) {
      try {
        const res = await fetch('/api/import-docx', {
          method: 'POST',
          body: arrayBuffer
        });
        if (res.ok) {
          const data = await res.json();
          if (data.html) {
            el.documentPage.innerHTML = data.html;
            updateDocStats();
            markSaved();
            showToast('Document opened', 'success');
            return;
          }
        }
      } catch (e) {
        console.warn('Server import failed, falling back to client Mammoth', e);
      }
    }

    // Client-side Mammoth.js fallback
    try {
      if (window.mammoth) {
        const result = await window.mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        el.documentPage.innerHTML = result.value || '<p>Empty document</p>';
        updateDocStats();
        markSaved();
        showToast('Document opened', 'success');
      } else {
        throw new Error('Mammoth library not loaded');
      }
    } catch (err) {
      showToast('Could not open this .docx file: ' + err.message, 'error');
    }
  }

  // --- DOCX EXPORT (Client-side OpenXML builder + Server Pandoc) ---
  function escapeXml(unsafe) {
    return (unsafe || '').replace(/[<>&'"]/g, function (c) {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '\'': return '&apos;';
        case '"': return '&quot;';
      }
    });
  }

  function generateClientOpenXmlDocx(containerElem) {
    const zip = new window.JSZip();

    // 1. [Content_Types].xml
    zip.file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>'
    );

    // 2. _rels/.rels
    zip.folder('_rels').file('.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
    );

    // 3. word/_rels/document.xml.rels
    zip.folder('word').folder('_rels').file('document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>'
    );

    // 4. word/styles.xml
    zip.folder('word').file('styles.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:color w:val="1E293B"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="48"/><w:color w:val="0F172A"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1E293B"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="334155"/></w:rPr></w:style>' +
      '</w:styles>'
    );

    // 5. Build document.xml body from DOM
    let bodyXml = '';

    function parseNode(node, currentStyle = {}) {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent;
        if (!text) return '';
        let rPr = '';
        if (currentStyle.bold) rPr += '<w:b/>';
        if (currentStyle.italic) rPr += '<w:i/>';
        if (currentStyle.underline) rPr += '<w:u w:val="single"/>';
        if (currentStyle.strike) rPr += '<w:strike/>';
        if (currentStyle.size) rPr += `<w:sz w:val="${currentStyle.size}"/>`;
        if (currentStyle.color) {
          const hex = currentStyle.color.replace('#', '');
          if (/^[0-9A-Fa-f]{6}$/.test(hex)) rPr += `<w:color w:val="${hex}"/>`;
        }
        return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return '';

      const tag = node.tagName.toLowerCase();
      const nextStyle = { ...currentStyle };

      if (tag === 'strong' || tag === 'b') nextStyle.bold = true;
      if (tag === 'em' || tag === 'i') nextStyle.italic = true;
      if (tag === 'u') nextStyle.underline = true;
      if (tag === 's' || tag === 'strike') nextStyle.strike = true;

      // Check inline style font color
      if (node.style && node.style.color) {
        nextStyle.color = node.style.color;
      }

      if (tag === 'h1') {
        let inner = '';
        node.childNodes.forEach(child => inner += parseNode(child, nextStyle));
        return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${inner}</w:p>`;
      }
      if (tag === 'h2') {
        let inner = '';
        node.childNodes.forEach(child => inner += parseNode(child, nextStyle));
        return `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${inner}</w:p>`;
      }
      if (tag === 'h3') {
        let inner = '';
        node.childNodes.forEach(child => inner += parseNode(child, nextStyle));
        return `<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>${inner}</w:p>`;
      }
      if (tag === 'p' || tag === 'div') {
        let inner = '';
        node.childNodes.forEach(child => inner += parseNode(child, nextStyle));
        let align = '';
        if (node.style.textAlign === 'center') align = '<w:jc w:val="center"/>';
        else if (node.style.textAlign === 'right') align = '<w:jc w:val="right"/>';
        else if (node.style.textAlign === 'justify') align = '<w:jc w:val="both"/>';
        return `<w:p><w:pPr>${align}</w:pPr>${inner}</w:p>`;
      }
      if (tag === 'ul' || tag === 'ol') {
        let inner = '';
        node.childNodes.forEach(li => {
          if (li.tagName && li.tagName.toLowerCase() === 'li') {
            let liInner = '';
            li.childNodes.forEach(child => liInner += parseNode(child, nextStyle));
            inner += `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>• </w:t></w:r>${liInner}</w:p>`;
          }
        });
        return inner;
      }
      if (tag === 'table') {
        let tblXml = '<w:tbl><w:tblPr><w:tblBorders>' +
          '<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>' +
          '</w:tblBorders></w:tblPr>';

        const rows = node.querySelectorAll('tr');
        rows.forEach(tr => {
          tblXml += '<w:tr>';
          const cells = tr.querySelectorAll('th, td');
          cells.forEach(tc => {
            let tcInner = '';
            tc.childNodes.forEach(child => tcInner += parseNode(child, nextStyle));
            if (!tcInner) tcInner = '<w:p/>';
            tblXml += `<w:tc><w:p>${tcInner}</w:p></w:tc>`;
          });
          tblXml += '</w:tr>';
        });
        tblXml += '</w:tbl>';
        return tblXml;
      }

      let childrenXml = '';
      node.childNodes.forEach(child => {
        childrenXml += parseNode(child, nextStyle);
      });
      return childrenXml;
    }

    containerElem.childNodes.forEach(node => {
      bodyXml += parseNode(node);
    });

    if (!bodyXml.trim()) {
      bodyXml = '<w:p><w:r><w:t>Document Content</w:t></w:r></w:p>';
    }

    zip.folder('word').file('document.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' + bodyXml + '</w:body>' +
      '</w:document>'
    );

    return zip.generateAsync({ type: 'blob' });
  }

  async function exportDocx() {
    const filename = (state.docTitle || 'document').replace(/\.[^/.]+$/, '') + '.docx';
    const html = el.documentPage.innerHTML;

    try {
      let blob = null;

      // Option 1: Pandoc Server Export (Highest Fidelity)
      if (state.pandocOnline) {
        try {
          const res = await fetch('/api/export-docx', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ html: html, filename: filename })
          });
          if (res.ok) blob = await res.blob();
        } catch (e) {
          console.warn('Server export failed, using client builder', e);
        }
      }

      // Option 2: Pure Client-Side OpenXML Generator (Works Offline!)
      if (!blob) blob = await generateClientOpenXmlDocx(el.documentPage);

      await saveBlob(blob, filename);
    } catch (err) {
      showToast('Could not save this .docx file: ' + err.message, 'error');
    }
  }

  // Saves through the native panel in the app, or as a download in a browser.
  // Reports success only once the file exists. Cancelling is not an error.
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

  function getActiveSheetData() {
    return state.sheet.sheets[state.sheet.activeSheet] || null;
  }

  function evaluateFormula(expr, sheetData, visited = new Set()) {
    if (!expr.startsWith('=')) return expr;
    const raw = expr.substring(1).trim();

    // Check SUM function: =SUM(A1:A5) or =SUM(A1, B1)
    const sumMatch = raw.match(/^SUM\(([^)]+)\)$/i);
    if (sumMatch) {
      const args = sumMatch[1].split(',');
      let total = 0;
      args.forEach(arg => {
        const rangeParts = arg.trim().split(':');
        if (rangeParts.length === 2) {
          const start = addressToCoord(rangeParts[0]);
          const end = addressToCoord(rangeParts[1]);
          if (start && end) {
            const minR = Math.min(start.r, end.r);
            const maxR = Math.max(start.r, end.r);
            const minC = Math.min(start.c, end.c);
            const maxC = Math.max(start.c, end.c);
            for (let r = minR; r <= maxR; r++) {
              for (let c = minC; c <= maxC; c++) {
                const addr = coordToAddress(r, c);
                const val = getCellValue(addr, sheetData, visited);
                const num = parseFloat(val);
                if (!isNaN(num)) total += num;
              }
            }
          }
        } else {
          const val = getCellValue(arg.trim(), sheetData, visited);
          const num = parseFloat(val);
          if (!isNaN(num)) total += num;
        }
      });
      return total;
    }

    // Check AVERAGE function
    const avgMatch = raw.match(/^AVERAGE\(([^)]+)\)$/i);
    if (avgMatch) {
      const rangeParts = avgMatch[1].trim().split(':');
      let total = 0, count = 0;
      if (rangeParts.length === 2) {
        const start = addressToCoord(rangeParts[0]);
        const end = addressToCoord(rangeParts[1]);
        if (start && end) {
          for (let r = Math.min(start.r, end.r); r <= Math.max(start.r, end.r); r++) {
            for (let c = Math.min(start.c, end.c); c <= Math.max(start.c, end.c); c++) {
              const addr = coordToAddress(r, c);
              const val = getCellValue(addr, sheetData, visited);
              const num = parseFloat(val);
              if (!isNaN(num)) {
                total += num;
                count++;
              }
            }
          }
        }
      }
      return count > 0 ? (total / count) : 0;
    }

    // Check COUNT function
    const countMatch = raw.match(/^COUNT\(([^)]+)\)$/i);
    if (countMatch) {
      const rangeParts = countMatch[1].trim().split(':');
      let count = 0;
      if (rangeParts.length === 2) {
        const start = addressToCoord(rangeParts[0]);
        const end = addressToCoord(rangeParts[1]);
        if (start && end) {
          for (let r = Math.min(start.r, end.r); r <= Math.max(start.r, end.r); r++) {
            for (let c = Math.min(start.c, end.c); c <= Math.max(start.c, end.c); c++) {
              const addr = coordToAddress(r, c);
              const val = getCellValue(addr, sheetData, visited);
              if (val !== '' && !isNaN(parseFloat(val))) count++;
            }
          }
        }
      }
      return count;
    }

    // Basic arithmetic evaluation: replace cell coordinates (e.g. A1, B2) with evaluated numeric values
    let evalStr = raw.replace(/\b([A-Z]+[0-9]+)\b/gi, (match) => {
      const val = getCellValue(match.toUpperCase(), sheetData, visited);
      const num = parseFloat(val);
      return isNaN(num) ? '0' : num;
    });

    try {
      // Safe arithmetic evaluator (only math operators and numbers allowed)
      if (/^[0-9+\-*/().\s]+$/.test(evalStr)) {
        // eslint-disable-next-line no-new-func
        const res = Function(`"use strict"; return (${evalStr});`)();
        return isNaN(res) ? '#VALUE!' : res;
      }
    } catch (e) {
      return '#ERROR!';
    }

    return raw;
  }

  function getCellValue(addr, sheetData, visited = new Set()) {
    if (visited.has(addr)) return '#CYCLE!';
    visited.add(addr);

    const cell = sheetData.cells[addr];
    if (!cell) return '';

    if (cell.f) {
      return evaluateFormula(cell.f, sheetData, visited);
    }
    return cell.v !== undefined ? cell.v : '';
  }

  function formatCellValue(val, fmt) {
    if (val === '' || val === null || val === undefined) return '';
    const num = parseFloat(val);
    if (isNaN(num)) return val;

    switch (fmt) {
      case 'currency':
        return '$' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'percent':
        return (num * 100).toFixed(1) + '%';
      case 'decimal':
        return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'integer':
        return Math.round(num).toLocaleString('en-US');
      case 'text':
        return String(val);
      case 'general':
      default:
        return typeof val === 'number' ? num.toLocaleString('en-US') : val;
    }
  }

  function renderGrid() {
    const sheetData = getActiveSheetData();
    if (!sheetData) return;

    const rowCount = sheetData.rowCount;
    const colCount = sheetData.colCount;

    let html = '<thead><tr><th class="grid-corner"></th>';
    for (let c = 0; c < colCount; c++) {
      html += `<th class="col-header" data-col="${c}">${indexToCol(c)}</th>`;
    }
    html += '</tr></thead><tbody>';

    for (let r = 0; r < rowCount; r++) {
      html += `<tr><th class="row-header" data-row="${r}">${r + 1}</th>`;
      for (let c = 0; c < colCount; c++) {
        const addr = coordToAddress(r, c);
        const cell = sheetData.cells[addr] || {};
        const displayVal = cell.f ? evaluateFormula(cell.f, sheetData) : (cell.v || '');
        const formatted = formatCellValue(displayVal, (cell.s && cell.s.format) || 'general');

        let cellStyle = '';
        if (cell.s) {
          if (cell.s.bold) cellStyle += 'font-weight:bold;';
          if (cell.s.italic) cellStyle += 'font-style:italic;';
          if (cell.s.align) cellStyle += `text-align:${cell.s.align};`;
          if (cell.s.color) cellStyle += `color:${cell.s.color};`;
          if (cell.s.bg) cellStyle += `background-color:${cell.s.bg};`;
        }

        const isSelected = state.sheet.selectedCell.r === r && state.sheet.selectedCell.c === c;
        html += `<td class="grid-cell ${isSelected ? 'selected' : ''}" data-row="${r}" data-col="${c}" data-addr="${addr}" style="${cellStyle}">` +
          `<div class="cell-content">${escapeXml(String(formatted))}</div>` +
          `</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody>';

    el.spreadsheetGrid.innerHTML = html;
    updateFormulaBar();
    updateSheetStats();
    renderSheetTabs();
  }

  function selectCell(r, c) {
    state.sheet.selectedCell = { r, c };
    const prevSelected = el.spreadsheetGrid.querySelector('.grid-cell.selected');
    if (prevSelected) prevSelected.classList.remove('selected');

    const td = el.spreadsheetGrid.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
    if (td) {
      td.classList.add('selected');
      td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    updateFormulaBar();
    updateSheetStats();
  }

  function updateFormulaBar() {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const addr = coordToAddress(r, c);
    el.cellAddress.textContent = addr;

    const cell = (sheetData && sheetData.cells[addr]) || {};
    el.formulaInput.value = cell.f || cell.v || '';
    syncSheetToolbar(cell, r, c);
  }

  // Reflect the selected cell in the toolbar and in the row and column headers.
  function syncSheetToolbar(cell, r, c) {
    const style = cell.s || {};
    setPressed('sheetBold', !!style.bold);
    setPressed('sheetItalic', !!style.italic);
    setPressed('sheetAlignLeft', style.align === 'left');
    setPressed('sheetAlignCenter', style.align === 'center');
    setPressed('sheetAlignRight', style.align === 'right');
    document.getElementById('sheetNumberFormat').value = style.format || 'general';

    el.spreadsheetGrid.querySelectorAll('.is-active').forEach((th) => th.classList.remove('is-active'));
    const colHeader = el.spreadsheetGrid.querySelector(`.col-header[data-col="${c}"]`);
    const rowHeader = el.spreadsheetGrid.querySelector(`.row-header[data-row="${r}"]`);
    if (colHeader) colHeader.classList.add('is-active');
    if (rowHeader) rowHeader.classList.add('is-active');
  }

  function setCellVal(r, c, inputVal) {
    const sheetData = getActiveSheetData();
    if (!sheetData) return;
    writeCellInput(sheetData, r, c, inputVal);
    markSheetModified();
    renderGrid();
  }

  function writeCellInput(sheetData, r, c, inputVal) {
    const addr = coordToAddress(r, c);

    if (!inputVal || inputVal.trim() === '') {
      delete sheetData.cells[addr];
    } else {
      if (!sheetData.cells[addr]) {
        sheetData.cells[addr] = { s: {} };
      }
      const cell = sheetData.cells[addr];
      if (inputVal.startsWith('=')) {
        cell.f = inputVal.toUpperCase();
        cell.v = evaluateFormula(cell.f, sheetData);
      } else {
        delete cell.f;
        cell.v = isNaN(Number(inputVal)) ? inputVal : Number(inputVal);
      }
    }
  }

  // Keeps empty rows and columns off the printed page. The grid's next render clears the marks.
  function markPrintRange() {
    let lastRow = -1;
    let lastCol = -1;
    Object.keys(getActiveSheetData().cells).forEach((addr) => {
      const { r, c } = addressToCoord(addr);
      lastRow = Math.max(lastRow, r);
      lastCol = Math.max(lastCol, c);
    });

    el.spreadsheetGrid.querySelectorAll('tbody tr').forEach((row, r) => {
      row.classList.toggle('print-hidden', r > lastRow);
    });
    el.spreadsheetGrid.querySelectorAll('[data-col]').forEach((cell) => {
      cell.classList.toggle('print-hidden', Number(cell.dataset.col) > lastCol);
    });
  }

  // --- Spreadsheet Clipboard ---
  const CELL_CLIPBOARD_TYPE = 'application/x-officelite-cell';

  function gridHasFocus() {
    if (state.mode !== 'sheet' || state.sheet.editingCell) return false;
    if (document.querySelector('.modal-overlay.active')) return false;
    const active = document.activeElement;
    return !(active && active.closest('input, select, textarea, [contenteditable="true"]'));
  }

  function copySelectedCell(clipboardData) {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const cell = sheetData.cells[coordToAddress(r, c)] || {};
    const value = cell.f ? evaluateFormula(cell.f, sheetData) : (cell.v ?? '');
    const format = (cell.s && cell.s.format) || 'general';
    // Plain text for other apps, the full cell for pasting back into OfficeLite.
    clipboardData.setData('text/plain', String(formatCellValue(value, format)));
    clipboardData.setData(CELL_CLIPBOARD_TYPE, JSON.stringify(cell));
  }

  function pasteIntoGrid(clipboardData) {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const rawCell = clipboardData.getData(CELL_CLIPBOARD_TYPE);

    if (rawCell) {
      sheetData.cells[coordToAddress(r, c)] = JSON.parse(rawCell);
    } else {
      // Rows and columns copied from Excel, Numbers, or a web table arrive as tab separated text.
      const rows = clipboardData.getData('text/plain').replace(/\r?\n$/, '').split(/\r?\n/);
      rows.forEach((row, rowOffset) => {
        row.split('\t').forEach((value, colOffset) => {
          sheetData.rowCount = Math.max(sheetData.rowCount, r + rowOffset + 1);
          sheetData.colCount = Math.max(sheetData.colCount, c + colOffset + 1);
          writeCellInput(sheetData, r + rowOffset, c + colOffset, value);
        });
      });
    }

    markSheetModified();
    renderGrid();
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
      copySelectedCell(e.clipboardData);
    });

    document.addEventListener('cut', (e) => {
      if (!gridHasFocus()) return;
      e.preventDefault();
      copySelectedCell(e.clipboardData);
      const { r, c } = state.sheet.selectedCell;
      setCellVal(r, c, '');
    });

    document.addEventListener('paste', (e) => {
      if (!gridHasFocus()) return;
      e.preventDefault();
      pasteIntoGrid(e.clipboardData);
    });
  }

  function updateSheetStats() {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const addr = coordToAddress(r, c);
    el.sheetStatCell.textContent = `Cell: ${addr}`;

    const cell = (sheetData && sheetData.cells[addr]) || {};
    const val = cell.f ? evaluateFormula(cell.f, sheetData) : cell.v;
    const num = parseFloat(val);
    el.sheetStatSum.textContent = !isNaN(num) ? `Value: ${num.toLocaleString()}` : `Value: ${val || '(empty)'}`;

    // Count non-empty cells
    const count = Object.keys(sheetData.cells).length;
    el.sheetStatCount.textContent = `Cells filled: ${count}`;
  }

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
        if (name !== state.sheet.activeSheet) {
          state.sheet.activeSheet = name;
          renderGrid();
        }
      };

      tab.onkeydown = (e) => {
        if (e.target !== tab) return;
        if (e.key === 'Enter') tab.click();
        else if (e.key === 'F2') startTabRename(tab, name);
      };

      el.tabsList.appendChild(tab);
    });
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
    let newName = name;
    if (!newName) {
      let i = state.sheet.sheetNames.length + 1;
      while (state.sheet.sheetNames.includes(`Sheet${i}`)) i++;
      newName = `Sheet${i}`;
    }
    state.sheet.sheetNames.push(newName);
    state.sheet.sheets[newName] = {
      cells: {},
      rowCount: 40,
      colCount: 20
    };
    state.sheet.activeSheet = newName;
    markSheetModified();
    renderGrid();
    showToast(`Added ${newName}`);
  }

  function renameSheet(oldName, newName) {
    if (state.sheet.sheetNames.includes(newName)) {
      showToast(`A sheet named ${newName} already exists`, 'error');
      renderSheetTabs();
      return;
    }
    const idx = state.sheet.sheetNames.indexOf(oldName);
    state.sheet.sheetNames[idx] = newName;
    state.sheet.sheets[newName] = state.sheet.sheets[oldName];
    delete state.sheet.sheets[oldName];
    if (state.sheet.activeSheet === oldName) state.sheet.activeSheet = newName;
    markSheetModified();
    renderGrid();
  }

  function deleteSheet(name) {
    if (state.sheet.sheetNames.length <= 1) return;
    confirmDialog({
      title: `Delete ${name}?`,
      message: 'Everything on this sheet will be removed. You cannot undo this.',
      confirmLabel: 'Delete sheet'
    }).then((confirmed) => {
      if (!confirmed) return;
      state.sheet.sheetNames = state.sheet.sheetNames.filter(n => n !== name);
      delete state.sheet.sheets[name];
      if (state.sheet.activeSheet === name) {
        state.sheet.activeSheet = state.sheet.sheetNames[0];
      }
      markSheetModified();
      renderGrid();
    });
  }

  // --- Spreadsheet Ribbon and Grid Events ---
  function initSheetEvents() {
    // Add Sheet button
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

    // Formula Input change
    el.formulaInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        const { r, c } = state.sheet.selectedCell;
        setCellVal(r, c, el.formulaInput.value);
        selectCell(Math.min(r + 1, getActiveSheetData().rowCount - 1), c);
      }
    });

    // Grid Cell Click & Double Click
    el.spreadsheetGrid.addEventListener('click', (e) => {
      const td = e.target.closest('td.grid-cell');
      if (td) {
        const r = parseInt(td.getAttribute('data-row'), 10);
        const c = parseInt(td.getAttribute('data-col'), 10);
        selectCell(r, c);
      }
    });

    el.spreadsheetGrid.addEventListener('dblclick', (e) => {
      const td = e.target.closest('td.grid-cell');
      if (td) {
        startCellEdit(td);
      }
    });

    // Keyboard navigation in grid
    document.addEventListener('keydown', (e) => {
      if (state.mode !== 'sheet') return;
      if (state.sheet.editingCell) return;
      if (e.target.closest && e.target.closest('input, select, button, [role="tab"], .modal-overlay')) return;
      if (document.querySelector('.modal-overlay.active')) return;

      const { r, c } = state.sheet.selectedCell;
      const sheetData = getActiveSheetData();

      switch (e.key) {
        case 'ArrowUp':
          if (r > 0) selectCell(r - 1, c);
          e.preventDefault();
          break;
        case 'ArrowDown':
          if (r < sheetData.rowCount - 1) selectCell(r + 1, c);
          e.preventDefault();
          break;
        case 'ArrowLeft':
          if (c > 0) selectCell(r, c - 1);
          e.preventDefault();
          break;
        case 'ArrowRight':
          if (c < sheetData.colCount - 1) selectCell(r, c + 1);
          e.preventDefault();
          break;
        case 'Tab':
          e.preventDefault();
          if (e.shiftKey) {
            if (c > 0) selectCell(r, c - 1);
          } else {
            if (c < sheetData.colCount - 1) selectCell(r, c + 1);
          }
          break;
        case 'Enter':
          e.preventDefault();
          if (e.shiftKey) {
            if (r > 0) selectCell(r - 1, c);
          } else {
            if (r < sheetData.rowCount - 1) selectCell(r + 1, c);
          }
          break;
        case 'Delete':
        case 'Backspace':
          setCellVal(r, c, '');
          e.preventDefault();
          break;
        default:
          if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
            const td = el.spreadsheetGrid.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
            if (td) {
              // Otherwise the browser also types this key into the editor that just took focus.
              e.preventDefault();
              startCellEdit(td, e.key);
            }
          }
          break;
      }
    });

    // Cell Formatting buttons
    document.getElementById('sheetBold').onclick = () => toggleCellFormat('bold');
    document.getElementById('sheetItalic').onclick = () => toggleCellFormat('italic');

    document.getElementById('sheetTextColor').oninput = (e) => {
      setCellFormatProperty('color', e.target.value);
    };

    document.getElementById('sheetFillColor').oninput = (e) => {
      setCellFormatProperty('bg', e.target.value);
    };

    document.getElementById('sheetAlignLeft').onclick = () => setCellFormatProperty('align', 'left');
    document.getElementById('sheetAlignCenter').onclick = () => setCellFormatProperty('align', 'center');
    document.getElementById('sheetAlignRight').onclick = () => setCellFormatProperty('align', 'right');

    document.getElementById('sheetNumberFormat').onchange = (e) => {
      setCellFormatProperty('format', e.target.value);
    };

    // Row & Column Operations
    document.getElementById('sheetInsertRowAbove').onclick = () => insertRow(state.sheet.selectedCell.r);
    document.getElementById('sheetInsertRowBelow').onclick = () => insertRow(state.sheet.selectedCell.r + 1);
    document.getElementById('sheetDeleteRow').onclick = () => deleteRow(state.sheet.selectedCell.r);
    document.getElementById('sheetInsertColLeft').onclick = () => insertCol(state.sheet.selectedCell.c);
    document.getElementById('sheetInsertColRight').onclick = () => insertCol(state.sheet.selectedCell.c + 1);
    document.getElementById('sheetDeleteCol').onclick = () => deleteCol(state.sheet.selectedCell.c);

    // AutoSum shortcut
    document.getElementById('sheetAutoSum').onclick = () => {
      const { r, c } = state.sheet.selectedCell;
      if (r > 0) {
        const topAddr = coordToAddress(0, c);
        const aboveAddr = coordToAddress(r - 1, c);
        setCellVal(r, c, `=SUM(${topAddr}:${aboveAddr})`);
      }
    };
  }

  function startCellEdit(td, initialChar = null) {
    state.sheet.editingCell = td;
    const r = parseInt(td.getAttribute('data-row'), 10);
    const c = parseInt(td.getAttribute('data-col'), 10);
    const addr = coordToAddress(r, c);
    const sheetData = getActiveSheetData();
    const cell = sheetData.cells[addr] || {};

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-editor';
    input.value = initialChar !== null ? initialChar : (cell.f || cell.v || '');

    td.classList.add('editing');
    td.appendChild(input);
    input.focus();
    if (initialChar === null) input.select();

    let finished = false;
    const finish = (commit) => {
      // Removing a focused input can fire blur, so this may be re-entered.
      if (finished) return;
      finished = true;
      const newVal = input.value;
      input.remove();
      td.classList.remove('editing');
      state.sheet.editingCell = null;
      if (commit) setCellVal(r, c, newVal);
    };

    input.addEventListener('blur', () => finish(true));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        finish(true);
        selectCell(Math.min(r + 1, sheetData.rowCount - 1), c);
      } else if (e.key === 'Escape') {
        finish(false);
      }
    });
  }

  function toggleCellFormat(prop) {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const addr = coordToAddress(r, c);
    if (!sheetData.cells[addr]) sheetData.cells[addr] = { s: {} };
    if (!sheetData.cells[addr].s) sheetData.cells[addr].s = {};

    sheetData.cells[addr].s[prop] = !sheetData.cells[addr].s[prop];
    markSheetModified();
    renderGrid();
  }

  function setCellFormatProperty(prop, value) {
    const sheetData = getActiveSheetData();
    const { r, c } = state.sheet.selectedCell;
    const addr = coordToAddress(r, c);
    if (!sheetData.cells[addr]) sheetData.cells[addr] = { s: {} };
    if (!sheetData.cells[addr].s) sheetData.cells[addr].s = {};

    sheetData.cells[addr].s[prop] = value;
    markSheetModified();
    renderGrid();
  }

  function insertRow(atRow) {
    const sheetData = getActiveSheetData();
    const newCells = {};
    for (const [addr, cell] of Object.entries(sheetData.cells)) {
      const coord = addressToCoord(addr);
      if (coord.r >= atRow) {
        newCells[coordToAddress(coord.r + 1, coord.c)] = cell;
      } else {
        newCells[addr] = cell;
      }
    }
    sheetData.cells = newCells;
    sheetData.rowCount++;
    markSheetModified();
    renderGrid();
    showToast(`Inserted row ${atRow + 1}`);
  }

  function deleteRow(atRow) {
    const sheetData = getActiveSheetData();
    if (sheetData.rowCount <= 1) return;
    const newCells = {};
    for (const [addr, cell] of Object.entries(sheetData.cells)) {
      const coord = addressToCoord(addr);
      if (coord.r === atRow) continue;
      if (coord.r > atRow) {
        newCells[coordToAddress(coord.r - 1, coord.c)] = cell;
      } else {
        newCells[addr] = cell;
      }
    }
    sheetData.cells = newCells;
    sheetData.rowCount--;
    markSheetModified();
    renderGrid();
    showToast(`Deleted row ${atRow + 1}`);
  }

  function insertCol(atCol) {
    const sheetData = getActiveSheetData();
    const newCells = {};
    for (const [addr, cell] of Object.entries(sheetData.cells)) {
      const coord = addressToCoord(addr);
      if (coord.c >= atCol) {
        newCells[coordToAddress(coord.r, coord.c + 1)] = cell;
      } else {
        newCells[addr] = cell;
      }
    }
    sheetData.cells = newCells;
    sheetData.colCount++;
    markSheetModified();
    renderGrid();
    showToast(`Inserted column ${indexToCol(atCol)}`);
  }

  function deleteCol(atCol) {
    const sheetData = getActiveSheetData();
    if (sheetData.colCount <= 1) return;
    const newCells = {};
    for (const [addr, cell] of Object.entries(sheetData.cells)) {
      const coord = addressToCoord(addr);
      if (coord.c === atCol) continue;
      if (coord.c > atCol) {
        newCells[coordToAddress(coord.r, coord.c - 1)] = cell;
      } else {
        newCells[addr] = cell;
      }
    }
    sheetData.cells = newCells;
    sheetData.colCount--;
    markSheetModified();
    renderGrid();
    showToast(`Deleted column ${indexToCol(atCol)}`);
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
          cells[cellAddr] = {
            v: cellObj.v !== undefined ? cellObj.v : '',
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

        state.sheet.sheets[sheetName] = {
          cells: cells,
          rowCount: Math.max(range.e.r + 10, 40),
          colCount: Math.max(range.e.c + 5, 20)
        };
      });

      state.sheet.activeSheet = wb.SheetNames[0];
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
    text: { id: 49 }
  };

  const BUILT_IN_FORMATS = { 2: 'decimal', 3: 'integer', 4: 'decimal', 9: 'percent', 10: 'percent', 49: 'text' };

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
    if (/[dmyhs]/i.test(tokens)) return null;
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
      const value = cell.f ? evaluateFormula(cell.f, sheetData) : (cell.v ?? '');
      const text = String(formatCellValue(value, (cell.s && cell.s.format) || 'general'));
      const column = widths[addressToCoord(addr).c];
      column.wch = Math.max(column.wch, Math.min(text.length + 2, 60));
    }
    return widths;
  }

  async function exportXlsx() {
    try {
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
          if (cell.f) {
            cellObj.f = cell.f.replace(/^=/, '');
            const val = evaluateFormula(cell.f, sheetData);
            cellObj.v = val;
            cellObj.t = typeof val === 'number' ? 'n' : 's';
          } else {
            cellObj.v = cell.v !== undefined ? cell.v : '';
            cellObj.t = typeof cell.v === 'number' ? 'n' : 's';
          }
          ws[addr] = cellObj;
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
      confirmDialog({
        title: `Create a blank ${noun}?`,
        message: `This replaces the current ${noun}. Changes you have not saved will be lost.`,
        confirmLabel: `Create blank ${noun}`
      }).then((confirmed) => {
        if (!confirmed) return;
        if (state.mode === 'doc') {
          el.documentPage.innerHTML = '';
          state.docTitle = 'Untitled Document';
          el.docNameInput.value = state.docTitle;
          updateDocStats();
          el.documentPage.focus();
        } else {
          state.sheet = {
            activeSheet: 'Sheet1',
            sheetNames: ['Sheet1'],
            sheets: {
              'Sheet1': { cells: {}, rowCount: 40, colCount: 20 }
            },
            selectedCell: { r: 1, c: 0 },
            editingCell: null
          };
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

  function handleIncomingFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'docx' || ext === 'doc') {
      switchMode('doc');
      loadDocxFile(file);
    } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
      switchMode('sheet');
      loadSpreadsheetFile(file);
    } else if (ext === 'html' || ext === 'txt') {
      const reader = new FileReader();
      reader.onload = (re) => {
        switchMode('doc');
        if (ext === 'html') el.documentPage.innerHTML = re.target.result;
        else el.documentPage.innerText = re.target.result;
        state.docTitle = file.name.replace(/\.[^/.]+$/, '');
        el.docNameInput.value = state.docTitle;
        updateDocStats();
        markSaved();
        showToast(`Opened ${file.name}`, 'success');
      };
      reader.readAsText(file);
    } else {
      showToast(`OfficeLite cannot open .${ext} files. Try a .docx, .xlsx, or .csv file.`, 'error');
    }
  }

  // --- Sample Document & Spreadsheet Loaders ---
  function loadSampleDocument() {
    switchMode('doc');
    state.docTitle = 'Sample Project Proposal';
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
          sheet: state.sheet
        };
        localStorage.setItem('officelite_autosave', JSON.stringify(payload));
      } catch (e) {
        // storage quota exceeded or disabled
      }
    }, 4000);

    // Check if previous session exists
    try {
      const saved = localStorage.getItem('officelite_autosave');
      if (saved) {
        const data = JSON.parse(saved);
        if (data.docHtml && data.docHtml.length > 50) {
          el.documentPage.innerHTML = data.docHtml;
          state.docTitle = data.docTitle || 'Untitled Document';
          el.docNameInput.value = state.docTitle;
          updateDocStats();
        }
        if (data.sheet && data.sheet.sheetNames) {
          state.sheet = data.sheet;
          state.sheetTitle = data.sheetTitle || 'Untitled Spreadsheet';
        }
      } else {
        // Load default starting document
        loadSampleDocument();
      }
    } catch (e) {
      loadSampleDocument();
    }
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
    checkServerStatus();
    initAutoSave();
    updateDocStats();
    state.docModified = false;
    state.sheetModified = false;
    updateSaveStatus();
  }

  // Start app when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
