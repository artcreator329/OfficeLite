/**
 * OfficeLite - .docx reader and writer.
 *
 * Converts between the editor's HTML and Word's file format inside the app, so
 * every computer gets the same result. htmlToDocx needs a live DOM and JSZip.
 * docxToHtml needs DOMParser and JSZip.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OfficeLiteDocx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  };
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  /** The font menu's values, in points. styles.css draws <font size> at these sizes. */
  const FONT_SIZE_PT = { 1: 10, 2: 11, 3: 12, 4: 14, 5: 18, 6: 24, 7: 32 };

  const BODY_FONT = 'Calibri';
  const MONO_FONT = 'Courier New';
  const TEXT_COLOR = '1E293B';
  const STRONG_COLOR = '0F172A';
  const QUOTE_COLOR = '475569';
  const RULE_COLOR = 'CBD5E1';
  const ACCENT_COLOR = '1D4ED8';
  const TH_FILL = 'F1F5F9';
  const PRE_FILL = 'F8FAFC';

  /**
   * How styles.css draws each block, in Word's units: half points and twips.
   * The writer's styles.xml uses these, and the reader only adds inline styles
   * where a file differs from them, so a file made here reopens as clean HTML.
   */
  const BLOCK_DEFAULTS = {
    p: { font: BODY_FONT, sz: 24, color: TEXT_COLOR, bold: false, before: 0, after: 180, line: 360, left: 0 },
    h1: { font: BODY_FONT, sz: 45, color: STRONG_COLOR, bold: true, before: 480, after: 180, line: 288, left: 0 },
    h2: { font: BODY_FONT, sz: 36, color: STRONG_COLOR, bold: true, before: 360, after: 120, line: 320, left: 0 },
    h3: { font: BODY_FONT, sz: 30, color: TEXT_COLOR, bold: true, before: 240, after: 120, line: 336, left: 0 },
    h4: { font: BODY_FONT, sz: 24, color: TEXT_COLOR, bold: true, before: 240, after: 120, line: 360, left: 0 },
    h5: { font: BODY_FONT, sz: 24, color: TEXT_COLOR, bold: true, before: 240, after: 120, line: 360, left: 0 },
    h6: { font: BODY_FONT, sz: 24, color: TEXT_COLOR, bold: true, before: 240, after: 120, line: 360, left: 0 },
    li: { font: BODY_FONT, sz: 24, color: TEXT_COLOR, bold: false, before: 0, after: 60, line: 360, left: 0 },
    blockquote: { font: BODY_FONT, sz: 24, color: QUOTE_COLOR, bold: false, before: 240, after: 240, line: 360, left: 240 },
    pre: { font: MONO_FONT, sz: 21, color: TEXT_COLOR, bold: false, before: 180, after: 180, line: 342, left: 0, bg: PRE_FILL },
    cell: { font: BODY_FONT, sz: 21, color: TEXT_COLOR, bold: false, before: 0, after: 180, line: 342, left: 0 },
    th: { font: BODY_FONT, sz: 21, color: TEXT_COLOR, bold: true, before: 0, after: 180, line: 342, left: 0 }
  };

  const DEFAULT_SECTION = {
    pgSz: { w: 12240, h: 15840, orient: null },
    pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 }
  };

  const HIGHLIGHT_COLORS = {
    black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000',
    yellow: 'FFFF00', white: 'FFFFFF', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000',
    darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0'
  };

  const CSS_SIZE_KEYWORDS = {
    'xx-small': 6.75, 'x-small': 7.5, small: 9.75, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24,
    'xxx-large': 36, '-webkit-xxx-large': 36
  };

  const LIST_FORMATS = {
    decimal: 'decimal',
    'lower-alpha': 'lowerLetter', 'lower-latin': 'lowerLetter', a: 'lowerLetter',
    'upper-alpha': 'upperLetter', 'upper-latin': 'upperLetter', A: 'upperLetter',
    'lower-roman': 'lowerRoman', i: 'lowerRoman',
    'upper-roman': 'upperRoman', I: 'upperRoman'
  };
  const LIST_STYLE_TYPES = {
    lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman'
  };
  const BULLETS = { disc: '•', circle: '◦', square: '▪' };
  const BULLET_ORDER = ['disc', 'circle', 'square'];

  const IMAGE_TYPES = {
    'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/tiff': 'tiff'
  };
  const IMAGE_MIME_BY_EXT = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif',
    bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', webp: 'image/webp'
  };

  // --- Shared helpers ---

  const CONTROL_CHARACTERS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

  /** Escapes text for XML and removes the control characters Word refuses to open. */
  function escapeXml(text) {
    return String(text == null ? '' : text)
      .replace(CONTROL_CHARACTERS, '')
      .replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
  }

  function escapeHtml(text) {
    return String(text == null ? '' : text).replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[c]));
  }

  let colorProbe = null;

  /** Returns RRGGBB in capitals, or null for transparent, inherit, and anything unreadable. */
  function cssColorToHex(value) {
    if (!value) return null;
    const text = String(value).trim().toLowerCase();
    if (!text || text === 'transparent' || text === 'inherit' || text === 'initial' || text === 'currentcolor' || text === 'auto') return null;

    let match = text.match(/^#?([0-9a-f]{6})$/);
    if (match) return match[1].toUpperCase();
    match = text.match(/^#([0-9a-f])([0-9a-f])([0-9a-f])$/);
    if (match) return (match[1] + match[1] + match[2] + match[2] + match[3] + match[3]).toUpperCase();

    match = text.match(/^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+%?))?\s*\)$/);
    if (match) {
      if (match[4] !== undefined && parseFloat(match[4]) === 0) return null;
      return [match[1], match[2], match[3]]
        .map((part) => Math.max(0, Math.min(255, Math.round(parseFloat(part)))).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
    }

    if (typeof document !== 'undefined') {
      if (!colorProbe) colorProbe = document.createElement('canvas').getContext('2d');
      if (colorProbe) {
        colorProbe.fillStyle = '#000000';
        colorProbe.fillStyle = text;
        const resolved = colorProbe.fillStyle;
        if (resolved !== '#000000' || text === 'black') return cssColorToHex(resolved);
      }
    }
    return null;
  }

  /** Converts a CSS length to twips. Returns null for auto, percentages, and anything else. */
  function cssLengthToTwips(value, emTwips = 240) {
    if (value === null || value === undefined) return null;
    const text = String(value).trim().toLowerCase();
    if (text === '0') return 0;
    const match = text.match(/^(-?[\d.]+)(pt|px|in|cm|mm|em|rem)$/);
    if (!match) return null;
    const amount = parseFloat(match[1]);
    const perUnit = { pt: 20, px: 15, in: 1440, cm: 567, mm: 56.7, em: emTwips, rem: 240 };
    return Math.round(amount * perUnit[match[2]]);
  }

  /** Converts a CSS font size to half points. */
  function cssFontSizeToHalfPoints(value, parentHalfPoints = 24) {
    if (!value) return null;
    const text = String(value).trim().toLowerCase();
    if (CSS_SIZE_KEYWORDS[text]) return Math.round(CSS_SIZE_KEYWORDS[text] * 2);
    if (text === 'smaller') return Math.round(parentHalfPoints / 1.2);
    if (text === 'larger') return Math.round(parentHalfPoints * 1.2);
    let match = text.match(/^([\d.]+)%$/);
    if (match) return Math.round(parentHalfPoints * parseFloat(match[1]) / 100);
    match = text.match(/^([\d.]+)r?em$/);
    if (match) return Math.round(parentHalfPoints * parseFloat(match[1]));
    const twips = cssLengthToTwips(text);
    return twips === null ? null : Math.round(twips / 10);
  }

  /** The first family in a CSS font list, without quotes or generic names. */
  function firstFontFamily(value) {
    if (!value) return null;
    const first = String(value).split(',')[0].trim().replace(/^['"]|['"]$/g, '').trim();
    if (!first || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|-apple-system|inherit|initial)$/i.test(first)) return null;
    return first;
  }

  function bookmarkName(name) {
    const clean = String(name || '').replace(/[^A-Za-z0-9_]/g, '_').slice(0, 40);
    return /^[A-Za-z_]/.test(clean) ? clean : ('_' + clean).slice(0, 40);
  }

  function isSafeHref(href) {
    return /^(https?:|mailto:|tel:|#)/i.test(String(href || '').trim());
  }

  // ==========================================
  // --- WRITER: editor HTML to .docx ---
  // ==========================================

  const BLOCK_TAGS = new Set([
    'address', 'article', 'aside', 'blockquote', 'center', 'dd', 'details', 'dir', 'div', 'dl', 'dt', 'fieldset',
    'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'li',
    'main', 'menu', 'nav', 'ol', 'p', 'pre', 'section', 'summary', 'table', 'ul'
  ]);
  const BLOCK_SELECTOR = [...BLOCK_TAGS].join(',');
  const SKIPPED_TAGS = new Set(['script', 'style', 'meta', 'link', 'title', 'head', 'noscript', 'template', 'colgroup', 'col']);
  const LINE_BREAK = '<w:r><w:br/></w:r>';

  function normalizeSection(section) {
    const number = (value, fallback) => {
      const parsed = Number(value);
      return Number.isFinite(parsed) ? Math.round(parsed) : fallback;
    };
    const size = (section && section.pgSz) || {};
    const margin = (section && section.pgMar) || {};
    const pgMar = {};
    Object.keys(DEFAULT_SECTION.pgMar).forEach((key) => { pgMar[key] = number(margin[key], DEFAULT_SECTION.pgMar[key]); });
    return {
      pgSz: {
        w: number(size.w, DEFAULT_SECTION.pgSz.w),
        h: number(size.h, DEFAULT_SECTION.pgSz.h),
        orient: size.orient === 'landscape' ? 'landscape' : null
      },
      pgMar
    };
  }

  function runProperties(style) {
    let xml = '';
    if (style.link) xml += '<w:rStyle w:val="Hyperlink"/>';
    if (style.font) {
      const font = escapeXml(style.font);
      xml += `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`;
    }
    if (style.bold === true) xml += '<w:b/>';
    else if (style.bold === false) xml += '<w:b w:val="0"/>';
    if (style.italic === true) xml += '<w:i/>';
    else if (style.italic === false) xml += '<w:i w:val="0"/>';
    if (style.caps) xml += '<w:caps/>';
    if (style.strike) xml += '<w:strike/>';
    if (style.color) xml += `<w:color w:val="${style.color}"/>`;
    if (style.sz) xml += `<w:sz w:val="${style.sz}"/><w:szCs w:val="${style.sz}"/>`;
    if (style.underline) xml += '<w:u w:val="single"/>';
    if (style.bg) xml += `<w:shd w:val="clear" w:color="auto" w:fill="${style.bg}"/>`;
    if (style.vert) xml += `<w:vertAlign w:val="${style.vert}"/>`;
    return xml ? `<w:rPr>${xml}</w:rPr>` : '';
  }

  function isBoldWeight(weight) {
    if (weight === 'bold' || weight === 'bolder') return true;
    const numeric = parseInt(weight, 10);
    return Number.isNaN(numeric) ? null : numeric >= 600;
  }

  /** Adds what an element says about its text to the style it inherited. */
  function inlineStyleOf(element, inherited) {
    const style = { ...inherited };
    const tag = element.tagName.toLowerCase();

    if (tag === 'b' || tag === 'strong') style.bold = true;
    if (tag === 'i' || tag === 'em' || tag === 'cite' || tag === 'var') style.italic = true;
    if (tag === 'u' || tag === 'ins') style.underline = true;
    if (tag === 's' || tag === 'strike' || tag === 'del') style.strike = true;
    if (tag === 'sup') style.vert = 'superscript';
    if (tag === 'sub') style.vert = 'subscript';
    if (tag === 'code' || tag === 'kbd' || tag === 'samp' || tag === 'tt') style.font = MONO_FONT;
    if (tag === 'mark') style.bg = 'FFFF00';

    if (tag === 'font') {
      const color = cssColorToHex(element.getAttribute('color'));
      if (color) style.color = color;
      const face = firstFontFamily(element.getAttribute('face'));
      if (face) style.font = face;
      const size = FONT_SIZE_PT[element.getAttribute('size')];
      if (size) style.sz = size * 2;
    }

    const css = element.style;
    if (css) {
      if (css.fontWeight) {
        if (css.fontWeight === 'normal') style.bold = false;
        else {
          const bold = isBoldWeight(css.fontWeight);
          if (bold !== null) style.bold = bold;
        }
      }
      if (css.fontStyle === 'italic' || css.fontStyle === 'oblique') style.italic = true;
      else if (css.fontStyle === 'normal') style.italic = false;

      const decoration = `${css.textDecorationLine || ''} ${css.textDecoration || ''}`;
      if (decoration.includes('underline')) style.underline = true;
      if (decoration.includes('line-through')) style.strike = true;

      const color = cssColorToHex(css.color);
      if (color) style.color = color;
      const background = cssColorToHex(css.backgroundColor);
      if (background) style.bg = background;
      const size = cssFontSizeToHalfPoints(css.fontSize, inherited.sz || 24);
      if (size) style.sz = size;
      const family = firstFontFamily(css.fontFamily);
      if (family) style.font = family;
      if (css.verticalAlign === 'super') style.vert = 'superscript';
      if (css.verticalAlign === 'sub') style.vert = 'subscript';
      if (css.textTransform === 'uppercase') style.caps = true;
    }
    return style;
  }

  /** A block's background fills the block, so its text does not repeat it. */
  function blockStyleOf(element, inherited) {
    return { ...inlineStyleOf(element, inherited), bg: inherited.bg };
  }

  function blockPropertiesOf(element) {
    const css = element.style || {};
    const props = {};

    let align = css.textAlign || element.getAttribute('align') || '';
    if (element.tagName.toLowerCase() === 'center') align = 'center';
    align = align.replace(/^-webkit-/, '').toLowerCase();
    const alignments = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', justify: 'both' };
    if (alignments[align]) props.align = alignments[align];

    const marginLeft = cssLengthToTwips(css.marginLeft);
    const paddingLeft = cssLengthToTwips(css.paddingLeft);
    if (marginLeft !== null || paddingLeft !== null) props.left = (marginLeft || 0) + (paddingLeft || 0);
    const marginRight = cssLengthToTwips(css.marginRight);
    if (marginRight !== null) props.right = marginRight;
    const textIndent = cssLengthToTwips(css.textIndent);
    if (textIndent !== null) props.first = textIndent;

    const before = cssLengthToTwips(css.marginTop);
    if (before !== null) props.before = before;
    const after = cssLengthToTwips(css.marginBottom);
    if (after !== null) props.after = after;

    const lineHeight = css.lineHeight;
    if (lineHeight && lineHeight !== 'normal') {
      if (/^[\d.]+$/.test(lineHeight)) props.line = Math.round(parseFloat(lineHeight) * 240);
      else if (/^[\d.]+%$/.test(lineHeight)) props.line = Math.round(parseFloat(lineHeight) * 2.4);
      else {
        const exact = cssLengthToTwips(lineHeight);
        if (exact) {
          props.line = exact;
          props.lineExact = true;
        }
      }
    }

    const background = cssColorToHex(css.backgroundColor);
    if (background) props.bg = background;
    return props;
  }

  class DocxWriter {
    constructor(section, images) {
      this.section = section;
      this.images = images;
      this.relationships = [{ type: `${REL}/styles`, target: 'styles.xml' }, { type: `${REL}/settings`, target: 'settings.xml' }];
      this.abstractNums = [];
      this.nums = [];
      this.bookmarkId = 0;
      this.drawingId = 0;
      this.mediaByKey = new Map();
    }

    get contentWidth() {
      return this.section.pgSz.w - this.section.pgMar.left - this.section.pgMar.right;
    }

    addRelationship(type, target, external) {
      const existing = this.relationships.findIndex((rel) => rel.type === type && rel.target === target);
      if (existing >= 0) return `rId${existing + 1}`;
      this.relationships.push({ type, target, external });
      return `rId${this.relationships.length}`;
    }

    // --- Blocks ---

    paragraph(content, inherited, own) {
      let props = '';
      const left = (inherited.left || 0) + (own.left || 0);
      let indent = left;

      if (inherited.styleId) props += `<w:pStyle w:val="${inherited.styleId}"/>`;
      if (inherited.num) {
        if (!inherited.num.used) {
          inherited.num.used = true;
          props += `<w:numPr><w:ilvl w:val="${inherited.num.level}"/><w:numId w:val="${inherited.num.id}"/></w:numPr>`;
        } else {
          indent += 360 * (inherited.num.level + 1);
        }
      }
      if (inherited.bg || own.bg) props += `<w:shd w:val="clear" w:color="auto" w:fill="${own.bg || inherited.bg}"/>`;

      const spacing = [];
      if (own.before !== undefined) spacing.push(`w:before="${own.before}"`);
      if (own.after !== undefined) spacing.push(`w:after="${own.after}"`);
      if (own.line !== undefined) spacing.push(`w:line="${own.line}" w:lineRule="${own.lineExact ? 'exact' : 'auto'}"`);
      if (spacing.length) props += `<w:spacing ${spacing.join(' ')}/>`;

      const ind = [];
      if (indent) ind.push(`w:left="${indent}"`);
      if (own.right) ind.push(`w:right="${own.right}"`);
      if (own.first > 0) ind.push(`w:firstLine="${own.first}"`);
      if (own.first < 0) ind.push(`w:hanging="${-own.first}"`);
      if (ind.length) props += `<w:ind ${ind.join(' ')}/>`;

      const align = own.align || inherited.align;
      if (align) props += `<w:jc w:val="${align}"/>`;

      return `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${content}</w:p>`;
    }

    container(element, style, inherited, own, out) {
      let para = null;
      const flush = () => {
        if (!para) return;
        if (para.xml.endsWith(LINE_BREAK)) para.xml = para.xml.slice(0, -LINE_BREAK.length);
        if (para.content) out.push(this.paragraph(para.xml, inherited, own));
        para = null;
      };

      element.childNodes.forEach((node) => {
        if (node.nodeType === 3) {
          if (!para && !inherited.preserve && !/[^ \t\r\n\f]/.test(node.nodeValue)) return;
          para = para || { xml: '', content: false, endsSpace: true };
          this.text(node.nodeValue, style, inherited.preserve, para);
          return;
        }
        if (node.nodeType !== 1) return;

        const tag = node.tagName.toLowerCase();
        if (SKIPPED_TAGS.has(tag)) return;
        if (BLOCK_TAGS.has(tag)) {
          flush();
          this.block(node, style, inherited, out);
        } else if (node.querySelector(BLOCK_SELECTOR)) {
          flush();
          this.container(node, inlineStyleOf(node, style), inherited, {}, out);
        } else {
          para = para || { xml: '', content: false, endsSpace: true };
          this.inline(node, style, inherited, para);
        }
      });
      flush();
    }

    block(element, inheritedStyle, inherited, out) {
      const tag = element.tagName.toLowerCase();
      const style = blockStyleOf(element, inheritedStyle);
      const props = blockPropertiesOf(element);
      const next = { ...inherited };
      if (props.align) next.align = props.align;
      if (props.bg) next.bg = props.bg;

      const own = {
        before: props.before, after: props.after, line: props.line, lineExact: props.lineExact,
        first: props.first, right: props.right, left: props.left
      };
      const startedAt = out.length;
      const ensureParagraph = () => {
        if (out.length === startedAt) out.push(this.paragraph('', next, own));
      };

      if (tag === 'hr') {
        if (element.hasAttribute('data-page-break')) {
          out.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
        } else {
          out.push(`<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${RULE_COLOR}"/></w:pBdr></w:pPr></w:p>`);
        }
        return;
      }

      const heading = tag.match(/^h([1-6])$/);
      if (heading) {
        next.styleId = `Heading${heading[1]}`;
        this.container(element, style, next, own, out);
        ensureParagraph();
        return;
      }

      if (tag === 'ul' || tag === 'ol' || tag === 'menu' || tag === 'dir') {
        this.list(element, style, next, out);
        return;
      }
      if (tag === 'table') {
        this.table(element, style, next, out);
        return;
      }

      if (tag === 'blockquote') {
        const css = element.style || {};
        const isPlainIndent = css.borderStyle === 'none' || css.border === 'none' || /^none/.test(css.borderLeft || '');
        if (isPlainIndent) {
          next.left = (inherited.left || 0) + (props.left || 0);
          this.container(element, style, next, { before: props.before, after: props.after }, out);
        } else {
          next.styleId = 'Quote';
          this.container(element, style, next, own, out);
        }
        ensureParagraph();
        return;
      }

      if (tag === 'pre') {
        next.styleId = 'Preformatted';
        next.preserve = true;
        this.container(element, style, next, own, out);
        ensureParagraph();
        return;
      }

      if (tag === 'li') {
        this.container(element, style, next, own, out);
        ensureParagraph();
        return;
      }

      // Paragraphs and generic containers
      if (tag === 'p' && inherited.styleId === 'TableText' && own.after === undefined) own.after = BLOCK_DEFAULTS.cell.after;
      const isLeaf = !element.querySelector(BLOCK_SELECTOR);
      if (tag !== 'p' && isLeaf && own.after === undefined && !inherited.styleId && !inherited.num) own.after = 0;
      if (tag === 'p' || isLeaf) {
        this.container(element, style, next, own, out);
        ensureParagraph();
      } else {
        // Indents and alignment reach the children. Spacing belongs to the container.
        if (own.left) next.left = (inherited.left || 0) + own.left;
        this.container(element, style, next, {}, out);
      }
    }

    list(element, style, inherited, out) {
      const levels = [];
      const add = (list, depth) => {
        if (levels[depth]) return;
        const css = (list.style && list.style.listStyleType) || '';
        if (list.tagName.toLowerCase() === 'ol') {
          levels[depth] = { format: LIST_FORMATS[css] || LIST_FORMATS[list.getAttribute('type')] || 'decimal' };
        } else {
          levels[depth] = { format: 'bullet', bullet: BULLETS[css] || BULLETS[BULLET_ORDER[Math.min(depth, 2)]] };
        }
      };
      add(element, 0);
      element.querySelectorAll('ul,ol').forEach((list) => {
        let depth = 0;
        for (let node = list.parentElement; node && node !== element; node = node.parentElement) {
          if (/^(UL|OL)$/.test(node.tagName)) depth += 1;
        }
        add(list, depth + 1);
      });

      const fallback = element.tagName.toLowerCase() === 'ol'
        ? { format: 'decimal' }
        : { format: 'bullet', bullet: BULLETS.disc };
      const abstractId = this.abstractNums.length;
      const start = parseInt(element.getAttribute('start'), 10);
      const startsAt = start && start !== 1 ? start : 1;
      let levelsXml = '';
      for (let i = 0; i < 9; i++) {
        const level = levels[i] || (fallback.format === 'bullet'
          ? { format: 'bullet', bullet: BULLETS[BULLET_ORDER[Math.min(i, 2)]] }
          : fallback);
        const text = level.format === 'bullet' ? level.bullet : `%${i + 1}.`;
        levelsXml += `<w:lvl w:ilvl="${i}"><w:start w:val="${i === 0 ? startsAt : 1}"/><w:numFmt w:val="${level.format}"/>` +
          `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
          `<w:pPr><w:ind w:left="${360 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
      }
      this.abstractNums.push(`<w:abstractNum w:abstractNumId="${abstractId}"><w:multiLevelType w:val="hybridMultilevel"/>${levelsXml}</w:abstractNum>`);

      const id = this.nums.length + 1;
      const override = startsAt !== 1
        ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${startsAt}"/></w:lvlOverride>`
        : '';
      this.nums.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${abstractId}"/>${override}</w:num>`);

      this.listItems(element, id, 0, style, inherited, out);
    }

    listItems(list, id, level, style, inherited, out) {
      [...list.children].forEach((child) => {
        const tag = child.tagName.toLowerCase();
        if (tag === 'ul' || tag === 'ol') {
          this.listItems(child, id, level + 1, inlineStyleOf(child, style), inherited, out);
          return;
        }
        if (tag !== 'li') return;

        const itemStyle = blockStyleOf(child, style);
        const props = blockPropertiesOf(child);
        const next = { ...inherited, num: { id, level, used: false } };
        if (props.align) next.align = props.align;
        next.styleId = inherited.styleId === 'TableText' ? 'TableText' : 'ListParagraph';
        const own = { before: props.before, after: props.after, line: props.line, lineExact: props.lineExact };
        if (next.styleId === 'TableText' && own.after === undefined) own.after = BLOCK_DEFAULTS.li.after;

        this.listItemContent(child, id, level, itemStyle, next, own, out);
        if (!next.num.used) {
          out.push(this.paragraph('', next, own));
        }
      });
    }

    listItemContent(item, id, level, style, inherited, own, out) {
      let para = null;
      const flush = () => {
        if (!para) return;
        if (para.xml.endsWith(LINE_BREAK)) para.xml = para.xml.slice(0, -LINE_BREAK.length);
        if (para.content) out.push(this.paragraph(para.xml, inherited, own));
        para = null;
      };

      item.childNodes.forEach((node) => {
        if (node.nodeType === 3) {
          if (!para && !/[^ \t\r\n\f]/.test(node.nodeValue)) return;
          para = para || { xml: '', content: false, endsSpace: true };
          this.text(node.nodeValue, style, false, para);
        } else if (node.nodeType === 1) {
          const tag = node.tagName.toLowerCase();
          if (SKIPPED_TAGS.has(tag)) return;
          if (tag === 'ul' || tag === 'ol') {
            flush();
            this.listItems(node, id, level + 1, inlineStyleOf(node, style), inherited, out);
          } else if (BLOCK_TAGS.has(tag)) {
            flush();
            this.block(node, style, tag === 'table' ? { ...inherited, num: null } : inherited, out);
          } else if (node.querySelector(BLOCK_SELECTOR)) {
            flush();
            this.listItemContent(node, id, level, inlineStyleOf(node, style), inherited, own, out);
          } else {
            para = para || { xml: '', content: false, endsSpace: true };
            this.inline(node, style, inherited, para);
          }
        }
      });
      flush();
    }

    table(table, style, inherited, out) {
      if (out.length && out[out.length - 1].startsWith('<w:tbl>')) out.push('<w:p/>');

      if (table.caption) {
        this.container(table.caption, style, { align: 'center' }, {}, out);
      }

      const rows = [...table.rows];
      const occupied = [];
      const grid = rows.map(() => []);
      let columnCount = 0;

      rows.forEach((row, r) => {
        let c = 0;
        [...row.cells].forEach((cell) => {
          while (occupied[r] && occupied[r][c]) {
            grid[r].push(occupied[r][c]);
            c += occupied[r][c].colSpan;
          }
          const colSpan = Math.max(1, cell.colSpan || 1);
          const rowSpan = Math.max(1, cell.rowSpan || 1);
          grid[r].push({ cell, colSpan, rowSpan, col: c });
          for (let down = 1; down < rowSpan && r + down < rows.length; down++) {
            occupied[r + down] = occupied[r + down] || [];
            occupied[r + down][c] = { continues: true, colSpan, col: c };
          }
          c += colSpan;
        });
        while (occupied[r] && occupied[r][c]) {
          grid[r].push(occupied[r][c]);
          c += occupied[r][c].colSpan;
        }
        columnCount = Math.max(columnCount, c);
      });
      if (!columnCount) return;

      // Column widths come from the first row that states them.
      const css = table.style || {};
      let tableWidth = this.contentWidth;
      let tblW = '<w:tblW w:w="5000" w:type="pct"/>';
      const explicitWidth = cssLengthToTwips(css.width);
      if (explicitWidth) {
        tableWidth = explicitWidth;
        tblW = `<w:tblW w:w="${explicitWidth}" w:type="dxa"/>`;
      } else if (/%$/.test(css.width || '')) {
        const percent = parseFloat(css.width);
        tableWidth = Math.round(this.contentWidth * percent / 100);
        tblW = `<w:tblW w:w="${Math.round(percent * 50)}" w:type="pct"/>`;
      }

      const widths = new Array(columnCount).fill(null);
      grid.forEach((row) => row.forEach((entry) => {
        if (entry.continues || entry.colSpan !== 1 || widths[entry.col] !== null) return;
        const width = cssLengthToTwips(entry.cell.style && entry.cell.style.width);
        if (width) widths[entry.col] = width;
      }));
      const stated = widths.filter((width) => width !== null);
      const remaining = Math.max(0, tableWidth - stated.reduce((sum, width) => sum + width, 0));
      const missing = columnCount - stated.length;
      for (let i = 0; i < columnCount; i++) {
        if (widths[i] === null) widths[i] = missing ? Math.max(1, Math.round(remaining / missing)) : Math.round(tableWidth / columnCount);
      }

      const domCells = rows.flatMap((row) => [...row.cells]);
      const borderless = domCells.length > 0 && domCells.every((cell) => cell.style.borderStyle === 'none' || cell.style.border === 'none');
      const border = (side) => (borderless
        ? `<w:${side} w:val="nil"/>`
        : `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${RULE_COLOR}"/>`);
      let xml = `<w:tbl><w:tblPr>${tblW}<w:tblBorders>` +
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
        '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/>' +
        '<w:bottom w:w="120" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
        `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`;

      grid.forEach((row, r) => {
        const domRow = rows[r];
        const inHead = domRow.parentElement && domRow.parentElement.tagName === 'THEAD';
        const allHeaders = domRow.cells.length > 0 && [...domRow.cells].every((cell) => cell.tagName === 'TH');
        xml += `<w:tr>${inHead || allHeaders ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}`;

        row.forEach((entry) => {
          const span = entry.colSpan;
          let width = 0;
          for (let i = 0; i < span; i++) width += widths[entry.col + i] || 0;
          let cellProps = `<w:tcW w:w="${width}" w:type="dxa"/>`;
          if (span > 1) cellProps += `<w:gridSpan w:val="${span}"/>`;

          if (entry.continues) {
            xml += `<w:tc><w:tcPr>${cellProps}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
            return;
          }
          if (entry.rowSpan > 1) cellProps += '<w:vMerge w:val="restart"/>';

          const cell = entry.cell;
          const cellCss = cell.style || {};
          const isHeader = cell.tagName === 'TH';
          if (cellCss.borderStyle === 'none' || cellCss.border === 'none') {
            const nil = (side) => `<w:${side} w:val="nil"/>`;
            cellProps += `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map(nil).join('')}</w:tcBorders>`;
          }
          const declared = cellCss.backgroundColor;
          const fill = cssColorToHex(declared) || (isHeader && !declared ? TH_FILL : null);
          if (fill) cellProps += `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
          const vertical = { middle: 'center', bottom: 'bottom' }[cellCss.verticalAlign];
          if (vertical) cellProps += `<w:vAlign w:val="${vertical}"/>`;

          const cellStyle = blockStyleOf(cell, style);
          if (isHeader && cellStyle.bold === undefined) cellStyle.bold = true;
          const props = blockPropertiesOf(cell);
          const cellInherited = { styleId: 'TableText', align: props.align };
          const content = [];
          this.container(cell, cellStyle, cellInherited, {}, content);
          if (!content.length || !content[content.length - 1].startsWith('<w:p')) content.push('<w:p/>');
          xml += `<w:tc><w:tcPr>${cellProps}</w:tcPr>${content.join('')}</w:tc>`;
        });
        xml += '</w:tr>';
      });
      out.push(`${xml}</w:tbl>`);
    }

    // --- Inline content ---

    text(value, style, preserve, para) {
      if (preserve) {
        value.split(/(\r\n|\n|\t)/).forEach((part) => {
          if (!part) return;
          if (part === '\t') para.xml += `<w:r>${runProperties(style)}<w:tab/></w:r>`;
          else if (part === '\n' || part === '\r\n') para.xml += LINE_BREAK;
          else para.xml += `<w:r>${runProperties(style)}<w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>`;
          para.content = true;
        });
        para.endsSpace = false;
        return;
      }

      // Control characters go first, so the spaces around them still collapse.
      // A no-break space after a space is how HTML writes two spaces in a row.
      let text = value
        .replace(CONTROL_CHARACTERS, '')
        .replace(/[ \t\r\n\f]+/g, ' ')
        .replace(/ (\u00A0+)/g, (match, spaces) => ' '.repeat(spaces.length + 1));
      if (para.endsSpace && text[0] === ' ') text = text.slice(1);
      if (!text) return;
      para.endsSpace = text.endsWith(' ');
      para.xml += `<w:r>${runProperties(style)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
      para.content = true;
    }

    inline(element, inheritedStyle, inherited, para) {
      const tag = element.tagName.toLowerCase();
      if (SKIPPED_TAGS.has(tag)) return;

      if (tag === 'br') {
        para.xml += LINE_BREAK;
        para.content = true;
        para.endsSpace = true;
        return;
      }
      if (tag === 'img') {
        para.xml += this.drawing(element);
        para.content = true;
        para.endsSpace = false;
        return;
      }

      const style = inlineStyleOf(element, inheritedStyle);
      const css = element.style;
      const preserve = inherited.preserve || (css && /^pre/.test(css.whiteSpace || ''));
      const childInherited = preserve === inherited.preserve ? inherited : { ...inherited, preserve };

      let closing = '';
      if (tag === 'a') {
        const name = element.getAttribute('id') || element.getAttribute('name');
        if (name) {
          const id = this.bookmarkId++;
          para.xml += `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(bookmarkName(name))}"/><w:bookmarkEnd w:id="${id}"/>`;
        }
        const href = (element.getAttribute('href') || '').trim();
        if (href && isSafeHref(href)) {
          style.link = true;
          if (href.startsWith('#')) {
            para.xml += `<w:hyperlink w:anchor="${escapeXml(bookmarkName(href.slice(1)))}" w:history="1">`;
          } else {
            const rel = this.addRelationship(`${REL}/hyperlink`, href, true);
            para.xml += `<w:hyperlink r:id="${rel}" w:history="1">`;
          }
          closing = '</w:hyperlink>';
        }
      }

      element.childNodes.forEach((node) => {
        if (node.nodeType === 3) this.text(node.nodeValue, style, preserve, para);
        else if (node.nodeType === 1) this.inline(node, style, childInherited, para);
      });
      para.xml += closing;
    }

    drawing(img) {
      const info = this.images.get(img);
      if (!info) return '';

      let media = this.mediaByKey.get(info.key);
      if (!media) {
        media = { name: `media/image${this.mediaByKey.size + 1}.${info.ext}`, bytes: info.bytes, ext: info.ext };
        media.rel = this.addRelationship(`${REL}/image`, media.name);
        this.mediaByKey.set(info.key, media);
      }

      let width = info.width;
      let height = info.height;
      const maxWidth = this.contentWidth / 15;
      if (width > maxWidth) {
        height = height * maxWidth / width;
        width = maxWidth;
      }
      const cx = Math.max(1, Math.round(width * 9525));
      const cy = Math.max(1, Math.round(height * 9525));
      const id = ++this.drawingId;
      const alt = escapeXml(img.getAttribute('alt') || '');

      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:docPr id="${id}" name="Picture ${id}" descr="${alt}"/><wp:cNvGraphicFramePr/>` +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
        `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${media.rel}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    }
  }

  // --- Images ---

  function dataUriToBytes(src) {
    const match = src.match(/^data:([^;,]*)((?:;[^,]*)?),(.*)$/s);
    if (!match) return null;
    const mime = (match[1] || 'text/plain').toLowerCase();
    if (/;base64/i.test(match[2])) {
      const binary = atob(match[3].replace(/\s/g, ''));
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      return { mime, bytes };
    }
    return { mime, bytes: new TextEncoder().encode(decodeURIComponent(match[3])) };
  }

  /** Waits for load events with a time limit, so one bad image cannot stop a save. */
  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      const timer = setTimeout(() => reject(new Error('This image took too long to load')), 5000);
      image.onload = () => {
        clearTimeout(timer);
        resolve(image);
      };
      image.onerror = () => {
        clearTimeout(timer);
        reject(new Error('This image could not be loaded'));
      };
      image.src = src;
    });
  }

  async function imageToPng(img, width, height) {
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(width));
    canvas.height = Math.max(1, Math.round(height));
    canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) throw new Error('This image could not be converted');
    return new Uint8Array(await blob.arrayBuffer());
  }

  async function loadImages(root) {
    const images = new Map();
    const failed = [];

    for (const img of root.querySelectorAll('img')) {
      const src = img.getAttribute('src') || '';
      try {
        let probe = img;
        if (!img.complete || !img.naturalWidth) {
          probe = await loadImage(src);
        }
        const css = img.style || {};
        const statedWidth = (cssLengthToTwips(css.width) || 0) / 15 || parseFloat(img.getAttribute('width')) || 0;
        const statedHeight = (cssLengthToTwips(css.height) || 0) / 15 || parseFloat(img.getAttribute('height')) || 0;
        const natural = { width: probe.naturalWidth || 96, height: probe.naturalHeight || 96 };
        const width = statedWidth || (statedHeight ? natural.width * statedHeight / natural.height : natural.width);
        const height = statedHeight || (statedWidth ? natural.height * statedWidth / natural.width : natural.height);

        let data;
        if (src.startsWith('data:')) {
          data = dataUriToBytes(src);
        } else {
          const response = await fetch(src);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          data = {
            mime: (response.headers.get('Content-Type') || '').split(';')[0].toLowerCase(),
            bytes: new Uint8Array(await response.arrayBuffer())
          };
        }

        let ext = IMAGE_TYPES[data && data.mime];
        let bytes = data && data.bytes;
        if (!ext) {
          bytes = await imageToPng(probe, natural.width, natural.height);
          ext = 'png';
        }
        images.set(img, { key: src, ext, bytes, width, height });
      } catch (e) {
        failed.push(src);
      }
    }
    return { images, failed };
  }

  // --- Package parts ---

  function buildStyles() {
    const fonts = (font) => `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`;
    const paragraphStyle = (id, name, spec, { before = '', after = '', flags = '' } = {}) => {
      const pPr = `${before}<w:spacing w:before="${spec.before}" w:after="${spec.after}" w:line="${spec.line}" w:lineRule="auto"/>${spec.left ? `<w:ind w:left="${spec.left}"/>` : ''}${after}`;
      const rPr = `${spec.font !== BODY_FONT ? fonts(spec.font) : ''}${spec.bold ? '<w:b/>' : ''}<w:color w:val="${spec.color}"/><w:sz w:val="${spec.sz}"/><w:szCs w:val="${spec.sz}"/>`;
      return `<w:style w:type="paragraph" w:styleId="${id}"${flags}><w:name w:val="${name}"/>` +
        `${flags ? '' : '<w:basedOn w:val="Normal"/>'}<w:qFormat/><w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;
    };
    const headings = [1, 2, 3, 4, 5, 6].map((n) => paragraphStyle(`Heading${n}`, `heading ${n}`, BLOCK_DEFAULTS[`h${n}`], {
      before: '<w:keepNext/>',
      after: `<w:outlineLvl w:val="${n - 1}"/>`
    })).join('');

    return `${XML_HEADER}<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr>${fonts(BODY_FONT)}` +
      `<w:color w:val="${TEXT_COLOR}"/><w:sz w:val="${BLOCK_DEFAULTS.p.sz}"/><w:szCs w:val="${BLOCK_DEFAULTS.p.sz}"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="${BLOCK_DEFAULTS.p.after}" w:line="${BLOCK_DEFAULTS.p.line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      paragraphStyle('Normal', 'Normal', BLOCK_DEFAULTS.p, { flags: ' w:default="1"' }) +
      headings +
      paragraphStyle('Quote', 'Quote', BLOCK_DEFAULTS.blockquote, { before: `<w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="${ACCENT_COLOR}"/></w:pBdr>` }) +
      paragraphStyle('Preformatted', 'Preformatted', BLOCK_DEFAULTS.pre, { before: `<w:shd w:val="clear" w:color="auto" w:fill="${PRE_FILL}"/>` }) +
      paragraphStyle('ListParagraph', 'List Paragraph', BLOCK_DEFAULTS.li) +
      paragraphStyle('TableText', 'Table Text', { ...BLOCK_DEFAULTS.cell, after: 0 }) +
      `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="${ACCENT_COLOR}"/><w:u w:val="single"/></w:rPr></w:style>` +
      '</w:styles>';
  }

  /**
   * Writes the editor's content as a .docx file.
   * Returns { blob, warnings }. warnings lists anything that could not be saved.
   */
  async function htmlToDocx(rootElement, options = {}) {
    const JSZip = options.JSZip || self.JSZip;
    const section = normalizeSection(options.section);
    const { images, failed } = await loadImages(rootElement);
    const writer = new DocxWriter(section, images);

    const out = [];
    writer.container(rootElement, {}, {}, {}, out);
    if (!out.length) out.push('<w:p/>');

    const { pgSz, pgMar } = section;
    const sectPr = `<w:sectPr><w:pgSz w:w="${pgSz.w}" w:h="${pgSz.h}"${pgSz.orient ? ` w:orient="${pgSz.orient}"` : ''}/>` +
      `<w:pgMar w:top="${pgMar.top}" w:right="${pgMar.right}" w:bottom="${pgMar.bottom}" w:left="${pgMar.left}" w:header="${pgMar.header}" w:footer="${pgMar.footer}" w:gutter="${pgMar.gutter}"/></w:sectPr>`;
    const documentXml = `${XML_HEADER}<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"><w:body>${out.join('')}${sectPr}</w:body></w:document>`;

    const zip = new JSZip();
    const hasLists = writer.nums.length > 0;
    if (hasLists) writer.addRelationship(`${REL}/numbering`, 'numbering.xml');

    const extensions = new Set([...writer.mediaByKey.values()].map((media) => media.ext));
    const imageDefaults = [...extensions].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext}"/>`).join('');
    const mainType = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
    zip.file('[Content_Types].xml', `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      `<Default Extension="xml" ContentType="application/xml"/>${imageDefaults}` +
      `<Override PartName="/word/document.xml" ContentType="${mainType}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${mainType}.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${mainType}.settings+xml"/>` +
      `${hasLists ? `<Override PartName="/word/numbering.xml" ContentType="${mainType}.numbering+xml"/>` : ''}</Types>`);
    zip.file('_rels/.rels', `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/document.xml', documentXml);
    zip.file('word/styles.xml', buildStyles());
    zip.file('word/settings.xml', `${XML_HEADER}<w:settings xmlns:w="${NS.w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
    if (hasLists) {
      zip.file('word/numbering.xml', `${XML_HEADER}<w:numbering xmlns:w="${NS.w}">${writer.abstractNums.join('')}${writer.nums.join('')}</w:numbering>`);
    }
    writer.mediaByKey.forEach((media) => zip.file(`word/${media.name}`, media.bytes));
    zip.file('word/_rels/document.xml.rels', `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      writer.relationships.map((rel, i) => `<Relationship Id="rId${i + 1}" Type="${rel.type}" Target="${escapeXml(rel.target)}"${rel.external ? ' TargetMode="External"' : ''}/>`).join('') +
      '</Relationships>');

    const warnings = [];
    if (failed.length) warnings.push(failed.length === 1 ? '1 image could not be saved' : `${failed.length} images could not be saved`);
    const blob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    return { blob, warnings };
  }

  // ==========================================
  // --- READER: .docx to editor HTML ---
  // ==========================================

  function elementsOf(node) {
    const out = [];
    for (let child = node && node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === 1) out.push(child);
    }
    return out;
  }
  const kid = (node, name) => elementsOf(node).find((child) => child.localName === name) || null;
  const kids = (node, name) => elementsOf(node).filter((child) => child.localName === name);

  function attr(node, name) {
    if (!node || !node.attributes) return null;
    for (let i = 0; i < node.attributes.length; i++) {
      if (node.attributes[i].localName === name) return node.attributes[i].value;
    }
    return null;
  }

  function intAttr(node, name) {
    const value = parseInt(attr(node, name), 10);
    return Number.isNaN(value) ? undefined : value;
  }

  /** Word's on and off properties: <w:b/> is on, <w:b w:val="0"/> is off. */
  function toggle(node) {
    if (!node) return undefined;
    const value = attr(node, 'val');
    return !(value === '0' || value === 'false' || value === 'off');
  }

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('This file has damaged XML');
    return doc;
  }

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = `${base}/${target}`.split('/');
    const resolved = [];
    parts.forEach((part) => {
      if (part === '..') resolved.pop();
      else if (part && part !== '.') resolved.push(part);
    });
    return resolved.join('/');
  }

  function mergeDefined(base, extra) {
    const merged = { ...base };
    Object.keys(extra || {}).forEach((key) => {
      if (extra[key] !== undefined) merged[key] = extra[key];
    });
    return merged;
  }

  /** What to draw with when a computer does not have the file's font. */
  function genericFamilyOf(font) {
    if (/courier|consol|mono|menlo|code/i.test(font)) return 'monospace';
    if (/times|georgia|cambria|garamond|palatino|serif|book|minion|baskerville/i.test(font) && !/sans/i.test(font)) return 'serif';
    return 'sans-serif';
  }

  class DocxReader {
    constructor(zip) {
      this.zip = zip;
      this.warnings = new Set();
      this.styles = {};
      this.resolvedStyles = {};
      this.docRun = {};
      this.docPara = {};
      this.theme = {};
      this.abstractNums = {};
      this.nums = {};
      this.counters = {};
      this.rels = {};
      this.images = {};
      this.footnotes = {};
      this.endnotes = {};
      this.noteOrder = [];
      this.tableStyle = null;
    }

    async part(path) {
      const file = this.zip.file(path);
      return file ? parseXml(await file.async('string')) : null;
    }

    async load() {
      let mainPath = 'word/document.xml';
      const packageRels = await this.part('_rels/.rels');
      if (packageRels) {
        const main = kids(packageRels.documentElement, 'Relationship').find((rel) => /\/officeDocument$/.test(attr(rel, 'Type') || ''));
        if (main) mainPath = resolvePath('', attr(main, 'Target'));
      }
      this.document = await this.part(mainPath);
      if (!this.document) throw new Error('This is not a Word document');

      const dir = mainPath.includes('/') ? mainPath.slice(0, mainPath.lastIndexOf('/')) : '';
      const name = mainPath.slice(dir ? dir.length + 1 : 0);
      const relsDoc = await this.part(`${dir}/_rels/${name}.rels`);
      const targets = { styles: 'styles.xml', numbering: 'numbering.xml', footnotes: 'footnotes.xml', endnotes: 'endnotes.xml', theme: 'theme/theme1.xml' };

      if (relsDoc) {
        kids(relsDoc.documentElement, 'Relationship').forEach((rel) => {
          const type = (attr(rel, 'Type') || '').split('/').pop();
          const target = attr(rel, 'Target') || '';
          const external = attr(rel, 'TargetMode') === 'External';
          this.rels[attr(rel, 'Id')] = { type, target, external };
          if (targets[type] && !external) targets[type] = target;
        });
      }

      const sibling = (key) => this.part(resolvePath(dir, targets[key]));
      const commentsRel = Object.values(this.rels).find((rel) => rel.type === 'comments' && !rel.external);
      if (commentsRel) {
        const comments = await this.part(resolvePath(dir, commentsRel.target));
        if (comments && kids(comments.documentElement, 'comment').length) this.warnings.add('comments');
      }
      this.readTheme(await sibling('theme'));
      this.readStyles(await sibling('styles'));
      this.readNumbering(await sibling('numbering'));
      this.footnotes = this.readNotes(await sibling('footnotes'), 'footnote');
      this.endnotes = this.readNotes(await sibling('endnotes'), 'endnote');

      for (const [id, rel] of Object.entries(this.rels)) {
        if (rel.type !== 'image' || rel.external) continue;
        const path = resolvePath(dir, rel.target);
        const file = this.zip.file(path);
        const ext = path.split('.').pop().toLowerCase();
        const mime = IMAGE_MIME_BY_EXT[ext];
        if (file && mime) this.images[id] = `data:${mime};base64,${await file.async('base64')}`;
      }
    }

    readTheme(doc) {
      if (!doc) return;
      const major = doc.getElementsByTagNameNS('*', 'majorFont')[0];
      const minor = doc.getElementsByTagNameNS('*', 'minorFont')[0];
      if (major && kid(major, 'latin')) this.theme.major = attr(kid(major, 'latin'), 'typeface');
      if (minor && kid(minor, 'latin')) this.theme.minor = attr(kid(minor, 'latin'), 'typeface');
    }

    readRun(rPr) {
      if (!rPr) return {};
      const run = {};
      const fonts = kid(rPr, 'rFonts');
      if (fonts) {
        const themed = attr(fonts, 'asciiTheme') || attr(fonts, 'hAnsiTheme');
        const family = themed
          ? (themed.startsWith('major') ? this.theme.major : this.theme.minor)
          : (attr(fonts, 'ascii') || attr(fonts, 'hAnsi'));
        if (family) run.font = family;
      }
      run.bold = toggle(kid(rPr, 'b'));
      run.italic = toggle(kid(rPr, 'i'));
      const strike = toggle(kid(rPr, 'strike'));
      const doubleStrike = toggle(kid(rPr, 'dstrike'));
      if (strike !== undefined || doubleStrike !== undefined) run.strike = Boolean(strike || doubleStrike);
      run.caps = toggle(kid(rPr, 'caps'));
      run.hidden = toggle(kid(rPr, 'vanish'));

      const color = kid(rPr, 'color');
      if (color) {
        const value = attr(color, 'val');
        run.color = value && value !== 'auto' ? cssColorToHex(value) : null;
      }
      const size = intAttr(kid(rPr, 'sz'), 'val');
      if (size) run.sz = size;

      const highlight = kid(rPr, 'highlight');
      const shading = kid(rPr, 'shd');
      if (highlight) {
        const named = attr(highlight, 'val');
        run.bg = named === 'none' ? null : HIGHLIGHT_COLORS[named] || null;
      } else if (shading) {
        const fill = attr(shading, 'fill');
        run.bg = fill && fill !== 'auto' ? cssColorToHex(fill) : null;
      }
      const underline = kid(rPr, 'u');
      if (underline) run.underline = attr(underline, 'val') !== 'none';
      const vertical = kid(rPr, 'vertAlign');
      if (vertical) run.vert = attr(vertical, 'val');
      return run;
    }

    readParagraph(pPr) {
      if (!pPr) return {};
      const para = {};
      const jc = attr(kid(pPr, 'jc'), 'val');
      if (jc) para.align = { start: 'left', end: 'right', both: 'justify', distribute: 'justify' }[jc] || jc;

      const ind = kid(pPr, 'ind');
      if (ind) {
        const left = intAttr(ind, 'left') ?? intAttr(ind, 'start');
        const right = intAttr(ind, 'right') ?? intAttr(ind, 'end');
        const first = intAttr(ind, 'firstLine');
        const hanging = intAttr(ind, 'hanging');
        if (left !== undefined) para.left = left;
        if (right !== undefined) para.right = right;
        if (first !== undefined) para.first = first;
        if (hanging !== undefined) para.first = -hanging;
      }

      const spacing = kid(pPr, 'spacing');
      if (spacing) {
        para.before = intAttr(spacing, 'before');
        para.after = intAttr(spacing, 'after');
        para.line = intAttr(spacing, 'line');
        const rule = attr(spacing, 'lineRule');
        if (rule) para.lineRule = rule;
      }
      if (toggle(kid(pPr, 'contextualSpacing'))) para.contextual = true;
      if (toggle(kid(pPr, 'pageBreakBefore'))) para.pageBreakBefore = true;

      const shading = kid(pPr, 'shd');
      if (shading) {
        const fill = attr(shading, 'fill');
        para.bg = fill && fill !== 'auto' ? cssColorToHex(fill) : null;
      }
      const borders = kid(pPr, 'pBdr');
      if (borders && kid(borders, 'bottom') && attr(kid(borders, 'bottom'), 'val') !== 'none') para.bottomBorder = true;

      const numPr = kid(pPr, 'numPr');
      if (numPr) {
        para.numId = attr(kid(numPr, 'numId'), 'val');
        const level = intAttr(kid(numPr, 'ilvl'), 'val');
        if (level !== undefined) para.ilvl = level;
      }
      const outline = intAttr(kid(pPr, 'outlineLvl'), 'val');
      if (outline !== undefined) para.outline = outline;
      return para;
    }

    readStyles(doc) {
      if (!doc) return;
      const root = doc.documentElement;
      const defaults = kid(root, 'docDefaults');
      if (defaults) {
        this.docRun = this.readRun(kid(kid(defaults, 'rPrDefault'), 'rPr'));
        this.docPara = this.readParagraph(kid(kid(defaults, 'pPrDefault'), 'pPr'));
      }
      kids(root, 'style').forEach((style) => {
        const id = attr(style, 'styleId');
        const tblPr = kid(style, 'tblPr');
        this.styles[id] = {
          id,
          type: attr(style, 'type'),
          name: (attr(kid(style, 'name'), 'val') || '').toLowerCase(),
          basedOn: attr(kid(style, 'basedOn'), 'val'),
          run: this.readRun(kid(style, 'rPr')),
          para: this.readParagraph(kid(style, 'pPr')),
          borders: tblPr ? kid(tblPr, 'tblBorders') : null
        };
        if (attr(style, 'default') === '1' && attr(style, 'type') === 'paragraph') this.defaultParagraphStyle = id;
      });
    }

    readNumbering(doc) {
      if (!doc) return;
      kids(doc.documentElement, 'abstractNum').forEach((abstract) => {
        const levels = {};
        kids(abstract, 'lvl').forEach((level) => {
          levels[intAttr(level, 'ilvl')] = {
            format: attr(kid(level, 'numFmt'), 'val') || 'decimal',
            start: intAttr(kid(level, 'start'), 'val') ?? 1,
            text: attr(kid(level, 'lvlText'), 'val') || ''
          };
        });
        this.abstractNums[attr(abstract, 'abstractNumId')] = levels;
      });
      kids(doc.documentElement, 'num').forEach((num) => {
        const overrides = {};
        kids(num, 'lvlOverride').forEach((override) => {
          const start = intAttr(kid(override, 'startOverride'), 'val');
          if (start !== undefined) overrides[intAttr(override, 'ilvl')] = start;
        });
        this.nums[attr(num, 'numId')] = { abstract: attr(kid(num, 'abstractNumId'), 'val'), overrides };
      });
    }

    readNotes(doc, element) {
      const notes = {};
      if (!doc) return notes;
      kids(doc.documentElement, element).forEach((note) => {
        const type = attr(note, 'type');
        if (type === 'separator' || type === 'continuationSeparator' || type === 'continuationNotice') return;
        notes[attr(note, 'id')] = note;
      });
      return notes;
    }

    /** A style with everything it inherits through basedOn. */
    style(id) {
      if (!id || !this.styles[id]) return null;
      if (this.resolvedStyles[id]) return this.resolvedStyles[id];
      const own = this.styles[id];
      this.resolvedStyles[id] = own;
      const parent = own.basedOn && own.basedOn !== id ? this.style(own.basedOn) : null;
      const resolved = {
        ...own,
        run: mergeDefined(parent ? parent.run : {}, own.run),
        para: mergeDefined(parent ? parent.para : {}, own.para),
        borders: own.borders || (parent && parent.borders) || null,
        ancestry: [own.name, ...(parent ? parent.ancestry || [parent.name] : [])],
        ids: [own.id, ...(parent ? parent.ids || [parent.id] : [])]
      };
      this.resolvedStyles[id] = resolved;
      return resolved;
    }

    /** Which HTML block a paragraph style stands for. */
    blockTagOf(style, para) {
      if (style) {
        const ancestry = style.ancestry || [style.name];
        for (const name of ancestry) {
          const heading = name.match(/^heading ([1-6])$/);
          if (heading) return `h${heading[1]}`;
          if (name === 'title') return 'h1';
          if (name === 'quote' || name === 'intense quote' || name === 'block text' || name === 'blockquote') return 'blockquote';
          if (name === 'preformatted' || name === 'html preformatted' || name === 'source code' || name === 'code') return 'pre';
        }
      }
      if (para.outline !== undefined && para.outline >= 0 && para.outline <= 5) return `h${para.outline + 1}`;
      return 'p';
    }

    // --- Inline content ---

    effectiveRun(rPr, styleRun, inLink) {
      const direct = this.readRun(rPr);
      const charStyleId = rPr ? attr(kid(rPr, 'rStyle'), 'val') : null;
      const charStyle = this.style(charStyleId);
      const isLinkStyle = charStyle && charStyle.name === 'hyperlink';
      let run = styleRun;
      if (charStyle && !(isLinkStyle && inLink)) run = mergeDefined(run, charStyle.run);
      return mergeDefined(run, direct);
    }

    collectInline(node, state) {
      elementsOf(node).forEach((child) => {
        switch (child.localName) {
          case 'r':
            this.collectRun(child, state);
            break;
          case 'hyperlink': {
            const rel = this.rels[attr(child, 'id')];
            const anchor = attr(child, 'anchor');
            const href = rel ? rel.target : anchor ? `#${anchor}` : null;
            const previous = state.link;
            state.link = href && isSafeHref(href) ? href : null;
            this.collectInline(child, state);
            state.link = previous;
            break;
          }
          case 'bookmarkStart': {
            const name = attr(child, 'name');
            if (name && name !== '_GoBack') state.segments.push({ html: `<a id="${escapeHtml(name)}"></a>` });
            break;
          }
          case 'ins':
          case 'moveTo':
          case 'smartTag':
          case 'fldSimple':
          case 'dir':
          case 'bdo':
            this.collectInline(child, state);
            break;
          case 'sdt':
            if (kid(child, 'sdtContent')) this.collectInline(kid(child, 'sdtContent'), state);
            break;
          case 'customXml':
            this.collectInline(child, state);
            break;
          case 'AlternateContent':
            this.collectInline(kid(child, 'Choice') || kid(child, 'Fallback') || child, state);
            break;
          case 'oMath':
          case 'oMathPara': {
            const text = [...child.getElementsByTagNameNS('*', 't')].map((t) => t.textContent).join('');
            if (text) {
              state.segments.push({ text, run: state.styleRun, link: state.link });
              this.warnings.add('equation layout');
            }
            break;
          }
          default:
            break;
        }
      });
    }

    collectRun(runNode, state) {
      const rPr = kid(runNode, 'rPr');
      const run = this.effectiveRun(rPr, state.styleRun, Boolean(state.link));
      if (run.hidden) return;
      const addText = (text) => {
        if (text) state.segments.push({ text, run, link: state.link });
      };

      elementsOf(runNode).forEach((child) => {
        switch (child.localName) {
          case 't':
            addText(child.textContent);
            break;
          case 'tab':
            state.segments.push({ tab: true, run, link: state.link });
            break;
          case 'noBreakHyphen':
            addText('‑');
            break;
          case 'br':
          case 'cr': {
            const type = attr(child, 'type');
            if (type === 'page') state.segments.push({ pageBreak: true });
            else if (type !== 'column') state.segments.push({ html: '<br>', link: state.link });
            break;
          }
          case 'sym': {
            const code = parseInt(attr(child, 'char'), 16);
            if (code && code < 0xE000) addText(String.fromCodePoint(code));
            break;
          }
          case 'drawing':
          case 'pict':
          case 'object':
            this.collectDrawing(child, state);
            break;
          case 'AlternateContent': {
            const choice = kid(child, 'Choice');
            const fallback = kid(child, 'Fallback');
            const usable = choice && (choice.getElementsByTagNameNS('*', 'blip').length || choice.getElementsByTagNameNS('*', 'txbxContent').length)
              ? choice
              : fallback || choice;
            if (usable) this.collectDrawing(usable, state);
            break;
          }
          case 'footnoteReference':
          case 'endnoteReference': {
            const kind = child.localName === 'footnoteReference' ? 'footnote' : 'endnote';
            this.noteOrder.push({ kind, id: attr(child, 'id') });
            state.segments.push({ html: `<sup>[${this.noteOrder.length}]</sup>` });
            this.warnings.add('footnote placement');
            break;
          }
          case 'fldChar': {
            const type = attr(child, 'fldCharType');
            if (type === 'begin') state.field = '';
            else if (type === 'separate' && state.field !== null) {
              const match = state.field.match(/HYPERLINK\s+"([^"]+)"/);
              if (match && isSafeHref(match[1])) {
                state.fieldLink = true;
                state.link = match[1];
              }
              state.field = null;
            } else if (type === 'end') {
              if (state.fieldLink) state.link = null;
              state.fieldLink = false;
              state.field = null;
            }
            break;
          }
          case 'instrText':
            if (state.field !== null && state.field !== undefined) state.field += child.textContent;
            break;
          default:
            break;
        }
      });
    }

    collectDrawing(node, state) {
      const blips = [...node.getElementsByTagNameNS('*', 'blip')];
      const legacy = [...node.getElementsByTagNameNS('*', 'imagedata')];
      const extent = node.getElementsByTagNameNS('*', 'extent')[0];
      const docPr = node.getElementsByTagNameNS('*', 'docPr')[0];
      const size = extent
        ? ` style="width:${Math.round(intAttr(extent, 'cx') / 9525)}px;height:${Math.round(intAttr(extent, 'cy') / 9525)}px"`
        : '';
      const alt = docPr ? attr(docPr, 'descr') || '' : '';

      [...blips.map((blip) => attr(blip, 'embed')), ...legacy.map((image) => attr(image, 'id'))].forEach((id) => {
        if (!id) return;
        if (this.images[id]) {
          state.segments.push({ html: `<img src="${this.images[id]}" alt="${escapeHtml(alt)}"${size}>`, link: state.link });
        } else {
          state.segments.push({ html: '[Image in a format OfficeLite cannot show]' });
          this.warnings.add('images in formats OfficeLite cannot show');
        }
      });

      if (node.getElementsByTagNameNS('*', 'chart').length || node.getElementsByTagNameNS('*', 'relIds').length) {
        this.warnings.add('charts and diagrams');
      }

      const boxes = [...node.getElementsByTagNameNS('*', 'txbxContent')];
      if (boxes.length) {
        state.textBoxes.push(...boxes);
        this.warnings.add('text box placement');
      }
    }

    runCss(run, defaults) {
      const css = [];
      const font = run.font || 'Times New Roman';
      if (font.toLowerCase() !== defaults.font.toLowerCase()) {
        const clean = font.replace(/[^\w\s\-.]/g, '');
        if (clean) css.push(`font-family:'${clean}',${genericFamilyOf(clean)}`);
      }
      const size = run.sz || 20;
      if (size !== defaults.sz) css.push(`font-size:${size / 2}pt`);
      if (run.color && run.color !== defaults.color) css.push(`color:#${run.color.toLowerCase()}`);
      if (run.bg) css.push(`background-color:#${run.bg.toLowerCase()}`);
      if (run.caps) css.push('text-transform:uppercase');
      if (!run.bold && defaults.bold) css.push('font-weight:400');
      return css.join(';');
    }

    wrapRun(html, run, defaults) {
      let out = html;
      const css = this.runCss(run, defaults);
      if (css) out = `<span style="${css}">${out}</span>`;
      if (run.vert === 'superscript') out = `<sup>${out}</sup>`;
      if (run.vert === 'subscript') out = `<sub>${out}</sub>`;
      if (run.strike) out = `<s>${out}</s>`;
      if (run.underline) out = `<u>${out}</u>`;
      if (run.italic) out = `<i>${out}</i>`;
      if (run.bold && !defaults.bold) out = `<b>${out}</b>`;
      return out;
    }

    renderSegments(segments, defaults, isPre) {
      const out = [];
      let linkOpen = null;
      let pending = null;

      const flushText = () => {
        if (!pending) return;
        const inner = isPre
          ? escapeHtml(pending.text)
          : escapeHtml(pending.text).replace(/ {2,}/g, (spaces) => ` ${'&nbsp;'.repeat(spaces.length - 1)}`).replace(/\u00A0/g, '&nbsp;');
        out.push(this.wrapRun(inner, pending.run, defaults));
        pending = null;
      };
      const setLink = (href) => {
        if (href === linkOpen) return;
        flushText();
        if (linkOpen) out.push('</a>');
        if (href) out.push(`<a href="${escapeHtml(href)}">`);
        linkOpen = href;
      };

      segments.forEach((segment) => {
        setLink(segment.link || null);
        if (segment.tab) {
          flushText();
          out.push(this.wrapRun('<span style="white-space:pre">\t</span>', segment.run, defaults));
        } else if (segment.html) {
          flushText();
          out.push(isPre && segment.html === '<br>' ? '\n' : segment.html);
        } else if (segment.text) {
          const key = JSON.stringify(segment.run);
          if (pending && pending.key === key) pending.text += segment.text;
          else {
            flushText();
            pending = { text: segment.text, run: segment.run, key };
          }
        }
      });
      setLink(null);
      flushText();
      return out.join('');
    }

    blockCss(para, defaults, isListItem) {
      const css = [];
      if (para.align && para.align !== 'left') css.push(`text-align:${para.align}`);
      if (!isListItem) {
        const left = para.left || 0;
        if (left !== defaults.left) css.push(`margin-left:${left / 20}pt`);
        if (para.first) css.push(`text-indent:${para.first / 20}pt`);
        if (para.right) css.push(`margin-right:${para.right / 20}pt`);
      }
      const before = para.contextual ? 0 : para.before || 0;
      const after = para.contextual ? 0 : para.after || 0;
      if (before !== defaults.before) css.push(`margin-top:${before / 20}pt`);
      if (after !== defaults.after) css.push(`margin-bottom:${after / 20}pt`);
      const line = para.line === undefined ? 240 : para.line;
      if (para.lineRule === 'exact' || para.lineRule === 'atLeast') css.push(`line-height:${line / 20}pt`);
      else if (line !== defaults.line) css.push(`line-height:${(line / 240).toFixed(2)}`);
      if (para.bg && para.bg !== defaults.bg) css.push(`background-color:#${para.bg.toLowerCase()}`);
      return css.join(';');
    }

    // --- Lists ---

    levelOf(numId, ilvl) {
      const num = this.nums[numId];
      const abstract = num && this.abstractNums[num.abstract];
      return (abstract && abstract[ilvl]) || null;
    }

    /** Advances a list's counter and returns the number this paragraph shows. */
    advanceCounter(numId, ilvl) {
      const counters = this.counters[numId] = this.counters[numId] || {};
      const level = this.levelOf(numId, ilvl);
      const num = this.nums[numId];
      const start = num && num.overrides[ilvl] !== undefined ? num.overrides[ilvl] : (level ? level.start : 1);
      counters[ilvl] = counters[ilvl] === undefined ? start : counters[ilvl] + 1;
      Object.keys(counters).forEach((deeper) => {
        if (Number(deeper) > ilvl) delete counters[deeper];
      });
      return counters[ilvl];
    }

    headingNumber(numId, ilvl) {
      const level = this.levelOf(numId, ilvl);
      if (!level || !level.text || level.format === 'bullet' || level.format === 'none') return '';
      this.advanceCounter(numId, ilvl);
      const counters = this.counters[numId];
      return `${level.text.replace(/%(\d)/g, (match, n) => {
        const value = counters[Number(n) - 1];
        return value === undefined ? '' : String(value);
      })} `;
    }

    // --- Blocks ---

    paragraph(node, inCell, inTh) {
      const pPr = kid(node, 'pPr');
      const directPara = this.readParagraph(pPr);
      const styleId = (pPr && attr(kid(pPr, 'pStyle'), 'val')) || this.defaultParagraphStyle;
      const style = this.style(styleId);
      const table = this.tableStyle;

      let para = mergeDefined(this.docPara, table ? table.para : null);
      para = mergeDefined(para, style ? style.para : null);
      para = mergeDefined(para, directPara);

      let styleRun = mergeDefined(this.docRun, table ? table.run : null);
      styleRun = mergeDefined(styleRun, style ? style.run : null);

      const tag = this.blockTagOf(style, para);

      const level = para.numId && para.numId !== '0' ? this.levelOf(para.numId, para.ilvl || 0) : null;
      const isList = Boolean(level) && tag === 'p' && level.format !== 'none';

      const defaultsKey = isList ? 'li' : tag === 'p' && inCell ? (inTh ? 'th' : 'cell') : tag;
      const defaults = BLOCK_DEFAULTS[defaultsKey];

      const state = { segments: [], styleRun, link: null, field: undefined, fieldLink: false, textBoxes: [] };
      this.collectInline(node, state);

      if (level && tag !== 'p') {
        const prefix = this.headingNumber(para.numId, para.ilvl || 0);
        if (prefix) state.segments.unshift({ text: prefix, run: styleRun });
      }

      // A paragraph can hold page breaks. Each one splits it.
      const pieces = [[]];
      state.segments.forEach((segment) => {
        if (segment.pageBreak) {
          pieces.push([]);
        } else pieces[pieces.length - 1].push(segment);
      });

      return { tag, para, isList, level, defaults, state, pieces };
    }

    renderParagraph(info) {
      const { tag, para, defaults, state, pieces } = info;
      const css = this.blockCss(para, defaults, false);
      const attrs = css ? ` style="${css}"` : '';
      const hasText = (piece) => piece.length > 0;

      const parts = [];
      if (para.pageBreakBefore) parts.push('<hr data-page-break>');

      if (pieces.length === 1 && !hasText(pieces[0]) && para.bottomBorder && tag === 'p') {
        parts.push('<hr>');
      } else {
        pieces.forEach((piece, index) => {
          if (index > 0) parts.push('<hr data-page-break>');
          if (hasText(piece) || pieces.length === 1) {
            const inner = this.renderSegments(piece, defaults, tag === 'pre');
            parts.push(`<${tag}${attrs}>${inner || '<br>'}</${tag}>`);
          }
        });
      }
      state.textBoxes.forEach((box) => parts.push(this.blocks(elementsOf(box), false)));
      return parts.join('');
    }

    /** Converts a run of body or cell children, grouping list paragraphs into lists. */
    blocks(nodes, inCell, inTh) {
      const html = [];
      const stack = [];

      const closeTo = (depth) => {
        while (stack.length > depth) {
          const top = stack.pop();
          html.push(`</li></${top.tag}>`);
        }
      };

      const addListItem = (info) => {
        const numId = info.para.numId;
        const ilvl = info.para.ilvl || 0;
        const format = info.level.format;
        const listTag = format === 'bullet' ? 'ul' : 'ol';

        closeTo(ilvl + 1);
        const number = this.advanceCounter(numId, ilvl);
        const top = stack[stack.length - 1];

        if (stack.length === ilvl + 1 && top.numId === numId && top.tag === listTag) {
          html.push('</li>');
        } else {
          if (stack.length === ilvl + 1) closeTo(ilvl);
          while (stack.length < ilvl + 1) {
            const opening = stack.length === ilvl;
            const levelFormat = opening ? format : 'decimal';
            const tagName = opening ? listTag : 'ul';
            const attrs = [];
            if (tagName === 'ol') {
              const listStyle = LIST_STYLE_TYPES[levelFormat];
              if (listStyle) attrs.push(`style="list-style-type:${listStyle}"`);
              if (opening && number !== 1) attrs.push(`start="${number}"`);
            }
            html.push(`<${tagName}${attrs.length ? ` ${attrs.join(' ')}` : ''}>`);
            stack.push({ tag: tagName, numId });
            if (!opening) html.push('<li>');
          }
        }

        const css = this.blockCss(info.para, info.defaults, true);
        const flat = info.pieces.reduce((all, piece) => all.concat(piece), []);
        html.push(`<li${css ? ` style="${css}"` : ''}>${this.renderSegments(flat, info.defaults, false) || '<br>'}`);
        info.state.textBoxes.forEach((box) => html.push(this.blocks(elementsOf(box), false)));
      };

      const visit = (node) => {
        switch (node.localName) {
          case 'p': {
            const info = this.paragraph(node, inCell, inTh);
            if (info.isList) {
              addListItem(info);
            } else {
              closeTo(0);
              html.push(this.renderParagraph(info));
            }
            break;
          }
          case 'tbl':
            closeTo(0);
            html.push(this.table(node));
            break;
          case 'sdt':
            if (kid(node, 'sdtContent')) elementsOf(kid(node, 'sdtContent')).forEach(visit);
            break;
          case 'customXml':
          case 'ins':
          case 'moveTo':
            elementsOf(node).forEach(visit);
            break;
          case 'AlternateContent':
            elementsOf(kid(node, 'Choice') || kid(node, 'Fallback') || node).forEach(visit);
            break;
          default:
            break;
        }
      };

      nodes.forEach(visit);
      closeTo(0);
      return html.join('');
    }

    table(node) {
      const tblPr = kid(node, 'tblPr');
      const previousStyle = this.tableStyle;
      const styleId = tblPr ? attr(kid(tblPr, 'tblStyle'), 'val') : null;
      const tableStyle = this.style(styleId);
      this.tableStyle = tableStyle;

      const directBorders = tblPr ? kid(tblPr, 'tblBorders') : null;
      const borders = directBorders || (tableStyle && tableStyle.borders);
      const hasBorders = Boolean(borders) && elementsOf(borders).some((side) => {
        const value = attr(side, 'val');
        return value && value !== 'nil' && value !== 'none';
      });

      const rows = [];
      const addRows = (parent) => elementsOf(parent).forEach((child) => {
        if (child.localName === 'tr') rows.push(child);
        else if (child.localName === 'sdt' && kid(child, 'sdtContent')) addRows(kid(child, 'sdtContent'));
        else if (child.localName === 'customXml') addRows(child);
      });
      addRows(node);

      const grid = rows.map((row) => {
        const cells = [];
        const trPr = kid(row, 'trPr');
        let column = (trPr && intAttr(kid(trPr, 'gridBefore'), 'val')) || 0;
        const addCells = (parent) => elementsOf(parent).forEach((child) => {
          if (child.localName === 'tc') {
            const tcPr = kid(child, 'tcPr');
            const span = (tcPr && intAttr(kid(tcPr, 'gridSpan'), 'val')) || 1;
            const merge = tcPr ? kid(tcPr, 'vMerge') : null;
            cells.push({
              tc: child,
              tcPr,
              column,
              span,
              continues: Boolean(merge) && attr(merge, 'val') !== 'restart',
              restarts: Boolean(merge) && attr(merge, 'val') === 'restart'
            });
            column += span;
          } else if (child.localName === 'sdt' && kid(child, 'sdtContent')) {
            addCells(kid(child, 'sdtContent'));
          } else if (child.localName === 'customXml') {
            addCells(child);
          }
        });
        addCells(row);
        return { cells, isHeader: Boolean(trPr && kid(trPr, 'tblHeader')) };
      });

      let tableCss = '';
      const tblW = tblPr ? kid(tblPr, 'tblW') : null;
      if (tblW) {
        const type = attr(tblW, 'type');
        const width = intAttr(tblW, 'w');
        if (type === 'dxa' && width) tableCss = `width:${Math.round(width / 15)}px`;
        else if (type === 'pct' && width && width !== 5000) tableCss = `width:${width / 50}%`;
      }

      const out = [`<table${tableCss ? ` style="${tableCss}"` : ''}>`];
      let headerOpen = false;
      let bodyOpen = false;

      grid.forEach((row, r) => {
        const isHeaderRow = row.isHeader && !bodyOpen;
        if (isHeaderRow && !headerOpen) {
          out.push('<thead>');
          headerOpen = true;
        } else if (!isHeaderRow && !bodyOpen) {
          if (headerOpen) out.push('</thead>');
          out.push('<tbody>');
          bodyOpen = true;
        }

        out.push('<tr>');
        row.cells.forEach((cell) => {
          if (cell.continues) return;
          let rowSpan = 1;
          if (cell.restarts) {
            for (let below = r + 1; below < grid.length; below++) {
              const match = grid[below].cells.find((other) => other.column === cell.column);
              if (!match || !match.continues) break;
              rowSpan += 1;
            }
          }

          const css = [];
          const tcPr = cell.tcPr;
          const tcW = tcPr ? kid(tcPr, 'tcW') : null;
          if (tcW && attr(tcW, 'type') === 'dxa' && intAttr(tcW, 'w')) css.push(`width:${Math.round(intAttr(tcW, 'w') / 15)}px`);

          const cellBorders = tcPr ? kid(tcPr, 'tcBorders') : null;
          const cellNoBorders = cellBorders && elementsOf(cellBorders).length > 0 && elementsOf(cellBorders).every((side) => {
            const value = attr(side, 'val');
            return value === 'nil' || value === 'none';
          });
          if (!hasBorders && !(cellBorders && !cellNoBorders)) css.push('border:none');
          else if (cellNoBorders) css.push('border:none');

          const shading = tcPr ? kid(tcPr, 'shd') : null;
          const fill = shading ? attr(shading, 'fill') : null;
          const hex = fill && fill !== 'auto' ? cssColorToHex(fill) : null;
          if (hex) css.push(`background-color:#${hex.toLowerCase()}`);
          else if (isHeaderRow) css.push('background-color:transparent');

          const vertical = tcPr ? attr(kid(tcPr, 'vAlign'), 'val') : null;
          if (vertical === 'center') css.push('vertical-align:middle');
          if (vertical === 'bottom') css.push('vertical-align:bottom');

          const tag = isHeaderRow ? 'th' : 'td';
          const attrs = [];
          if (cell.span > 1) attrs.push(`colspan="${cell.span}"`);
          if (rowSpan > 1) attrs.push(`rowspan="${rowSpan}"`);
          if (css.length) attrs.push(`style="${css.join(';')}"`);
          // A lone paragraph with no spacing draws the same as text in the cell.
          const content = this.blocks(elementsOf(cell.tc), true, isHeaderRow)
            .replace(/^<p style="margin-bottom:0pt">((?:(?!<\/?p[ >]).)*)<\/p>$/s, '$1');
          out.push(`<${tag}${attrs.length ? ` ${attrs.join(' ')}` : ''}>${content}</${tag}>`);
        });
        out.push('</tr>');
      });

      if (bodyOpen) out.push('</tbody>');
      else if (headerOpen) out.push('</thead>');
      out.push('</table>');

      this.tableStyle = previousStyle;
      return out.join('');
    }

    readSection(body) {
      const sectPr = kid(body, 'sectPr');
      if (!sectPr) return null;
      const size = kid(sectPr, 'pgSz');
      const margin = kid(sectPr, 'pgMar');
      const usesPart = (name) => this.document.getElementsByTagNameNS('*', name).length > 0;
      if (usesPart('headerReference') || usesPart('footerReference')) this.warnings.add('headers and footers');
      if (kid(sectPr, 'cols') && intAttr(kid(sectPr, 'cols'), 'num') > 1) this.warnings.add('multiple columns');

      const section = {};
      if (size) section.pgSz = { w: intAttr(size, 'w'), h: intAttr(size, 'h'), orient: attr(size, 'orient') };
      if (margin) {
        section.pgMar = {};
        Object.keys(DEFAULT_SECTION.pgMar).forEach((key) => { section.pgMar[key] = intAttr(margin, key); });
      }
      return normalizeSection(section);
    }

    convert() {
      const body = kid(this.document.documentElement, 'body');
      if (!body) throw new Error('This Word document has no content');
      let html = this.blocks(elementsOf(body), false);

      // Each note is read once. Notes that mention other notes join the end of the list.
      if (this.noteOrder.length) {
        const notes = [];
        for (let i = 0; i < this.noteOrder.length; i++) {
          const { kind, id } = this.noteOrder[i];
          const note = (kind === 'footnote' ? this.footnotes : this.endnotes)[id];
          if (!note) continue;
          const content = this.blocks(elementsOf(note), false);
          notes.push(content.replace(/^<p([^>]*)>/, `<p$1><sup>[${i + 1}]</sup> `));
        }
        if (notes.length) html += `<hr>${notes.join('')}`;
      }
      return { html, section: this.readSection(body), warnings: [...this.warnings] };
    }
  }

  /**
   * Reads a .docx file into editor HTML.
   * Returns { html, section, warnings }. warnings names what the editor cannot keep.
   */
  async function docxToHtml(arrayBuffer, options = {}) {
    const JSZip = options.JSZip || self.JSZip;
    const zip = await JSZip.loadAsync(arrayBuffer);
    const reader = new DocxReader(zip);
    await reader.load();
    return reader.convert();
  }

  return {
    htmlToDocx,
    docxToHtml,
    FONT_SIZE_PT,
    _helpers: { escapeXml, escapeHtml, cssColorToHex, cssLengthToTwips, cssFontSizeToHalfPoints, firstFontFamily, bookmarkName, isSafeHref, normalizeSection }
  };
});
