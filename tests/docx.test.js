'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { _helpers: h, FONT_SIZE_PT } = require('../docx.js');

test('colors become capital hex, and transparent becomes nothing', () => {
  assert.strictEqual(h.cssColorToHex('#cc3300'), 'CC3300');
  assert.strictEqual(h.cssColorToHex('#c30'), 'CC3300');
  assert.strictEqual(h.cssColorToHex('rgb(204, 51, 0)'), 'CC3300');
  assert.strictEqual(h.cssColorToHex('rgba(204, 51, 0, 0)'), null);
  assert.strictEqual(h.cssColorToHex('transparent'), null);
  assert.strictEqual(h.cssColorToHex(''), null);
});

test('lengths become twips', () => {
  assert.strictEqual(h.cssLengthToTwips('12pt'), 240);
  assert.strictEqual(h.cssLengthToTwips('40px'), 600);
  assert.strictEqual(h.cssLengthToTwips('1in'), 1440);
  assert.strictEqual(h.cssLengthToTwips('-18pt'), -360);
  assert.strictEqual(h.cssLengthToTwips('0'), 0);
  assert.strictEqual(h.cssLengthToTwips('auto'), null);
  assert.strictEqual(h.cssLengthToTwips('50%'), null);
});

test('font sizes become half points', () => {
  assert.strictEqual(h.cssFontSizeToHalfPoints('14pt'), 28);
  assert.strictEqual(h.cssFontSizeToHalfPoints('16px'), 24);
  assert.strictEqual(h.cssFontSizeToHalfPoints('x-large'), 36);
  assert.strictEqual(h.cssFontSizeToHalfPoints('150%', 24), 36);
  assert.strictEqual(h.cssFontSizeToHalfPoints('2em', 24), 48);
});

test('the font menu sizes are the sizes it names', () => {
  assert.deepStrictEqual(Object.values(FONT_SIZE_PT), [10, 11, 12, 14, 18, 24, 32]);
});

test('font lists give their first real family', () => {
  assert.strictEqual(h.firstFontFamily("'Times New Roman', serif"), 'Times New Roman');
  assert.strictEqual(h.firstFontFamily('Georgia, serif'), 'Georgia');
  assert.strictEqual(h.firstFontFamily('sans-serif'), null);
});

test('text that would break a Word file is cleaned', () => {
  assert.strictEqual(h.escapeXml('a<b>&"\'c'), 'a&lt;b&gt;&amp;&quot;&apos;c');
  assert.strictEqual(h.escapeXml('bad\u0001\u0008text'), 'badtext');
  assert.strictEqual(h.escapeXml('line\ttab'), 'line\ttab');
});

test('only safe links are kept', () => {
  assert.ok(h.isSafeHref('https://example.com'));
  assert.ok(h.isSafeHref('mailto:a@example.com'));
  assert.ok(h.isSafeHref('#section'));
  assert.ok(!h.isSafeHref('javascript:alert(1)'));
  assert.ok(!h.isSafeHref('file:///etc/passwd'));
});

test('page setup falls back to safe numbers', () => {
  const section = h.normalizeSection({ pgSz: { w: 11906, h: 16838 }, pgMar: { top: 'x', left: 720 } });
  assert.strictEqual(section.pgSz.w, 11906);
  assert.strictEqual(section.pgMar.left, 720);
  assert.strictEqual(section.pgMar.top, 1440);
});
