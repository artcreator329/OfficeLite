const test = require('node:test');
const assert = require('node:assert');
const F = require('../formula.js');

function workbook(sheets) {
  const objects = {};
  for (const [name, cells] of Object.entries(sheets)) {
    objects[name] = { cells: {} };
    for (const [addr, value] of Object.entries(cells)) {
      objects[name].cells[addr] = typeof value === 'string' && value.startsWith('=') ? { f: value } : { v: value };
    }
  }
  const lookup = (name) => objects[Object.keys(objects).find((k) => k.toLowerCase() === name.toLowerCase())] || null;
  return { objects, engine: new F.Engine(lookup) };
}

function calc(formula, cells = {}) {
  const { objects, engine } = workbook({ Sheet1: cells });
  return engine.evaluate(formula, objects.Sheet1);
}

test('the same cell can be referenced twice', () => {
  assert.strictEqual(calc('=A1+A1', { A1: 5 }), 10);
  assert.strictEqual(calc('=SUM(A1:A2,A1)', { A1: 5, A2: 1 }), 11);
});

test('two formulas can share a source cell', () => {
  assert.strictEqual(calc('=SUM(B1:B2)', { A1: 5, B1: '=A1', B2: '=A1' }), 10);
});

test('arithmetic and precedence', () => {
  assert.strictEqual(calc('=2+3*4'), 14);
  assert.strictEqual(calc('=(2+3)*4'), 20);
  assert.strictEqual(calc('=2^3^2'), 64);
  assert.strictEqual(calc('=-2^2'), 4);
  assert.strictEqual(calc('=50%'), 0.5);
  assert.strictEqual(calc('=SUM(A1:A2)*2', { A1: 5, A2: 7 }), 24);
  assert.strictEqual(calc('=$A$1+A$1+$A1', { A1: 2 }), 6);
  assert.strictEqual(calc('= a1 + 1', { A1: 2 }), 3);
});

test('errors', () => {
  assert.strictEqual(calc('=1/0'), '#DIV/0!');
  assert.strictEqual(calc('=NOSUCH(1)'), '#NAME?');
  assert.strictEqual(calc('=A1+1', { A1: 'text' }), '#VALUE!');
  assert.strictEqual(calc('=1+'), '#ERROR!');
  assert.strictEqual(calc('=(1'), '#ERROR!');
  assert.strictEqual(calc('=A1+1', { A1: '=1/0' }), '#DIV/0!');
  assert.strictEqual(calc('=#REF!+1'), '#REF!');
  assert.strictEqual(calc('=NoSheet!A1'), '#REF!');
  assert.strictEqual(calc('=SQRT(-1)'), '#NUM!');
});

test('circular references are caught', () => {
  const { objects, engine } = workbook({ S: { A1: '=B1', B1: '=A1', C1: '=C1+1', D1: '=A1+1' } });
  assert.strictEqual(engine.value(objects.S, 'A1'), '#CIRC!');
  assert.strictEqual(engine.value(objects.S, 'C1'), '#CIRC!');
  assert.strictEqual(engine.value(objects.S, 'D1'), '#CIRC!');
});

test('aggregates ignore text and blanks in ranges', () => {
  const cells = { A1: 1, A2: 'x', A3: 3, A5: 5 };
  assert.strictEqual(calc('=SUM(A1:A5)', cells), 9);
  assert.strictEqual(calc('=AVERAGE(A1:A5)', cells), 3);
  assert.strictEqual(calc('=COUNT(A1:A5)', cells), 3);
  assert.strictEqual(calc('=COUNTA(A1:A5)', cells), 4);
  assert.strictEqual(calc('=COUNTBLANK(A1:A5)', cells), 1);
  assert.strictEqual(calc('=MIN(A1:A5)', cells), 1);
  assert.strictEqual(calc('=MAX(A1:A5,10)', cells), 10);
  assert.strictEqual(calc('=MEDIAN(A1:A5)', cells), 3);
  assert.strictEqual(calc('=PRODUCT(A1,A3)', cells), 3);
  assert.strictEqual(calc('=SUM(A2)', cells), 0);
  assert.strictEqual(calc('=AVERAGE(A2)', cells), '#DIV/0!');
  assert.strictEqual(calc('=SUM(B1:B2)', { B1: '=1/0' }), '#DIV/0!');
});

test('logic takes only the branch it needs', () => {
  assert.strictEqual(calc('=IF(A1>3,"big","small")', { A1: 5 }), 'big');
  assert.strictEqual(calc('=if(A1>3,"big","small")', { A1: 1 }), 'small');
  assert.strictEqual(calc('=IF(TRUE,1,1/0)'), 1);
  assert.strictEqual(calc('=IF(FALSE,1)'), false);
  assert.strictEqual(calc('=IFERROR(1/0,"none")'), 'none');
  assert.strictEqual(calc('=IFERROR(2,"none")'), 2);
  assert.strictEqual(calc('=AND(1,TRUE)'), true);
  assert.strictEqual(calc('=OR(0,FALSE)'), false);
  assert.strictEqual(calc('=NOT(0)'), true);
  assert.strictEqual(calc('=A1="yes"', { A1: 'YES' }), true);
  assert.strictEqual(calc('=A1=0', {}), true);
  assert.strictEqual(calc('=2<>3'), true);
});

test('strings keep their case and handle quotes', () => {
  assert.strictEqual(calc('="Mixed"&" Case"'), 'Mixed Case');
  assert.strictEqual(calc('="say ""hi"""'), 'say "hi"');
  assert.strictEqual(calc('=UPPER("a")&LOWER("B")&LEN("abc")'), 'Ab3');
  assert.strictEqual(calc('=LEFT("hello",2)&RIGHT("hello",2)&MID("hello",2,3)'), 'heloell');
  assert.strictEqual(calc('=TRIM("  a   b ")'), 'a b');
  assert.strictEqual(calc('=CONCAT(A1:A2,"!")', { A1: 'x', A2: 1 }), 'x1!');
});

test('rounding', () => {
  assert.strictEqual(calc('=ROUND(2.5,0)'), 3);
  assert.strictEqual(calc('=ROUND(-2.5,0)'), -3);
  assert.strictEqual(calc('=ROUND(1.005,2)'), 1.01);
  assert.strictEqual(calc('=ROUND(1234,-2)'), 1200);
  assert.strictEqual(calc('=ROUNDUP(1.01,1)'), 1.1);
  assert.strictEqual(calc('=ROUNDDOWN(1.99,0)'), 1);
  assert.strictEqual(calc('=INT(-1.5)'), -2);
  assert.strictEqual(calc('=MOD(-1,3)'), 2);
  assert.strictEqual(calc('=ABS(-2)'), 2);
});

test('conditional and lookup functions', () => {
  const cells = { A1: 'a', B1: 1, A2: 'b', B2: 2, A3: 'a', B3: 3 };
  assert.strictEqual(calc('=SUMIF(A1:A3,"a",B1:B3)', cells), 4);
  assert.strictEqual(calc('=SUMIF(B1:B3,">1")', cells), 5);
  assert.strictEqual(calc('=COUNTIF(A1:A3,"a")', cells), 2);
  assert.strictEqual(calc('=COUNTIF(A1:A3,"<>a")', cells), 1);
  assert.strictEqual(calc('=COUNTIF(B1:B3,2)', cells), 1);
  assert.strictEqual(calc('=COUNTIF(A1:A3,"?")', cells), 3);
  assert.strictEqual(calc('=AVERAGEIF(A1:A3,"a",B1:B3)', cells), 2);
  assert.strictEqual(calc('=VLOOKUP("b",A1:B3,2,FALSE)', cells), 2);
  assert.strictEqual(calc('=VLOOKUP("z",A1:B3,2,FALSE)', cells), '#N/A');
  assert.strictEqual(calc('=VLOOKUP("b",A1:B3,9,FALSE)', cells), '#REF!');
  assert.strictEqual(calc('=MATCH("b",A1:A3,0)', cells), 2);
  assert.strictEqual(calc('=INDEX(A1:B3,3,2)', cells), 3);
});

test('dates', () => {
  assert.strictEqual(calc('=DATE(2026,1,31)'), 46053);
  assert.strictEqual(calc('=YEAR(46053)&"-"&MONTH(46053)&"-"&DAY(46053)'), '2026-1-31');
  assert.strictEqual(F.formatValue(46053, 'date'), '2026-01-31');
});

test('cross sheet references', () => {
  const { objects, engine } = workbook({ One: { A1: 2 }, 'My Sheet': { A1: 3, B1: '=One!A1+1' } });
  assert.strictEqual(engine.evaluate('=One!A1+1', objects['My Sheet']), 3);
  assert.strictEqual(engine.evaluate("='My Sheet'!A1+SUM('My Sheet'!A1:A1)", objects.One), 6);
  assert.strictEqual(engine.evaluate('=one!A1', objects.One), 2);
  assert.strictEqual(engine.value(objects['My Sheet'], 'B1'), 3);
});

test('a value saved by Excel is kept for functions this engine lacks', () => {
  const { objects, engine } = workbook({ S: {} });
  objects.S.cells.A1 = { f: '=XIRR(B1:B2,C1:C2)', v: 0.125 };
  objects.S.cells.A2 = { f: '=XIRR(B1:B2,C1:C2)' };
  assert.strictEqual(engine.value(objects.S, 'A1'), 0.125);
  assert.strictEqual(engine.value(objects.S, 'A2'), '#NAME?');
});

test('reset picks up changed values', () => {
  const { objects, engine } = workbook({ S: { A1: 1, B1: '=A1*2' } });
  assert.strictEqual(engine.value(objects.S, 'B1'), 2);
  objects.S.cells.A1.v = 5;
  engine.reset();
  assert.strictEqual(engine.value(objects.S, 'B1'), 10);
});

test('copying a formula shifts relative references only', () => {
  assert.strictEqual(F.shiftRelative('=A1+$B$2+C$3+$D4', 1, 1), '=B2+$B$2+D$3+$D5');
  assert.strictEqual(F.shiftRelative('=SUM(A1:A3)', 0, 2), '=SUM(C1:C3)');
  assert.strictEqual(F.shiftRelative('=A1', -1, 0), '=#REF!');
  assert.strictEqual(F.shiftRelative('="A1"&A1', 1, 0), '="A1"&A2');
  assert.strictEqual(F.shiftRelative('=LOG10(A1)+TOTAL1', 1, 0), '=LOG10(A2)+TOTAL1');
  assert.strictEqual(F.shiftRelative("='My Sheet'!A1+One!B2", 1, 0), "='My Sheet'!A2+One!B3");
});

test('inserting rows and columns moves references', () => {
  const adjust = (f, axis, at, count) => F.adjustReferences(f, 'S', 'S', axis, at, count);
  assert.strictEqual(adjust('=SUM(G1:G2)', 'row', 0, 1), '=SUM(G2:G3)');
  assert.strictEqual(adjust('=SUM(G1:G3)', 'row', 1, 2), '=SUM(G1:G5)');
  assert.strictEqual(adjust('=SUM(G1:G3)', 'row', 3, 1), '=SUM(G1:G3)');
  assert.strictEqual(adjust('=$A$5', 'row', 0, 1), '=$A$6');
  assert.strictEqual(adjust('=B1+C1', 'col', 2, 1), '=B1+D1');
  assert.strictEqual(F.adjustReferences('=Other!A1+A1', 'S', 'S', 'row', 0, 1), '=Other!A1+A2');
  assert.strictEqual(F.adjustReferences('=S!A1+A1', 'Other', 'S', 'row', 0, 1), '=S!A2+A1');
});

test('deleting rows and columns shrinks ranges or leaves #REF!', () => {
  const adjust = (f, axis, at, count) => F.adjustReferences(f, 'S', 'S', axis, at, count);
  assert.strictEqual(adjust('=A5', 'row', 1, -2), '=A3');
  assert.strictEqual(adjust('=A2', 'row', 1, -1), '=#REF!');
  assert.strictEqual(adjust('=SUM(A1:A5)', 'row', 1, -2), '=SUM(A1:A3)');
  assert.strictEqual(adjust('=SUM(A2:A5)', 'row', 0, -2), '=SUM(A1:A3)');
  assert.strictEqual(adjust('=SUM(A1:A3)', 'row', 2, -3), '=SUM(A1:A2)');
  assert.strictEqual(adjust('=SUM(A2:A3)', 'row', 0, -4), '=SUM(#REF!)');
});

test('renaming and deleting sheets updates references', () => {
  assert.strictEqual(F.renameSheetReferences("=One!A1+'One'!B2+A1", 'One', 'Two Words'), "='Two Words'!A1+'Two Words'!B2+A1");
  assert.strictEqual(F.renameSheetReferences('="One!A1"', 'One', 'Two'), '="One!A1"');
  assert.strictEqual(F.removeSheetReferences('=One!A1+A1', 'One'), '=One!#REF!+A1');
  assert.strictEqual(calc('=One!#REF!+1', {}), '#REF!');
});

test('parsing what people type', () => {
  assert.deepStrictEqual(F.parseInput('42'), { value: 42 });
  assert.deepStrictEqual(F.parseInput('-3.5'), { value: -3.5 });
  assert.deepStrictEqual(F.parseInput('1,234.50'), { value: 1234.5 });
  assert.deepStrictEqual(F.parseInput('$1,200'), { value: 1200, format: 'currency' });
  assert.deepStrictEqual(F.parseInput('12.5%'), { value: 0.125, format: 'percent' });
  assert.deepStrictEqual(F.parseInput('12 Main St'), { value: '12 Main St' });
  assert.deepStrictEqual(F.parseInput('1,2'), { value: '1,2' });
  assert.deepStrictEqual(F.parseInput('$'), { value: '$' });
  assert.deepStrictEqual(F.parseInput('2026-01-31'), { value: '2026-01-31' });
});

test('formatting values', () => {
  assert.strictEqual(F.formatValue(1234.5678, 'general'), '1234.5678');
  assert.strictEqual(F.formatValue(0.1 + 0.2, 'general'), '0.3');
  assert.strictEqual(F.formatValue(1234.5, 'currency'), '$1,234.50');
  assert.strictEqual(F.formatValue(-5, 'currency'), '-$5.00');
  assert.strictEqual(F.formatValue(0.125, 'percent'), '12.5%');
  assert.strictEqual(F.formatValue(1234.567, 'decimal'), '1,234.57');
  assert.strictEqual(F.formatValue(1234.5, 'integer'), '1,235');
  assert.strictEqual(F.formatValue('12 Main St', 'currency'), '12 Main St');
  assert.strictEqual(F.formatValue(true, 'general'), 'TRUE');
  assert.strictEqual(F.formatValue('', 'currency'), '');
  assert.strictEqual(F.formatValue(5, 'text'), '5');
});
