/**
 * OfficeLite formula engine: parses and evaluates spreadsheet formulas, rewrites cell
 * references when rows, columns, or sheets change, and formats cell values.
 * It has no DOM dependencies, so tests/formula.test.js can run it under Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OfficeLiteFormula = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // --- Coordinates ---
  function colToIndex(letters) {
    let index = 0;
    for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
    return index - 1;
  }

  function indexToCol(index) {
    let letters = '';
    for (let n = index; n >= 0; n = Math.floor(n / 26) - 1) {
      letters = String.fromCharCode((n % 26) + 65) + letters;
    }
    return letters;
  }

  function toAddress(r, c) {
    return `${indexToCol(c)}${r + 1}`;
  }

  // --- Values ---
  class FormulaError {
    constructor(code) {
      this.code = code;
    }
  }

  const ERR = {
    VALUE: new FormulaError('#VALUE!'),
    DIV0: new FormulaError('#DIV/0!'),
    NAME: new FormulaError('#NAME?'),
    REF: new FormulaError('#REF!'),
    NA: new FormulaError('#N/A'),
    NUM: new FormulaError('#NUM!'),
    CIRC: new FormulaError('#CIRC!'),
    PARSE: new FormulaError('#ERROR!')
  };

  const ERROR_LITERALS = {
    '#VALUE!': ERR.VALUE, '#DIV/0!': ERR.DIV0, '#NAME?': ERR.NAME, '#REF!': ERR.REF,
    '#N/A': ERR.NA, '#NUM!': ERR.NUM, '#CIRC!': ERR.CIRC
  };

  class RangeValue {
    constructor(engine, sheet, r1, c1, r2, c2) {
      Object.assign(this, { engine, sheet, r1, c1, r2, c2 });
    }

    get(r, c) {
      return this.engine.cellValue(this.sheet, toAddress(r, c));
    }

    get rows() {
      return this.r2 - this.r1 + 1;
    }

    get cols() {
      return this.c2 - this.c1 + 1;
    }

    forEach(callback) {
      for (let r = this.r1; r <= this.r2; r++) {
        for (let c = this.c1; c <= this.c2; c++) callback(this.get(r, c), r, c);
      }
    }
  }

  function scalar(value) {
    if (value instanceof RangeValue) {
      if (value.rows === 1 && value.cols === 1) return scalar(value.get(value.r1, value.c1));
      throw ERR.VALUE;
    }
    if (value instanceof FormulaError) throw value;
    return value;
  }

  function toNumber(value) {
    const v = scalar(value);
    if (v === null || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const n = Number(v);
    if (String(v).trim() === '' || Number.isNaN(n)) throw ERR.VALUE;
    return n;
  }

  function toText(value) {
    const v = scalar(value);
    if (v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return String(Number(v.toPrecision(15)));
    return String(v);
  }

  function toBool(value) {
    const v = scalar(value);
    if (typeof v === 'boolean') return v;
    if (typeof v === 'string') {
      if (v.toUpperCase() === 'TRUE') return true;
      if (v.toUpperCase() === 'FALSE') return false;
      throw ERR.VALUE;
    }
    return toNumber(v) !== 0;
  }

  // Ranges and single cell references contribute only their numbers. Values typed directly must convert.
  function numbersIn(args) {
    const out = [];
    args.forEach((arg) => {
      if (arg instanceof RangeValue) {
        arg.forEach((v) => {
          if (v instanceof FormulaError) throw v;
          if (typeof v === 'number') out.push(v);
        });
      } else if (arg !== null) {
        out.push(toNumber(arg));
      }
    });
    return out;
  }

  // --- Tokenizer ---
  const TOKEN = new RegExp([
    '(\\s+)',
    '(\\d+\\.?\\d*(?:[eE][+-]?\\d+)?|\\.\\d+(?:[eE][+-]?\\d+)?)',
    '("(?:[^"]|"")*")',
    "('(?:[^']|'')+'!|[A-Za-z_][A-Za-z0-9_.]*!)",
    '(\\$?[A-Za-z]{1,3}\\$?[0-9]+(?![A-Za-z0-9_(.]))',
    '(#[A-Z0-9/]+[!?]|#N/A)',
    '([A-Za-z_][A-Za-z0-9_.]*)',
    '(<>|<=|>=|[-+*/^&=<>%(),;:])'
  ].join('|'), 'y');

  function tokenize(text) {
    const tokens = [];
    let pos = 0;
    while (pos < text.length) {
      TOKEN.lastIndex = pos;
      const m = TOKEN.exec(text);
      if (!m) throw ERR.PARSE;
      pos = TOKEN.lastIndex;
      if (m[1]) continue;
      if (m[2]) tokens.push({ type: 'num', value: Number(m[2]) });
      else if (m[3]) tokens.push({ type: 'str', value: m[3].slice(1, -1).replace(/""/g, '"') });
      else if (m[4]) tokens.push({ type: 'sheet', value: unquoteSheet(m[4].slice(0, -1)) });
      else if (m[5]) tokens.push({ type: 'ref', value: m[5] });
      else if (m[6]) tokens.push({ type: 'err', value: m[6] });
      else if (m[7]) tokens.push({ type: 'name', value: m[7] });
      else tokens.push({ type: 'op', value: m[8] });
    }
    return tokens;
  }

  function unquoteSheet(raw) {
    return raw.startsWith("'") ? raw.slice(1, -1).replace(/''/g, "'") : raw;
  }

  function parseRefParts(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/.exec(ref);
    return { c: colToIndex(m[1]), r: Number(m[2]) - 1 };
  }

  // --- Parser: builds a tree so IF and IFERROR can skip the branches they do not take ---
  function parse(text) {
    const tokens = tokenize(text);
    let i = 0;
    const peekOp = (...ops) => tokens[i] && tokens[i].type === 'op' && ops.includes(tokens[i].value);
    const expectOp = (op) => {
      if (!peekOp(op)) throw ERR.PARSE;
      i++;
    };

    function binaryLevel(next, ...ops) {
      return () => {
        let left = next();
        while (peekOp(...ops)) {
          const op = tokens[i++].value;
          left = { type: 'bin', op, left, right: next() };
        }
        return left;
      };
    }

    const parsePower = () => {
      let left = parseUnary();
      while (peekOp('^')) {
        i++;
        left = { type: 'bin', op: '^', left, right: parseUnary() };
      }
      return left;
    };

    function parseUnary() {
      if (peekOp('-', '+')) {
        const op = tokens[i++].value;
        return { type: 'un', op, arg: parseUnary() };
      }
      return parsePostfix();
    }

    function parsePostfix() {
      let node = parsePrimary();
      while (peekOp('%')) {
        i++;
        node = { type: 'pct', arg: node };
      }
      return node;
    }

    function parseRef(sheet) {
      const token = tokens[i];
      if (!token || token.type !== 'ref') {
        if (token && token.type === 'err') {
          i++;
          return { type: 'err', value: ERROR_LITERALS[token.value] || ERR.REF };
        }
        throw ERR.PARSE;
      }
      i++;
      const start = parseRefParts(token.value);
      if (peekOp(':')) {
        i++;
        if (!tokens[i] || tokens[i].type !== 'ref') throw ERR.PARSE;
        const end = parseRefParts(tokens[i++].value);
        return {
          type: 'range', sheet,
          r1: Math.min(start.r, end.r), c1: Math.min(start.c, end.c),
          r2: Math.max(start.r, end.r), c2: Math.max(start.c, end.c)
        };
      }
      return { type: 'cell', sheet, r: start.r, c: start.c };
    }

    function parsePrimary() {
      const token = tokens[i];
      if (!token) throw ERR.PARSE;
      if (token.type === 'num' || token.type === 'str') {
        i++;
        return { type: token.type, value: token.value };
      }
      if (token.type === 'err') {
        i++;
        return { type: 'err', value: ERROR_LITERALS[token.value] || ERR.PARSE };
      }
      if (token.type === 'sheet') {
        i++;
        return parseRef(token.value);
      }
      if (token.type === 'ref') return parseRef(null);
      if (token.type === 'name') {
        i++;
        const upper = token.value.toUpperCase().replace(/^_XLFN\./, '');
        if (peekOp('(')) {
          i++;
          const args = [];
          if (!peekOp(')')) {
            args.push(parseExpression());
            while (peekOp(',', ';')) {
              i++;
              args.push(parseExpression());
            }
          }
          expectOp(')');
          return { type: 'call', name: upper, args };
        }
        if (upper === 'TRUE') return { type: 'bool', value: true };
        if (upper === 'FALSE') return { type: 'bool', value: false };
        return { type: 'err', value: ERR.NAME };
      }
      if (peekOp('(')) {
        i++;
        const inner = parseExpression();
        expectOp(')');
        return inner;
      }
      throw ERR.PARSE;
    }

    const parseAdditive = binaryLevel(binaryLevel(parsePower, '*', '/'), '+', '-');
    const parseConcat = binaryLevel(parseAdditive, '&');
    const parseExpression = binaryLevel(parseConcat, '=', '<>', '<', '>', '<=', '>=');

    const tree = parseExpression();
    if (i !== tokens.length) throw ERR.PARSE;
    return tree;
  }

  // --- Functions ---
  function roundTo(value, digits, mode) {
    const factor = 10 ** digits;
    const scaled = Math.abs(value) * factor;
    const precise = Number(scaled.toPrecision(15));
    const rounded = mode === 'up' ? Math.ceil(precise) : mode === 'down' ? Math.floor(precise) : Math.round(precise);
    return (Math.sign(value) * rounded) / factor;
  }

  function matchesCriteria(criteria) {
    const raw = scalar(criteria);
    if (typeof raw === 'number' || typeof raw === 'boolean') {
      return (v) => v === raw;
    }
    const text = raw === null ? '' : String(raw);
    const m = /^(<>|<=|>=|=|<|>)?(.*)$/s.exec(text);
    const op = m[1] || '=';
    const operand = m[2];
    const operandIsNumber = operand.trim() !== '' && !Number.isNaN(Number(operand));

    if (operandIsNumber) {
      const n = Number(operand);
      const tests = {
        '=': (v) => v === n, '<>': (v) => v !== n, '<': (v) => v < n,
        '>': (v) => v > n, '<=': (v) => v <= n, '>=': (v) => v >= n
      };
      return (v) => (op === '<>' ? v !== n : typeof v === 'number' && tests[op](v));
    }

    const pattern = new RegExp(`^${operand.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i');
    const matches = (v) => pattern.test(v === null ? '' : String(v));
    if (op === '=') return matches;
    if (op === '<>') return (v) => !matches(v);
    return () => false;
  }

  function serialToDate(serial) {
    return new Date(Math.round((serial - 25569) * 86400000));
  }

  function dateToSerial(utcMillis) {
    return utcMillis / 86400000 + 25569;
  }

  function criteriaSum(args) {
    const [range, criteria, sumRange = range] = args;
    if (!(range instanceof RangeValue) || !(sumRange instanceof RangeValue)) throw ERR.VALUE;
    const test = matchesCriteria(criteria);
    const hits = [];
    range.forEach((v, r, c) => {
      if (v instanceof FormulaError || !test(v)) return;
      const value = sumRange.get(sumRange.r1 + (r - range.r1), sumRange.c1 + (c - range.c1));
      if (value instanceof FormulaError) throw value;
      if (typeof value === 'number') hits.push(value);
    });
    return hits;
  }

  function looseEquals(a, b) {
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    return a === b;
  }

  const FUNCTIONS = {
    SUM: (a) => numbersIn(a).reduce((sum, n) => sum + n, 0),
    AVERAGE: (a) => {
      const nums = numbersIn(a);
      if (!nums.length) throw ERR.DIV0;
      return nums.reduce((sum, n) => sum + n, 0) / nums.length;
    },
    COUNT: (a) => {
      let count = 0;
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { if (typeof v === 'number') count++; });
        else if (typeof arg === 'number') count++;
      });
      return count;
    },
    COUNTA: (a) => {
      let count = 0;
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { if (v !== null && v !== '') count++; });
        else if (arg !== null) count++;
      });
      return count;
    },
    COUNTBLANK: (a) => {
      let count = 0;
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { if (v === null || v === '') count++; });
      });
      return count;
    },
    MIN: (a) => {
      const nums = numbersIn(a);
      return nums.length ? Math.min(...nums) : 0;
    },
    MAX: (a) => {
      const nums = numbersIn(a);
      return nums.length ? Math.max(...nums) : 0;
    },
    MEDIAN: (a) => {
      const nums = numbersIn(a).sort((x, y) => x - y);
      if (!nums.length) throw ERR.NUM;
      const mid = Math.floor(nums.length / 2);
      return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
    },
    PRODUCT: (a) => numbersIn(a).reduce((product, n) => product * n, 1),
    ABS: (a) => Math.abs(toNumber(a[0])),
    INT: (a) => Math.floor(toNumber(a[0])),
    TRUNC: (a) => {
      const factor = 10 ** (a[1] === undefined ? 0 : toNumber(a[1]));
      return Math.trunc(toNumber(a[0]) * factor) / factor;
    },
    SIGN: (a) => Math.sign(toNumber(a[0])),
    PI: () => Math.PI,
    SQRT: (a) => {
      const n = toNumber(a[0]);
      if (n < 0) throw ERR.NUM;
      return Math.sqrt(n);
    },
    POWER: (a) => toNumber(a[0]) ** toNumber(a[1]),
    MOD: (a) => {
      const divisor = toNumber(a[1]);
      if (divisor === 0) throw ERR.DIV0;
      const n = toNumber(a[0]);
      return n - divisor * Math.floor(n / divisor);
    },
    ROUND: (a) => roundTo(toNumber(a[0]), a[1] === undefined ? 0 : toNumber(a[1])),
    ROUNDUP: (a) => roundTo(toNumber(a[0]), a[1] === undefined ? 0 : toNumber(a[1]), 'up'),
    ROUNDDOWN: (a) => roundTo(toNumber(a[0]), a[1] === undefined ? 0 : toNumber(a[1]), 'down'),
    AND: (a) => {
      let result = true;
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { if (v !== null && v !== '' && !toBool(v)) result = false; });
        else if (!toBool(arg)) result = false;
      });
      return result;
    },
    OR: (a) => {
      let result = false;
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { if (v !== null && v !== '' && toBool(v)) result = true; });
        else if (toBool(arg)) result = true;
      });
      return result;
    },
    NOT: (a) => !toBool(a[0]),
    LEN: (a) => toText(a[0]).length,
    UPPER: (a) => toText(a[0]).toUpperCase(),
    LOWER: (a) => toText(a[0]).toLowerCase(),
    TRIM: (a) => toText(a[0]).trim().replace(/\s+/g, ' '),
    LEFT: (a) => toText(a[0]).slice(0, a[1] === undefined ? 1 : toNumber(a[1])),
    RIGHT: (a) => {
      const count = a[1] === undefined ? 1 : toNumber(a[1]);
      const text = toText(a[0]);
      return count === 0 ? '' : text.slice(-count);
    },
    MID: (a) => {
      const start = toNumber(a[1]);
      if (start < 1) throw ERR.VALUE;
      return toText(a[0]).substr(start - 1, toNumber(a[2]));
    },
    CONCAT: (a) => {
      let out = '';
      a.forEach((arg) => {
        if (arg instanceof RangeValue) arg.forEach((v) => { out += toText(v); });
        else out += toText(arg);
      });
      return out;
    },
    SUMIF: (a) => criteriaSum(a).reduce((sum, n) => sum + n, 0),
    AVERAGEIF: (a) => {
      const hits = criteriaSum(a);
      if (!hits.length) throw ERR.DIV0;
      return hits.reduce((sum, n) => sum + n, 0) / hits.length;
    },
    COUNTIF: (a) => {
      if (!(a[0] instanceof RangeValue)) throw ERR.VALUE;
      const test = matchesCriteria(a[1]);
      let count = 0;
      a[0].forEach((v) => {
        if (!(v instanceof FormulaError) && test(v)) count++;
      });
      return count;
    },
    VLOOKUP: (a) => {
      const [wanted, table, column, approximate] = a;
      if (!(table instanceof RangeValue)) throw ERR.VALUE;
      const key = scalar(wanted);
      const offset = toNumber(column) - 1;
      if (offset < 0 || offset >= table.cols) throw ERR.REF;
      const exact = approximate !== undefined && !toBool(approximate);
      let hit = -1;
      for (let r = table.r1; r <= table.r2; r++) {
        const candidate = table.get(r, table.c1);
        if (looseEquals(candidate, key)) {
          hit = r;
          break;
        }
        if (!exact && typeof key === 'number' && typeof candidate === 'number' && candidate < key) hit = r;
      }
      if (hit < 0) throw ERR.NA;
      return scalar(table.get(hit, table.c1 + offset));
    },
    MATCH: (a) => {
      if (!(a[1] instanceof RangeValue)) throw ERR.VALUE;
      const key = scalar(a[0]);
      const mode = a[2] === undefined ? 1 : toNumber(a[2]);
      const list = [];
      a[1].forEach((v) => list.push(v));
      if (mode === 0) {
        const found = list.findIndex((v) => looseEquals(v, key));
        if (found < 0) throw ERR.NA;
        return found + 1;
      }
      let best = -1;
      list.forEach((v, index) => {
        if (typeof v === 'number' && (mode > 0 ? v <= key : v >= key)) best = index;
      });
      if (best < 0) throw ERR.NA;
      return best + 1;
    },
    INDEX: (a) => {
      if (!(a[0] instanceof RangeValue)) throw ERR.VALUE;
      const table = a[0];
      let row = toNumber(a[1]);
      let col = a[2] === undefined ? 1 : toNumber(a[2]);
      if (a[2] === undefined && table.rows === 1) {
        col = row;
        row = 1;
      }
      if (row < 1 || col < 1 || row > table.rows || col > table.cols) throw ERR.REF;
      return scalar(table.get(table.r1 + row - 1, table.c1 + col - 1));
    },
    TODAY: () => {
      const now = new Date();
      return dateToSerial(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    },
    DATE: (a) => dateToSerial(Date.UTC(toNumber(a[0]), toNumber(a[1]) - 1, toNumber(a[2]))),
    YEAR: (a) => serialToDate(toNumber(a[0])).getUTCFullYear(),
    MONTH: (a) => serialToDate(toNumber(a[0])).getUTCMonth() + 1,
    DAY: (a) => serialToDate(toNumber(a[0])).getUTCDate()
  };
  FUNCTIONS.CONCATENATE = FUNCTIONS.CONCAT;

  // --- Engine ---
  class Engine {
    /** @param {(name: string) => object|null} resolveSheet finds a sheet ({ cells }) by name */
    constructor(resolveSheet) {
      this.resolveSheet = resolveSheet;
      this.parsed = new Map();
      this.reset();
    }

    /** Forgets computed values. Call after any cell changes. */
    reset() {
      this.cache = new WeakMap();
      this.stacks = new WeakMap();
    }

    parseCached(text) {
      if (!this.parsed.has(text)) {
        let tree;
        try {
          tree = parse(text);
        } catch (e) {
          tree = { type: 'err', value: e instanceof FormulaError ? e : ERR.PARSE };
        }
        this.parsed.set(text, tree);
      }
      return this.parsed.get(text);
    }

    /** Internal value of a cell: a number, string, boolean, null, or FormulaError. */
    cellValue(sheet, addr) {
      const cell = sheet.cells[addr];
      if (!cell) return null;
      if (!cell.f) return cell.v === undefined || cell.v === '' ? null : cell.v;

      let sheetCache = this.cache.get(sheet);
      if (!sheetCache) {
        sheetCache = new Map();
        this.cache.set(sheet, sheetCache);
      }
      if (sheetCache.has(addr)) return sheetCache.get(addr);

      if (!this.stacks.has(sheet)) this.stacks.set(sheet, new Set());
      const stack = this.stacks.get(sheet);
      if (stack.has(addr)) return ERR.CIRC;
      stack.add(addr);

      let result;
      try {
        result = this.evaluateText(cell.f, sheet);
      } catch (e) {
        // A chain of references too deep for the call stack.
        result = e instanceof RangeError ? ERR.CIRC : ERR.PARSE;
      }
      stack.delete(addr);

      // Excel files can use functions this engine lacks. Keep the value Excel saved in the file.
      if (result === ERR.NAME && cell.v !== undefined && cell.v !== '') result = cell.v;
      sheetCache.set(addr, result);
      return result;
    }

    evaluateText(formula, sheet) {
      const text = String(formula).replace(/^=/, '');
      const tree = this.parseCached(text);
      try {
        const value = this.evalNode(tree, sheet);
        if (value instanceof RangeValue) return value.rows === 1 && value.cols === 1 ? value.get(value.r1, value.c1) : ERR.VALUE;
        return value;
      } catch (e) {
        if (e instanceof FormulaError) return e;
        throw e;
      }
    }

    /** Evaluates a formula as if it were in `sheet`. Errors come back as strings such as "#DIV/0!". */
    evaluate(formula, sheet) {
      return Engine.toExternal(this.evaluateText(formula, sheet));
    }

    /** Value of a cell for display or export. */
    value(sheet, addr) {
      return Engine.toExternal(this.cellValue(sheet, addr));
    }

    static toExternal(value) {
      if (value instanceof FormulaError) return value.code;
      return value === null ? '' : value;
    }

    sheetFor(name, current) {
      if (!name) return current;
      const sheet = this.resolveSheet(name);
      if (!sheet) throw ERR.REF;
      return sheet;
    }

    argument(node, sheet) {
      // Aggregates treat a single cell reference like a one cell range.
      if (node.type === 'cell') return new RangeValue(this, this.sheetFor(node.sheet, sheet), node.r, node.c, node.r, node.c);
      return this.evalNode(node, sheet);
    }

    evalNode(node, sheet) {
      switch (node.type) {
        case 'num':
        case 'str':
        case 'bool':
          return node.value;
        case 'err':
          throw node.value;
        case 'cell': {
          const value = this.cellValue(this.sheetFor(node.sheet, sheet), toAddress(node.r, node.c));
          if (value instanceof FormulaError) throw value;
          return value;
        }
        case 'range':
          return new RangeValue(this, this.sheetFor(node.sheet, sheet), node.r1, node.c1, node.r2, node.c2);
        case 'un': {
          const n = toNumber(this.evalNode(node.arg, sheet));
          return node.op === '-' ? -n : n;
        }
        case 'pct':
          return toNumber(this.evalNode(node.arg, sheet)) / 100;
        case 'bin':
          return this.evalBinary(node, sheet);
        case 'call':
          return this.evalCall(node, sheet);
        default:
          throw ERR.PARSE;
      }
    }

    evalBinary(node, sheet) {
      const left = this.evalNode(node.left, sheet);
      const right = this.evalNode(node.right, sheet);

      if (node.op === '&') return toText(left) + toText(right);
      if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) return Engine.compare(node.op, scalar(left), scalar(right));

      const a = toNumber(left);
      const b = toNumber(right);
      let result;
      switch (node.op) {
        case '+': result = a + b; break;
        case '-': result = a - b; break;
        case '*': result = a * b; break;
        case '/':
          if (b === 0) throw ERR.DIV0;
          result = a / b;
          break;
        default: result = a ** b;
      }
      if (!Number.isFinite(result)) throw ERR.NUM;
      return result;
    }

    static compare(op, rawA, rawB) {
      // A blank cell compares as 0 against numbers and as "" against text.
      const a = rawA === null ? (typeof rawB === 'string' ? '' : 0) : rawA;
      const b = rawB === null ? (typeof rawA === 'string' ? '' : 0) : rawB;
      let order;
      if (typeof a === typeof b) {
        const x = typeof a === 'string' ? a.toLowerCase() : a;
        const y = typeof b === 'string' ? b.toLowerCase() : b;
        order = x < y ? -1 : x > y ? 1 : 0;
      } else {
        // Excel orders mixed types: numbers, then text, then booleans.
        const rank = { number: 0, string: 1, boolean: 2 };
        order = rank[typeof a] - rank[typeof b];
      }
      return { '=': order === 0, '<>': order !== 0, '<': order < 0, '>': order > 0, '<=': order <= 0, '>=': order >= 0 }[op];
    }

    evalCall(node, sheet) {
      if (node.name === 'IF') {
        if (node.args.length < 2) throw ERR.VALUE;
        if (toBool(this.evalNode(node.args[0], sheet))) return this.evalNode(node.args[1], sheet);
        return node.args[2] ? this.evalNode(node.args[2], sheet) : false;
      }
      if (node.name === 'IFERROR') {
        try {
          const value = this.evalNode(node.args[0], sheet);
          return value instanceof RangeValue ? scalar(value) : value;
        } catch (e) {
          if (!(e instanceof FormulaError)) throw e;
          return node.args[1] ? this.evalNode(node.args[1], sheet) : '';
        }
      }
      const fn = FUNCTIONS[node.name];
      if (!fn) throw ERR.NAME;
      const result = fn(node.args.map((arg) => this.argument(arg, sheet)));
      if (typeof result === 'number' && !Number.isFinite(result)) throw ERR.NUM;
      return result;
    }
  }

  // --- Rewriting references ---
  const REFERENCE = /"(?:[^"]|"")*"|((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+)(?::(\$?[A-Za-z]{1,3}\$?[0-9]+))?(?![A-Za-z0-9_(.])|[A-Za-z_][A-Za-z0-9_.]*/g;

  function parseRef(text) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/.exec(text);
    return { absC: !!m[1], c: colToIndex(m[2]), absR: !!m[3], r: Number(m[4]) - 1 };
  }

  function formatRef(ref) {
    return `${ref.absC ? '$' : ''}${indexToCol(ref.c)}${ref.absR ? '$' : ''}${ref.r + 1}`;
  }

  /** Calls `visit` for each reference. It returns [start, end] to rewrite, or null for #REF!. */
  function mapReferences(formula, visit) {
    return formula.replace(REFERENCE, (match, prefix, a, b) => {
      if (!a) return match;
      const sheet = prefix ? unquoteSheet(prefix.slice(0, -1)) : null;
      const result = visit(sheet, parseRef(a), b ? parseRef(b) : null);
      if (result === undefined) return match;
      if (result === null) return `${prefix || ''}#REF!`;
      const [start, end] = result;
      return `${prefix || ''}${formatRef(start)}${end ? `:${formatRef(end)}` : ''}`;
    });
  }

  /** Adjusts relative references for a formula copied by (dr, dc) cells. */
  function shiftRelative(formula, dr, dc) {
    return mapReferences(formula, (sheet, start, end) => {
      const move = (ref) => ({
        ...ref,
        r: ref.absR ? ref.r : ref.r + dr,
        c: ref.absC ? ref.c : ref.c + dc
      });
      const a = move(start);
      const b = end ? move(end) : null;
      if (a.r < 0 || a.c < 0 || (b && (b.r < 0 || b.c < 0))) return null;
      return [a, b];
    });
  }

  /**
   * Adjusts a formula after rows or columns change on `targetSheet`.
   * `count` > 0 inserts that many before `at`. `count` < 0 deletes that many starting at `at`.
   */
  function adjustReferences(formula, ownSheet, targetSheet, axis, at, count) {
    const key = axis === 'row' ? 'r' : 'c';
    const sameSheet = (a, b) => a.toLowerCase() === b.toLowerCase();

    return mapReferences(formula, (sheet, start, end) => {
      if (!sameSheet(sheet === null ? ownSheet : sheet, targetSheet)) return undefined;

      if (count > 0) {
        const grow = (ref) => (ref[key] >= at ? { ...ref, [key]: ref[key] + count } : ref);
        return [grow(start), end ? grow(end) : null];
      }

      const removed = -count;
      const after = at + removed;
      if (!end) {
        if (start[key] >= at && start[key] < after) return null;
        return [start[key] >= after ? { ...start, [key]: start[key] - removed } : start, null];
      }

      // Ranges shrink around what was deleted, and only vanish if all of it went.
      const low = Math.min(start[key], end[key]);
      const high = Math.max(start[key], end[key]);
      if (low >= at && high < after) return null;
      const newLow = low < at ? low : low < after ? at : low - removed;
      const newHigh = high < at ? high : high < after ? at - 1 : high - removed;
      return [{ ...start, [key]: newLow }, { ...end, [key]: newHigh }];
    });
  }

  function quoteSheet(name) {
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
  }

  function renameSheetReferences(formula, oldName, newName) {
    return formula.replace(REFERENCE, (match, prefix) => {
      if (!prefix || unquoteSheet(prefix.slice(0, -1)).toLowerCase() !== oldName.toLowerCase()) return match;
      return `${quoteSheet(newName)}!${match.slice(prefix.length)}`;
    });
  }

  function removeSheetReferences(formula, name) {
    return mapReferences(formula, (sheet) => {
      return sheet !== null && sheet.toLowerCase() === name.toLowerCase() ? null : undefined;
    });
  }

  // --- Input and display ---
  /** Reads what someone typed: "$1,200.50" becomes 1200.5 as currency, "12.5%" becomes 0.125 as percent. */
  function parseInput(text) {
    const trimmed = String(text).trim();
    const m = /^(-?)(\$?)((?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?)(%?)$/.exec(trimmed);
    if (!m || !m[3] || (m[2] && m[4])) return { value: text };
    const number = Number(m[3].replace(/,/g, '')) * (m[1] ? -1 : 1);
    if (m[4]) return { value: number / 100, format: 'percent' };
    if (m[2]) return { value: number, format: 'currency' };
    return { value: number };
  }

  function formatValue(value, format) {
    if (value === '' || value === null || value === undefined) return '';
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';

    const isNumeric = typeof value === 'number' || (String(value).trim() !== '' && !Number.isNaN(Number(value)));
    if (!isNumeric || format === 'text') return String(value);
    const n = Number(value);

    switch (format) {
      case 'currency':
        return `${n < 0 ? '-' : ''}$${Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      case 'percent':
        return `${(n * 100).toFixed(1)}%`;
      case 'decimal':
        return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'integer':
        return Math.round(n).toLocaleString('en-US');
      case 'date': {
        const date = serialToDate(n);
        return Number.isNaN(date.getTime()) ? String(n) : date.toISOString().slice(0, 10);
      }
      default:
        return String(Number(n.toPrecision(11)));
    }
  }

  return {
    Engine, parseInput, formatValue, shiftRelative, adjustReferences, renameSheetReferences,
    removeSheetReferences, colToIndex, indexToCol, toAddress
  };
});
