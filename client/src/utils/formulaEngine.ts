import { SheetData } from '../types/spreadsheet';
import { parseCellAddress, parseRange, iterateRange } from './cellRange';

type EvalResult = number | string;

export function evaluateCell(value: string, cells: SheetData): string {
  if (!value.startsWith('=')) return value;
  try {
    const result = evalExpr(value.slice(1).trim(), cells);
    if (typeof result === 'number') {
      const rounded = Math.round(result * 1e10) / 1e10;
      return isFinite(rounded) ? String(rounded) : '#DIV/0!';
    }
    return String(result);
  } catch (e) {
    const msg = e instanceof Error ? e.message : '';
    return msg.startsWith('#') ? msg : '#ERROR!';
  }
}

function evalExpr(expr: string, cells: SheetData): EvalResult {
  expr = expr.trim();
  if (!expr) return '';
  if (expr.startsWith('"') && expr.endsWith('"') && expr.length >= 2) return expr.slice(1, -1);
  const funcCall = parseFunctionCall(expr);
  if (funcCall) return callFunction(funcCall.name, funcCall.argsStr, cells);
  if (/^[A-Z]+\d+$/i.test(expr)) return resolveCellRef(expr, cells);
  return evalCompOrArith(expr, cells);
}

function parseFunctionCall(expr: string): { name: string; argsStr: string } | null {
  const nameMatch = expr.match(/^([A-Z][A-Z0-9_]*)\(/i);
  if (!nameMatch) return null;
  const name = nameMatch[1];
  const rest = expr.slice(name.length);
  let depth = 0;
  for (let i = 0; i < rest.length; i++) {
    if (rest[i] === '(') depth++;
    else if (rest[i] === ')') {
      depth--;
      if (depth === 0) {
        if (i !== rest.length - 1) return null;
        return { name: name.toUpperCase(), argsStr: rest.slice(1, i) };
      }
    }
  }
  return null;
}

function splitArgs(argsStr: string): string[] {
  const args: string[] = [];
  let depth = 0, inString = false, current = '';
  for (let i = 0; i < argsStr.length; i++) {
    const ch = argsStr[i];
    if (ch === '"') inString = !inString;
    else if (!inString) {
      if (ch === '(') depth++;
      else if (ch === ')') depth--;
      else if (ch === ',' && depth === 0) { args.push(current.trim()); current = ''; continue; }
    }
    current += ch;
  }
  if (current.trim()) args.push(current.trim());
  return args;
}

function callFunction(name: string, argsStr: string, cells: SheetData): EvalResult {
  const args = splitArgs(argsStr);
  switch (name) {
    case 'SUM':     return aggRange(args[0], cells, v => v.reduce((a, b) => a + b, 0));
    case 'AVERAGE':
    case 'AVG':     return aggRange(args[0], cells, v => v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0);
    case 'MAX':     return aggRange(args[0], cells, v => v.length ? Math.max(...v) : 0);
    case 'MIN':     return aggRange(args[0], cells, v => v.length ? Math.min(...v) : 0);
    case 'COUNT':   return aggRange(args[0], cells, v => v.length);
    case 'IF': {
      if (args.length < 2) throw new Error('#VALUE!');
      const c = evalExpr(args[0], cells);
      const truthy = typeof c === 'number' ? c !== 0 : c !== '' && c !== 'FALSE' && c !== '0';
      return truthy ? evalExpr(args[1], cells) : (args[2] ? evalExpr(args[2], cells) : 0);
    }
    case 'IFERROR': {
      if (args.length < 2) throw new Error('#VALUE!');
      try {
        const v = evalExpr(args[0], cells);
        return (typeof v === 'string' && v.startsWith('#')) ? evalExpr(args[1], cells) : v;
      } catch { return evalExpr(args[1], cells); }
    }
    case 'COUNTIF':
      if (args.length < 2) throw new Error('#VALUE!');
      return countIf(args[0], args[1], cells);
    case 'VLOOKUP':
      if (args.length < 3) throw new Error('#VALUE!');
      return vlookup(evalExpr(args[0], cells), args[1], Number(evalExpr(args[2], cells)), cells);
    case 'LEN':         return String(evalExpr(args[0] ?? '', cells)).length;
    case 'UPPER':       return String(evalExpr(args[0] ?? '', cells)).toUpperCase();
    case 'LOWER':       return String(evalExpr(args[0] ?? '', cells)).toLowerCase();
    case 'TRIM':        return String(evalExpr(args[0] ?? '', cells)).trim();
    case 'CONCATENATE': return args.map(a => String(evalExpr(a, cells))).join('');
    case 'ABS':         return Math.abs(toNum(evalExpr(args[0] ?? '0', cells)));
    case 'ROUND': {
      const v = toNum(evalExpr(args[0] ?? '0', cells));
      const p = args[1] ? toNum(evalExpr(args[1], cells)) : 0;
      return Math.round(v * Math.pow(10, p)) / Math.pow(10, p);
    }
    case 'FLOOR':   return Math.floor(toNum(evalExpr(args[0] ?? '0', cells)));
    case 'CEILING':
    case 'CEIL':    return Math.ceil(toNum(evalExpr(args[0] ?? '0', cells)));
    case 'TODAY':   return new Date().toLocaleDateString();
    case 'NOW':     return new Date().toLocaleString();
    default: throw new Error('#NAME?');
  }
}

/** Strip currency symbols and thousand-separator commas so "$1,000" → "1000", "€50,000.50" → "50000.50" */
function stripNonNumeric(s: string): string {
  // Remove leading/trailing whitespace, currency symbols, and thousand-separator commas
  return s.trim().replace(/^[£$€¥₹₩₫₪₱฿₴₸₺₼₽R\s]+|[£$€¥₹₩₫₪₱฿₴₸₺₼₽R\s]+$/g, '').replace(/,/g, '');
}

/** Parse a numeric string that may contain currency symbols and thousand-separator commas */
export function parseNum(s: string): number {
  // Handle parenthesized negatives: ($1,000) → -1000
  const trimmed = s.trim();
  if (/^\(.*\)$/.test(trimmed)) {
    return -Number(stripNonNumeric(trimmed.slice(1, -1)));
  }
  return Number(stripNonNumeric(trimmed));
}

/** Check if a display value looks like a number (including comma-formatted, currency-prefixed) */
export function isNumericValue(s: string): boolean {
  if (s === '' || s.trim() === '') return false;
  const n = parseNum(s);
  return !isNaN(n) && isFinite(n);
}

function toNum(v: EvalResult): number {
  if (typeof v === 'number') return v;
  const n = parseNum(String(v));
  return isNaN(n) ? 0 : n;
}

function aggRange(rangeStr: string, cells: SheetData, fn: (v: number[]) => number): number {
  try {
    const range = parseRange(rangeStr.trim());
    const vals: number[] = [];
    for (const key of iterateRange(range)) {
      const cell = cells.get(key);
      const raw = cell?.displayValue ?? cell?.value ?? '';
      const n = parseNum(raw);
      if (!isNaN(n) && raw !== '') vals.push(n);
    }
    return fn(vals);
  } catch { throw new Error('#VALUE!'); }
}

function countIf(rangeStr: string, criteria: string, cells: SheetData): number {
  const range = parseRange(rangeStr.trim());
  const crit = criteria.trim();
  const cmpM = crit.match(/^([><=!]{1,2})(.+)$/);
  let count = 0;
  for (const key of iterateRange(range)) {
    const cell = cells.get(key);
    const raw = cell?.displayValue ?? cell?.value ?? '';
    if (cmpM) {
      const [, op, val] = cmpM;
      const cn = parseNum(raw), vn = parseNum(val);
      if (!isNaN(cn) && !isNaN(vn)) { if (cmpOp(cn, vn, op)) count++; }
      else { if (cmpOp(raw, val.replace(/^"(.*)"$/, '$1'), op)) count++; }
    } else {
      if (raw === crit.replace(/^"(.*)"$/, '$1')) count++;
    }
  }
  return count;
}

function vlookup(lookupVal: EvalResult, rangeStr: string, colIdx: number, cells: SheetData): EvalResult {
  const range = parseRange(rangeStr.trim());
  for (let r = range.start.row; r <= range.end.row; r++) {
    const fc = cells.get(`${r}:${range.start.col}`);
    const fv = fc?.displayValue ?? fc?.value ?? '';
    const match = fv === String(lookupVal) ||
      (!isNaN(parseNum(String(lookupVal))) && !isNaN(parseNum(fv)) && parseNum(fv) === parseNum(String(lookupVal)));
    if (match) {
      const tc = cells.get(`${r}:${range.start.col + colIdx - 1}`);
      const tv = tc?.displayValue ?? tc?.value ?? '';
      const n = parseNum(tv);
      return (isNaN(n) || tv === '') ? tv : n;
    }
  }
  throw new Error('#N/A');
}

function resolveCellRef(ref: string, cells: SheetData): EvalResult {
  const addr = parseCellAddress(ref);
  const cell = cells.get(`${addr.row}:${addr.col}`);
  const val = cell?.displayValue ?? cell?.value ?? '';
  const n = parseNum(val);
  return (isNaN(n) || val === '') ? val : n;
}

function findTopLevel(expr: string, ops: string[]): { idx: number; op: string } | null {
  let depth = 0, inStr = false;
  for (let i = 0; i < expr.length; i++) {
    const ch = expr[i];
    if (ch === '"') inStr = !inStr;
    if (!inStr) {
      if (ch === '(') depth++;
      else if (ch === ')') depth--;
      else if (depth === 0) {
        for (const op of ops) {
          if (expr.slice(i, i + op.length) === op) return { idx: i, op };
        }
      }
    }
  }
  return null;
}

function evalCompOrArith(expr: string, cells: SheetData): EvalResult {
  const amp = findTopLevel(expr, ['&']);
  if (amp) {
    return String(evalExpr(expr.slice(0, amp.idx).trim(), cells)) +
           String(evalExpr(expr.slice(amp.idx + 1).trim(), cells));
  }
  const cmp = findTopLevel(expr, ['<>', '>=', '<=', '>', '<', '=']);
  if (cmp) {
    const l = evalExpr(expr.slice(0, cmp.idx).trim(), cells);
    const r = evalExpr(expr.slice(cmp.idx + cmp.op.length).trim(), cells);
    return cmpOp(l, r, cmp.op) ? 1 : 0;
  }
  const resolved = expr.replace(/[A-Z]+\d+/gi, (ref) => {
    try {
      const addr = parseCellAddress(ref);
      const cell = cells.get(`${addr.row}:${addr.col}`);
      const val = cell?.displayValue ?? cell?.value ?? '';
      const n = parseNum(val);
      return isNaN(n) ? '0' : String(n);
    } catch { return '0'; }
  });
  return safeArithmetic(resolved);
}

function cmpOp(a: EvalResult, b: EvalResult, op: string): boolean {
  switch (op) {
    case '>':  return a > b;
    case '<':  return a < b;
    case '>=': return a >= b;
    case '<=': return a <= b;
    case '=':
    case '==': return String(a) === String(b);
    case '<>':
    case '!=': return String(a) !== String(b);
    default:   return false;
  }
}

// Simple recursive descent arithmetic parser — avoids eval()
function safeArithmetic(expr: string): number | string {
  const tokens = tokenize(expr.replace(/\s+/g, ''));
  if (tokens.length === 0) return 0;
  try {
    const state = { pos: 0 };
    const result = parseExpr(tokens, state);
    return isFinite(result) ? result : '#DIV/0!';
  } catch {
    return '#VALUE!';
  }
}

function parseExpr(tokens: string[], state: { pos: number }): number {
  let left = parseTerm(tokens, state);
  while (state.pos < tokens.length && (tokens[state.pos] === '+' || tokens[state.pos] === '-')) {
    const op = tokens[state.pos++];
    const right = parseTerm(tokens, state);
    left = op === '+' ? left + right : left - right;
  }
  return left;
}

function parseTerm(tokens: string[], state: { pos: number }): number {
  let left = parseFactor(tokens, state);
  while (state.pos < tokens.length && (tokens[state.pos] === '*' || tokens[state.pos] === '/')) {
    const op = tokens[state.pos++];
    const right = parseFactor(tokens, state);
    if (op === '/') {
      if (right === 0) throw new Error('Division by zero');
      left = left / right;
    } else {
      left = left * right;
    }
  }
  return left;
}

function parseFactor(tokens: string[], state: { pos: number }): number {
  if (tokens[state.pos] === '(') {
    state.pos++;
    const val = parseExpr(tokens, state);
    if (tokens[state.pos] === ')') state.pos++;
    return val;
  }
  if (tokens[state.pos] === '-') {
    state.pos++;
    return -parseFactor(tokens, state);
  }
  const num = Number(tokens[state.pos++]);
  if (isNaN(num)) throw new Error(`Not a number: ${tokens[state.pos - 1]}`);
  return num;
}

function tokenize(expr: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  while (i < expr.length) {
    if ('+-*/()'.includes(expr[i])) {
      tokens.push(expr[i++]);
    } else if (expr[i] >= '0' && expr[i] <= '9' || expr[i] === '.') {
      let num = '';
      while (i < expr.length && (expr[i] >= '0' && expr[i] <= '9' || expr[i] === '.')) {
        num += expr[i++];
      }
      tokens.push(num);
    } else {
      i++;
    }
  }
  return tokens;
}

export function recomputeFormulas(cells: SheetData): SheetData {
  const updated = new Map(cells);
  Array.from(updated.entries()).forEach(([key, cell]) => {
    if (cell.value.startsWith('=')) {
      updated.set(key, { ...cell, displayValue: evaluateCell(cell.value, updated) });
    }
  });
  return updated;
}
