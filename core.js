'use strict';
/* =====================================================================
   District Performance Reports — core data logic
   Everything here runs in the browser. No data leaves the machine.
   ===================================================================== */

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Math.random().toString(36).slice(2, 9);
const clone = o => JSON.parse(JSON.stringify(o));
const LS = {
  get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
};
const K_CUR = 'dpr.current', K_PROF = 'dpr.profiles', K_TPL = 'dpr.templates', K_ACTIVE = 'dpr.activeProfile';

function defaultConfig() {
  return {
    v: 1,
    sheetName: null,
    headerStart: null,      // Excel row number where headings begin (null = auto)
    headerRows: null,       // number of heading rows (null = auto)
    naming: 'smart',        // smart | full
    aliases: {},            // column key -> readable name chosen by the user
    keyLabels: {},          // column key -> last known name (used to explain mapping on new files)
    summaryMode: 'auto',    // auto | none | manual
    summaryRowsManual: '',  // e.g. "4" or "4, 80"
    district: { mode: 'auto', column: null, list: [], keyColumn: null, pairs: '', verifiedSig: null },
    checks: [],
    calcs: [],
    selected: [],           // [{key, label}] in report order
    ranking: { by: null, better: null, order: 'best', by2: null, better2: 'higher', ties: 'competition', show: 'all', n: 10, min: '', max: '', excluded: [], blanks: 'end', totalLabel: 'State total', baseCol: null, baseMin: '' },
    message: { subject: '', body: '', period: '', date: '', department: '', designation: '', signatory: '', position: 'first', districtTemplate: '', mode: 'exact', listN: 5, draftLang: 'hi', source: '' },
    report: { title: 'District performance report', programme: '', orgName: '', footer: '', logo: '', format: 'detailed', tbN: 5, orientation: 'auto', showTotal: true, showStats: true, showNotes: true, highlight: true, numFmt: 'in' }
  };
}
function deepMerge(base, over) {
  if (!over || typeof over !== 'object' || Array.isArray(over)) return over === undefined ? base : over;
  const out = Array.isArray(base) ? [] : { ...base };
  for (const k of Object.keys(over)) {
    const b = base ? base[k] : undefined;
    out[k] = (b && typeof b === 'object' && !Array.isArray(b) && over[k] && typeof over[k] === 'object' && !Array.isArray(over[k])) ? deepMerge(b, over[k]) : over[k];
  }
  return out;
}

/* ---------------- value helpers ---------------- */
const isBlank = v => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
const MISSING_TOKENS = new Set(['-', '–', '—', 'na', 'n/a', 'nil', 'null', '#n/a']);
function parseNum(v) {
  if (isBlank(v)) return { n: null, bad: false, blank: true };
  if (typeof v === 'number') return Number.isFinite(v) ? { n: v, bad: false } : { n: null, bad: true };
  if (typeof v === 'boolean') return { n: null, bad: true };
  let s = String(v).trim();
  if (MISSING_TOKENS.has(s.toLowerCase())) return { n: null, bad: false, blank: true };
  s = s.replace(/[,\s\u00A0]/g, '').replace(/%$/, '');
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return { n: Number(s), bad: false };
  return { n: null, bad: true };
}
const colLetter = c => (typeof XLSX !== 'undefined') ? XLSX.utils.encode_col(c) : String(c);
const normKey = s => s.toLowerCase().replace(/\s+/g, ' ').trim();
function round(v, d) { if (v === null || !Number.isFinite(v)) return null; const f = Math.pow(10, d); return Math.round((v + Number.EPSILON * Math.sign(v)) * f) / f; }

/* ---------------- workbook -> grid ---------------- */
function sheetToGrid(ws) {
  if (!ws || !ws['!ref']) return { aoa: [], merges: [], r0: 0, c0: 0 };
  const R = XLSX.utils.decode_range(ws['!ref']);
  const aoa = [];
  for (let r = R.s.r; r <= R.e.r; r++) {
    const row = [];
    for (let c = R.s.c; c <= R.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell && cell.v !== undefined ? cell.v : null);
    }
    aoa.push(row);
  }
  const merges = (ws['!merges'] || []).map(m => ({ r1: m.s.r - R.s.r, c1: m.s.c - R.s.c, r2: m.e.r - R.s.r, c2: m.e.c - R.s.c }));
  return { aoa, merges, r0: R.s.r, c0: R.s.c };
}

/* ---------------- formula engine (no eval) ---------------- */
function tokenize(expr) {
  const out = []; const re = /\s*(?:(\{\{[^}]+\}\})|(\d+\.?\d*|\.\d+)|([-+*/()×÷−]))/y; let m; re.lastIndex = 0;
  while (re.lastIndex < expr.length) {
    const at = re.lastIndex; m = re.exec(expr);
    if (!m) { if (expr.slice(at).trim() === '') break; throw new Error('Unexpected text near “' + expr.slice(at, at + 12).trim() + '”'); }
    if (m[1]) out.push({ t: 'ref', k: m[1].slice(2, -2) });
    else if (m[2]) out.push({ t: 'num', v: Number(m[2]) });
    else out.push({ t: 'op', v: ({ '×': '*', '÷': '/', '−': '-' })[m[3]] || m[3] });
  }
  return out;
}
function parseExpr(expr) {
  const tk = tokenize(expr); let i = 0;
  const peek = () => tk[i], take = () => tk[i++];
  function E() { let n = T(); while (peek() && peek().t === 'op' && '+-'.includes(peek().v)) { const o = take().v; n = { o, a: n, b: T() }; } return n; }
  function T() { let n = F(); while (peek() && peek().t === 'op' && '*/'.includes(peek().v)) { const o = take().v; n = { o, a: n, b: F() }; } return n; }
  function F() {
    const k = take(); if (!k) throw new Error('Formula ends too early');
    if (k.t === 'num') return { num: k.v };
    if (k.t === 'ref') return { ref: k.k };
    if (k.t === 'op' && (k.v === '-' || k.v === '+')) { const f = F(); return k.v === '-' ? { o: 'neg', a: f } : f; }
    if (k.t === 'op' && k.v === '(') { const e = E(); const c = take(); if (!c || c.v !== ')') throw new Error('Missing closing bracket'); return e; }
    throw new Error('Unexpected “' + k.v + '”');
  }
  if (!tk.length) throw new Error('Formula is empty');
  const ast = E(); if (i < tk.length) throw new Error('Unexpected “' + (tk[i].v ?? '') + '”');
  return ast;
}
function astRefs(a, acc = new Set()) { if (!a) return acc; if (a.ref) acc.add(a.ref); if (a.a) astRefs(a.a, acc); if (a.b) astRefs(a.b, acc); return acc; }
function evalAst(a, get, ctx) {
  if (a.num !== undefined) return a.num;
  if (a.ref !== undefined) return get(a.ref);
  if (a.o === 'neg') { const v = evalAst(a.a, get, ctx); return v === null ? null : -v; }
  const x = evalAst(a.a, get, ctx), y = evalAst(a.b, get, ctx);
  if (x === null || y === null) return null;
  if (a.o === '+') return x + y; if (a.o === '-') return x - y; if (a.o === '*') return x * y;
  if (y === 0) { ctx.zero = true; return null; } return x / y;
}

/* ---------------- calculation types ---------------- */
const CALC_TYPES = {
  percent:   { label: 'Percentage — A ÷ B × 100', ab: true, mult: 100, suffix: '%', dec: 1 },
  ratio:     { label: 'Ratio — A ÷ B', ab: true, mult: 1, suffix: '', dec: 2 },
  divide:    { label: 'Division with a multiplier — A ÷ B × k', ab: true, mult: 1000, suffix: '', dec: 1 },
  diff:      { label: 'Difference — A − B', ab: true, mult: 1, suffix: '', dec: 0 },
  ppdiff:    { label: 'Percentage-point difference — A − B (both in %)', ab: true, mult: 1, suffix: ' pp', dec: 1 },
  pctchange: { label: 'Percentage change — (A − B) ÷ B × 100', ab: true, mult: 100, suffix: '%', dec: 1, aL: 'Current value (A)', bL: 'Earlier value (B)' },
  product:   { label: 'Product — A × B', ab: true, mult: 1, suffix: '', dec: 0 },
  sum:       { label: 'Sum of several columns', multi: true, mult: 1, suffix: '', dec: 0 },
  average:   { label: 'Average of several columns', multi: true, mult: 1, suffix: '', dec: 1 },
  share:     { label: 'Share of the all-district total — A ÷ ΣA × 100', aOnly: true, mult: 100, suffix: '%', dec: 2 },
  custom:    { label: 'Custom formula', custom: true, mult: 1, suffix: '', dec: 1 }
};
function newCalcDraft(type = 'percent') {
  const T = CALC_TYPES[type];
  return { id: null, name: '', type, a: '', b: '', cols: [], expr: '', exprText: '', multiplier: T.mult, decimals: T.dec, suffix: T.suffix, missing: 'blank' };
}
function calcRefs(c) {
  const T = CALC_TYPES[c.type] || {};
  if (T.multi) return [...(c.cols || [])];
  if (T.aOnly) return [c.a];
  if (T.custom) { try { return [...astRefs(parseExpr(c.expr || ''))]; } catch (e) { return []; } }
  return [c.a, c.b];
}
/* Evaluate one calculation for every row. Returns {vals: Map(row->v), stats} */
function evalCalc(c, rows, districtRows, getVal) {
  const T = CALC_TYPES[c.type]; const stats = { zero: 0, missing: 0, over100: 0, negative: 0, n: 0 };
  const vals = new Map(); const mult = Number.isFinite(+c.multiplier) && c.multiplier !== '' ? +c.multiplier : 1;
  let ast = null;
  if (T.custom) ast = parseExpr(c.expr || '');
  let colTotal = null;
  if (T.aOnly) { colTotal = 0; for (const r of districtRows) { const v = getVal(r, c.a); if (v !== null) colTotal += v; } }
  const miss = c.missing || 'blank';
  const g = (r, k) => { const v = getVal(r, k); return v === null && miss === 'zero' ? 0 : v; };
  for (const r of rows) {
    let v = null; const ctx = {};
    if (T.ab) {
      const a = g(r, c.a), b = g(r, c.b);
      if (a === null || b === null) { stats.missing++; }
      else if (c.type === 'percent' || c.type === 'ratio' || c.type === 'divide') { if (b === 0) stats.zero++; else v = a / b; }
      else if (c.type === 'pctchange') { if (b === 0) stats.zero++; else v = (a - b) / b; }
      else if (c.type === 'diff' || c.type === 'ppdiff') v = a - b;
      else if (c.type === 'product') v = a * b;
    } else if (T.multi) {
      const xs = (c.cols || []).map(k => getVal(r, k));
      const has = xs.filter(x => x !== null);
      if (!xs.length) v = null;
      else if (miss === 'blank' && has.length < xs.length) stats.missing++;
      else {
        const arr = miss === 'zero' ? xs.map(x => x ?? 0) : has;
        if (!arr.length) stats.missing++;
        else v = c.type === 'sum' ? arr.reduce((s, x) => s + x, 0) : arr.reduce((s, x) => s + x, 0) / arr.length;
      }
    } else if (T.aOnly) {
      const a = g(r, c.a);
      if (a === null) stats.missing++; else if (!colTotal) stats.zero++; else v = a / colTotal;
    } else if (T.custom) {
      v = evalAst(ast, k => g(r, k), ctx);
      if (v === null) { if (ctx.zero) stats.zero++; else stats.missing++; }
    }
    if (v !== null) { v = v * mult; if (!Number.isFinite(v)) { v = null; stats.zero++; } }
    if (v !== null) { stats.n++; if ((c.suffix || '').trim() === '%' && v > 100.0000001) stats.over100++; if (v < 0) stats.negative++; }
    vals.set(r, v);
  }
  return { vals, stats };
}

/* ======================================================================
   process(): rebuild every derived structure from raw grid + config
   ====================================================================== */
function process() {
  const cfg = S.cfg; const D = { issues: [], cols: [], colMap: {}, rows: [], summary: [], dataRows: [], calcStats: {}, checksOut: [] };
  S.D = D;
  const g = S.sheets && S.sheets[S.sheetName];
  if (!g || !g.aoa.length) { D.empty = true; return D; }
  const aoa = g.aoa;
  const ncol = Math.max(0, ...aoa.map(r => r.length));
  const usedCols = []; for (let c = 0; c < ncol; c++) if (aoa.some(r => !isBlank(r[c]))) usedCols.push(c);
  const xlRow = r => r + g.r0 + 1, xlCol = c => colLetter(c + g.c0);
  D.xlRow = xlRow;

  /* ---- detect title rows, header start, header row count ---- */
  const nonBlank = r => usedCols.filter(c => !isBlank(aoa[r][c])).length;
  const isDataRow = r => { const ne = usedCols.filter(c => !isBlank(aoa[r][c])); if (ne.length < 2) return false; const nums = ne.filter(c => { const p = parseNum(aoa[r][c]); return p.n !== null; }); return nums.length / ne.length >= 0.5; };
  let first = aoa.findIndex(r => r.some(v => !isBlank(v))); if (first < 0) first = 0;
  let hs = first; const titleRows = [];
  while (hs < aoa.length && usedCols.length > 3 && nonBlank(hs) === 1 && !isDataRow(hs)) { titleRows.push(hs); hs++; }
  let ds = -1; for (let r = hs; r < aoa.length; r++) if (isDataRow(r)) { ds = r; break; }
  const autoRows = ds > hs ? Math.min(ds - hs, 8) : 1;
  D.auto = { headerStart: xlRow(hs), headerRows: autoRows, titleRows: titleRows.map(xlRow), titles: titleRows.map(r => usedCols.map(c => aoa[r][c]).find(v => !isBlank(v))) };
  let H0 = cfg.headerStart ? Math.max(0, cfg.headerStart - g.r0 - 1) : hs;
  const HN = Math.max(1, Math.min(10, cfg.headerRows || autoRows));
  if (H0 + HN > aoa.length) H0 = Math.max(0, aoa.length - HN);
  D.headerStart = xlRow(H0); D.headerRows = HN; D.headerOverride = !!(cfg.headerStart || cfg.headerRows);
  if (D.headerOverride && (D.headerStart !== D.auto.headerStart || HN !== D.auto.headerRows))
    D.issues.push({ lvl: 'warn', step: 2, msg: `Heading rows are set manually (rows ${D.headerStart}–${D.headerStart + HN - 1}). Automatic detection suggests rows ${D.auto.headerStart}–${D.auto.headerStart + D.auto.headerRows - 1}.` });

  /* ---- header grid with merged cells expanded ---- */
  const Hg = []; for (let i = 0; i < HN; i++) Hg.push(usedCols.map(() => ''));
  const ci = {}; usedCols.forEach((c, j) => ci[c] = j);
  for (let i = 0; i < HN; i++) usedCols.forEach((c, j) => { const v = aoa[H0 + i] ? aoa[H0 + i][c] : null; Hg[i][j] = isBlank(v) ? '' : String(v).trim().replace(/\s+/g, ' '); });
  const mergesInHeader = g.merges.filter(m => m.r2 >= H0 && m.r1 < H0 + HN);
  for (const m of mergesInHeader) {
    const v = aoa[m.r1] ? aoa[m.r1][m.c1] : null; const s = isBlank(v) ? '' : String(v).trim().replace(/\s+/g, ' ');
    for (let r = Math.max(m.r1, H0); r <= Math.min(m.r2, H0 + HN - 1); r++) for (let c = m.c1; c <= m.c2; c++) if (ci[c] !== undefined) Hg[r - H0][ci[c]] = s;
  }
  D.mergeInfo = mergesInHeader.length ? 'merged cells' : (HN > 1 ? 'no merged cells — blank group headings filled from the left' : 'single heading row');
  if (!g.merges.length && HN > 1) {
    for (let i = 0; i < HN - 1; i++) for (let j = 1; j < usedCols.length; j++)
      if (!Hg[i][j] && Hg[i][j - 1] && Hg.slice(i + 1).some(rw => rw[j])) Hg[i][j] = Hg[i][j - 1];
  }
  /* ---- column paths and names ---- */
  const cols = usedCols.map((c, j) => {
    const parts = []; for (let i = 0; i < HN; i++) { const s = Hg[i][j]; if (s && s !== parts[parts.length - 1]) parts.push(s); }
    return { c, j, letter: xlCol(c), parts };
  });
  // group runs on the top level, used by smart naming
  const runOf = j => { const t = cols[j].parts[0]; if (!t) return [j]; let a = j, b = j; while (a > 0 && cols[a - 1].parts[0] === t) a--; while (b < cols.length - 1 && cols[b + 1].parts[0] === t) b++; const r = []; for (let k = a; k <= b; k++) r.push(k); return r; };
  const seenKeys = {};
  for (const col of cols) {
    const p = col.parts;
    let key = p.length ? normKey(p.join(' > ')) : 'col:' + col.letter;
    if (seenKeys[key]) { seenKeys[key]++; key += ' #' + seenKeys[key]; } else seenKeys[key] = 1;
    col.key = key;
    col.full = p.length ? p.join(' - ') : 'Column ' + col.letter;
    let smart = col.full;
    if (p.length >= 3) smart = p[p.length - 2] + ' - ' + p[p.length - 1];
    else if (p.length === 2) {
      const run = runOf(col.j);
      if (run.length === 1) smart = `${p[0]} (${p[1]})`;
      else if (run.some(k => cols[k].parts.length >= 3)) smart = p[1];
      else smart = `${p[0]} - ${p[1]}`;
    }
    col.smart = smart;
    col.group = p.length > 1 ? p[0] : '';
  }
  // smart names must be unique; fall back to full path for clashes
  const cnt = {}; cols.forEach(c => cnt[c.smart] = (cnt[c.smart] || 0) + 1);
  cols.forEach(c => { if (cnt[c.smart] > 1) c.smart = c.full; });
  for (const col of cols) {
    col.auto = cfg.naming === 'full' ? col.full : col.smart;
    col.name = (cfg.aliases[col.key] || '').trim() || col.auto;
    col.kind = 'orig';
  }

  /* ---- data rows ---- */
  const rows = [];
  for (let r = H0 + HN; r < aoa.length; r++) {
    if (usedCols.every(c => isBlank(aoa[r][c]))) continue;
    const row = { src: xlRow(r), v: {}, raw: {}, bad: {}, blank: {}, name: null, isSummary: false };
    for (const col of cols) { const raw = aoa[r][col.c]; row.raw[col.key] = raw; const p = parseNum(raw); row.v[col.key] = p.n; if (p.bad) row.bad[col.key] = true; if (p.blank) row.blank[col.key] = true; }
    rows.push(row);
  }
  D.rows = rows;
  for (const col of cols) {
    let nb = 0, num = 0, bad = 0, blank = 0; const texts = new Set();
    for (const r of rows) { const raw = r.raw[col.key]; if (r.blank[col.key]) blank++; if (isBlank(raw)) continue; nb++; if (r.v[col.key] !== null) num++; else { if (r.bad[col.key]) bad++; texts.add(String(raw).trim()); } }
    col.numeric = nb > 0 && num / nb >= 0.7;
    col.blanks = blank; col.invalid = col.numeric ? bad : 0; col.textUnique = texts.size; col.nonBlank = nb;
    col.decimals = col.numeric && rows.some(r => r.v[col.key] !== null && !Number.isInteger(r.v[col.key])) ? 2 : 0;
    col.suffix = '';
  }
  D.cols = cols; cols.forEach(c => { D.colMap[c.key] = c; cfg.keyLabels[c.key] = c.name; });
  D.numCols = cols.filter(c => c.numeric);
  D.textCols = cols.filter(c => !c.numeric && c.nonBlank > 0);

  /* ---- summary / total rows ---- */
  const KW = /(^|\b)(grand\s*total|total|state|overall|sum)(\b|$)|कुल|योग|राज्य/i;
  const hasKW = r => D.textCols.some(c => { const s = r.raw[c.key]; return !isBlank(s) && KW.test(String(s)); });
  const sumMatch = (row, others) => {
    let tested = 0, ok = 0;
    for (const c of D.numCols) { const v = row.v[c.key]; if (v === null) continue; let s = 0; for (const o of others) s += o.v[c.key] || 0; tested++; if (Math.abs(v - s) <= Math.max(0.5, Math.abs(s) * 1e-6)) ok++; }
    return { tested, ok, match: tested > 0 && others.length >= 2 && ok / tested >= 0.8 };
  };
  let summary = [];
  if (cfg.summaryMode === 'manual') {
    const want = new Set(String(cfg.summaryRowsManual || '').split(/[,\s]+/).map(Number).filter(Boolean));
    summary = rows.filter(r => want.has(r.src));
    if (want.size && summary.length !== want.size) D.issues.push({ lvl: 'warn', step: 2, msg: 'Some summary row numbers you entered are not data rows in this sheet.' });
  } else if (cfg.summaryMode === 'auto' && rows.length > 2) {
    const kwRows = rows.filter(hasKW);
    const cands = new Set([rows[0], rows[rows.length - 1], ...kwRows]);
    for (const cand of cands) {
      const others = rows.filter(r => r !== cand && !kwRows.includes(r));
      const m = sumMatch(cand, others);
      if (m.match || kwRows.includes(cand)) { cand.sumCheck = m; summary.push(cand); }
    }
  }
  summary.forEach(r => { r.isSummary = true; r.sumCheck = r.sumCheck || sumMatch(r, rows.filter(x => !summary.includes(x))); });
  D.summary = summary; D.total = summary[0] || null;
  D.dataRows = rows.filter(r => !r.isSummary);
  for (const s of summary) if (!s.sumCheck.match) D.issues.push({ lvl: 'warn', step: 2, msg: `Row ${s.src} is marked as a summary row but its values do not equal the sum of the district rows (${s.sumCheck.ok} of ${s.sumCheck.tested} columns match).` });

  /* ---- data quality ---- */
  for (const col of D.numCols) {
    const badRows = D.dataRows.filter(r => r.bad[col.key]).map(r => r.src);
    if (badRows.length) D.issues.push({ lvl: 'warn', step: 2, msg: `“${col.name}” has ${badRows.length} cell(s) that are not valid numbers (rows ${badRows.slice(0, 8).join(', ')}${badRows.length > 8 ? '…' : ''}). They are treated as missing.` });
    const blankRows = D.dataRows.filter(r => r.v[col.key] === null && !r.bad[col.key]).map(r => r.src);
    if (blankRows.length) D.issues.push({ lvl: 'warn', step: 2, msg: `“${col.name}” is blank in ${blankRows.length} district row(s) (rows ${blankRows.slice(0, 8).join(', ')}${blankRows.length > 8 ? '…' : ''}).` });
  }
  for (const col of D.numCols) if (D.dataRows.some(r => r.v[col.key] !== null && r.v[col.key] < 0)) D.issues.push({ lvl: 'info', step: 2, msg: `“${col.name}” contains negative values.` });

  /* ---- district identification ---- */
  resolveDistricts(D);

  /* ---- user-defined sum checks ---- */
  for (const ch of cfg.checks) {
    const okRefs = [ch.total, ...ch.parts].every(k => D.colMap[k]);
    if (!okRefs) { D.checksOut.push({ ch, broken: true }); continue; }
    const bad = [];
    for (const r of D.dataRows.concat(D.summary)) {
      const t = r.v[ch.total]; if (t === null) continue;
      const s = ch.parts.reduce((a, k) => a + (r.v[k] || 0), 0);
      if (Math.abs(s - t) > 0.5) bad.push({ r, s, t, d: s - t });
    }
    D.checksOut.push({ ch, bad });
    if (bad.length) D.issues.push({ lvl: 'warn', step: 2, msg: `Check “${ch.parts.map(k => D.colMap[k].name).join(' + ')} = ${D.colMap[ch.total].name}” fails in ${bad.length} row(s). Choose denominators with care.` });
  }

  /* ---- calculated indicators ---- */
  const allRows = D.dataRows.concat(D.summary);
  const getVal = (r, k) => { const v = r.v[k]; return v === undefined ? null : v; };
  const known = new Set(cols.map(c => c.key));
  for (const c of cfg.calcs) {
    const key = 'calc:' + c.id; const T = CALC_TYPES[c.type];
    const col = { key, name: c.name || 'Untitled calculation', kind: 'calc', numeric: true, decimals: +c.decimals || 0, suffix: c.suffix || '', calc: c, letter: '', group: 'Calculated indicators' };
    let refs = []; let err = null;
    try { refs = calcRefs(c); if (T.custom) parseExpr(c.expr || ''); } catch (e) { err = e.message; }
    const missing = refs.filter(k => !k || !known.has(k));
    if (!T) err = 'Unknown calculation type';
    else if (!refs.length) err = 'No columns chosen';
    else if (missing.length) err = 'Uses a column that is not in this file';
    if (err) { col.broken = err; for (const r of allRows) r.v[key] = null; D.issues.push({ lvl: 'err', step: 4, msg: `Calculation “${col.name}” cannot run: ${err}.` }); }
    else {
      const { vals, stats } = evalCalc(c, allRows, D.dataRows, getVal);
      for (const r of allRows) r.v[key] = vals.get(r);
      // district-only stats
      const dz = { ...stats }; D.calcStats[key] = stats;
      if (stats.zero) D.issues.push({ lvl: 'info', step: 4, msg: `“${col.name}”: ${stats.zero} row(s) left blank because the denominator is zero.` });
      if (stats.over100) D.issues.push({ lvl: 'warn', step: 4, msg: `“${col.name}”: ${stats.over100} row(s) exceed 100%. The numerator is larger than the denominator — check that these columns belong together.` });
    }
    col.formula = describeCalc(c);
    D.cols.push(col); D.colMap[key] = col; known.add(key);
  }
  D.numericAll = D.cols.filter(c => c.numeric);

  /* ---- ranking ---- */
  computeRanking(D);
  return D;
}

function resolveDistricts(D) {
  const cfg = S.cfg, dc = cfg.district, rows = D.dataRows;
  const KW = /distr|dist\b|जनपद|जिला|zila|janpad|name|नाम/i;
  const cands = D.textCols.filter(c => c.textUnique >= Math.max(2, rows.length * 0.6));
  D.districtCandidates = cands;
  const suggested = cands.find(c => KW.test(c.full)) || cands.sort((a, b) => b.textUnique - a.textUnique)[0] || null;
  let mode = dc.mode;
  if (mode === 'auto') mode = suggested ? 'column' : 'list';
  D.dmode = mode; D.dsuggested = suggested;
  const errs = [];
  rows.forEach(r => r.name = null);
  if (mode === 'column') {
    const key = (dc.column && D.colMap[dc.column]) ? dc.column : (suggested && suggested.key);
    D.dcolumn = key;
    if (!key) errs.push('No column with district names was found in this sheet. Paste a district list instead.');
    else {
      const col = D.colMap[key];
      if (col.numeric) errs.push(`“${col.name}” contains numbers, not names. A row number or serial number is not a district name.`);
      rows.forEach(r => { const s = r.raw[key]; r.name = isBlank(s) ? null : String(s).trim(); });
    }
  } else if (mode === 'list') {
    const list = (dc.list || []).map(s => String(s).trim()).filter(Boolean);
    D.dlistCount = list.length;
    if (!list.length) errs.push(`This sheet has no district names. Its ${rows.length} data rows are identified only by position. Paste the district list in the same order as the rows in this file.`);
    else if (list.length !== rows.length) errs.push(`The pasted list has ${list.length} names but the sheet has ${rows.length} district rows. They must match exactly.`);
    else {
      rows.forEach((r, i) => r.name = list[i]);
      if (dc.verifiedSig !== S.fileSig) { D.needsVerify = true; errs.push('Confirm that the pasted list follows the same row order as this file before ranking.'); }
    }
  } else if (mode === 'key') {
    const key = dc.keyColumn && D.colMap[dc.keyColumn] ? dc.keyColumn : null;
    const map = parsePairs(dc.pairs);
    D.dpairs = map.size;
    if (!key) errs.push('Choose the column that holds the district code.');
    else if (!map.size) errs.push('Paste the code-to-name list (one “code, name” pair per line).');
    else {
      const miss = [];
      rows.forEach(r => { const k = keyNorm(r.raw[key]); r.name = map.get(k) || null; if (!r.name) miss.push(r.src); });
      if (miss.length) errs.push(`${miss.length} row(s) have a code that is not in your list (rows ${miss.slice(0, 8).join(', ')}${miss.length > 8 ? '…' : ''}).`);
    }
  }
  const named = rows.filter(r => r.name);
  if (mode === 'column' && named.length < rows.length && D.dcolumn) errs.push(`${rows.length - named.length} district row(s) have no name.`);
  const seen = {}; const dups = new Set();
  named.forEach(r => { const k = r.name.toLowerCase(); if (seen[k]) dups.add(r.name); seen[k] = 1; });
  if (dups.size) errs.push('Duplicate district names: ' + [...dups].slice(0, 6).join(', ') + (dups.size > 6 ? '…' : '') + '. Each district must appear once.');
  D.districtErrors = errs;
  D.districtReady = errs.length === 0 && rows.length > 0;
  errs.forEach(m => D.issues.unshift({ lvl: 'err', step: 2, msg: m }));
  for (const s of D.summary) s.name = cfg.ranking.totalLabel || 'State total';
}
const keyNorm = v => { if (isBlank(v)) return ''; const p = parseNum(v); return p.n !== null ? String(p.n) : String(v).trim().toLowerCase(); };
function parsePairs(text) {
  const m = new Map();
  for (const line of String(text || '').split(/\r?\n/)) {
    const t = line.trim(); if (!t) continue;
    const i = t.search(/[\t,;]/); if (i < 0) continue;
    const k = keyNorm(t.slice(0, i)), n = t.slice(i + 1).trim().replace(/^"|"$/g, '');
    if (k && n) m.set(k, n);
  }
  return m;
}

function describeCalc(c) {
  const N = k => '[' + ((S.D && S.D.colMap[k] && S.D.colMap[k].name) || S.cfg.keyLabels[k] || '?') + ']';
  const m = +c.multiplier, mt = (Number.isFinite(m) && m !== 1) ? ' × ' + m : '';
  switch (c.type) {
    case 'percent': case 'ratio': case 'divide': return `${N(c.a)} ÷ ${N(c.b)}${mt}`;
    case 'pctchange': return `(${N(c.a)} − ${N(c.b)}) ÷ ${N(c.b)}${mt}`;
    case 'diff': case 'ppdiff': return `${N(c.a)} − ${N(c.b)}${mt}`;
    case 'product': return `${N(c.a)} × ${N(c.b)}${mt}`;
    case 'sum': return (c.cols || []).map(N).join(' + ') + (mt ? ` then${mt}` : '');
    case 'average': return 'Average of ' + (c.cols || []).map(N).join(', ') + mt;
    case 'share': return `${N(c.a)} ÷ total of ${N(c.a)} across all districts${mt}`;
    case 'custom': return exprToDisplay(c.expr || '') + mt;
  }
  return '';
}
function exprToDisplay(stored) { return String(stored).replace(/\{\{([^}]+)\}\}/g, (_, k) => '[' + ((S.D && S.D.colMap[k] && S.D.colMap[k].name) || S.cfg.keyLabels[k] || k) + ']'); }
function exprToStored(text, allowedCols) {
  const errs = [];
  const out = String(text).replace(/\[([^\]]+)\]/g, (_, name) => {
    const n = normKey(name); const col = allowedCols.find(c => normKey(c.name) === n);
    if (!col) { errs.push(name); return '{{?' + name + '}}'; }
    return '{{' + col.key + '}}';
  });
  return { stored: out, errs };
}

/* ======================================================================
   Ranking
   ====================================================================== */
function colDecimals(key) { const c = S.D.colMap[key]; return c ? c.decimals : 2; }
function computeRanking(D) {
  const R = S.cfg.ranking; const out = { list: [], visible: [], blanks: [], ready: false };
  D.rank = out;
  const by = R.by && D.colMap[R.by] && D.colMap[R.by].numeric ? R.by : null;
  out.by = by;
  if (!D.districtReady) { out.reason = 'Ranking is paused until every district row has a reliable name (step 2).'; out.visible = D.dataRows.map(r => ({ row: r, rank: null })); return; }
  if (!by) { out.reason = 'Choose the ranking indicator in step 5.'; out.visible = D.dataRows.map(r => ({ row: r, rank: null })); return; }
  if (!R.better) { out.reason = 'Say whether a higher or a lower value is better for this indicator (step 5).'; out.visible = D.dataRows.map(r => ({ row: r, rank: null })); return; }
  const by2 = R.by2 && D.colMap[R.by2] && D.colMap[R.by2].numeric && R.by2 !== by ? R.by2 : null;
  const ex = new Set((R.excluded || []).map(s => s.toLowerCase()));
  let pool = D.dataRows.filter(r => !ex.has((r.name || '').toLowerCase()));
  out.excludedCount = D.dataRows.length - pool.length;
  const bc = R.baseCol && D.colMap[R.baseCol] && D.colMap[R.baseCol].numeric ? R.baseCol : null;
  const bm = R.baseMin === '' || R.baseMin === null || R.baseMin === undefined ? null : +R.baseMin;
  out.baseOut = [];
  if (bc && bm !== null) { out.baseOut = pool.filter(r => !(r.v[bc] !== null && r.v[bc] >= bm)); pool = pool.filter(r => r.v[bc] !== null && r.v[bc] >= bm); out.baseCol = bc; out.baseMin = bm; }
  const d1 = colDecimals(by), d2 = by2 ? colDecimals(by2) : 0;
  const s1 = R.better === 'lower' ? 1 : -1, s2 = R.better2 === 'lower' ? 1 : -1;
  const withV = pool.filter(r => r.v[by] !== null), blanks = pool.filter(r => r.v[by] === null);
  withV.sort((a, b) => {
    const x = (a.v[by] - b.v[by]) * s1; if (x) return x;
    if (by2) { const va = a.v[by2], vb = b.v[by2]; if (va === null && vb !== null) return 1; if (vb === null && va !== null) return -1; if (va !== null && vb !== null && va !== vb) return (va - vb) * s2; }
    return (a.name || '').localeCompare(b.name || '');
  });
  const keyOf = r => R.ties === 'unique' ? null : [round(r.v[by], d1), by2 ? round(r.v[by2], d2) : 0].join('|');
  let rank = 0, dense = 0, prev = undefined;
  const list = withV.map((r, i) => {
    const k = keyOf(r);
    if (R.ties === 'unique' || k !== prev) { rank = i + 1; dense++; }
    prev = k;
    return { row: r, rank: R.ties === 'dense' ? dense : rank, value: r.v[by] };
  });
  out.list = list; out.blanks = blanks; out.ready = true; out.by2 = by2;
  // thresholds filter display only (ranks already assigned among all included districts)
  const mn = R.min === '' || R.min === null ? null : +R.min, mx = R.max === '' || R.max === null ? null : +R.max;
  let shown = list.filter(e => (mn === null || e.value >= mn) && (mx === null || e.value <= mx));
  out.thresholdHidden = list.length - shown.length;
  const n = Math.max(1, +R.n || 5);
  const topN = (arr, N) => { if (arr.length <= N) return arr.slice(); const cut = arr[N - 1].rank; return arr.filter((e, i) => i < N || e.rank === cut); };
  const botN = (arr, N) => { const rev = arr.slice().reverse(); if (rev.length <= N) return rev; const cut = rev[N - 1].rank; return rev.filter((e, i) => i < N || e.rank === cut); };
  out.topN = topN; out.botN = botN;
  let sel = shown;
  if (R.show === 'top') sel = topN(shown, n);
  else if (R.show === 'bottom') sel = botN(shown, n).reverse();
  if (R.order === 'worst') sel = sel.slice().reverse();
  const blankEntries = R.blanks === 'end' && (R.show === 'all') ? blanks.map(r => ({ row: r, rank: null, value: null })) : [];
  out.visible = sel.concat(blankEntries);
  out.shown = shown;
  // statistics over all ranked districts
  const vals = list.map(e => e.value);
  if (vals.length) {
    const sorted = vals.slice().sort((a, b) => a - b);
    out.stats = {
      n: vals.length, best: list[0], worst: list[list.length - 1],
      mean: vals.reduce((s, v) => s + v, 0) / vals.length,
      median: sorted.length % 2 ? sorted[(sorted.length - 1) / 2] : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2,
      state: D.total ? D.total.v[by] : null
    };
  }
}

/* ---------------- formatting ---------------- */
const _nf = {};
function fmtNum(v, d, fmt) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  fmt = fmt || S.cfg.report.numFmt;
  if (fmt === 'plain') return round(v, d).toFixed(d);
  const k = fmt + d; if (!_nf[k]) _nf[k] = new Intl.NumberFormat(fmt === 'in' ? 'en-IN' : 'en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
  return _nf[k].format(round(v, d));
}
function fmtCol(v, key) { const c = S.D.colMap[key]; if (!c) return '—'; if (!c.numeric) return isBlank(v) ? '—' : String(v); const s = fmtNum(v, c.decimals); return s === '—' ? s : s + (c.suffix || ''); }
function cellVal(row, key) { const c = S.D.colMap[key]; if (!c) return null; return c.numeric ? row.v[key] : row.raw[key]; }
