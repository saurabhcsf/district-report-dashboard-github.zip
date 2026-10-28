/* =====================================================================
   UI layer
   ===================================================================== */
const S = {
  fileName: '', fileSig: '', sheets: null, sheetNames: [], sheetName: null,
  step: 1, cfg: deepMerge(defaultConfig(), LS.get(K_CUR, {})), D: null,
  compat: null, draft: newCalcDraft(), activeProfile: LS.get(K_ACTIVE, ''), ui: {}
};
const STEPS = [
  ['Upload Excel', 'xlsx, xls or csv'],
  ['Confirm data', 'Headings, totals, districts'],
  ['Select columns', 'What the report shows'],
  ['Calculations', 'Percentages and more'],
  ['Rank & message', 'Order and instructions'],
  ['Preview & download', 'PDF and table export']
];

function toast(msg, ms = 2600) { const t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg; document.body.appendChild(t); setTimeout(() => t.remove(), ms); }
function persist() { LS.set(K_CUR, S.cfg); }
function getPath(o, p) { return p.split('.').reduce((a, k) => a == null ? a : a[k], o); }
function setPath(o, p, v) { const ks = p.split('.'); const last = ks.pop(); const t = ks.reduce((a, k) => a[k] ??= {}, o); t[last] = v; }

/* ---------------- render orchestration ---------------- */
function render() {
  if (S.sheets) process();
  renderTop(); renderRail(); renderMain(); persist();
}
function refresh() { // light refresh: keeps focus in the field being typed in
  if (S.sheets) process();
  renderRail();
  $$('[data-block]').forEach(el => { const f = BLOCKS[el.dataset.block]; if (f) el.innerHTML = f(); });
  postRender(); persist();
}
function hasData() { return !!(S.sheets && S.D && !S.D.empty); }

function renderTop() {
  const prof = LS.get(K_PROF, {}); const names = Object.keys(prof).sort();
  $('#top').innerHTML = `
    <div class="brand">District Performance Reports<small>PM SHRI and other district reviews</small></div>
    ${S.fileName ? `<span class="filechip" title="${esc(S.fileName)}">${esc(S.fileName)}${S.sheetNames.length > 1 ? ' · ' + esc(S.sheetName) : ''}</span>` : ''}
    <span class="spacer"></span>
    <label class="row small" style="gap:6px">Saved setup
      <select data-act-change="loadProfile" aria-label="Saved setup">
        <option value="">${names.length ? '— choose —' : 'none saved yet'}</option>
        ${names.map(n => `<option ${n === S.activeProfile ? 'selected' : ''}>${esc(n)}</option>`).join('')}
      </select></label>
    <button class="btn sm" data-act="saveProfile">Save setup</button>
    <button class="btn sm" data-act="saveProfileAs">Save as…</button>
    <button class="btn sm" data-act="exportProfile" title="Download this setup as a .json file">Export</button>
    <button class="btn sm" data-act="importProfile">Import</button>
    ${S.activeProfile ? `<button class="btn sm" data-act="deleteProfile">Delete</button>` : ''}
    ${S.fileName ? `<button class="btn sm" data-act="pickFile">Replace file</button>` : ''}`;
}

function stepStatus(i) {
  if (i === 1) return S.fileName ? 'done' : '';
  if (!hasData()) return 'off';
  const D = S.D;
  if (S.compat && i >= 2) return i === 2 ? 'block' : 'off';
  if (i === 2) return !D.districtReady ? 'block' : D.issues.some(x => x.step === 2 && x.lvl === 'warn') ? 'warn' : 'done';
  if (i === 3) return S.cfg.selected.length ? 'done' : '';
  if (i === 4) return D.cols.some(c => c.broken) ? 'block' : S.cfg.calcs.length ? 'done' : '';
  if (i === 5) return D.rank.ready ? 'done' : (D.districtReady ? '' : 'block');
  return '';
}
function stepSub(i) {
  if (!hasData() || i === 1) return i === 1 && S.fileName ? esc(S.fileName) : STEPS[i - 1][1];
  const D = S.D;
  if (S.compat && i === 2) return 'Confirm saved setup';
  if (i === 2) return D.districtReady ? `${D.dataRows.length} districts identified` : 'District names needed';
  if (i === 3) return `${S.cfg.selected.length} column(s) chosen`;
  if (i === 4) return S.cfg.calcs.length ? `${S.cfg.calcs.length} calculation(s)` : 'Optional';
  if (i === 5) return D.rank.ready ? 'By ' + esc(D.colMap[D.rank.by].name) : 'Not ranked yet';
  return 'Check, then save PDF';
}
function renderRail() {
  $('#rail').innerHTML = `<ol>${STEPS.map((s, k) => {
    const i = k + 1, st = stepStatus(i);
    return `<li class="${st} ${S.step === i ? 'on' : ''}"><button data-act="go" data-step="${i}" ${st === 'off' ? 'aria-disabled="true"' : ''} ${S.step === i ? 'aria-current="step"' : ''}>
      <span class="n">${st === 'done' && S.step !== i ? '✓' : i}</span><span><span class="t">${s[0]}</span><span class="s">${stepSub(i)}</span></span></button></li>`;
  }).join('')}</ol>
  <div class="privacy">Files are read inside this browser. Nothing is uploaded or sent anywhere.</div>`;
}

function renderMain() {
  const m = $('#main'); const f = [null, step1, step2, step3, step4, step5, step6][S.step];
  m.innerHTML = (S.step > 1 && !hasData()) ? `<div class="panel"><h2>No file loaded</h2><p class="lede">Upload the Excel file first.</p><button class="btn primary" data-act="go" data-step="1">Go to upload</button></div>` : f();
  postRender();
}
function postRender() {
  // drag and drop wiring, preview scaling
  wireDnD();
  const st = $('.stage .sheet'); if (st) { const stage = $('.stage'); const w = st.offsetWidth; const avail = stage.clientWidth - 44; st.style.zoom = w > avail ? (avail / w).toFixed(3) : 1; }
}

/* ---------------- shared bits ---------------- */
const BLOCKS = {};
function issuesHTML(stepFilter) {
  const D = S.D; const list = D.issues.filter(x => !stepFilter || stepFilter.includes(x.step));
  if (!list.length) return `<div class="note ok">No problems found.</div>`;
  const order = { err: 0, warn: 1, info: 2 };
  return `<div class="issues">${list.slice().sort((a, b) => order[a.lvl] - order[b.lvl]).map(x => `<div class="note ${x.lvl === 'err' ? 'err' : x.lvl === 'warn' ? 'warn' : ''}"><b>${x.lvl === 'err' ? 'Must fix' : x.lvl === 'warn' ? 'Check' : 'Note'}:</b> ${esc(x.msg)}</div>`).join('')}</div>`;
}
function optCols(cols, sel, blank = '— choose —') {
  return `<option value="">${esc(blank)}</option>` + cols.map(c => `<option value="${esc(c.key)}" ${c.key === sel ? 'selected' : ''}>${esc(c.name)}${c.kind === 'calc' ? ' (calculated)' : ''}</option>`).join('');
}
function selHTML(bind, opts, val, attrs = '') { return `<select data-bind="${bind}" ${attrs}>${opts.map(([v, l]) => `<option value="${esc(v)}" ${String(v) === String(val) ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`; }
function segHTML(bind, opts, val, name) { return `<span class="seg" role="radiogroup">${opts.map(([v, l]) => `<label><input type="radio" name="${name}" data-bind="${bind}" value="${esc(v)}" ${String(v) === String(val) ? 'checked' : ''}>${esc(l)}</label>`).join('')}</span>`; }
function rankChip(e) {
  if (!e || e.rank === null) return '<span class="rk">—</span>';
  const L = S.D.rank.list; const n = L.length; const botCut = n > 6 ? L[n - 3].rank : Infinity;
  const cls = e.rank <= 3 ? 'top' : (e.rank >= botCut ? 'bot' : '');
  return `<span class="rk ${cls}">${e.rank}</span>`;
}

/* ---------------- live table (steps 3-5) ---------------- */
BLOCKS.live = function () {
  const D = S.D, R = D.rank; const cols = reportColumns('detailed');
  let h = '';
  if (R.stats) {
    const by = R.by;
    h += `<div class="stats">
      <div class="stat"><span>Districts ranked</span><b>${R.stats.n}</b>${R.excludedCount ? `<em>${R.excludedCount} excluded</em>` : ''}${R.baseOut && R.baseOut.length ? `<em>${R.baseOut.length} below minimum base</em>` : ''}${R.blanks.length ? `<em>${R.blanks.length} with no value</em>` : ''}</div>
      <div class="stat"><span>Best</span><b>${esc(R.stats.best.row.name)}</b><em>${fmtCol(R.stats.best.value, by)}</em></div>
      <div class="stat"><span>Weakest</span><b>${esc(R.stats.worst.row.name)}</b><em>${fmtCol(R.stats.worst.value, by)}</em></div>
      <div class="stat"><span>District average</span><b>${fmtCol(R.stats.mean, by)}</b><em>median ${fmtCol(R.stats.median, by)}</em></div>
      ${R.stats.state !== null && R.stats.state !== undefined ? `<div class="stat"><span>${esc(S.cfg.ranking.totalLabel || 'State total')}</span><b>${fmtCol(R.stats.state, by)}</b></div>` : ''}
    </div>`;
  } else if (R.reason) h += `<div class="note ${D.districtReady ? '' : 'err'}">${esc(R.reason)}</div>`;
  if (!cols.length) return h + `<div class="note">No columns chosen yet. Pick them in step 3.</div>`;
  h += `<div class="tw"><table class="t"><thead><tr><th>Rank</th><th>District</th>${cols.map(c => `<th class="${D.colMap[c.key].numeric ? 'num' : ''}">${esc(c.label)}${D.colMap[c.key].kind === 'calc' ? ' <span class="badge calc">calc</span>' : ''}</th>`).join('')}</tr></thead><tbody>`;
  for (const e of R.visible) h += `<tr><td>${rankChip(e)}</td><td>${esc(e.row.name || '(row ' + e.row.src + ')')}</td>${cols.map(c => `<td class="${D.colMap[c.key].numeric ? 'num' : ''}">${esc(fmtCol(cellVal(e.row, c.key), c.key))}</td>`).join('')}</tr>`;
  if (D.total && S.cfg.report.showTotal) h += `<tr class="tot"><td></td><td>${esc(D.total.name)}</td>${cols.map(c => `<td class="${D.colMap[c.key].numeric ? 'num' : ''}">${esc(fmtCol(cellVal(D.total, c.key), c.key))}</td>`).join('')}</tr>`;
  return h + `</tbody></table></div><p class="tiny muted">Showing ${R.visible.length} of ${D.dataRows.length} district rows. Original source rows are untouched; this view is computed from them.</p>`;
};
function livePanel(title) { return `<div class="panel"><div class="livehead"><h2>${title || 'Live preview'}</h2><span class="small muted">Updates as you change settings</span></div><div data-block="live">${BLOCKS.live()}</div></div>`; }

/* ======================================================================
   Step 1 — upload
   ====================================================================== */
function step1() {
  const D = S.D;
  return `<div class="panel"><h2>Upload the Excel file</h2>
  <p class="lede">Use the export for the reporting period you are reviewing. When a new period starts, upload the new file here and your saved setup is reused after you confirm the columns.</p>
  <div class="drop" id="drop" data-act="pickFile" tabindex="0" role="button" aria-label="Choose a file to upload">
    <h3>${S.fileName ? 'Replace the file' : 'Drop the file here, or click to choose'}</h3>
    <div class="muted">.xlsx, .xls or .csv</div></div>
  ${S.fileName ? `
  <div class="sub">Loaded</div>
  <div class="facts">
    <div><b>${esc(S.fileName)}</b><span>File</span></div>
    <div><b>${esc(S.sheetName)}</b><span>Worksheet</span></div>
    ${hasData() ? `<div><b>${D.rows.length + D.headerRows}</b><span>Rows incl. headings</span></div><div><b>${D.cols.filter(c => c.kind === 'orig').length}</b><span>Columns</span></div>` : ''}
  </div>
  ${S.sheetNames.length > 1 ? `<label class="f" style="max-width:320px;margin-top:10px">Worksheet<select data-act-change="sheet">${S.sheetNames.map(n => `<option ${n === S.sheetName ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label>` : ''}
  ${hasData() ? '' : `<div class="note err">This worksheet is empty. Choose another worksheet or file.</div>`}
  <div class="row" style="margin-top:16px"><button class="btn primary" data-act="go" data-step="2">Continue to confirm data</button></div>` : ''}
  </div>`;
}

/* ======================================================================
   Step 2 — confirm data and district mapping
   ====================================================================== */
function compatPanel() {
  const C = S.compat; const D = S.D;
  const origCols = D.cols.filter(c => c.kind === 'orig');
  return `<div class="panel" style="border-color:var(--marigold)"><h2>Reuse your saved setup?</h2>
  <p class="lede">${C.total - C.missing.length} of ${C.total} column(s) used by your setup were found in this file by their heading.${C.missing.length ? ' The ones below were not found. Map each one to a column in the new file, or drop it. Calculations will not run on a guessed column.' : ' Confirm to apply your saved columns, calculations, ranking and message.'}</p>
  ${C.missing.length ? `<div class="tw" style="max-height:none"><table class="t"><thead><tr><th>In your saved setup</th><th>Use this column in the new file</th></tr></thead><tbody>
  ${C.missing.map(m => `<tr><td>${esc(m.label)}<div class="tiny muted">${esc(m.key)}</div></td><td><select data-compat="${esc(m.key)}"><option value="">— choose —</option><option value="__drop" ${C.map[m.key] === '__drop' ? 'selected' : ''}>Drop it from the setup</option>${origCols.map(c => `<option value="${esc(c.key)}" ${C.map[m.key] === c.key ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select></td></tr>`).join('')}
  </tbody></table></div>` : ''}
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="applyCompat">Use saved setup</button><button class="btn" data-act="freshCompat">Start fresh for this file</button></div></div>`;
}

BLOCKS.issues2 = () => issuesHTML([2]);
BLOCKS.districtSpot = function () {
  const D = S.D; const rows = D.dataRows; if (!rows.length) return '';
  const nums = D.numCols.filter(c => c.kind === 'orig').slice(0, 3);
  const pick = rows.length <= 10 ? rows : rows.slice(0, 5).concat([null], rows.slice(-5));
  return `<div class="tiny muted" style="margin:10px 0 4px">Spot-check: compare a few rows with the source portal before confirming.</div>
  <div class="tw" style="max-height:none"><table class="t"><thead><tr><th>Excel row</th><th>District name</th>${nums.map(c => `<th class="num">${esc(c.name)}</th>`).join('')}</tr></thead><tbody>
  ${pick.map(r => r ? `<tr><td>${r.src}</td><td>${r.name ? esc(r.name) : '<span class="badge err">missing</span>'}</td>${nums.map(c => `<td class="num">${fmtNum(r.v[c.key], c.decimals)}</td>`).join('')}</tr>` : `<tr><td colspan="${2 + nums.length}" class="muted">…</td></tr>`).join('')}
  </tbody></table></div>`;
};
BLOCKS.listCount = function () { const D = S.D; const n = (S.cfg.district.list || []).filter(s => String(s).trim()).length; const ok = n === D.dataRows.length; return `<span class="badge ${ok ? '' : 'err'}">${n} of ${D.dataRows.length} names</span>`; };

function step2() {
  const D = S.D, cfg = S.cfg, dc = cfg.district;
  let h = S.compat ? compatPanel() : '';
  h += `<div class="panel"><h2>Confirm the data</h2>
  <p class="lede">Check how the headings were read, which rows are totals, and how districts are identified. Your source file is never changed.</p>
  <div class="facts">
    <div><b>${esc(S.sheetName)}</b><span>Worksheet</span></div>
    <div><b>${D.headerRows}</b><span>Heading rows (${D.headerStart}–${D.headerStart + D.headerRows - 1})</span></div>
    <div><b>${D.dataRows.length}</b><span>District rows</span></div>
    <div><b>${D.summary.length ? D.summary.map(r => 'Row ' + r.src).join(', ') : 'None'}</b><span>Summary / total row</span></div>
    <div><b>${D.cols.filter(c => c.kind === 'orig').length}</b><span>Columns (${D.numCols.filter(c => c.kind === 'orig').length} numeric)</span></div>
  </div>
  ${D.auto.titleRows.length ? `<p class="small muted">Title row(s) above the headings were skipped: ${D.auto.titles.map(esc).join('; ')}.</p>` : ''}

  <div class="sub">Problems to review</div><div data-block="issues2">${BLOCKS.issues2()}</div>

  <div class="sub">Headings</div>
  <p class="small muted" style="margin-top:0">Read from ${esc(D.mergeInfo)}. Grouped headings are combined into one readable name. Edit any name; your edits are remembered for future files with the same headings.</p>
  <div class="grid" style="margin-bottom:10px">
    <label class="f">Headings start at Excel row<input type="number" min="1" data-bind="headerStart" data-num value="${cfg.headerStart || ''}" placeholder="auto (${D.auto.headerStart})"></label>
    <label class="f">Number of heading rows<input type="number" min="1" max="10" data-bind="headerRows" data-num value="${cfg.headerRows || ''}" placeholder="auto (${D.auto.headerRows})"></label>
    <label class="f">Naming style${selHTML('naming', [['smart', 'Short readable names'], ['full', 'Full heading path']], cfg.naming, 'data-full')}</label>
  </div>
  ${D.headerOverride ? `<button class="btn sm" data-act="autoHeaders">Return to automatic heading detection</button>` : ''}
  <div class="tw" style="margin-top:8px"><table class="t"><thead><tr><th>Col</th><th>Heading in the file</th><th style="min-width:240px">Name used in the dashboard</th><th>Type</th><th class="num">Blank</th><th class="num">Invalid</th></tr></thead><tbody>
  ${D.cols.filter(c => c.kind === 'orig').map(c => `<tr><td>${c.letter}</td><td class="small">${c.parts.map(esc).join(' <span class="muted">›</span> ') || '<span class="muted">(no heading)</span>'}</td>
    <td><input type="text" data-alias="${esc(c.key)}" value="${esc(cfg.aliases[c.key] || '')}" placeholder="${esc(c.auto)}" aria-label="Name for column ${c.letter}"></td>
    <td>${c.numeric ? 'Number' : '<span class="badge">Text</span>'}</td><td class="num">${c.blanks || ''}</td><td class="num ${c.invalid ? 'bad' : ''}">${c.invalid || ''}</td></tr>`).join('')}
  </tbody></table></div>

  <div class="sub">Summary / total row</div>
  <p class="small muted" style="margin-top:0">A summary row is kept out of the ranking. It can be shown at the bottom of the report as the state total.</p>
  <div class="row">${segHTML('summaryMode', [['auto', 'Detect automatically'], ['manual', 'I will choose'], ['none', 'There is none']], cfg.summaryMode, 'sm')}
    ${cfg.summaryMode === 'manual' ? `<label class="f" style="flex-direction:row;align-items:center;gap:8px">Excel row(s)<input type="text" data-bind="summaryRowsManual" value="${esc(cfg.summaryRowsManual)}" placeholder="e.g. 4" style="width:110px"></label>` : ''}
    <label class="f" style="flex-direction:row;align-items:center;gap:8px">Label in reports<input type="text" data-bind="ranking.totalLabel" value="${esc(cfg.ranking.totalLabel)}" style="width:180px"></label></div>
  ${D.summary.map(r => `<div class="note ${r.sumCheck.match ? 'ok' : 'warn'} small">Row ${r.src}: ${r.sumCheck.match ? `every checked column equals the sum of the ${D.dataRows.length} district rows (${r.sumCheck.ok} of ${r.sumCheck.tested} columns), so it is treated as the overall total.` : `values do not equal the district sum in ${r.sumCheck.tested - r.sumCheck.ok} column(s).`}</div>`).join('')}

  <div class="sub">District names</div>
  ${!D.textCols.length ? `<div class="note warn"><b>This file has no district names.</b> Every column contains numbers, so the ${D.dataRows.length} district rows can only be told apart by their position. Rankings stay paused until you supply names below.</div>` : ''}
  <div class="row" style="margin-bottom:10px">${segHTML('district.mode', [['column', 'A column in this file'], ['list', 'Paste a list in row order'], ['key', 'Match by district code']], D.dmode, 'dm')}</div>
  ${D.dmode === 'column' ? `<label class="f" style="max-width:360px">Column with district names<select data-bind="district.column" data-full>${optCols(D.textCols, D.dcolumn)}</select>${!D.textCols.length ? '<span class="h">No text columns in this sheet.</span>' : ''}</label>` : ''}
  ${D.dmode === 'list' ? `<div class="grid two"><div><label class="f">District names, one per line, in the same order as the rows of this file
     <textarea data-bind="district.list" data-lines rows="10" placeholder="Agra&#10;Aligarh&#10;…">${esc((dc.list || []).join('\n'))}</textarea></label>
     <div class="row" style="margin-top:6px"><span data-block="listCount">${BLOCKS.listCount()}</span><button class="btn sm" data-act="mapFile">Load names from a file</button></div></div>
     <div><div class="note small">Use the district order of the portal export this file came from. Do not assume alphabetical order. The list is saved with your setup, but you confirm the order again for every new file.</div>
     <label class="chk" style="margin-top:6px"><input type="checkbox" data-act-change="verifyList" ${dc.verifiedSig === S.fileSig ? 'checked' : ''} ${(dc.list || []).filter(Boolean).length !== D.dataRows.length ? 'disabled' : ''}> I have checked that these names follow the row order of <b>${esc(S.fileName)}</b>.</label></div></div>` : ''}
  ${D.dmode === 'key' ? `<div class="grid two"><label class="f">Column holding the district code<select data-bind="district.keyColumn" data-full>${optCols(D.cols.filter(c => c.kind === 'orig'), dc.keyColumn)}</select></label>
     <label class="f">Code and name, one pair per line (comma or tab)<textarea data-bind="district.pairs" rows="8" placeholder="0901, Agra&#10;0902, Aligarh">${esc(dc.pairs)}</textarea><span class="h">${D.dpairs || 0} pairs read</span></label></div>` : ''}
  <div data-block="districtSpot">${BLOCKS.districtSpot()}</div>

  <div class="sub">Consistency checks (optional)</div>
  <p class="small muted" style="margin-top:0">Check that parts add up to a total before you use them as numerator and denominator.</p>
  ${D.checksOut.map((o, i) => o.broken ? `<div class="note err small">A saved check uses a column that is not in this file. <button class="btn sm" data-act="delCheck" data-i="${i}">Remove</button></div>` :
    `<div class="note ${o.bad.length ? 'warn' : 'ok'} small"><b>${o.ch.parts.map(k => esc(D.colMap[k].name)).join(' + ')} = ${esc(D.colMap[o.ch.total].name)}</b>: ${o.bad.length ? `fails in ${o.bad.length} row(s), e.g. ${o.bad.slice(0, 3).map(b => `row ${b.r.src} (${b.d > 0 ? '+' : ''}${fmtNum(b.d, 0)})`).join(', ')}` : 'holds in every row'}. <button class="btn sm ghost" data-act="delCheck" data-i="${i}">Remove</button></div>`).join('')}
  <details ${S.ui.checkOpen ? 'open' : ''} data-act-toggle="checkOpen"><summary class="small" style="cursor:pointer;font-weight:600">Add a check</summary>
    <div class="grid two" style="margin-top:8px"><label class="f">Total column<select id="chkTotal">${optCols(D.numCols.filter(c => c.kind === 'orig'), '')}</select></label>
    <div class="f">Parts that should add up to it<div class="list" style="max-height:160px">${D.numCols.filter(c => c.kind === 'orig').map(c => `<label class="it chk"><input type="checkbox" class="chkPart" value="${esc(c.key)}">${esc(c.name)}</label>`).join('')}</div></div></div>
    <button class="btn sm" data-act="addCheck" style="margin-top:8px">Add check</button></details>
  </div>

  <div class="panel"><div class="livehead"><h2>Data preview</h2><input type="search" placeholder="Search rows" data-filter="#rawTable tbody tr" aria-label="Search rows" style="width:220px"></div>
  <div class="tw"><table class="t" id="rawTable"><thead><tr><th>Excel row</th><th>Type</th><th>District</th>${D.cols.filter(c => c.kind === 'orig').map(c => `<th class="${c.numeric ? 'num' : ''}">${esc(c.name)}</th>`).join('')}</tr></thead><tbody>
  ${D.rows.slice(0, 1000).map(r => `<tr class="${r.isSummary ? 'tot' : ''}"><td>${r.src}</td><td>${r.isSummary ? '<span class="badge tot">Total</span>' : '<span class="badge">District</span>'}</td><td>${r.isSummary ? '' : r.name ? esc(r.name) : '<span class="badge err">no name</span>'}</td>
    ${D.cols.filter(c => c.kind === 'orig').map(c => `<td class="${c.numeric ? 'num' : ''} ${r.bad[c.key] ? 'bad' : ''}">${c.numeric ? (r.bad[c.key] ? esc(r.raw[c.key]) : fmtNum(r.v[c.key], c.decimals)) : esc(r.raw[c.key] ?? '')}</td>`).join('')}</tr>`).join('')}
  </tbody></table></div>${D.rows.length > 1000 ? '<p class="tiny muted">First 1,000 rows shown.</p>' : ''}
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="go" data-step="3">Continue to select columns</button></div></div>`;
  return h;
}

/* ======================================================================
   Step 3 — select columns
   ====================================================================== */
function step3() {
  if (S.compat) return `<div class="panel"><div class="note warn">Confirm or discard your saved setup in step 2 first.</div><button class="btn" data-act="go" data-step="2">Go to step 2</button></div>`;
  const D = S.D, sel = S.cfg.selected; const selSet = new Set(sel.map(s => s.key));
  const groups = []; let cur = null;
  for (const c of D.cols) { const g = c.kind === 'calc' ? 'Calculated indicators' : (c.group || 'Columns'); if (!cur || cur.g !== g) { cur = { g, cols: [] }; groups.push(cur); } cur.cols.push(c); }
  return `<div class="panel"><h2>Select columns for the report</h2>
  <p class="lede">Tick the columns the stakeholder asked for, then arrange and rename them. District and rank are always included.</p>
  <div class="pick"><div>
    <div class="row" style="margin-bottom:8px"><input type="search" placeholder="Search indicators" data-filter="#availList .it" aria-label="Search indicators" style="flex:1"><button class="btn sm" data-act="selAll">Select all</button><button class="btn sm" data-act="selNone">Clear</button></div>
    <div class="list" id="availList">${groups.map(gr => `<div class="grp">${esc(gr.g)}</div>` + gr.cols.map(c => `<label class="it chk" data-name="${esc(c.name.toLowerCase())}"><input type="checkbox" data-pick="${esc(c.key)}" ${selSet.has(c.key) ? 'checked' : ''}>${esc(c.name)}${c.kind === 'calc' ? ' <span class="badge calc">calc</span>' : ''}${!c.numeric ? ' <span class="badge">text</span>' : ''}</label>`).join('')).join('')}</div>
  </div><div>
    <div class="row" style="justify-content:space-between;margin-bottom:8px"><b>In the report, in this order</b><span class="small muted">Drag ⋮⋮ or use arrows</span></div>
    <div class="list" id="selList">${sel.length ? sel.map((s, i) => { const c = D.colMap[s.key]; return `<div class="it" draggable="true" data-idx="${i}"><span class="handle" aria-hidden="true">⋮⋮</span>
      <input type="text" data-label="${i}" value="${esc(s.label || '')}" placeholder="${esc(c ? c.name : s.key)}" aria-label="Heading for ${esc(c ? c.name : s.key)}">
      <button class="btn sm ghost" data-act="mv" data-i="${i}" data-d="-1" aria-label="Move up" ${i === 0 ? 'disabled' : ''}>↑</button><button class="btn sm ghost" data-act="mv" data-i="${i}" data-d="1" aria-label="Move down" ${i === sel.length - 1 ? 'disabled' : ''}>↓</button><button class="btn sm ghost danger" data-act="unpick" data-i="${i}" aria-label="Remove">✕</button></div>`; }).join('')
      : '<div class="it muted">Nothing selected yet.</div>'}</div>
    <p class="tiny muted">The heading box is what the report prints. Leave it empty to use the column name.</p>
  </div></div>
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="go" data-step="4">Continue to calculations</button></div></div>` + livePanel();
}

/* ======================================================================
   Step 4 — calculations
   ====================================================================== */
function draftAllowedCols() {
  const D = S.D; const out = D.cols.filter(c => c.kind === 'orig' && c.numeric);
  const calcs = S.cfg.calcs; const lim = S.draft.id ? calcs.findIndex(c => c.id === S.draft.id) : calcs.length;
  return out.concat(calcs.slice(0, lim < 0 ? calcs.length : lim).map(c => D.colMap['calc:' + c.id]).filter(c => c && !c.broken));
}
BLOCKS.draftPreview = function () {
  const d = S.draft, D = S.D, T = CALC_TYPES[d.type];
  let c = clone(d); let err = null;
  if (T.custom) { const r = exprToStored(d.exprText || '', draftAllowedCols()); if (r.errs.length) err = 'Unknown column name(s): ' + r.errs.map(e => '[' + e + ']').join(', ') + '. Use the “Insert column” list.'; c.expr = r.stored; }
  const refs = (() => { try { return calcRefs(c); } catch (e) { err = err || e.message; return []; } })();
  if (!err && (!refs.length || refs.some(k => !k))) err = 'Choose the column(s) for this calculation.';
  if (!err && T.custom) try { parseExpr(c.expr); } catch (e) { err = e.message; }
  if (err) return `<div class="note warn">${esc(err)}</div>`;
  const rows = D.dataRows.concat(D.summary);
  const { vals, stats } = evalCalc(c, rows, D.dataRows, (r, k) => r.v[k] ?? null);
  const dec = +d.decimals || 0; const fv = v => v === null ? '—' : fmtNum(v, dec) + (d.suffix || '');
  const sample = D.dataRows.slice(0, 6).concat(D.total ? [D.total] : []);
  const showCols = [...new Set(refs)].slice(0, 4);
  const dvals = D.dataRows.map(r => vals.get(r)).filter(v => v !== null);
  return `<div class="small"><b>Formula:</b> <span class="formula">${esc(d.name || 'Result')} = ${esc(describeCalc(c))}</span></div>
  <div class="row small" style="margin:8px 0">
    <span class="badge">${dvals.length} of ${D.dataRows.length} districts have a value</span>
    ${stats.zero ? `<span class="badge err">${stats.zero} blank: denominator is zero</span>` : ''}
    ${stats.missing ? `<span class="badge err">${stats.missing} blank: input missing</span>` : ''}
    ${stats.over100 ? `<span class="badge err">${stats.over100} above 100%</span>` : ''}
    ${dvals.length ? `<span class="badge">range ${fv(Math.min(...dvals))} to ${fv(Math.max(...dvals))}</span>` : ''}
  </div>
  ${stats.over100 ? `<div class="note warn small">Some results are above 100%, so the numerator is not a subset of the denominator in those rows. Check that these two columns measure the same group of children before using this indicator.</div>` : ''}
  <div class="tw" style="max-height:none"><table class="t"><thead><tr><th>District</th>${showCols.map(k => `<th class="num">${esc(D.colMap[k] ? D.colMap[k].name : k)}</th>`).join('')}<th class="num">${esc(d.name || 'Result')}</th></tr></thead><tbody>
  ${sample.map(r => `<tr class="${r.isSummary ? 'tot' : ''}"><td>${esc(r.name || 'Row ' + r.src)}</td>${showCols.map(k => `<td class="num">${fmtCol(r.v[k], k)}</td>`).join('')}<td class="num"><b>${fv(vals.get(r))}</b></td></tr>`).join('')}
  </tbody></table></div><p class="tiny muted">First six districts${D.total ? ' and the total row' : ''}. The total row is calculated from its own values, not averaged from districts.</p>`;
};
function step4() {
  if (S.compat) return `<div class="panel"><div class="note warn">Confirm or discard your saved setup in step 2 first.</div></div>`;
  const D = S.D, d = S.draft, T = CALC_TYPES[d.type], allowed = draftAllowedCols();
  const calcCards = S.cfg.calcs.map(c => { const col = D.colMap['calc:' + c.id]; const inSel = S.cfg.selected.some(s => s.key === col.key);
    return `<div class="calccard ${col.broken ? 'broken' : ''}"><div><b>${esc(c.name)}</b> <span class="badge">${esc(CALC_TYPES[c.type].label.split(' — ')[0])}</span>${col.broken ? ` <span class="badge err">${esc(col.broken)}</span>` : ''}
      <div class="formula">${esc(col.formula)}</div><div class="tiny muted" style="margin-top:3px">${c.decimals} decimal place(s)${c.suffix ? ', shown with “' + esc(c.suffix.trim()) + '”' : ''}; missing inputs ${c.missing === 'zero' ? 'count as 0' : c.missing === 'skip' ? 'are skipped' : 'leave the result blank'}${inSel ? '' : ' · not in the report'}</div></div>
      <div class="row" style="flex-wrap:nowrap"><button class="btn sm" data-act="editCalc" data-id="${c.id}">Edit</button><button class="btn sm danger" data-act="delCalc" data-id="${c.id}">Delete</button></div></div>`; }).join('');
  return `<div class="panel"><h2>Calculations</h2>
  <p class="lede">Create percentages and other indicators from the columns. You choose every numerator and denominator; nothing is assumed. Division by zero and missing inputs give a blank, never a misleading number.</p>
  ${calcCards || '<p class="small muted">No calculations yet.</p>'}
  <div class="builder" style="margin-top:14px"><h3 style="font-size:16px;margin-bottom:10px">${d.id ? 'Edit calculation' : 'New calculation'}</h3>
  <div class="grid">
    <label class="f">Name of the indicator<input type="text" data-sbind="draft.name" value="${esc(d.name)}" placeholder="e.g. Aadhaar verification (%)"></label>
    <label class="f">Calculation<select data-act-change="draftType">${Object.entries(CALC_TYPES).map(([k, t]) => `<option value="${k}" ${k === d.type ? 'selected' : ''}>${esc(t.label)}</option>`).join('')}</select></label>
  </div>
  <div class="grid" style="margin-top:10px">
    ${T.ab || T.aOnly ? `<label class="f">${esc(T.aL || (T.aOnly ? 'Column (A)' : 'Numerator / first column (A)'))}<select data-sbind="draft.a" data-full>${optCols(allowed, d.a)}</select></label>` : ''}
    ${T.ab ? `<label class="f">${esc(T.bL || 'Denominator / second column (B)')}<select data-sbind="draft.b" data-full>${optCols(allowed, d.b)}</select></label>` : ''}
  </div>
  ${T.multi ? `<div class="f" style="margin-top:10px">Columns to combine<div class="list" style="max-height:180px">${allowed.map(c => `<label class="it chk"><input type="checkbox" data-draftcol="${esc(c.key)}" ${d.cols.includes(c.key) ? 'checked' : ''}>${esc(c.name)}</label>`).join('')}</div></div>` : ''}
  ${T.custom ? `<label class="f" style="margin-top:10px">Formula — use + − × ÷ ( ) numbers and [Column name]<textarea id="exprBox" data-sbind="draft.exprText" rows="3" placeholder="([Guardian Aadhaar - Verify] + [Guardian Aadhaar - Not Verify]) ÷ [Total Enrollment] × 100">${esc(d.exprText)}</textarea></label>
     <div class="row" style="margin-top:6px"><select id="exprIns" aria-label="Insert column">${optCols(allowed, '', 'Insert column…')}</select><button class="btn sm" data-act="insCol">Insert</button></div>` : ''}
  <div class="grid" style="margin-top:10px">
    <label class="f">Multiply result by<input type="number" step="any" data-sbind="draft.multiplier" data-num value="${esc(d.multiplier)}"><span class="h">${T.mult === 100 ? '100 turns a fraction into a percentage' : '1 leaves the result unchanged'}</span></label>
    <label class="f">Decimal places<input type="number" min="0" max="6" data-sbind="draft.decimals" data-num value="${esc(d.decimals)}"></label>
    <label class="f">Show after the number<input type="text" data-sbind="draft.suffix" value="${esc(d.suffix)}" placeholder="e.g. %"></label>
    <label class="f">When an input is missing${`<select data-sbind="draft.missing">${[['blank', 'Leave the result blank'], ['zero', 'Treat missing as 0'], ...(T.multi ? [['skip', 'Skip it, use the rest']] : [])].map(([v, l]) => `<option value="${v}" ${d.missing === v ? 'selected' : ''}>${l}</option>`).join('')}</select>`}</label>
  </div>
  <div class="sub">Preview before applying</div><div data-block="draftPreview">${BLOCKS.draftPreview()}</div>
  <div class="row" style="margin-top:12px"><button class="btn primary" data-act="saveCalc">${d.id ? 'Update calculation' : 'Add calculation'}</button>${d.id ? '<button class="btn" data-act="cancelCalc">Cancel editing</button>' : ''}</div>
  ${d.id ? '' : '<p class="tiny muted">A new calculation is added to the report columns automatically. You can remove it in step 3.</p>'}
  </div>
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="go" data-step="5">Continue to ranking</button></div></div>` + livePanel();
}

/* ======================================================================
   Step 5 — rank and message
   ====================================================================== */
function step5() {
  if (S.compat) return `<div class="panel"><div class="note warn">Confirm or discard your saved setup in step 2 first.</div></div>`;
  const D = S.D, R = S.cfg.ranking, M = S.cfg.message; const num = D.numericAll.filter(c => !c.broken);
  const tpls = LS.get(K_TPL, {}); const tn = Object.keys(tpls).sort();
  const ex = new Set((R.excluded || []).map(s => s.toLowerCase()));
  return `<div class="panel"><h2>Rank the districts</h2>
  <p class="lede">Choose the indicator, then say whether a higher or lower value is better. Rank 1 is always the best performer by that rule.</p>
  ${!D.districtReady ? `<div class="note err"><b>Ranking is paused.</b> ${esc(D.districtErrors[0] || '')} <button class="btn sm" data-act="go" data-step="2">Fix in step 2</button></div>` : ''}
  <div class="grid">
    <label class="f">Rank by<select data-bind="ranking.by" data-full>${optCols(num, R.by)}</select></label>
    <div class="f">Which is better?${segHTML('ranking.better', [['higher', 'Higher value'], ['lower', 'Lower value']], R.better, 'rb')}<span class="h">e.g. higher for verification %, lower for pending %</span></div>
    <div class="f">Order in the report${segHTML('ranking.order', [['best', 'Best first'], ['worst', 'Weakest first']], R.order, 'ro')}</div>
  </div>
  <div class="grid" style="margin-top:12px">
    <label class="f">Tie-break by (optional)<select data-bind="ranking.by2" data-full>${optCols(num.filter(c => c.key !== R.by), R.by2, 'No secondary indicator')}</select></label>
    ${R.by2 ? `<div class="f">Better for tie-break${segHTML('ranking.better2', [['higher', 'Higher'], ['lower', 'Lower']], R.better2, 'rb2')}</div>` : ''}
    <label class="f">Equal values${selHTML('ranking.ties', [['competition', 'Share a rank, skip the next (1, 2, 2, 4)'], ['dense', 'Share a rank, no gap (1, 2, 2, 3)'], ['unique', 'Always a unique rank (alphabetical last resort)']], R.ties, 'data-full')}<span class="h">Equality uses the value as displayed.</span></label>
  </div>
  <div class="grid" style="margin-top:12px">
    <label class="f">Show${selHTML('ranking.show', [['all', 'All districts'], ['top', 'Top N (best)'], ['bottom', 'Bottom N (weakest)']], R.show, 'data-full')}</label>
    ${R.show !== 'all' ? `<label class="f">N<select data-bind="ranking.n" data-num data-full>${[5, 10, 15, 20].map(n => `<option ${+R.n === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>` : ''}
    <label class="f">Only show values from<input type="number" step="any" data-bind="ranking.min" value="${esc(R.min)}" placeholder="no minimum"></label>
    <label class="f">up to<input type="number" step="any" data-bind="ranking.max" value="${esc(R.max)}" placeholder="no maximum"><span class="h">Ranks are still counted among all included districts.</span></label>
    <label class="f">Only rank districts where<select data-bind="ranking.baseCol" data-full>${optCols(D.numericAll.filter(c => c.kind === 'orig'), R.baseCol, 'No minimum base')}</select><span class="h">Stops a tiny denominator from topping a percentage ranking</span></label>
    ${R.baseCol ? `<label class="f">is at least<input type="number" step="any" data-bind="ranking.baseMin" value="${esc(R.baseMin)}" placeholder="e.g. 100"></label>` : ''}
    <label class="f">Districts with no value${selHTML('ranking.blanks', [['end', 'List at the end, unranked'], ['hide', 'Leave out']], R.blanks, 'data-full')}</label>
  </div>
  ${D.districtReady ? `<details style="margin-top:14px" ${S.ui.exOpen ? 'open' : ''} data-act-toggle="exOpen"><summary class="small" style="cursor:pointer;font-weight:600">Include or exclude districts (${D.dataRows.length - ex.size} of ${D.dataRows.length} included)</summary>
    <div class="row" style="margin:8px 0"><input type="search" placeholder="Search districts" data-filter="#exList label" aria-label="Search districts"><button class="btn sm" data-act="exAll">Include all</button><button class="btn sm" data-act="exNone">Exclude all</button></div>
    <div class="excl" id="exList">${D.dataRows.map(r => `<label data-name="${esc((r.name || '').toLowerCase())}"><input type="checkbox" data-incl="${esc(r.name)}" ${ex.has((r.name || '').toLowerCase()) ? '' : 'checked'}>${esc(r.name)}</label>`).join('')}</div></details>` : ''}
  </div>

  <div class="panel"><h2>Message to districts</h2>
  <p class="lede">Pick one of your saved sample messages. Send it exactly as written, or let the dashboard fill in this report's figures. The dashboard only prepares the report; it never sends anything.</p>
  <div data-block="library">${BLOCKS.library()}</div>
  <div class="sub">This report's message</div>
  <div class="row" style="margin-bottom:10px">${segHTML('message.mode', [['exact', 'Send exactly as written'], ['data', 'Fill in figures from the report']], M.mode, 'mmode')}
    <button class="btn sm ghost" data-act="clearMsg">Clear message</button></div>
  ${M.mode === 'data' ? `<div class="note small">Words in curly brackets are replaced with this report's figures when the PDF is made, and again automatically each time you upload a new file. Click a field to insert it where your cursor is.</div>
  <div class="row" style="gap:6px;margin:6px 0 10px">${PH_LIST.map(([k, l]) => `<button class="btn sm" data-act="insPh" data-ph="${k}" title="${esc(l)}">{${k}}</button>`).join('')}</div>
  <div class="builder" style="padding:12px 14px;margin-bottom:12px"><b>Write a draft from the data</b>
    <p class="small muted" style="margin:4px 0 8px">Keeps the opening and closing paragraphs of the current message (for example, from your sample) and writes the middle from the ranking: state figure, best and weakest districts, and how many are below the district average.</p>
    <div class="row">
      <label class="f" style="flex-direction:row;align-items:center;gap:6px">Language${selHTML('message.draftLang', [['hi', 'Hindi'], ['en', 'English'], ['both', 'Hindi and English']], M.draftLang)}</label>
      <label class="f" style="flex-direction:row;align-items:center;gap:6px">Districts in each list${selHTML('message.listN', [[3, '3'], [5, '5'], [10, '10']], M.listN, 'data-num')}</label>
      <button class="btn primary sm" data-act="draftMsg">Write draft</button></div>
    ${D.rank.ready ? '' : '<div class="note warn small">The ranking is not ready yet, so figures will show as “—” until you complete the ranking above.</div>'}</div>` : ''}
  <div class="grid two">
    <label class="f">Subject / heading<input type="text" data-bind="message.subject" value="${esc(M.subject)}" placeholder="विषय: …"></label>
    <label class="f">Reporting period<input type="text" data-bind="message.period" value="${esc(M.period)}" placeholder="e.g. October 2026"></label>
  </div>
  <label class="f" style="margin-top:10px">Message<textarea id="msgBody" data-bind="message.body" rows="9" lang="hi" placeholder="Enter the stakeholder's instruction here.">${esc(M.body)}</textarea></label>
  <div class="grid" style="margin-top:10px">
    <label class="f">Signed by<input type="text" data-bind="message.signatory" value="${esc(M.signatory)}"></label>
    <label class="f">Designation<input type="text" data-bind="message.designation" value="${esc(M.designation)}"></label>
    <label class="f">Department<input type="text" data-bind="message.department" value="${esc(M.department)}"></label>
    <label class="f">Date<input type="text" data-bind="message.date" value="${esc(M.date)}" placeholder="e.g. 28.10.2026"></label>
  </div>
  <div class="f" style="margin-top:12px">Where the message appears${segHTML('message.position', [['first', 'Top of first page'], ['cover', 'Separate cover page'], ['beforeTable', 'Just before the ranking table']], M.position, 'mp')}</div>
  <details style="margin-top:12px" ${S.ui.dOpen ? 'open' : ''} data-act-toggle="dOpen"><summary class="small" style="cursor:pointer;font-weight:600">District-specific message (used by the “one page per district” format)</summary>
    <label class="f" style="margin-top:8px">Message for each district page<textarea data-bind="message.districtTemplate" rows="5" lang="hi" placeholder="जनपद {district} का {indicator} {value} है, प्रदेश में स्थान {rank}/{total_districts}।">${esc(M.districtTemplate)}</textarea>
    <span class="h">Placeholders: {district} {rank} {total_districts} {indicator} {value} {state_value} {period}. Leave empty to reuse the main message.</span></label></details>
  <div class="row" style="margin-top:10px"><button class="btn sm" data-act="saveTpl">Save this message to the library</button></div>
  <div class="sub">Message preview${M.mode === 'data' ? ' (with this report’s figures)' : ''}</div><div class="stage" style="max-height:none;padding:14px"><div class="rpt" style="background:#fff;padding:10px 14px" data-block="msgPreview">${BLOCKS.msgPreview()}</div></div>
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="go" data-step="6">Continue to preview and download</button></div></div>` + livePanel();
}
BLOCKS.msgPreview = () => {
  const html = messageHTML(); if (!html) return '<p class="muted">The message is empty.</p>';
  const M = S.cfg.message; const unk = M.mode === 'data' ? unknownPh(M.subject + ' ' + M.body) : [];
  return (unk.length ? `<div class="note warn small">Not recognised, so printed as typed: ${unk.map(k => '{' + esc(k) + '}').join(', ')}</div>` : '') + html;
};
function libItems() { const t = LS.get(K_TPL, {}); return Object.keys(t).sort().map(n => ({ name: n, ...t[n] })); }
BLOCKS.library = function () {
  const items = libItems();
  return `<div class="row" style="margin-bottom:8px"><b>Sample message library</b><span class="small muted">${items.length} saved</span><span style="flex:1"></span>
    <button class="btn sm" data-act="uploadSamples">Upload samples (.txt, .docx)</button><button class="btn sm" data-act="pasteSample">Add a sample by pasting</button></div>
  ${items.length ? `<div class="list" style="max-height:260px">${items.map(it => `<div class="it" style="align-items:flex-start">
     <div style="flex:1;min-width:0"><b>${esc(it.name)}</b> <span class="badge ${it.mode === 'data' ? 'calc' : ''}">${it.mode === 'data' ? 'fills in figures' : 'exact text'}</span>
     <div class="tiny muted" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc((it.subject ? it.subject + ' — ' : '') + (it.body || '').replace(/\s+/g, ' ').slice(0, 160))}</div></div>
     <button class="btn sm primary" data-act="useSample" data-n="${esc(it.name)}">Use</button>
     <button class="btn sm" data-act="toggleSampleMode" data-n="${esc(it.name)}" title="Switch between exact text and filling in figures">${it.mode === 'data' ? 'Make exact' : 'Make fillable'}</button>
     <button class="btn sm ghost danger" data-act="delSample" data-n="${esc(it.name)}" aria-label="Delete ${esc(it.name)}">✕</button></div>`).join('')}</div>`
    : `<div class="note small">No samples yet. Upload a few messages you have sent before (one per file), or paste them in. They are saved in this browser and included when you export a setup.</div>`}`;
};
const PH_LIST = [
  ['period', 'Reporting period'], ['date_today', 'Today’s date'], ['indicator', 'Ranking indicator'], ['state_value', 'State total for the indicator'],
  ['average', 'District average'], ['total_districts', 'Number of districts ranked'], ['best_district', 'Best district'], ['best_value', 'Best value'],
  ['worst_district', 'Weakest district'], ['worst_value', 'Weakest value'], ['top_list', 'Numbered list of the best districts'], ['bottom_list', 'Numbered list of the weakest districts'],
  ['below_avg_count', 'Districts below the district average'], ['below_avg_list', 'Names of districts below the average']
];
function reportPlaceholders() {
  const D = S.D, R = D && D.rank, M = S.cfg.message, by = R && R.by; const dash = '—';
  const map = { period: M.period || dash, date_today: today() };
  if (!R || !R.ready || !R.stats) return { ...map, indicator: by ? D.colMap[by].name : dash };
  const f = v => fmtCol(v, by), N = Math.max(1, +M.listN || 5), st = R.stats;
  const line = (e, i) => `${i + 1}. ${e.row.name} – ${f(e.value)} (rank ${e.rank})`;
  const best = R.topN(R.list, N), worst = R.botN(R.list, N);
  const below = R.list.filter(e => (S.cfg.ranking.better === 'higher' ? e.value < st.mean : e.value > st.mean));
  return { ...map, indicator: D.colMap[by].name, state_value: st.state !== null && st.state !== undefined ? f(st.state) : dash, average: f(st.mean),
    total_districts: String(st.n), best_district: st.best.row.name, best_value: f(st.best.value), worst_district: st.worst.row.name, worst_value: f(st.worst.value),
    top_list: best.map(line).join('\n'), bottom_list: worst.map(line).join('\n'), below_avg_count: String(below.length), below_avg_list: below.map(e => e.row.name).join(', ') || dash };
}
function fillText(t, extra) { const map = { ...reportPlaceholders(), ...(extra || {}) }; return String(t || '').replace(/\{(\w+)\}/g, (m, k) => map[k] !== undefined ? map[k] : m); }
function unknownPh(t) { const known = new Set([...PH_LIST.map(p => p[0]), 'district', 'rank', 'value']); return [...new Set([...String(t).matchAll(/\{(\w+)\}/g)].map(m => m[1]).filter(k => !known.has(k)))]; }
function draftFromData() {
  const M = S.cfg.message; const lang = M.draftLang || 'hi';
  const paras = String(M.body || '').trim().split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);
  let open = '', close = '';
  if (paras.length >= 2) { open = paras[0]; close = paras[paras.length - 1]; }
  else if (paras.length === 1) { const ls = paras[0].split('\n'); if (ls.length >= 3) { open = ls[0]; close = ls[ls.length - 1]; } }
  const better = S.cfg.ranking.better === 'lower';
  const hi = `{period} की समीक्षा में {indicator} पर प्रदेश का आँकड़ा {state_value} है तथा जनपदों का औसत {average} है। कुल {total_districts} जनपदों की रैंकिंग संलग्न है।\n\nसर्वश्रेष्ठ प्रदर्शन करने वाले जनपद:\n{top_list}\n\nसबसे कमज़ोर प्रदर्शन वाले जनपद:\n{bottom_list}\n\n{below_avg_count} जनपद जनपदीय औसत से ${better ? 'ऊपर (खराब स्थिति में)' : 'नीचे'} हैं: {below_avg_list}।`;
  const en = `For {period}, the state figure for {indicator} is {state_value} and the district average is {average}. The ranking of all {total_districts} districts is attached.\n\nBest-performing districts:\n{top_list}\n\nWeakest-performing districts:\n{bottom_list}\n\n{below_avg_count} districts are ${better ? 'worse than' : 'below'} the district average: {below_avg_list}.`;
  const mid = lang === 'en' ? en : lang === 'both' ? hi + '\n\n' + en : hi;
  return [open, mid, close].filter(Boolean).join('\n\n');
}
async function readSampleFile(f) {
  if (/\.docx$/i.test(f.name)) {
    const z = XLSX.CFB.read(new Uint8Array(await f.arrayBuffer()), { type: 'array' });
    const ent = XLSX.CFB.find(z, 'word/document.xml') || XLSX.CFB.find(z, '/word/document.xml');
    if (!ent) throw new Error('no text found in ' + f.name);
    const xml = new TextDecoder('utf-8').decode(ent.content);
    const ta = document.createElement('textarea');
    const txt = xml.replace(/<w:tab\/>/g, '\t').replace(/<w:br[^>]*\/>/g, '\n').replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, '');
    ta.innerHTML = txt; return ta.value.replace(/\n{3,}/g, '\n\n').trim();
  }
  if (/\.(txt|text|md)$/i.test(f.name) || f.type.startsWith('text/')) return (await f.text()).replace(/\r\n/g, '\n').trim();
  throw new Error(f.name + ' is not a .txt or .docx file');
}
function splitSample(text) {
  // a first line starting with विषय / Subject becomes the subject
  const lines = text.split('\n'); const i = lines.findIndex(l => l.trim());
  if (i >= 0 && /^\s*(विषय|subject|sub\.?)\s*[:：-]/i.test(lines[i])) return { subject: lines[i].trim(), body: lines.slice(i + 1).join('\n').trim() };
  return { subject: '', body: text };
}
function addSamples(list) {
  const t = LS.get(K_TPL, {});
  for (const s of list) { let n = s.name, k = 2; while (t[n]) n = `${s.name} (${k++})`; t[n] = s; }
  LS.set(K_TPL, t);
}

/* ======================================================================
   Step 6 — report preview and download
   ====================================================================== */
BLOCKS.report = function () { const r = buildReport(); return `<div class="sheet ${r.orient}">${r.html}</div>`; };
BLOCKS.reportWarn = function () {
  const D = S.D, w = [];
  if (!D.districtReady) w.push(['err', 'District names are not confirmed, so the report has no ranks. Fix this in step 2.']);
  else if (!D.rank.ready) w.push(['warn', D.rank.reason]);
  if (S.cfg.report.format === 'district' && !D.rank.ready) w.push(['warn', 'The one-page-per-district format needs a ranking.']);
  if (D.cols.some(c => c.broken)) w.push(['err', 'A calculation cannot run. Fix it in step 4.']);
  if (!S.cfg.selected.length) w.push(['warn', 'No columns selected in step 3; only the ranking indicator is shown.']);
  return w.map(([l, m]) => `<div class="note ${l}">${esc(m)}</div>`).join('');
};
function step6() {
  if (S.compat) return `<div class="panel"><div class="note warn">Confirm or discard your saved setup in step 2 first.</div></div>`;
  const P = S.cfg.report;
  return `<div class="panel"><h2>Preview and download</h2>
  <p class="lede">Check the report below. “Save as PDF” opens your browser's print window: choose <b>Save as PDF</b> as the destination. Chrome or Edge gives page numbers in the footer and the most accurate Hindi.</p>
  <div class="fmt" role="radiogroup" aria-label="Report format">
    ${[['summary', 'Ranked summary', 'Rank, district, the ranking indicator and your calculated indicators.'], ['detailed', 'Detailed report', 'Rank, district and every column you selected, in your order.'], ['topbottom', 'Top and bottom performers', 'Two short tables: the strongest and the weakest districts.'], ['district', 'One page per district', 'A page for each district with its rank and the district-specific message.']]
      .map(([v, t, d]) => `<label><input type="radio" name="fmt" data-bind="report.format" data-full value="${v}" ${P.format === v ? 'checked' : ''} style="position:absolute;opacity:0"><b>${t}</b><span>${d}</span></label>`).join('')}
  </div>
  <div class="grid" style="margin-top:14px">
    <label class="f">Report title<input type="text" data-bind="report.title" value="${esc(P.title)}"></label>
    <label class="f">Programme or department<input type="text" data-bind="report.programme" value="${esc(P.programme)}" placeholder="e.g. PM SHRI · Samagra Shiksha, UP"></label>
    <label class="f">Organisation name (header)<input type="text" data-bind="report.orgName" value="${esc(P.orgName)}"></label>
    <label class="f">Footer text<input type="text" data-bind="report.footer" value="${esc(P.footer)}" placeholder="e.g. For official use only"></label>
    ${P.format === 'topbottom' ? `<label class="f">Districts in each table<select data-bind="report.tbN" data-num data-full>${[3, 5, 10, 15].map(n => `<option ${+P.tbN === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>` : ''}
    <label class="f">Page orientation${selHTML('report.orientation', [['auto', 'Automatic (landscape for wide tables)'], ['portrait', 'Portrait'], ['landscape', 'Landscape']], P.orientation, 'data-full')}</label>
    <label class="f">Number format${selHTML('report.numFmt', [['in', 'Indian (12,34,567)'], ['intl', 'International (1,234,567)'], ['plain', 'No separators']], P.numFmt, 'data-full')}</label>
    <div class="f">Logo<div class="row"><button class="btn sm" data-act="pickLogo">${P.logo ? 'Change logo' : 'Add logo'}</button>${P.logo ? '<button class="btn sm ghost" data-act="rmLogo">Remove</button>' : ''}</div></div>
  </div>
  <div class="row" style="margin-top:12px;gap:18px">
    <label class="chk"><input type="checkbox" data-bind="report.showTotal" data-bool data-full ${P.showTotal ? 'checked' : ''}>State total row</label>
    <label class="chk"><input type="checkbox" data-bind="report.showStats" data-bool data-full ${P.showStats ? 'checked' : ''}>Summary figures</label>
    <label class="chk"><input type="checkbox" data-bind="report.highlight" data-bool data-full ${P.highlight ? 'checked' : ''}>Shade ranks 1–3 and the weakest three ranks</label>
    <label class="chk"><input type="checkbox" data-bind="report.showNotes" data-bool data-full ${P.showNotes ? 'checked' : ''}>Method notes and formulas</label>
  </div>
  <div data-block="reportWarn" style="margin-top:10px">${BLOCKS.reportWarn()}</div>
  <div class="row" style="margin-top:14px"><button class="btn primary" data-act="print">Save as PDF</button><button class="btn" data-act="exportXlsx">Download table (.xlsx)</button><button class="btn" data-act="exportCsv">Download table (.csv)</button></div>
  </div>
  <div class="panel"><div class="livehead"><h2>Report preview</h2><span class="small muted">Page breaks in the PDF follow the paper size; table headings repeat on every page.</span></div>
  <div class="stage" data-block="report">${BLOCKS.report()}</div></div>`;
}

/* ======================================================================
   Report builder (shared by preview, PDF and exports)
   ====================================================================== */
function reportColumns(format) {
  const D = S.D, sel = S.cfg.selected.filter(s => D.colMap[s.key]);
  const lab = s => (s.label || '').trim() || D.colMap[s.key].name;
  const labelFor = k => { const s = sel.find(x => x.key === k); return s ? lab(s) : D.colMap[k].name; };
  let cols = [];
  if (format === 'summary' || format === 'topbottom' || format === 'district-key') {
    const add = k => { if (k && D.colMap[k] && !cols.some(x => x.key === k)) cols.push({ key: k, label: labelFor(k) }); };
    add(D.rank.by); add(D.rank.by2);
    sel.filter(s => D.colMap[s.key].kind === 'calc').forEach(s => add(s.key));
  } else cols = sel.map(s => ({ key: s.key, label: lab(s) }));
  if (!cols.length && D.rank.by) cols = [{ key: D.rank.by, label: D.colMap[D.rank.by].name }];
  return cols;
}
function today() { return new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }); }
function messageHTML(text, extra) {
  const M = S.cfg.message; let body = text !== undefined ? text : M.body; let subject = M.subject;
  if (M.mode === 'data' || extra) { body = fillText(body, extra); subject = fillText(subject, extra); }
  if (!subject && !body && !M.signatory) return '';
  const sig = [M.signatory && `<b>${esc(M.signatory)}</b>`, M.designation && esc(M.designation), M.department && esc(M.department), M.date && esc(M.date)].filter(Boolean);
  return `<div class="msg">${subject ? `<div class="subj">${esc(subject)}</div>` : ''}${body ? `<div class="body">${esc(body)}</div>` : ''}${sig.length ? `<div class="sig">${sig.join('<br>')}</div>` : ''}</div>`;
}
function reportHead() {
  const P = S.cfg.report;
  if (!P.logo && !P.orgName && !P.programme) return '';
  return `<div class="rh">${P.logo ? `<img src="${P.logo}" alt="">` : ''}<div>${P.orgName ? `<div class="org">${esc(P.orgName)}</div>` : ''}${P.programme ? `<div class="prog">${esc(P.programme)}</div>` : ''}</div></div>`;
}
function metaLine() {
  const M = S.cfg.message; const bits = [];
  if (M.period) bits.push('Reporting period: ' + esc(M.period));
  bits.push('Generated ' + today());
  return `<div class="meta">${bits.join(' &nbsp;|&nbsp; ')}</div>`;
}
function rptTable(entries, cols, { showRank = true, total = true } = {}) {
  const D = S.D, P = S.cfg.report, L = D.rank.list;
  const botCut = L.length > 6 ? L[L.length - 3].rank : Infinity;
  const topSet = new Set(L.filter(e => e.rank <= 3).map(e => e.row)), botSet = new Set(L.filter(e => e.rank >= botCut && e.rank > 3).map(e => e.row));
  const isNum = k => D.colMap[k].numeric;
  let h = `<table><thead><tr>${showRank ? '<th class="rkc">Rank</th>' : ''}<th>District</th>${cols.map(c => `<th class="${isNum(c.key) ? 'num' : ''}">${esc(c.label)}</th>`).join('')}</tr></thead><tbody>`;
  for (const e of entries) {
    const cls = P.highlight && e.rank ? (topSet.has(e.row) ? 'hi-top' : botSet.has(e.row) ? 'hi-bot' : '') : '';
    h += `<tr class="${cls}">${showRank ? `<td class="rkc">${e.rank ?? '—'}</td>` : ''}<td>${esc(e.row.name || 'Row ' + e.row.src)}</td>${cols.map(c => `<td class="${isNum(c.key) ? 'num' : ''}">${esc(fmtCol(cellVal(e.row, c.key), c.key))}</td>`).join('')}</tr>`;
  }
  if (total && P.showTotal && D.total) h += `<tr class="tot">${showRank ? '<td></td>' : ''}<td>${esc(D.total.name)}</td>${cols.map(c => `<td class="${isNum(c.key) ? 'num' : ''}">${esc(fmtCol(cellVal(D.total, c.key), c.key))}</td>`).join('')}</tr>`;
  return h + '</tbody></table>';
}
function kpiHTML() {
  const R = S.D.rank; if (!R.stats || !S.cfg.report.showStats) return '';
  const by = R.by, st = R.stats;
  return `<div class="kpis"><div class="kpi"><span>Districts ranked</span>${st.n}</div>
    <div class="kpi"><span>Best</span>${esc(st.best.row.name)} (${fmtCol(st.best.value, by)})</div>
    <div class="kpi"><span>Weakest</span>${esc(st.worst.row.name)} (${fmtCol(st.worst.value, by)})</div>
    <div class="kpi"><span>District average</span>${fmtCol(st.mean, by)}</div>
    ${st.state !== null && st.state !== undefined ? `<div class="kpi"><span>${esc(S.cfg.ranking.totalLabel || 'State total')}</span>${fmtCol(st.state, by)}</div>` : ''}</div>`;
}
function notesHTML(cols) {
  if (!S.cfg.report.showNotes) return '';
  const D = S.D, R = D.rank, C = S.cfg.ranking, n = [];
  if (R.ready) {
    n.push(`Ranked by <b>${esc(D.colMap[R.by].name)}</b>; ${C.better === 'higher' ? 'a higher' : 'a lower'} value is better and rank 1 is the best performer. Listed ${C.order === 'best' ? 'best first' : 'weakest first'}.`);
    if (R.by2) n.push(`Equal values are separated by ${esc(D.colMap[R.by2].name)} (${C.better2 === 'higher' ? 'higher' : 'lower'} is better).`);
    n.push(C.ties === 'unique' ? 'Every district has a unique rank; any remaining ties are ordered alphabetically.' : `Districts with equal values share a rank${C.ties === 'dense' ? ' with no gap after a tie' : ' and the next rank is skipped'}.`);
    if (R.excludedCount) n.push(`${R.excludedCount} district(s) excluded from the ranking: ${esc((C.excluded || []).join(', '))}.`);
    if (R.baseOut && R.baseOut.length) n.push(`${R.baseOut.length} district(s) not ranked because ${esc(D.colMap[R.baseCol].name)} is below ${fmtNum(R.baseMin, 0)}: ${R.baseOut.map(r => esc(r.name)).join(', ')}.`);
    if (R.thresholdHidden) n.push(`${R.thresholdHidden} district(s) outside the value filter are not shown; ranks count all included districts.`);
    if (R.blanks.length) n.push(`${R.blanks.length} district(s) have no value for the ranking indicator and are not ranked.`);
  }
  const keys = new Set(cols.map(c => c.key)); if (R.by) keys.add(R.by);
  for (const k of keys) { const c = D.colMap[k]; if (c && c.kind === 'calc') n.push(`<b>${esc(c.name)}</b> = ${esc(c.formula)}. Blank where the denominator is zero or an input is missing.`); }
  n.push(`Source: ${esc(S.fileName)}, sheet “${esc(S.sheetName)}”${D.total ? `; the ${esc(D.total.name.toLowerCase())} row is row ${D.total.src} of the source` : ''}.`);
  return `<div class="notes">${n.map(x => `<p>${x}</p>`).join('')}</div>`;
}
function districtPh(e) {
  const D = S.D, R = D.rank, by = R.by;
  return { district: e.row.name, rank: e.rank ?? '—', value: by ? fmtCol(e.row.v[by], by) : '' };
}
function buildReport() {
  const D = S.D, P = S.cfg.report, M = S.cfg.message, R = D.rank;
  const fmt = P.format;
  const cols = reportColumns(fmt === 'district' ? 'detailed' : fmt);
  const ncols = cols.length + 2;
  const orient = P.orientation !== 'auto' ? P.orientation : (fmt === 'district' ? 'portrait' : ncols > 7 ? 'landscape' : 'portrait');
  const fs = ncols <= 6 ? 10 : ncols <= 9 ? 9 : ncols <= 12 ? 8 : ncols <= 16 ? 7 : 6.5;
  const title = `<h1>${esc(P.title || 'District performance report')}</h1>`;
  const msg = messageHTML();
  let h = '';
  if (M.position === 'cover' && fmt !== 'district') h += `<div class="pb cover">${reportHead()}${title}${metaLine()}${msg}</div>`;
  const top = () => (M.position === 'cover' ? '' : reportHead() + title + metaLine()) + (M.position === 'first' ? msg : '');
  const ranked = R.ready;
  if (fmt === 'district') {
    if (!ranked) return { html: `<div class="rpt">${reportHead()}${title}<p>The one-page-per-district format needs a ranking. Complete step 5.</p></div>`, orient };
    const rows = R.visible.filter(e => e.rank !== null);
    const tpl = M.districtTemplate || M.body;
    h += rows.map((e, i) => `<div class="${i < rows.length - 1 ? 'pb' : ''} dcard">${reportHead()}<h1>${esc(e.row.name)}</h1>${metaLine()}
      <div class="dbig"><div><span>Rank</span><b>${e.rank} of ${R.list.length}</b></div><div><span>${esc(D.colMap[R.by].name)}</span><b>${fmtCol(e.value, R.by)}</b></div>${D.total ? `<div><span>${esc(D.total.name)}</span><b>${fmtCol(D.total.v[R.by], R.by)}</b></div>` : ''}</div>
      ${tpl || M.subject ? messageHTML(tpl, districtPh(e)) : ''}
      <table><thead><tr><th>Indicator</th><th class="num">${esc(e.row.name)}</th>${D.total ? `<th class="num">${esc(D.total.name)}</th>` : ''}</tr></thead><tbody>
      ${cols.map(c => `<tr><td>${esc(c.label)}</td><td class="num">${esc(fmtCol(cellVal(e.row, c.key), c.key))}</td>${D.total ? `<td class="num">${esc(fmtCol(cellVal(D.total, c.key), c.key))}</td>` : ''}</tr>`).join('')}</tbody></table></div>`).join('');
    return { html: `<div class="rpt" style="--rfs:10.5pt">${h}</div>`, orient };
  }
  h += top() + (ranked ? kpiHTML() : '') + (M.position === 'beforeTable' ? msg : '');
  if (fmt === 'topbottom') {
    if (!ranked) h += `<p>Top and bottom tables need a ranking. Complete step 5.</p>`;
    else {
      const N = +P.tbN || 5;
      const pool = R.shown;
      const best = R.topN(pool, N), worst = R.botN(pool, N);
      h += `<h2>Top ${N} districts</h2>${rptTable(best, cols, { total: false })}<h2>Bottom ${N} districts</h2>${rptTable(worst, cols, { total: true })}`;
    }
  } else {
    h += rptTable(R.visible, cols, { showRank: ranked });
  }
  h += notesHTML(cols);
  return { html: `<div class="rpt" style="--rfs:${fs}pt">${h}</div>`, orient };
}
function cssStr(s) { return '"' + String(s || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n]+/g, ' ') + '"'; }
function pageCSS(orient) {
  const P = S.cfg.report; const f = 'font-family:Mukta,"Nirmala UI","Noto Sans Devanagari",sans-serif;font-size:8pt;color:#4A536B;';
  return `@page{size:A4 ${orient};margin:13mm 11mm 15mm 11mm;
  @bottom-left{content:${cssStr(P.footer)};${f}}
  @bottom-center{content:${cssStr('Generated ' + today())};${f}}
  @bottom-right{content:"Page " counter(page) " of " counter(pages);${f}}}`;
}
function fileBase() { const P = S.cfg.report; return ((P.title || 'District report') + (S.cfg.message.period ? ' - ' + S.cfg.message.period : '')).replace(/[\\/:*?"<>|]+/g, '-').trim(); }
function preparePrint() { const r = buildReport(); $('#print-root').innerHTML = r.html; $('#page-style').textContent = pageCSS(r.orient); return r; }
function tableAoa() {
  const D = S.D, P = S.cfg.report; const cols = reportColumns(P.format === 'district' ? 'detailed' : P.format);
  const rnd = (v, k) => { const c = D.colMap[k]; return c.numeric ? (v === null ? null : round(v, c.decimals)) : (v ?? ''); };
  const head = ['Rank', 'District', ...cols.map(c => c.label + (D.colMap[c.key].suffix && D.colMap[c.key].suffix.trim() ? ` (${D.colMap[c.key].suffix.trim()})` : ''))];
  const rows = D.rank.visible.map(e => [e.rank ?? '', e.row.name || 'Row ' + e.row.src, ...cols.map(c => rnd(cellVal(e.row, c.key), c.key))]);
  if (D.total && P.showTotal) rows.push(['', D.total.name, ...cols.map(c => rnd(cellVal(D.total, c.key), c.key))]);
  return [head, ...rows];
}
function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); }

/* ======================================================================
   File loading and saved setups
   ====================================================================== */
async function loadFile(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) { toast('Choose an .xlsx, .xls or .csv file.'); return; }
  let wb;
  try {
    if (ext === 'csv') wb = XLSX.read(await file.text(), { type: 'string', raw: true });
    else wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  } catch (e) { toast('This file could not be read: ' + e.message, 5000); return; }
  S.sheets = {}; S.sheetNames = wb.SheetNames.slice();
  for (const n of S.sheetNames) S.sheets[n] = sheetToGrid(wb.Sheets[n]);
  S.fileName = file.name; S.fileSig = [file.name, file.size, file.lastModified].join('|');
  S.sheetName = S.cfg.sheetName && S.sheets[S.cfg.sheetName] ? S.cfg.sheetName : (S.sheetNames.find(n => S.sheets[n].aoa.length) || S.sheetNames[0]);
  S.cfg.sheetName = S.sheetName;
  process();
  S.compat = buildCompat();
  S.step = 2; render(); window.scrollTo(0, 0);
}
function configRefs(cfg) {
  const refs = new Set(); const add = k => { if (k && !String(k).startsWith('calc:')) refs.add(k); };
  cfg.selected.forEach(s => add(s.key));
  cfg.calcs.forEach(c => calcRefs(c).forEach(add));
  add(cfg.ranking.by); add(cfg.ranking.by2);
  cfg.checks.forEach(ch => { add(ch.total); ch.parts.forEach(add); });
  if (cfg.district.mode === 'column') add(cfg.district.column);
  if (cfg.district.mode === 'key') add(cfg.district.keyColumn);
  return refs;
}
function buildCompat() {
  const refs = configRefs(S.cfg); if (!refs.size || !S.D || S.D.empty) return null;
  const have = new Set(S.D.cols.filter(c => c.kind === 'orig').map(c => c.key));
  const missing = [...refs].filter(k => !have.has(k)).map(k => ({ key: k, label: S.cfg.keyLabels[k] || k }));
  return { total: refs.size, missing, map: {} };
}
function replaceKey(o, n) {
  const c = S.cfg;
  c.selected.forEach(s => { if (s.key === o) s.key = n; });
  const seen = new Set(); c.selected = c.selected.filter(s => !seen.has(s.key) && seen.add(s.key));
  c.calcs.forEach(k => { if (k.a === o) k.a = n; if (k.b === o) k.b = n; k.cols = (k.cols || []).map(x => x === o ? n : x); if (k.expr) k.expr = k.expr.split('{{' + o + '}}').join('{{' + n + '}}'); });
  if (c.ranking.by === o) c.ranking.by = n; if (c.ranking.by2 === o) c.ranking.by2 = n;
  c.checks.forEach(ch => { if (ch.total === o) ch.total = n; ch.parts = ch.parts.map(x => x === o ? n : x); });
  if (c.district.column === o) c.district.column = n; if (c.district.keyColumn === o) c.district.keyColumn = n;
  if (c.aliases[o] && !c.aliases[n]) c.aliases[n] = c.aliases[o];
}
function dropKey(o) {
  const c = S.cfg;
  c.selected = c.selected.filter(s => s.key !== o);
  if (c.ranking.by === o) c.ranking.by = null; if (c.ranking.by2 === o) c.ranking.by2 = null;
  c.checks = c.checks.filter(ch => ch.total !== o && !ch.parts.includes(o));
  if (c.district.column === o) c.district.column = null; if (c.district.keyColumn === o) c.district.keyColumn = null;
}
function snapshotCfg() { const c = clone(S.cfg); c.district.verifiedSig = null; return c; }

/* ======================================================================
   Events
   ====================================================================== */
const ACT = {
  go(el) { const i = +el.dataset.step; if (i > 1 && !hasData()) return; S.step = i; render(); window.scrollTo(0, 0); },
  pickFile() { $('#fileInput').click(); },
  autoHeaders() { S.cfg.headerStart = null; S.cfg.headerRows = null; render(); },
  applyCompat() {
    const C = S.compat; const un = C.missing.filter(m => !C.map[m.key]);
    if (un.length) { toast(`Choose a column or “drop” for ${un.length} item(s) first.`); return; }
    for (const m of C.missing) { if (C.map[m.key] === '__drop') dropKey(m.key); else replaceKey(m.key, C.map[m.key]); }
    S.compat = null; render(); toast('Saved setup applied to this file.');
  },
  freshCompat() {
    const keep = { message: S.cfg.message, report: S.cfg.report, aliases: S.cfg.aliases, keyLabels: S.cfg.keyLabels, district: { ...S.cfg.district }, summaryMode: S.cfg.summaryMode, naming: S.cfg.naming };
    S.cfg = deepMerge(defaultConfig(), keep); S.cfg.district.column = null; S.cfg.district.keyColumn = null;
    S.cfg.sheetName = S.sheetName; S.compat = null; S.activeProfile = ''; LS.set(K_ACTIVE, ''); render();
  },
  addCheck() {
    const total = $('#chkTotal').value; const parts = $$('.chkPart:checked').map(x => x.value).filter(k => k !== total);
    if (!total || parts.length < 2) { toast('Choose a total and at least two parts.'); return; }
    S.cfg.checks.push({ id: uid(), total, parts }); S.ui.checkOpen = false; render();
  },
  delCheck(el) { S.cfg.checks.splice(+el.dataset.i, 1); render(); },
  selAll() { const have = new Set(S.cfg.selected.map(s => s.key)); S.D.cols.forEach(c => { if (!have.has(c.key)) S.cfg.selected.push({ key: c.key, label: '' }); }); render(); },
  selNone() { S.cfg.selected = []; render(); },
  mv(el) { const i = +el.dataset.i, j = i + +el.dataset.d; const a = S.cfg.selected; if (j < 0 || j >= a.length) return; [a[i], a[j]] = [a[j], a[i]]; render(); },
  unpick(el) { S.cfg.selected.splice(+el.dataset.i, 1); render(); },
  insCol() { const sel = $('#exprIns'); const box = $('#exprBox'); const c = S.D.colMap[sel.value]; if (!c) return; const t = '[' + c.name + ']'; const p = box.selectionStart ?? box.value.length; box.value = box.value.slice(0, p) + t + box.value.slice(box.selectionEnd ?? p); S.draft.exprText = box.value; box.focus(); box.selectionStart = box.selectionEnd = p + t.length; refresh(); },
  saveCalc() {
    const d = S.draft, T = CALC_TYPES[d.type];
    if (!d.name.trim()) { toast('Give the indicator a name.'); return; }
    if (S.cfg.calcs.some(c => c.id !== d.id && c.name.trim().toLowerCase() === d.name.trim().toLowerCase())) { toast('Another calculation already has this name.'); return; }
    const c = clone(d); delete c.exprText; c.name = c.name.trim();
    if (T.custom) { const r = exprToStored(d.exprText || '', draftAllowedCols()); if (r.errs.length) { toast('Unknown column: [' + r.errs[0] + ']'); return; } try { parseExpr(r.stored); } catch (e) { toast(e.message); return; } c.expr = r.stored; }
    const refs = calcRefs(c); if (!refs.length || refs.some(k => !k)) { toast('Choose the columns first.'); return; }
    if (T.multi && refs.length < 2) { toast('Choose at least two columns.'); return; }
    c.multiplier = c.multiplier === '' || c.multiplier === null ? 1 : +c.multiplier; c.decimals = Math.max(0, Math.min(6, +c.decimals || 0));
    if (d.id) { const i = S.cfg.calcs.findIndex(x => x.id === d.id); S.cfg.calcs[i] = c; toast('Calculation updated.'); }
    else { c.id = uid(); S.cfg.calcs.push(c); S.cfg.selected.push({ key: 'calc:' + c.id, label: '' }); toast('Calculation added.'); }
    S.draft = newCalcDraft(); render();
  },
  editCalc(el) { const c = S.cfg.calcs.find(x => x.id === el.dataset.id); S.draft = clone(c); S.draft.exprText = c.type === 'custom' ? exprToDisplay(c.expr) : ''; render(); $('.builder').scrollIntoView({ behavior: 'smooth' }); },
  cancelCalc() { S.draft = newCalcDraft(); render(); },
  delCalc(el) {
    const id = el.dataset.id, key = 'calc:' + id; const c = S.cfg.calcs.find(x => x.id === id);
    const users = S.cfg.calcs.filter(x => x.id !== id && calcRefs(x).includes(key));
    if (users.length) { toast(`“${users[0].name}” uses this calculation. Edit or delete that one first.`, 4000); return; }
    if (!confirm(`Delete the calculation “${c.name}”?`)) return;
    S.cfg.calcs = S.cfg.calcs.filter(x => x.id !== id); S.cfg.selected = S.cfg.selected.filter(s => s.key !== key);
    if (S.cfg.ranking.by === key) S.cfg.ranking.by = null; if (S.cfg.ranking.by2 === key) S.cfg.ranking.by2 = null;
    if (S.draft.id === id) S.draft = newCalcDraft(); render();
  },
  exAll() { S.cfg.ranking.excluded = []; render(); },
  exNone() { S.cfg.ranking.excluded = S.D.dataRows.map(r => r.name); render(); },
  saveTpl() { const M = S.cfg.message; const n = prompt('Name for this message in the library:', (M.subject || M.body || '').replace(/\s+/g, ' ').slice(0, 40)); if (!n) return; const t = LS.get(K_TPL, {}); t[n.trim()] = clone(M); LS.set(K_TPL, t); render(); toast('Saved to the message library.'); },
  uploadSamples() { $('#sampleInput').click(); },
  pasteSample() { S.ui.pasteOpen = true; const n = prompt('Name for this sample (e.g. “Aadhaar pending – reminder”):'); if (!n) return; const text = prompt('Paste the full message text here (line breaks are kept if you paste from a document):'); if (!text) return; addSamples([{ name: n.trim(), ...splitSample(text.trim()), mode: /\{\w+\}/.test(text) ? 'data' : 'exact' }]); render(); toast('Sample saved.'); },
  useSample(el) { const it = LS.get(K_TPL, {})[el.dataset.n]; if (!it) return; const M = S.cfg.message; if ((M.body || M.subject) && !confirm('Replace the current message with “' + el.dataset.n + '”?')) return; S.cfg.message = { ...M, ...clone(it), position: it.position || M.position, period: M.period || it.period || '', source: el.dataset.n }; render(); toast('Message loaded. Edit it if needed.'); },
  toggleSampleMode(el) { const t = LS.get(K_TPL, {}); const it = t[el.dataset.n]; if (!it) return; it.mode = it.mode === 'data' ? 'exact' : 'data'; LS.set(K_TPL, t); refresh(); },
  delSample(el) { if (!confirm('Delete “' + el.dataset.n + '” from the library?')) return; const t = LS.get(K_TPL, {}); delete t[el.dataset.n]; LS.set(K_TPL, t); refresh(); },
  insPh(el) { const box = $('#msgBody'); const t = '{' + el.dataset.ph + '}'; const p = box.selectionStart ?? box.value.length; box.value = box.value.slice(0, p) + t + box.value.slice(box.selectionEnd ?? p); S.cfg.message.body = box.value; box.focus(); box.selectionStart = box.selectionEnd = p + t.length; refresh(); },
  draftMsg() { const M = S.cfg.message; if (M.body && !confirm('Rewrite the middle of the message from the data? The first and last paragraphs are kept.')) return; M.body = draftFromData(); M.mode = 'data'; render(); toast('Draft written. Edit it freely; figures update with each new file.'); },
  clearMsg() { if (!confirm('Clear the message fields?')) return; const pos = S.cfg.message.position; S.cfg.message = { ...defaultConfig().message, position: pos }; render(); },
  print() {
    if (S.D.cols.some(c => c.broken)) { toast('Fix the broken calculation in step 4 first.'); return; }
    preparePrint(); const old = document.title; document.title = fileBase();
    setTimeout(() => { window.print(); setTimeout(() => document.title = old, 500); }, 60);
  },
  exportXlsx() { const ws = XLSX.utils.aoa_to_sheet(tableAoa()); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report'); XLSX.writeFile(wb, fileBase() + '.xlsx'); },
  exportCsv() { const csv = tableAoa().map(r => r.map(v => { const s = v === null ? '' : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',')).join('\r\n'); download(new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' }), fileBase() + '.csv'); },
  pickLogo() { $('#logoInput').click(); },
  rmLogo() { S.cfg.report.logo = ''; render(); },
  mapFile() { $('#mapInput').click(); },
  saveProfile() { if (!S.activeProfile) return ACT.saveProfileAs(); const p = LS.get(K_PROF, {}); p[S.activeProfile] = snapshotCfg(); LS.set(K_PROF, p) ? toast(`Setup “${S.activeProfile}” saved.`) : toast('Could not save: browser storage is unavailable. Use Export instead.', 4000); },
  saveProfileAs() { const n = prompt('Name for this setup (e.g. "PM SHRI DBT – Aadhaar review"):', S.activeProfile || ''); if (!n) return; S.activeProfile = n.trim(); LS.set(K_ACTIVE, S.activeProfile); ACT.saveProfile(); renderTop(); },
  deleteProfile() { if (!confirm(`Delete the saved setup “${S.activeProfile}”? This does not change the current screen.`)) return; const p = LS.get(K_PROF, {}); delete p[S.activeProfile]; LS.set(K_PROF, p); S.activeProfile = ''; LS.set(K_ACTIVE, ''); renderTop(); },
  exportProfile() { const data = { app: 'district-performance-reports', version: 1, name: S.activeProfile || 'setup', config: snapshotCfg(), templates: LS.get(K_TPL, {}) }; download(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), (S.activeProfile || 'district-report-setup').replace(/[^\w\- ]+/g, '') + '.json'); },
  importProfile() { $('#cfgInput').click(); }
};
function applyLoadedConfig(cfg, name) {
  S.cfg = deepMerge(defaultConfig(), cfg); S.activeProfile = name || ''; LS.set(K_ACTIVE, S.activeProfile);
  if (S.sheets) { if (S.cfg.sheetName && S.sheets[S.cfg.sheetName]) S.sheetName = S.cfg.sheetName; else S.cfg.sheetName = S.sheetName; process(); S.compat = buildCompat(); if (S.compat) S.step = 2; }
  S.draft = newCalcDraft(); render(); toast(`Setup “${name}” loaded${S.compat ? ' — confirm the columns in step 2' : ''}.`);
}
const CHG = {
  loadProfile(el) { const p = LS.get(K_PROF, {}); if (el.value && p[el.value]) applyLoadedConfig(p[el.value], el.value); },
  sheet(el) { S.sheetName = el.value; S.cfg.sheetName = el.value; process(); S.compat = buildCompat(); render(); },
  verifyList(el) { S.cfg.district.verifiedSig = el.checked ? S.fileSig : null; render(); },
  draftType(el) { const prev = S.draft; S.draft = { ...newCalcDraft(el.value), id: prev.id, name: prev.name, a: prev.a, b: prev.b, cols: prev.cols, exprText: prev.exprText }; render(); },
  loadTpl() {}
};
function readVal(el) {
  if (el.dataset.bool !== undefined) return el.checked;
  if (el.dataset.lines !== undefined) return el.value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (el.dataset.num !== undefined) return el.value === '' ? null : +el.value;
  return el.value;
}
function onField(e) {
  const el = e.target;
  if (el.dataset.bind !== undefined) {
    if (el.type === 'radio' && !el.checked) return;
    const before = JSON.stringify(S.cfg.district);
    setPath(S.cfg, el.dataset.bind, readVal(el));
    if (el.dataset.bind === 'district.list' && JSON.stringify(S.cfg.district) !== before) S.cfg.district.verifiedSig = null;
    const structural = el.dataset.full !== undefined || el.type === 'radio' || el.tagName === 'SELECT' || ['headerStart', 'headerRows'].includes(el.dataset.bind);
    if (structural && e.type === 'change') render(); else if (!structural) refresh();
    return;
  }
  if (el.dataset.sbind !== undefined) {
    let v = readVal(el); setPath(S, el.dataset.sbind, v);
    if (el.tagName === 'SELECT' && e.type === 'change') render(); else if (el.tagName !== 'SELECT') refresh();
    return;
  }
  if (el.dataset.alias !== undefined) { const v = el.value.trim(); if (v) S.cfg.aliases[el.dataset.alias] = v; else delete S.cfg.aliases[el.dataset.alias]; refresh(); return; }
  if (el.dataset.label !== undefined) { S.cfg.selected[+el.dataset.label].label = el.value; refresh(); return; }
  if (e.type !== 'change') return;
  if (el.dataset.pick !== undefined) { const k = el.dataset.pick; if (el.checked) { if (!S.cfg.selected.some(s => s.key === k)) S.cfg.selected.push({ key: k, label: '' }); } else S.cfg.selected = S.cfg.selected.filter(s => s.key !== k); render(); return; }
  if (el.dataset.draftcol !== undefined) { const k = el.dataset.draftcol; S.draft.cols = el.checked ? [...new Set([...S.draft.cols, k])] : S.draft.cols.filter(x => x !== k); refresh(); return; }
  if (el.dataset.incl !== undefined) { const n = el.dataset.incl; const ex = S.cfg.ranking.excluded.filter(x => x.toLowerCase() !== n.toLowerCase()); if (!el.checked) ex.push(n); S.cfg.ranking.excluded = ex; refresh(); return; }
  if (el.dataset.compat !== undefined) { S.compat.map[el.dataset.compat] = el.value; return; }
  if (el.dataset.actChange) { CHG[el.dataset.actChange](el); }
}
document.addEventListener('input', e => { const el = e.target; if (el.dataset.filter !== undefined) { const q = el.value.trim().toLowerCase(); $$(el.dataset.filter).forEach(r => { const t = (r.dataset.name || r.textContent).toLowerCase(); r.style.display = !q || t.includes(q) ? '' : 'none'; }); return; } if (el.type === 'checkbox' || el.type === 'radio' || el.tagName === 'SELECT') return; onField(e); });
document.addEventListener('change', e => { const el = e.target; if (el.type === 'checkbox' || el.type === 'radio' || el.tagName === 'SELECT' || el.dataset.bind !== undefined) onField(e); });
document.addEventListener('click', e => { const el = e.target.closest('[data-act]'); if (!el || el.disabled || el.getAttribute('aria-disabled') === 'true') return; const f = ACT[el.dataset.act]; if (f) { e.preventDefault(); f(el); } });
document.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && e.target.id === 'drop') { e.preventDefault(); ACT.pickFile(); } });
document.addEventListener('toggle', e => { const el = e.target; if (el.dataset && el.dataset.actToggle) S.ui[el.dataset.actToggle] = el.open; }, true);
$('#fileInput').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) loadFile(f); });
$('#cfgInput').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { const data = JSON.parse(await f.text()); if (!data.config) throw new Error('not a setup file'); const name = data.name || f.name.replace(/\.json$/i, ''); const p = LS.get(K_PROF, {}); p[name] = data.config; LS.set(K_PROF, p); if (data.templates) LS.set(K_TPL, { ...LS.get(K_TPL, {}), ...data.templates }); applyLoadedConfig(data.config, name); }
  catch (err) { toast('This is not a valid setup file: ' + err.message, 4000); }
});
$('#sampleInput').addEventListener('change', async e => {
  const files = [...e.target.files]; e.target.value = ''; if (!files.length) return; const ok = [], bad = [];
  for (const f of files) { try { const text = await readSampleFile(f); if (!text) throw new Error('empty'); ok.push({ name: f.name.replace(/\.[^.]+$/, ''), ...splitSample(text), mode: /\{\w+\}/.test(text) ? 'data' : 'exact' }); } catch (err) { bad.push(f.name); } }
  addSamples(ok); render(); toast(`${ok.length} sample(s) added${bad.length ? '; could not read ' + bad.join(', ') : ''}.`, 4000);
});
$('#logoInput').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; if (f.size > 600000) { toast('Use a logo smaller than 600 KB.'); return; } const r = new FileReader(); r.onload = () => { S.cfg.report.logo = r.result; render(); }; r.readAsDataURL(f); });
$('#mapInput').addEventListener('change', async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const wb = /\.csv$|\.txt$/i.test(f.name) ? XLSX.read(await f.text(), { type: 'string', raw: true }) : XLSX.read(await f.arrayBuffer(), { type: 'array' });
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    // take the first column that contains mostly text
    const ncol = Math.max(...aoa.map(r => r.length)); let best = 0, bestScore = -1;
    for (let c = 0; c < ncol; c++) { const sc = aoa.filter(r => r[c] !== '' && isNaN(+r[c])).length; if (sc > bestScore) { best = c; bestScore = sc; } }
    let names = aoa.map(r => String(r[best] ?? '').trim()).filter(Boolean);
    if (names.length === S.D.dataRows.length + 1 && /district|जनपद|जिला|name/i.test(names[0])) names = names.slice(1);
    S.cfg.district.list = names; S.cfg.district.mode = 'list'; S.cfg.district.verifiedSig = null; render(); toast(`${names.length} names loaded. Check the order, then confirm.`);
  } catch (err) { toast('Could not read that file: ' + err.message); }
});
const drop = () => $('#drop');
document.addEventListener('dragover', e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); const d = drop(); if (d) d.classList.add('over'); } });
document.addEventListener('dragleave', e => { const d = drop(); if (d && e.target === d) d.classList.remove('over'); });
document.addEventListener('drop', e => { if (e.dataTransfer && e.dataTransfer.files.length) { e.preventDefault(); const d = drop(); if (d) d.classList.remove('over'); loadFile(e.dataTransfer.files[0]); } });
function wireDnD() {
  const list = $('#selList'); if (!list) return; let from = null;
  $$('.it[draggable]', list).forEach(it => {
    it.addEventListener('dragstart', e => { from = +it.dataset.idx; it.classList.add('drag'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'x'); });
    it.addEventListener('dragend', () => it.classList.remove('drag'));
    it.addEventListener('dragover', e => { if (from === null) return; e.preventDefault(); e.stopPropagation(); it.classList.add('dropline'); });
    it.addEventListener('dragleave', () => it.classList.remove('dropline'));
    it.addEventListener('drop', e => { e.preventDefault(); e.stopPropagation(); const to = +it.dataset.idx; if (from === null || from === to) return; const a = S.cfg.selected; const [m] = a.splice(from, 1); a.splice(to, 0, m); from = null; render(); });
  });
}
window.addEventListener('afterprint', () => { $('#print-root').innerHTML = ''; });
window.addEventListener('resize', () => postRender());

/* test / automation hook */
window.__app = { S, process, render, refresh, buildReport, preparePrint, tableAoa, loadFile, ACT, CHG };
render();
