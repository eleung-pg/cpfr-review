// Pure CPFR review logic. No Office.js here, so it runs in the task pane and in Node tests.
// Inputs are plain 2D arrays read from A1 of a prepped tab (row i of the sheet = values[i-1]).
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CPFR = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const COL_B = 1, COL_C = 2, COL_D = 3, COL_E = 4, COL_F = 5, COL_G = 6, COL_H = 7, COL_J = 9, COL_K = 10;
  const YELLOW = '#FFFF00';

  const isNum = v => typeof v === 'number' && isFinite(v);
  const num = v => (isNum(v) ? v : 0);
  const isBlank = v => v === '' || v === null || v === undefined;

  function colLetter(i) { // 0-based index -> A, B, ... AA
    let s = '';
    for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }
  const q = name => "'" + String(name).replace(/'/g, "''") + "'";
  const serialToDate = s => new Date(Math.round((s - 25569) * 86400000));
  const fmtMD = s => { const d = serialToDate(s); return (d.getUTCMonth() + 1) + '/' + d.getUTCDate(); };

  // ---- Layout detection: blocks start at every "Beginning Stock" in column J ------------------
  function detectLayout(values) {
    const starts = [];
    for (let i = 1; i < values.length; i++) if (String(values[i][COL_J]).trim() === 'Beginning Stock') starts.push(i + 1);
    if (!starts.length) throw new Error('No "Beginning Stock" rows found in column J - is this a prepped CPFR tab?');

    const warnings = [];
    let n;
    if (starts.length > 1) {
      const gaps = {};
      for (let i = 1; i < starts.length; i++) { const g = starts[i] - starts[i - 1]; gaps[g] = (gaps[g] || 0) + 1; }
      const keys = Object.keys(gaps).sort((a, b) => gaps[b] - gaps[a]);
      n = +keys[0];
      if (keys.length > 1) warnings.push('Block spacing is irregular (gaps: ' + keys.join(', ') + '); using ' + n + '.');
    } else {
      n = 0; // labelled rows, then one separator row
      while (!isBlank((values[starts[0] - 1 + n] || [])[COL_J])) n++;
      n += 1;
    }

    const r0 = starts[0], jOff = {}, fOff = {};
    for (let o = 0; o < n; o++) {
      const row = values[r0 - 1 + o];
      if (!row) break;
      const j = String(row[COL_J]).trim();
      if (j && !(j in jOff)) jOff[j] = o;
      const f = String(row[COL_F]).trim();
      if (f === 'Average Demand') fOff.avg = o;
      else if (f === 'Average 4 Wks Sell Through') fOff.sell = o;
      else if (f.startsWith('% Demand @')) fOff.trend = o;
    }
    const need = { 'Beginning Stock': 'begin', 'Dependent Demand': 'dd', 'Total Outflow': 'out', 'PR Override Input': 'pr', 'SCP PO Schedule': 'sched', 'True Demand': 'trueDemand' };
    const off = {};
    for (const label of Object.keys(need)) {
      if (!(label in jOff)) throw new Error('Block is missing the "' + label + '" row in column J.');
      off[need[label]] = jOff[label];
    }
    for (const k of ['avg', 'sell', 'trend']) if (fOff[k] === undefined) throw new Error('Could not find the ' + k + ' metric label in column F of the first block.');
    Object.assign(off, fOff);

    let lastCol = COL_K - 1;
    for (let c = COL_K; c < (values[0] || []).length; c++) { if (isNum(values[0][c])) lastCol = c; else break; }
    if (lastCol < COL_K) throw new Error('No week dates found in row 1 from column K.');

    const blocks = starts.map((r, idx) => {
      const row = values[r - 1];
      return { idx, row: r, tmo: row[COL_B], item: row[COL_C], desc: row[COL_D], dc: row[COL_E] };
    });
    return { n, off, blocks, dates: values[0], firstCol: COL_K, lastCol, weeks: lastCol - COL_K + 1, lastRow: values.length, warnings };
  }

  function blockAt(layout, row) { // sheet row -> {block, offset} or null
    const bl = layout.blocks;
    let lo = 0, hi = bl.length - 1, f = -1;
    while (lo <= hi) { const m = (lo + hi) >> 1; if (bl[m].row <= row) { f = m; lo = m + 1; } else hi = m - 1; }
    if (f < 0) return null;
    const offset = row - bl[f].row;
    return offset < layout.n ? { block: bl[f], offset } : null;
  }

  const cell = (values, block, off, col) => (values[block.row - 1 + off] || [])[col];

  // ---- Sanity: hand-typed SCP PO Schedule cells ------------------------------------------------
  function findHardcodedPOs(layout, values, formulas) {
    const out = [];
    for (const b of layout.blocks) {
      const r = b.row + layout.off.sched;
      for (let c = layout.firstCol; c <= layout.lastCol; c++) {
        const f = formulas[r - 1][c], v = values[r - 1][c];
        const isFormula = typeof f === 'string' && f.startsWith('=');
        if (!isFormula && !isBlank(v)) out.push({ addr: colLetter(c) + r, value: v, tmo: b.tmo, item: b.item, dc: b.dc });
      }
    }
    return out;
  }

  // ---- Review 1: yellow PR Override comments ----------------------------------------------------
  // fills: 2D array of colour strings for rows 2..lastRow, columns firstCol..lastCol.
  function classify(text) {
    const t = String(text);
    if (/new\s*po/i.test(t)) return 'NEW PO';
    if (/transfer/i.test(t)) return 'DC transfer';
    if (/new\s*cp/i.test(t)) return 'NEW CP';
    return 'Other';
  }

  function yellowComments(layout, values, fills) {
    const perBlock = new Map(); // block idx -> Map(col -> text)
    const emptyYellow = [], outsidePR = [];
    for (let r = 2; r <= layout.lastRow; r++) {
      const frow = fills[r - 2];
      if (!frow) continue;
      for (let c = layout.firstCol; c <= layout.lastCol; c++) {
        if (String(frow[c - layout.firstCol] || '').toUpperCase() !== YELLOW) continue;
        const v = values[r - 1][c];
        const at = blockAt(layout, r);
        const addr = colLetter(c) + r;
        if (isBlank(v)) { emptyYellow.push(addr); continue; }
        if (!at || at.offset !== layout.off.pr) { outsidePR.push({ addr, value: v }); continue; }
        if (!perBlock.has(at.block.idx)) perBlock.set(at.block.idx, new Map());
        perBlock.get(at.block.idx).set(c, String(v).trim());
      }
    }
    // Merge blocks of the same item with identical (week -> comment) sets, e.g. DC80 / DC77.
    const groups = new Map();
    for (const idx of [...perBlock.keys()].sort((a, b) => a - b)) {
      const b = layout.blocks[idx];
      const cols = [...perBlock.get(idx).entries()].sort((x, y) => x[0] - y[0]);
      const key = String(b.item) + '|' + JSON.stringify(cols);
      if (!groups.has(key)) groups.set(key, { blocks: [], cols });
      groups.get(key).blocks.push(b);
    }
    const rows = [...groups.values()].map(g => {
      const first = g.blocks[0];
      return {
        tmo: first.tmo, item: first.item, desc: first.desc,
        blocks: g.blocks, dcs: g.blocks.map(b => b.dc),
        cols: g.cols.map(([c, text]) => ({ col: c, text })),
        cells: g.blocks.flatMap(b => g.cols.map(([c]) => colLetter(c) + (b.row + layout.off.pr))),
      };
    }).sort((a, b) => a.blocks[0].idx - b.blocks[0].idx);

    const types = {}; let questions = [];
    for (const row of rows) for (const { col, text } of row.cols) {
      const t = classify(text); types[t] = (types[t] || 0) + 1;
      if (text.endsWith('?')) questions.push({ tmo: row.tmo, item: row.item, dc: row.dcs.join(' / '), week: col, text });
    }
    return { rows, types, questions, emptyYellow, outsidePR, commentCount: rows.reduce((s, r) => s + r.cols.length, 0) };
  }

  // ---- Review 2: WoW swings ----------------------------------------------------------------------
  function tierOf(wow, big, watch) { return Math.abs(wow) >= big ? 'Big swing' : Math.abs(wow) >= watch ? 'Watch' : ''; }

  function demandSwings(layout, values, opts, commentItems) {
    const o = Object.assign({ big: 0.2, watch: 0.1, minUnits: 800, smallBase: 100 }, opts);
    const all = [], notComparable = [];
    for (const b of layout.blocks) {
      const thisWk = cell(values, b, layout.off.avg, COL_G);
      const lastWk = cell(values, b, layout.off.avg, COL_H);
      const wow = cell(values, b, layout.off.sell, COL_H);
      if (!isNum(thisWk) || !isNum(lastWk) || !isNum(wow)) {
        notComparable.push({ block: b, reason: String(isNum(thisWk) ? wow : thisWk) + (isNum(thisWk) && !isNum(lastWk) ? ' (no prev)' : '') });
        continue;
      }
      all.push({ block: b, thisWk, lastWk, unit: thisWk - lastWk, wow, tier: tierOf(wow, o.big, o.watch) });
    }
    const byItem = new Map();
    for (const r of all) { const k = String(r.block.item); if (!byItem.has(k)) byItem.set(k, []); byItem.get(k).push(r); }
    for (const r of all) {
      const notes = [];
      if (!r.tier) notes.push('Under ' + Math.round(o.watch * 100) + '% but large unit ' + (r.unit < 0 ? 'drop' : 'rise') + ' (' + Math.round(r.unit).toLocaleString('en-US') + '/wk)');
      const sibs = byItem.get(String(r.block.item));
      if (sibs.length > 1) {
        const sig = sibs.filter(s => Math.abs(s.wow) >= o.watch);
        const ups = sig.filter(s => s.wow > 0).length, downs = sig.filter(s => s.wow < 0).length;
        if (ups > 1 && !downs) notes.push('Both DCs up');
        else if (downs > 1 && !ups) notes.push('Both DCs cut - looks like a real forecast change');
        else if (ups && downs) notes.push('DCs moving opposite - possible shift between DCs');
      }
      if (Math.max(r.thisWk, r.lastWk) < o.smallBase) notes.push('Small base');
      if (commentItems && commentItems.has(String(r.block.item))) notes.push('Has yellow PR Override comments');
      r.note = notes.join('; ');
    }
    const rows = all.filter(r => r.tier || Math.abs(r.unit) >= o.minUnits)
      .sort((a, b) => Math.abs(b.wow) - Math.abs(a.wow));
    return { rows, notComparable, params: o };
  }

  // ---- Review 3: stock-out risk ------------------------------------------------------------------
  const CAT_ORDER = ['Stock-out only at trend', 'Earlier at trend', 'Short at current forecast', 'Forecast above trend', '', 'Skipped'];

  function stockRisk(layout, values, opts) {
    const gapThreshold = opts && opts.gapThreshold !== undefined ? opts.gapThreshold : -0.25;
    const off = layout.off, dates = values[0];
    const rows = layout.blocks.map(b => {
      const forecast = cell(values, b, off.avg, COL_G), trend = cell(values, b, off.trend, COL_G);
      const valid = isNum(forecast) && isNum(trend) && forecast > 0 && trend > 0;
      const r = { block: b, forecast, trend, valid, begin: num(cell(values, b, off.begin, COL_K)), gap: valid ? trend / forecast - 1 : null };
      if (valid) {
        let cur = r.begin, tr = r.begin;
        r.curPath = []; r.trendPath = []; r.firstCur = null; r.firstTrend = null;
        for (let c = layout.firstCol; c <= layout.lastCol; c++) {
          const dd = num(cell(values, b, off.dd, c)), out = num(cell(values, b, off.out, c)), po = num(cell(values, b, off.sched, c));
          cur = cur - Math.max(dd, out) + po;
          tr = tr - Math.max(dd, trend) + po;
          r.curPath.push(cur); r.trendPath.push(tr);
          if (cur < 0 && r.firstCur === null) r.firstCur = dates[c];
          if (tr < 0 && r.firstTrend === null) r.firstTrend = dates[c];
        }
        r.lowest = Math.min(...r.trendPath);
      }
      const { firstCur: fc, firstTrend: ft } = r;
      r.category = !valid ? 'Skipped'
        : (ft !== null && fc === null) ? 'Stock-out only at trend'
        : (ft !== null && fc !== null && ft < fc) ? 'Earlier at trend'
        : fc !== null ? 'Short at current forecast'
        : r.gap <= gapThreshold ? 'Forecast above trend' : '';
      return r;
    });
    const counts = {};
    for (const r of rows) counts[r.category || 'OK'] = (counts[r.category || 'OK'] || 0) + 1;
    const sorted = rows.slice().sort((a, b) => {
      const ca = CAT_ORDER.indexOf(a.category), cb = CAT_ORDER.indexOf(b.category);
      if (ca !== cb) return ca - cb;
      const da = a.firstTrend === null ? Infinity : a.firstTrend, db = b.firstTrend === null ? Infinity : b.firstTrend;
      return da - db || a.block.idx - b.block.idx;
    });
    return { rows: sorted, counts, gapThreshold };
  }

  return { detectLayout, blockAt, findHardcodedPOs, yellowComments, classify, demandSwings, tierOf, stockRisk, colLetter, q, fmtMD, isNum, serialToDate, CAT_ORDER, YELLOW };
});
