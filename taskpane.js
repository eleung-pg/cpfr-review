/* global Office, Excel, CPFR */
// Task-pane glue: read a prepped CPFR tab, run the pure logic in logic.js, render findings, and
// (optionally) write the three linked review tabs. All formulas link back to the source tab.
(function () {
  'use strict';
  const L = CPFR;
  const $ = id => document.getElementById(id);
  const esc = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = (v, d = 0) => (typeof v === 'number' ? v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }) : esc(v));
  const pct = v => (typeof v === 'number' ? (v * 100).toFixed(1) + '%' : '');
  const REVIEW_SUFFIXES = [' Comments', ' Demand Swings', ' Stock-Out Risk'];
  const ARROW = '→', PM = '±';

  function setStatus(msg, isErr) { const s = $('status'); s.textContent = msg; s.className = isErr ? 'err' : ''; }
  const tabName = (src, suffix) => src.slice(0, 31 - suffix.length) + suffix;

  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) { setStatus('This add-in only runs in Excel.', true); return; }
    loadSheets();
    initAutoOpen();
    $('runAll').onclick = () => run([1, 2, 3]);
    $('run1').onclick = () => run([1]);
    $('run2').onclick = () => run([2]);
    $('run3').onclick = () => run([3]);
  });

  // Saved in the workbook: Excel reopens this pane with the file (needs the Office.AutoShowTaskpaneWithDocument TaskpaneId).
  const AUTO = 'Office.AutoShowTaskpaneWithDocument';
  function initAutoOpen() {
    const settings = Office.context.document.settings, box = $('autoOpen');
    box.checked = settings.get(AUTO) === true;
    box.onchange = () => {
      settings.set(AUTO, box.checked);
      settings.saveAsync(r => setStatus(r.status === Office.AsyncResultStatus.Succeeded
        ? (box.checked ? 'This pane will open with this workbook (save the workbook to keep it).' : 'Auto-open turned off (save the workbook).')
        : 'Could not save the auto-open setting: ' + r.error.message, r.status !== Office.AsyncResultStatus.Succeeded));
    };
  }

  async function loadSheets() {
    try {
      await Excel.run(async ctx => {
        const sheets = ctx.workbook.worksheets; sheets.load('items/name');
        const active = ctx.workbook.worksheets.getActiveWorksheet(); active.load('name');
        await ctx.sync();
        const sel = $('sheet');
        const names = sheets.items.map(s => s.name).filter(n => !REVIEW_SUFFIXES.some(x => n.endsWith(x)) && !/\(Prev\)$/.test(n));
        sel.innerHTML = names.map(n => '<option>' + esc(n) + '</option>').join('');
        if (names.includes(active.name)) sel.value = active.name;
      });
    } catch (e) { setStatus('Could not list sheets: ' + e.message, true); }
  }

  // ---- reading ---------------------------------------------------------------------------------
  async function readSheet(name) {
    return Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(name);
      const ur = ws.getUsedRange(); ur.load('rowIndex,columnIndex,rowCount,columnCount');
      await ctx.sync();
      const rng = ws.getRangeByIndexes(0, 0, ur.rowIndex + ur.rowCount, ur.columnIndex + ur.columnCount);
      rng.load('values,formulas');
      await ctx.sync();
      return { values: rng.values, formulas: rng.formulas };
    });
  }

  async function readFills(name, layout) {
    return Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getItem(name);
      const out = [], total = layout.lastRow - 1, CH = 250;
      for (let s = 0; s < total; s += CH) {
        const cnt = Math.min(CH, total - s);
        const props = ws.getRangeByIndexes(1 + s, layout.firstCol, cnt, layout.weeks).getCellProperties({ format: { fill: { color: true } } });
        await ctx.sync();
        for (const row of props.value) out.push(row.map(c => (c.format && c.format.fill && c.format.fill.color) || ''));
      }
      return out;
    });
  }

  // ---- orchestration -----------------------------------------------------------------------------
  async function run(which) {
    const btns = document.querySelectorAll('button'); btns.forEach(b => (b.disabled = true));
    $('out').innerHTML = '';
    try {
      const src = $('sheet').value;
      if (!src) throw new Error('Pick a tab first.');
      const write = $('write').checked, replace = $('replace').checked;
      const big = +$('big').value / 100, watch = +$('watch').value / 100, minUnits = +$('minUnits').value, gapThreshold = +$('gap').value / 100;
      setStatus('Reading ' + src + '...');
      const { values, formulas } = await readSheet(src);
      const layout = L.detectLayout(values);
      let html = '<h2>Sanity checks</h2><div>' + layout.blocks.length + ' blocks of ' + layout.n + ' rows, ' + layout.weeks +
        ' weeks (' + L.fmtMD(values[0][layout.firstCol]) + ' ' + ARROW + ' ' + L.fmtMD(values[0][layout.lastCol]) + '). Read A1:' + L.colLetter(values[0].length - 1) + values.length + '.</div>';
      layout.warnings.forEach(w => (html += '<div class="warn">' + esc(w) + '</div>'));
      const hard = L.findHardcodedPOs(layout, values, formulas);
      html += hard.length
        ? '<div class="warn">' + hard.length + ' hand-typed SCP PO Schedule cell(s) (included in the projections, not reverted): ' + hard.map(h => esc(h.addr) + '=' + fmt(h.value)).join(', ') + '</div>'
        : '<div class="ok">No hand-typed SCP PO Schedule cells.</div>';
      const written = [];

      let yc = null;
      if (which.includes(1) || which.includes(2)) {
        setStatus('Reading fill colours (yellow comments)...');
        yc = L.yellowComments(layout, values, await readFills(src, layout));
      }
      if (which.includes(1)) {
        html += renderComments(layout, yc);
        if (write) written.push(await writeComments(src, layout, yc, replace));
      }
      if (which.includes(2)) {
        const sw = L.demandSwings(layout, values, { big, watch, minUnits }, new Set(yc.rows.map(r => String(r.item))));
        html += renderSwings(sw);
        if (write) written.push(await writeSwings(src, layout, sw, replace));
      }
      if (which.includes(3)) {
        const sr = L.stockRisk(layout, values, { gapThreshold });
        html += renderRisk(layout, sr);
        if (write) written.push(await writeRisk(src, layout, sr, replace));
      }
      $('out').innerHTML = html;
      const bad = written.filter(w => w.errors);
      setStatus((write ? 'Done. Wrote ' + written.map(w => '"' + w.name + '" (' + w.rows + ' rows, ' + w.errors + ' errors)').join(', ') + '.' : 'Done (read-only - no tabs written).') +
        (bad.length ? '\nCheck the tabs above for error cells.' : ''), bad.length > 0);
    } catch (e) {
      setStatus(e.message || String(e), true);
    } finally { btns.forEach(b => (b.disabled = false)); }
  }

  // ---- rendering ---------------------------------------------------------------------------------
  function table(head, rows, numCols) {
    return '<table><tr>' + head.map(h => '<th>' + esc(h) + '</th>').join('') + '</tr>' +
      rows.map(r => '<tr>' + r.map((c, i) => '<td' + ((numCols || []).includes(i) ? ' class="n"' : '') + '>' + c + '</td>').join('') + '</tr>').join('') + '</table>';
  }

  function renderComments(layout, yc) {
    let h = '<h2>1 · Yellow PR Override comments</h2><div>' + yc.commentCount + ' comments in ' + yc.rows.length + ' rows. ' +
      Object.keys(yc.types).map(t => '<span class="pill">' + esc(t) + ': ' + yc.types[t] + '</span>').join('') + '</div>';
    if (yc.emptyYellow.length) h += '<div class="warn">Empty yellow cells: ' + yc.emptyYellow.map(esc).join(', ') + '</div>';
    if (yc.outsidePR.length) h += '<div class="warn">Yellow outside PR Override rows: ' + yc.outsidePR.map(o => esc(o.addr)).join(', ') + '</div>';
    if (yc.questions.length) h += '<div class="bad">' + yc.questions.length + ' question(s) need a reply</div>' +
      table(['TMO#', 'Item', 'DC', 'Wk', 'Question'], yc.questions.map(x => [esc(x.tmo), esc(x.item), esc(x.dc), L.fmtMD(layout.dates[x.week]), esc(x.text)]));
    h += table(['TMO#', 'Item', 'DC', 'Comments'], yc.rows.map(r => [esc(r.tmo), esc(r.item), esc(r.dcs.join(' / ')),
      r.cols.map(c => L.fmtMD(layout.dates[c.col]) + ' ' + esc(c.text)).join('<br>')]));
    return h;
  }

  function renderSwings(sw) {
    let h = '<h2>2 · WoW Avg Demand swings</h2><div>' + sw.rows.length + ' flagged (Big ' + PM + Math.round(sw.params.big * 100) + '%, Watch ' + PM + Math.round(sw.params.watch * 100) + '%, or unit change ' + String.fromCharCode(8805) + ' ' + fmt(sw.params.minUnits) + '/wk).</div>';
    h += table(['TMO#', 'DC', 'This', 'Last', 'WoW', 'Tier', 'Note'],
      sw.rows.map(r => [esc(r.block.tmo), esc(r.block.dc), fmt(r.thisWk), fmt(r.lastWk), pct(r.wow), esc(r.tier), esc(r.note)]), [2, 3, 4]);
    if (sw.notComparable.length) h += '<div class="warn">Not compared: ' + sw.notComparable.map(n => esc(n.block.tmo + ' ' + n.block.dc + ' (' + n.reason + ')')).join(', ') + '</div>';
    return h;
  }

  function renderRisk(layout, sr) {
    let h = '<h2>3 · Stock-out risk at sell-through trend</h2><div>' +
      Object.keys(sr.counts).map(k => '<span class="pill">' + esc(k) + ': ' + sr.counts[k] + '</span>').join('') + '</div>';
    const section = cat => {
      const rows = sr.rows.filter(r => r.category === cat);
      if (!rows.length) return '';
      return '<h2>' + esc(cat) + ' (' + rows.length + ')</h2>' + table(['TMO#', 'Item', 'DC', 'Fcst/wk', 'Trend/wk', 'Gap', 'Short (cur)', 'Short (trend)', 'Low @ trend'],
        rows.map(r => [esc(r.block.tmo), esc(r.block.item), esc(r.block.dc), fmt(r.forecast), fmt(r.trend), pct(r.gap),
          r.firstCur === null ? '' : L.fmtMD(r.firstCur), r.firstTrend === null ? '' : L.fmtMD(r.firstTrend), fmt(r.lowest)]), [3, 4, 5, 8]);
    };
    ['Stock-out only at trend', 'Earlier at trend', 'Short at current forecast', 'Forecast above trend'].forEach(c => (h += section(c)));
    const skipped = sr.rows.filter(r => r.category === 'Skipped');
    if (skipped.length) h += '<div class="warn">Skipped (error or zero demand/trend): ' + skipped.map(r => esc(r.block.tmo + ' ' + r.block.dc)).join(', ') + '</div>';
    return h;
  }

  // ---- writing -----------------------------------------------------------------------------------
  async function makeSheet(ctx, name, replace) {
    const sheets = ctx.workbook.worksheets;
    const ex = sheets.getItemOrNullObject(name); ex.load('isNullObject');
    await ctx.sync();
    if (!ex.isNullObject) {
      if (!replace) throw new Error('Tab "' + name + '" already exists. Tick "Replace existing review tabs" to rebuild it (any notes you typed there will be lost).');
      ex.delete(); await ctx.sync();
    }
    return sheets.add(name);
  }

  function styleHeader(rng) { rng.format.font.bold = true; rng.format.fill.color = '#D9E1F2'; }
  function borders(rng) {
    const b = rng.format.borders.getItem('InsideHorizontal'); b.style = 'Continuous'; b.color = '#D9D9D9';
    const e = rng.format.borders.getItem('EdgeBottom'); e.style = 'Continuous'; e.color = '#D9D9D9';
  }
  function widths(ws, list) { list.forEach(([a, w]) => (ws.getRange(a).format.columnWidth = w)); }

  async function verifyTab(ctx, ws, name) {
    const ur = ws.getUsedRange(true); ur.load('values');
    await ctx.sync();
    let errors = 0;
    for (const row of ur.values) for (const v of row) if (typeof v === 'string' && /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!)/.test(v)) errors++;
    return { name, rows: ur.values.length - 1, errors };
  }

  async function writeComments(src, layout, yc, replace) {
    return Excel.run(async ctx => {
      const name = tabName(src, ' Comments');
      const ws = await makeSheet(ctx, name, replace);
      const S = L.q(src), off = layout.off;
      const data = [['TMO#', 'Item #', 'Description', 'DC', 'Week(s) ' + ARROW + ' Comment', 'Cell(s)']];
      for (const r of yc.rows) {
        const b0 = r.blocks[0];
        const parts = r.cols.map(({ col }) => 'TEXT(' + S + '!' + L.colLetter(col) + '$1,"m/d")&" "&' + S + '!' + L.colLetter(col) + (b0.row + off.pr));
        data.push(['=' + S + '!B' + b0.row, '=' + S + '!C' + b0.row, '=' + S + '!D' + b0.row,
          '=' + r.blocks.map(b => S + '!E' + b.row).join('&" / "&'), '=TEXTJOIN("; ",TRUE,' + parts.join(',') + ')', r.cells.join(', ')]);
      }
      ws.getRangeByIndexes(0, 5, data.length, 1).numberFormat = data.map(() => ['@']);
      const rng = ws.getRangeByIndexes(0, 0, data.length, 6);
      rng.formulas = data;
      styleHeader(ws.getRange('A1:F1'));
      if (data.length > 1) {
        const body = ws.getRangeByIndexes(1, 0, data.length - 1, 4); body.format.font.color = '#008000';
        ws.getRangeByIndexes(1, 1, data.length - 1, 1).numberFormat = data.slice(1).map(() => ['0']);
        ws.getRangeByIndexes(1, 4, data.length - 1, 1).format.wrapText = true;
        borders(rng);
      }
      widths(ws, [['A:A', 75], ['B:B', 100], ['C:C', 230], ['D:D', 85], ['E:E', 520], ['F:F', 130]]);
      ws.getRange('A2:F' + data.length).format.verticalAlignment = 'Top';
      ws.freezePanes.freezeRows(1);
      await ctx.sync();
      return verifyTab(ctx, ws, name);
    });
  }

  async function writeSwings(src, layout, sw, replace) {
    return Excel.run(async ctx => {
      const name = tabName(src, ' Demand Swings');
      const ws = await makeSheet(ctx, name, replace);
      const S = L.q(src), off = layout.off, n = sw.rows.length;
      const data = [['TMO#', 'Item #', 'Description', 'DC', 'Avg Demand This Wk', 'Avg Demand Last Wk', 'Unit Change', 'WoW %', 'Tier', 'Notes']];
      sw.rows.forEach((r, i) => {
        const row = i + 2, b = r.block.row;
        data.push(['=' + S + '!B' + b, '=' + S + '!C' + b, '=' + S + '!D' + b, '=' + S + '!E' + b,
          '=' + S + '!G' + (b + off.avg), '=' + S + '!H' + (b + off.avg), '=E' + row + '-F' + row, '=' + S + '!H' + (b + off.sell),
          '=IF(ABS(H' + row + ')>=$M$2,"Big swing",IF(ABS(H' + row + ')>=$M$3,"Watch",""))', r.note]);
      });
      ws.getRangeByIndexes(0, 0, data.length, 10).formulas = data;
      ws.getRange('L1:M3').values = [['Thresholds', ''], ['Big swing (' + PM + ')', sw.params.big], ['Watch (' + PM + ')', sw.params.watch]];
      ws.getRange('L1').format.font.bold = true;
      ws.getRange('M2:M3').numberFormat = [['0%'], ['0%']]; ws.getRange('M2:M3').format.font.color = '#0000FF';
      if (sw.notComparable.length) {
        ws.getRange('L5').values = [['Not compared (no prior week or zero demand)']]; ws.getRange('L5').format.font.bold = true;
        ws.getRangeByIndexes(5, 11, sw.notComparable.length, 2).values = sw.notComparable.map(x => [x.block.tmo + ' ' + x.block.dc, x.reason]);
      }
      styleHeader(ws.getRange('A1:J1'));
      if (n) {
        ws.getRange('A2:E' + (n + 1)).format.font.color = '#008000';
        ws.getRange('H2:H' + (n + 1)).format.font.color = '#008000';
        ws.getRange('B2:B' + (n + 1)).numberFormat = data.slice(1).map(() => ['0']);
        ws.getRange('E2:G' + (n + 1)).numberFormat = data.slice(1).map(() => Array(3).fill('#,##0.0;(#,##0.0);-'));
        ws.getRange('H2:H' + (n + 1)).numberFormat = data.slice(1).map(() => ['0.0%;(0.0%);-']);
        borders(ws.getRange('A1:J' + (n + 1)));
        const h = ws.getRange('H2:H' + (n + 1));
        const up = h.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        up.cellValue.format.font.color = '#006100'; up.cellValue.rule = { formula1: '0', operator: 'GreaterThan' };
        const dn = h.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        dn.cellValue.format.font.color = '#9C0006'; dn.cellValue.rule = { formula1: '0', operator: 'LessThan' };
        const tier = ws.getRange('I2:I' + (n + 1)).conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        tier.textComparison.format.fill.color = '#FCE4D6'; tier.textComparison.format.font.bold = true;
        tier.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: 'Big swing' };
      }
      widths(ws, [['A:A', 75], ['B:B', 100], ['C:C', 230], ['D:D', 55], ['E:G', 110], ['H:H', 65], ['I:I', 75], ['J:J', 380], ['K:K', 20], ['L:L', 130], ['M:M', 70]]);
      ws.freezePanes.freezeRows(1);
      await ctx.sync();
      return verifyTab(ctx, ws, name);
    });
  }

  async function writeRisk(src, layout, sr, replace) {
    return Excel.run(async ctx => {
      const name = tabName(src, ' Stock-Out Risk');
      const ws = await makeSheet(ctx, name, replace);
      const S = L.q(src), off = layout.off, W = layout.weeks;
      const FIRST = 6, LAST = FIRST + sr.rows.length - 1;
      const W0 = 13; // column N: first weekly stock column
      const wk = i => L.colLetter(W0 + i), src1 = i => L.colLetter(layout.firstCol + i);
      const hdr = '$N$5:$' + wk(W - 1) + '$5';
      const cats = ['Stock-out only at trend', 'Earlier at trend', 'Short at current forecast', 'Forecast above trend', 'Skipped'];

      const head = ['TMO#', 'Item #', 'Description', 'DC', 'Forecast/wk', 'Trend/wk', 'Gap %', 'Beginning Stock', 'Valid', 'First Short Wk (Current)', 'First Short Wk (Trend)', 'Lowest Stock @ Trend', 'Risk Category'];
      for (let i = 0; i < W; i++) head.push('=' + S + '!' + src1(i) + '1');
      const data = [head];
      sr.rows.forEach((r, k) => {
        const i = FIRST + k, b = r.block.row;
        const wkRange = 'N' + i + ':' + wk(W - 1) + i;
        const row = ['=' + S + '!B' + b, '=' + S + '!C' + b, '=' + S + '!D' + b, '=' + S + '!E' + b,
          '=IFERROR(' + S + '!G' + (b + off.avg) + ',"n/a")', '=IFERROR(' + S + '!G' + (b + off.trend) + ',"n/a")',
          '=IF(I' + i + ',F' + i + '/E' + i + '-1,"")', '=' + S + '!K' + (b + off.begin),
          '=AND(ISNUMBER(E' + i + '),ISNUMBER(F' + i + '),IFERROR(E' + i + '>0,FALSE),IFERROR(F' + i + '>0,FALSE))',
          '=IFERROR(INDEX(' + hdr + ',MATCH(TRUE,INDEX(' + S + '!K' + (b + off.trueDemand) + ':' + src1(W - 1) + (b + off.trueDemand) + '<0,0),0)),"")',
          '=IF(I' + i + ',IFERROR(INDEX(' + hdr + ',MATCH(TRUE,INDEX(' + wkRange + '<0,0),0)),""),"")',
          '=IF(I' + i + ',MIN(' + wkRange + '),"")',
          '=IF(NOT(I' + i + '),"Skipped",IF(AND(K' + i + '<>"",J' + i + '=""),"Stock-out only at trend",IF(AND(K' + i + '<>"",J' + i + '<>"",K' + i + '<J' + i + '),"Earlier at trend",IF(J' + i + '<>"","Short at current forecast",IF(G' + i + '<=$D$2,"Forecast above trend","")))))'];
        for (let w = 0; w < W; w++) {
          const prev = w === 0 ? S + '!K' + (b + off.begin) : wk(w - 1) + i;
          row.push('=IF($I' + i + ',' + prev + '-MAX(' + S + '!' + src1(w) + (b + off.dd) + ',$F' + i + ')+' + S + '!' + src1(w) + (b + off.sched) + ',"")');
        }
        data.push(row);
      });
      ws.getRange('A1').values = [[src + ' ' + String.fromCharCode(8212) + ' Stock-out risk if demand moves to sell-through trend']];
      ws.getRange('A1').format.font.bold = true; ws.getRange('A1').format.font.size = 13;
      ws.getRange('A2').values = [['Over-forecast gap threshold']];
      ws.getRange('D2').values = [[sr.gapThreshold]]; ws.getRange('D2').numberFormat = [['0%']]; ws.getRange('D2').format.font.color = '#0000FF';
      ws.getRange('A3').values = [['Method: weekly stock = prior ' + String.fromCharCode(8722) + ' MAX(Dependent Demand, Trend/wk) + SCP PO Schedule, starting from Beginning Stock. Current = ' + src + ' True Demand row. PO Commits not added. Skipped = error or zero forecast/trend.']];
      ws.getRange('A3').format.font.color = '#605E5C';
      ws.getRange('G1:K1').values = [cats];
      ws.getRange('G2:K2').formulas = [cats.map((c, i) => '=COUNTIF($M$' + FIRST + ':$M$' + LAST + ',' + L.colLetter(6 + i) + '1)')];
      ws.getRange('G1:K1').format.font.bold = true;
      ws.getRangeByIndexes(4, 0, data.length, head.length).formulas = data;

      const full = ws.getRangeByIndexes(4, 0, data.length, head.length);
      styleHeader(ws.getRangeByIndexes(4, 0, 1, head.length));
      ws.getRange('N5:' + wk(W - 1) + '5').numberFormat = [Array(W).fill('m/d')];
      ws.getRange('A' + FIRST + ':D' + LAST).format.font.color = '#008000';
      ws.getRange('H' + FIRST + ':H' + LAST).format.font.color = '#008000';
      const rows = LAST - FIRST + 1;
      const nf = (a, f, cols) => (ws.getRange(a + FIRST + ':' + a + LAST).numberFormat = Array.from({ length: rows }, () => [f]));
      nf('B', '0'); nf('E', '#,##0.0'); nf('F', '#,##0.0'); nf('G', '0.0%;(0.0%);-'); nf('H', '#,##0;(#,##0);-'); nf('J', 'm/d/yy'); nf('K', 'm/d/yy'); nf('L', '#,##0;(#,##0);-');
      ws.getRange('N' + FIRST + ':' + wk(W - 1) + LAST).numberFormat = Array.from({ length: rows }, () => Array(W).fill('#,##0;(#,##0);-'));
      borders(full);
      const colours = { 'Stock-out only at trend': ['#FFC7CE', '#9C0006'], 'Earlier at trend': ['#FCE4D6', '#833C0B'], 'Short at current forecast': ['#FFEB9C', '#7F6000'], 'Forecast above trend': ['#DDEBF7', '#1F4E78'], Skipped: ['#EDEDED', '#595959'] };
      for (const c of cats) {
        const cf = ws.getRange('M' + FIRST + ':M' + LAST).conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        cf.textComparison.format.fill.color = colours[c][0]; cf.textComparison.format.font.color = colours[c][1];
        cf.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: c };
      }
      const neg = ws.getRange('N' + FIRST + ':' + wk(W - 1) + LAST).conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      neg.cellValue.format.font.color = '#C00000'; neg.cellValue.rule = { formula1: '0', operator: 'LessThan' };
      widths(ws, [['A:A', 75], ['B:B', 100], ['C:C', 230], ['D:D', 55], ['E:H', 85], ['I:I', 50], ['J:K', 95], ['L:L', 95], ['M:M', 170], [wk(0) + ':' + wk(W - 1), 60]]);
      ws.freezePanes.freezeAt(ws.getRange('A1:D5'));
      await ctx.sync();
      return verifyTab(ctx, ws, name);
    });
  }
})();
