// Port of "diagram4/automation.py" (id: s4, "Risk Dashboard").
// Reads "OneTrust - Risk Export", computes Open/Closed/Total risk counts
// (Open = Evaluation+Identified+Treatment stages, Closed = Monitoring stage,
// year-filtered on Date closed when present), derives a period label from
// the max 2026 date found in Date created/Date closed — falling back to the
// real current date when neither column exists, matching today's shipped
// behavior exactly (confirmed with the user: no cross-run persistence, this
// report resets to the 2 seed history rows and recomputes "now" every run,
// exactly like the current Pyodide app, whose working directory is wiped
// before every click). Builds a Baseline'25 -> last-3-periods -> Target'26
// matrix and writes "Risk_Output.xlsx" (sheet "Dashboard" + "History Used")
// and "History.xlsx".
window.Reports.s4 = async function s4(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var YEAR = 2026;
  var OPEN_STAGES = { 'Evaluation': 1, 'Identified': 1, 'Treatment': 1 };
  var CLOSED_STAGE = 'Monitoring';
  var BASELINE_LABEL = "Baseline '25";
  var BASELINE_OPEN_RISK = 536, BASELINE_CLOSED_RISK = 42, BASELINE_TOTAL_RISK = 578;
  var TARGET_LABEL = "Target '26";
  var TARGET_PERCENT = 0.80;
  var SEED_ROWS = [
    { Month: "Q1 '26", 'Open risk as on date': 655, 'Closed Risk in 2026': 41, 'Total Risk': 696, 'Risk Created in 2026': 118 },
    { Month: "Apr '26", 'Open risk as on date': 694, 'Closed Risk in 2026': 42, 'Total Risk': 736, 'Risk Created in 2026': 158 },
  ];
  var MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  var sheet = E.readSheet(wb, 'OneTrust - Risk Export');
  var cleanedHeaders = sheet.headers.map(function (h) { return String(h).replace(/\r/g, '').replace(/\n/g, '').trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    sheet.headers.forEach(function (h, i) { out[cleanedHeaders[i]] = row[h]; });
    return out;
  });
  var lookup = {};
  cleanedHeaders.forEach(function (c) { lookup[c.toLowerCase()] = c; });
  function findRequiredColumn(name) {
    var key = lookup[name.toLowerCase()];
    if (!key) throw new Error("Required column '" + name + "' was not found.");
    return key;
  }
  function findOptionalColumn(name) { return lookup[name.toLowerCase()] || null; }

  var stageCol = findRequiredColumn('Stage');
  var dateCreatedCol = findOptionalColumn('Date created');
  var dateClosedCol = findOptionalColumn('Date closed');

  rows.forEach(function (r) {
    r[stageCol] = E.isBlank(r[stageCol]) ? '' : String(r[stageCol]).replace(/\r/g, '').replace(/\n/g, '').trim();
  });

  var openRisk = rows.filter(function (r) { return OPEN_STAGES.hasOwnProperty(r[stageCol]); }).length;
  var closedRisk;
  if (dateClosedCol) {
    closedRisk = rows.filter(function (r) { return r[stageCol] === CLOSED_STAGE && E.excelYear(r[dateClosedCol]) === YEAR; }).length;
  } else {
    closedRisk = rows.filter(function (r) { return r[stageCol] === CLOSED_STAGE; }).length;
  }
  var totalRisk = openRisk + closedRisk;
  var riskCreated = dateCreatedCol ? rows.filter(function (r) { return E.excelYear(r[dateCreatedCol]) === YEAR; }).length : 0;
  var targetValue = Math.round(totalRisk * TARGET_PERCENT);

  function getPeriodLabel() {
    var dates = [];
    [dateCreatedCol, dateClosedCol].forEach(function (col) {
      if (!col) return;
      var years = rows.map(function (r) { return E.excelDateInfo(r[col]); }).filter(function (d) { return d && d.year === YEAR; });
      if (years.length) {
        var maxD = years.reduce(function (a, b) {
          if (a.year !== b.year) return a.year > b.year ? a : b;
          if (a.month !== b.month) return a.month > b.month ? a : b;
          return a.day >= b.day ? a : b;
        });
        dates.push(maxD);
      }
    });
    if (!dates.length) throw new Error('No valid ' + YEAR + ' dates found in Date created/Date closed columns.');
    var latest = dates.reduce(function (a, b) {
      if (a.year !== b.year) return a.year > b.year ? a : b;
      if (a.month !== b.month) return a.month > b.month ? a : b;
      return a.day >= b.day ? a : b;
    });
    var abbr = MONTH_ABBR[latest.month - 1];
    return abbr === 'Mar' ? "Q1 '26" : (abbr + " '26");
  }

  var periodLabel;
  if (dateCreatedCol || dateClosedCol) {
    periodLabel = getPeriodLabel();
  } else {
    var now = new Date();
    var abbr = MONTH_ABBR[now.getMonth()];
    periodLabel = abbr === 'Mar' ? "Q1 '26" : (abbr + " '" + String(now.getFullYear()).slice(2));
  }

  var processedOn = (function () {
    var d = new Date();
    function pad(n) { return String(n).padStart(2, '0'); }
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  })();

  var metrics = {
    Month: periodLabel,
    'Open risk as on date': openRisk,
    'Closed Risk in 2026': closedRisk,
    'Total Risk': totalRisk,
    'Risk Created in 2026': riskCreated,
    'Target 80%': targetValue,
    'Closed %': totalRisk ? closedRisk / totalRisk : 0,
    'Input File': 'workbook.xlsx',
    'Processed On': processedOn,
  };

  // load_or_create_history: no persistence across runs (Decision 4) — always
  // seed fresh, matching the current app's actual (not intended) behavior.
  var HIST_COLS = ['Month', 'Open risk as on date', 'Closed Risk in 2026', 'Total Risk', 'Risk Created in 2026', 'Target 80%', 'Closed %', 'Input File', 'Processed On'];
  var history = SEED_ROWS.map(function (r) {
    var row = Object.assign({}, r);
    row['Target 80%'] = Math.round(row['Total Risk'] * TARGET_PERCENT);
    row['Closed %'] = row['Closed Risk in 2026'] / row['Total Risk'];
    row['Input File'] = 'Seed from sample image';
    row['Processed On'] = processedOn;
    return row;
  });

  // update_history: append only if this month isn't already present.
  var alreadyPresent = history.some(function (r) { return String(r.Month) === String(metrics.Month); });
  if (!alreadyPresent) history = history.concat([metrics]);

  // build_dashboard_frames
  function periodSortKey(label) {
    label = String(label);
    if (label.indexOf('Q1') === 0) return 3;
    var m = label.match(/^([A-Za-z]{3})/);
    var order = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
    return m ? (order[m[1]] || 99) : 99;
  }
  var sortedHistory = history.slice().sort(function (a, b) { return periodSortKey(a.Month) - periodSortKey(b.Month); });
  var graphHistory = sortedHistory.slice(-3);
  var periodLabels = graphHistory.map(function (r) { return r.Month; });
  var targetForTable = Math.round(totalRisk * TARGET_PERCENT);

  var wideColumns = [BASELINE_LABEL].concat(periodLabels, [TARGET_LABEL]);
  var rowOpen = [BASELINE_OPEN_RISK].concat(graphHistory.map(function (r) { return r['Open risk as on date']; }), ['']);
  var rowClosed = [BASELINE_CLOSED_RISK].concat(graphHistory.map(function (r) { return r['Closed Risk in 2026']; }), ['']);
  var rowTotal = [BASELINE_TOTAL_RISK].concat(graphHistory.map(function (r) { return r['Total Risk']; }), ['']);
  var rowBlank = wideColumns.map(function () { return ''; });
  var rowCreated = [0].concat(graphHistory.map(function (r) { return r['Risk Created in 2026']; }), ['']);
  var rowBaseline = [BASELINE_TOTAL_RISK].concat(periodLabels.map(function () { return BASELINE_TOTAL_RISK; }), ['']);

  var tableRows = [
    ['1. Open risk as on date'].concat(rowOpen),
    ['Closed Risk in 2026'].concat(rowClosed),
    ['Total Risk'].concat(rowTotal),
    [''].concat(rowBlank),
    ['Risk Created in 2026'].concat(rowCreated),
    [''].concat(rowBaseline),
  ];
  var tableHeaders = [''].concat(wideColumns);

  var calcLabels = [BASELINE_LABEL].concat(periodLabels, [TARGET_LABEL]);
  var calcValues = [BASELINE_TOTAL_RISK].concat(graphHistory.map(function (r) { return r['Open risk as on date']; }), [targetForTable]);
  var calcPercents = [''].concat(graphHistory.map(function (r) { return r['Closed %']; }), [TARGET_PERCENT]);

  // --- build the Dashboard grid (0-indexed rows/cols, matching xlsxwriter's ws.write(row,col,...)) ---
  function setCell0(grid, row, col, value) {
    while (grid.length <= row) grid.push([]);
    while (grid[row].length <= col) grid[row].push(null);
    grid[row][col] = value;
  }
  var g = [];
  tableHeaders.forEach(function (h, c) { setCell0(g, 0, c, h); });
  tableRows.forEach(function (row, r) { row.forEach(function (v, c) { setCell0(g, r + 1, c, v); }); });

  var startRow = 10;
  setCell0(g, startRow, 0, 'Calculation used for chart');
  calcLabels.forEach(function (label, r) {
    var excelRow = startRow + 1 + r;
    setCell0(g, excelRow, 0, label);
    setCell0(g, excelRow, 1, calcValues[r]);
    setCell0(g, excelRow, 2, calcPercents[r] === '' ? '' : calcPercents[r]);
  });
  setCell0(g, startRow + 1 + calcLabels.length, 0, 'Note');
  setCell0(g, startRow + 1 + calcLabels.length, 1, 'Target = 80% of latest Total Risk');

  var chartStart = startRow + 10;
  setCell0(g, chartStart, 0, 'Chart Label');
  setCell0(g, chartStart, 1, 'Chart Value');
  setCell0(g, chartStart, 2, 'Percent Label');
  calcLabels.forEach(function (label, r) {
    var excelRow = chartStart + 1 + r;
    setCell0(g, excelRow, 0, label);
    setCell0(g, excelRow, 1, calcValues[r]);
    setCell0(g, excelRow, 2, calcPercents[r] === '' ? '' : calcPercents[r]);
  });

  var historyHeaderRow = HIST_COLS;
  var historyDataRows = history.map(function (r) { return HIST_COLS.map(function (c) { return r[c]; }); });
  var historyGrid = [historyHeaderRow].concat(historyDataRows);

  var files = [
    { name: 'Risk_Output.xlsx', sheets: [{ name: 'Dashboard', grid: g }, { name: 'History Used', grid: historyGrid }] },
    { name: 'History.xlsx', sheets: [{ name: 'Sheet1', grid: historyGrid }] },
  ];

  // Preview chart comes from the History table (matches today's app, which
  // derives it from History.xlsx, not from the Dashboard's own matrix — see
  // ReportEngine.selectCharts' 's4' branch). The embedded chart on the
  // Dashboard sheet reproduces the original xlsxwriter native column chart
  // instead (black chart/plot area, gray Baseline/Target end-bars, gold
  // bars in between, white centered value labels, no legend/gridlines).
  var barColors = calcLabels.map(function (_, i) {
    return (i === 0 || i === calcLabels.length - 1) ? '#BFBFBF' : '#FFC000';
  });
  var progressTraces = [
    { x: calcLabels, y: calcValues, type: 'bar', marker: { color: barColors },
      text: calcValues.map(String), textposition: 'inside', insidetextanchor: 'middle', textangle: 0,
      textfont: { color: '#ffffff', size: 12 } },
  ];
  var progressLayout = {
    paper_bgcolor: '#000000', plot_bgcolor: '#000000',
    title: { text: "<b>Cumulative Risk Treatment<br>Progress</b>", font: { color: '#ffffff', size: 18 } },
    xaxis: { tickfont: { color: '#ffffff' }, linecolor: '#ffffff', showline: true },
    yaxis: { visible: false },
    showlegend: false,
    margin: { t: 70, r: 20, b: 50, l: 20 },
  };
  var progressChartPng = await B.renderStyledPng(progressTraces, progressLayout, 720, 430);

  var workbook = new ExcelJS.Workbook();
  var wsDash = workbook.addWorksheet('Dashboard');
  g.forEach(function (r) { wsDash.addRow(r); });
  wsDash.columns.forEach(function (c) { c.width = 16; });
  // Native, editable progress chart (data-linked to a hidden helper block)
  // replaces the baked PNG: single gold series with the first and last bars
  // recolored gray (Baseline/Target), white centered labels, no legend.
  var s4Placements = [];
  if (window.NativeChartInject && window.fflate && calcLabels.length) {
    var s4Points = calcLabels.map(function (_, i) {
      return (i === 0 || i === calcLabels.length - 1) ? { idx: i, color: 'BFBFBF' } : null;
    }).filter(Boolean);
    var s4Blk = window.NativeChartInject.buildDataBlock(wsDash, 'Dashboard', calcLabels, [
      { name: 'Progress', cache: calcValues, color: 'FFC000', points: s4Points },
    ], 20);
    s4Placements.push({
      sheetName: 'Dashboard', anchor: { fromCol: 6, fromRow: 1, toCol: 15, toRow: 20 }, // ~"G2"
      def: Object.assign({
        grouping: 'clustered', legend: false, title: 'Cumulative Risk Treatment Progress',
        chartBg: '000000', plotBg: '000000', axisColor: 'FFFFFF',
        dataLabels: { position: 'ctr', color: 'FFFFFF' },
      }, s4Blk),
    });
  }

  var wsHistUsed = workbook.addWorksheet('History Used');
  historyGrid.forEach(function (r) { wsHistUsed.addRow(r); });
  wsHistUsed.columns.forEach(function (c) { c.width = 16; });

  var riskOutputBuf = await workbook.xlsx.writeBuffer();
  if (s4Placements.length) {
    try { riskOutputBuf = window.NativeChartInject.inject(new Uint8Array(riskOutputBuf), s4Placements); }
    catch (e) { console.error('s4 native chart inject failed:', e); }
  }

  var historyWorkbook = new ExcelJS.Workbook();
  var wsHistOnly = historyWorkbook.addWorksheet('Sheet1');
  historyGrid.forEach(function (r) { wsHistOnly.addRow(r); });
  wsHistOnly.columns.forEach(function (c) { c.width = 16; });
  var historyBuf = await historyWorkbook.xlsx.writeBuffer();

  return {
    ok: true,
    files: [
      { name: 'Risk_Output.xlsx', bytes: riskOutputBuf, sheets: [{ name: 'Dashboard', grid: g }, { name: 'History Used', grid: historyGrid }] },
      { name: 'History.xlsx', bytes: historyBuf, sheets: [{ name: 'Sheet1', grid: historyGrid }] },
    ],
    chartImages: { 'Dashboard': progressChartPng },
  };
};
