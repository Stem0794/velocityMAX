// ============================================================
// ChartBuilder.gs – Create all charts on a single Dashboard sheet
// ============================================================

var DASHBOARD_SHEET = 'Dashboard';

/**
 * Master function: builds every chart on the Dashboard sheet.
 * Called automatically after data import.
 */
function buildAllCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure data sheets exist
  var velocitySheet = ss.getSheetByName('Weekly Velocity');
  var issuesSheet = ss.getSheetByName('Issues');
  var statusSheet = ss.getSheetByName('Status Breakdown');

  if (!velocitySheet && !issuesSheet && !statusSheet) {
    SpreadsheetApp.getUi().alert(
      'No data sheets found. Import issues first (VelocityMAX > Import Issues).'
    );
    return;
  }

  // Prepare scatter data helper sheet (needed before building charts)
  if (issuesSheet) {
    prepareScatterData_(ss, issuesSheet);
  }

  // Create / clear the Dashboard sheet
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  if (dashboard) {
    removeChartsFromSheet_(dashboard);
    dashboard.clear();
  } else {
    dashboard = ss.insertSheet(DASHBOARD_SHEET, 0);
  }

  // ---- Dashboard header ----
  var teamName = getSetting_('teamName') || '—';
  var projectName = getSetting_('projectName') || 'All Projects';
  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm'
  );

  dashboard.getRange('A1').setValue('VelocityMAX Dashboard');
  dashboard.getRange('A1').setFontSize(18).setFontWeight('bold');
  dashboard
    .getRange('A2')
    .setValue(
      'Team: ' + teamName +
      '  |  Project: ' + projectName +
      '  |  Updated: ' + timestamp
    );
  dashboard.getRange('A2').setFontSize(10).setFontColor('#666666');
  dashboard.setColumnWidth(1, 250);

  // ---- Build charts ----
  var chartRow = 4;

  // 1. Weekly Velocity (combo)
  if (velocitySheet && velocitySheet.getLastRow() > 1) {
    chartRow = addVelocityChart_(dashboard, velocitySheet, chartRow);
  }

  // 2. Points Velocity Trend (line + trendline)
  if (velocitySheet && velocitySheet.getLastRow() > 1) {
    chartRow = addVelocityTrendChart_(dashboard, velocitySheet, chartRow);
  }

  // 3. Issue Cycle Times (scatter)
  var scatterSheet = ss.getSheetByName('_CycleTimeScatter');
  if (scatterSheet && scatterSheet.getLastRow() > 1) {
    chartRow = addCycleTimeScatterChart_(dashboard, scatterSheet, chartRow);
  }

  // 4. Status Breakdown (bar)
  if (statusSheet && statusSheet.getLastRow() > 1) {
    chartRow = addStatusBreakdownChart_(dashboard, statusSheet, chartRow);
  }

  // Move Dashboard to first position
  ss.setActiveSheet(dashboard);
  ss.moveActiveSheet(1);
}

// ==================== Individual chart builders ====================

/**
 * Weekly Velocity – combo chart (bars for points, line for tickets).
 */
function addVelocityChart_(dashboard, srcSheet, startRow) {
  var numRows = srcSheet.getLastRow();

  // Section title
  dashboard.getRange('A' + startRow).setValue('Weekly Velocity');
  dashboard.getRange('A' + startRow).setFontSize(13).setFontWeight('bold');

  // Explanation
  dashboard
    .getRange('A' + (startRow + 1))
    .setValue(
      'How to read: Purple bars show story points completed each week (left axis). ' +
      'The red line tracks the number of tickets closed (right axis). ' +
      'Rising bars indicate the team is delivering more work per week.'
    );
  dashboard
    .getRange('A' + (startRow + 1))
    .setFontSize(9)
    .setFontColor('#888888')
    .setWrap(true);

  var chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(srcSheet.getRange(1, 1, numRows, 1)) // Week labels
    .addRange(srcSheet.getRange(1, 2, numRows, 1)) // Points
    .addRange(srcSheet.getRange(1, 3, numRows, 1)) // Ticket count
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(startRow + 2, 1, 0, 0)
    .setOption('title', 'Weekly Velocity')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('hAxis', { title: 'Week', textStyle: { fontSize: 11 } })
    .setOption('vAxes', {
      0: { title: 'Points Completed', textStyle: { fontSize: 11 } },
      1: { title: 'Tickets Completed', textStyle: { fontSize: 11 } },
    })
    .setOption('series', {
      0: {
        type: 'bars',
        targetAxisIndex: 0,
        color: '#5e6ad2',
        labelInLegend: 'Points Completed',
      },
      1: {
        type: 'line',
        targetAxisIndex: 1,
        color: '#e5484d',
        lineWidth: 3,
        labelInLegend: 'Tickets Completed',
      },
    })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('width', 900)
    .setOption('height', 450)
    .build();

  dashboard.insertChart(chart);
  return startRow + 29;
}

/**
 * Points Velocity Trend – line chart with linear trendline.
 */
function addVelocityTrendChart_(dashboard, srcSheet, startRow) {
  var numRows = srcSheet.getLastRow();

  dashboard.getRange('A' + startRow).setValue('Velocity Trend');
  dashboard.getRange('A' + startRow).setFontSize(13).setFontWeight('bold');

  dashboard
    .getRange('A' + (startRow + 1))
    .setValue(
      'How to read: The purple line shows points completed each week. ' +
      'The red dashed trend line reveals the overall direction — ' +
      'an upward slope means velocity is improving over time.'
    );
  dashboard
    .getRange('A' + (startRow + 1))
    .setFontSize(9)
    .setFontColor('#888888')
    .setWrap(true);

  var chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(srcSheet.getRange(1, 1, numRows, 1)) // Week
    .addRange(srcSheet.getRange(1, 2, numRows, 1)) // Points
    .setPosition(startRow + 2, 1, 0, 0)
    .setOption('title', 'Points Velocity Trend')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('hAxis', { title: 'Week', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { title: 'Points Completed', textStyle: { fontSize: 11 } })
    .setOption('series', {
      0: { color: '#5e6ad2', lineWidth: 3, labelInLegend: 'Points Completed' },
    })
    .setOption('trendlines', {
      0: {
        type: 'linear',
        color: '#e5484d',
        lineWidth: 2,
        opacity: 0.6,
        labelInLegend: 'Trend',
        showR2: true,
        visibleInLegend: true,
      },
    })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('width', 900)
    .setOption('height', 400)
    .build();

  dashboard.insertChart(chart);
  return startRow + 27;
}

/**
 * Issue Cycle Times – scatter plot of individual issues over time.
 */
function addCycleTimeScatterChart_(dashboard, scatterSheet, startRow) {
  var numRows = scatterSheet.getLastRow();

  dashboard
    .getRange('A' + startRow)
    .setValue('Individual Issue Cycle Times');
  dashboard.getRange('A' + startRow).setFontSize(13).setFontWeight('bold');

  dashboard
    .getRange('A' + (startRow + 1))
    .setValue(
      'How to read: Each dot is one completed issue. The Y-axis shows how many days ' +
      'it took from start to done. Dots high up took a long time; dots near the bottom ' +
      'were fast. The red trend line shows whether cycle times are improving (going down) ' +
      'or getting worse (going up).'
    );
  dashboard
    .getRange('A' + (startRow + 1))
    .setFontSize(9)
    .setFontColor('#888888')
    .setWrap(true);

  var chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(scatterSheet.getRange(1, 1, numRows, 1)) // Completed date
    .addRange(scatterSheet.getRange(1, 2, numRows, 1)) // Cycle time
    .setPosition(startRow + 2, 1, 0, 0)
    .setOption('title', 'Issue Cycle Times')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('hAxis', {
      title: 'Completion Date',
      textStyle: { fontSize: 11 },
    })
    .setOption('vAxis', {
      title: 'Cycle Time (days)',
      textStyle: { fontSize: 11 },
    })
    .setOption('series', {
      0: {
        color: '#5e6ad2',
        pointSize: 7,
        pointShape: 'circle',
        labelInLegend: 'Completed Issues',
      },
    })
    .setOption('trendlines', {
      0: {
        type: 'linear',
        color: '#e5484d',
        lineWidth: 2,
        opacity: 0.5,
        labelInLegend: 'Trend',
        visibleInLegend: true,
      },
    })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('width', 900)
    .setOption('height', 400)
    .build();

  dashboard.insertChart(chart);
  return startRow + 27;
}

/**
 * Status Breakdown – horizontal bar chart of avg/median hours per status.
 */
function addStatusBreakdownChart_(dashboard, srcSheet, startRow) {
  var numRows = srcSheet.getLastRow();

  dashboard.getRange('A' + startRow).setValue('Time Spent in Each Status');
  dashboard.getRange('A' + startRow).setFontSize(13).setFontWeight('bold');

  dashboard
    .getRange('A' + (startRow + 1))
    .setValue(
      'How to read: Each bar shows how long issues sit in a given workflow status. ' +
      'Purple = average across all issues, Green = median (less affected by outliers). ' +
      'Long bars highlight bottleneck statuses where work gets stuck.'
    );
  dashboard
    .getRange('A' + (startRow + 1))
    .setFontSize(9)
    .setFontColor('#888888')
    .setWrap(true);

  var chart = dashboard
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(srcSheet.getRange(1, 1, numRows, 1)) // Status names
    .addRange(srcSheet.getRange(1, 2, numRows, 1)) // Avg hours
    .addRange(srcSheet.getRange(1, 3, numRows, 1)) // Median hours
    .setPosition(startRow + 2, 1, 0, 0)
    .setOption('title', 'Average Time in Each Status')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('hAxis', { title: 'Hours', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { title: 'Status', textStyle: { fontSize: 11 } })
    .setOption('series', {
      0: { color: '#5e6ad2', labelInLegend: 'Avg Hours' },
      1: { color: '#30a46c', labelInLegend: 'Median Hours' },
    })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('width', 900)
    .setOption('height', 400)
    .build();

  dashboard.insertChart(chart);
  return startRow + 27;
}

// ==================== Scatter data preparation ====================

/**
 * Prepares the hidden _CycleTimeScatter sheet from Issues data.
 */
function prepareScatterData_(ss, issuesSheet) {
  var data = issuesSheet.getDataRange().getValues();
  if (data.length < 2) return;

  var headers = data[0];
  var completedCol = headers.indexOf('Completed');
  var cycleTimeCol = headers.indexOf('Cycle Time (days)');
  var pointsCol = headers.indexOf('Points');

  if (completedCol < 0 || cycleTimeCol < 0) return;

  var scatterSheetName = '_CycleTimeScatter';
  var scatterSheet = ss.getSheetByName(scatterSheetName);
  if (scatterSheet) {
    scatterSheet.clear();
  } else {
    scatterSheet = ss.insertSheet(scatterSheetName);
    scatterSheet.hideSheet();
  }

  var scatterHeaders = ['Completed Date', 'Cycle Time (days)', 'Points'];
  var scatterRows = [];
  for (var i = 1; i < data.length; i++) {
    var completed = data[i][completedCol];
    var cycleTime = data[i][cycleTimeCol];
    var points = pointsCol >= 0 ? data[i][pointsCol] : 1;
    if (completed && cycleTime) {
      scatterRows.push([new Date(completed), cycleTime, points || 1]);
    }
  }

  if (!scatterRows.length) return;

  var allData = [scatterHeaders].concat(scatterRows);
  scatterSheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);
}

// ==================== Legacy wrappers (menu still works) ====================

function buildVelocityChart() {
  buildAllCharts();
}

function buildCycleTimeChart() {
  buildAllCharts();
}

function buildStatusBreakdownChart() {
  buildAllCharts();
}

// ==================== Helpers ====================

function removeChartsFromSheet_(sheet) {
  var charts = sheet.getCharts();
  charts.forEach(function (c) {
    sheet.removeChart(c);
  });
}
