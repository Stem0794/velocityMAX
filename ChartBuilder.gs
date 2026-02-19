// ============================================================
// ChartBuilder.gs – Create Charts from the velocity data
// ============================================================

/**
 * Builds a velocity chart (points + tickets completed per week)
 * embedded in the "Weekly Velocity" sheet.
 */
function buildVelocityChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Weekly Velocity');
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No "Weekly Velocity" sheet found. Import issues first.'
    );
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data to chart.');
    return;
  }

  // Remove old charts on this sheet
  removeChartsFromSheet_(sheet);

  var numRows = data.length;

  // Combo chart: bars for points, line for ticket count
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(1, 1, numRows, 1)) // Week labels
    .addRange(sheet.getRange(1, 2, numRows, 1)) // Points
    .addRange(sheet.getRange(1, 3, numRows, 1)) // Ticket count
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(2, 6, 0, 0)
    .setOption('title', 'Weekly Velocity')
    .setOption('hAxis.title', 'Week')
    .setOption('vAxes', {
      0: { title: 'Points Completed' },
      1: { title: 'Tickets Completed' },
    })
    .setOption('series', {
      0: { type: 'bars', targetAxisIndex: 0, color: '#5e6ad2' },
      1: { type: 'line', targetAxisIndex: 1, color: '#e5484d' },
    })
    .setOption('legend', { position: 'top' })
    .setOption('width', 800)
    .setOption('height', 450)
    .build();

  sheet.insertChart(chart);

  // Also add a trend line chart for points
  var trendChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, numRows, 1)) // Week
    .addRange(sheet.getRange(1, 2, numRows, 1)) // Points
    .setPosition(28, 6, 0, 0)
    .setOption('title', 'Points Velocity Trend')
    .setOption('hAxis.title', 'Week')
    .setOption('vAxis.title', 'Points')
    .setOption('trendlines', { 0: { type: 'linear', color: '#e5484d', lineWidth: 2, opacity: 0.6 } })
    .setOption('series', { 0: { color: '#5e6ad2' } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 800)
    .setOption('height', 400)
    .build();

  sheet.insertChart(trendChart);

  SpreadsheetApp.getUi().alert('Velocity charts created on the "Weekly Velocity" sheet.');
}

/**
 * Builds a cycle-time chart from the "Weekly Velocity" sheet.
 */
function buildCycleTimeChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Weekly Velocity');
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No "Weekly Velocity" sheet found. Import issues first.'
    );
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data to chart.');
    return;
  }

  var numRows = data.length;

  // Also build a scatter plot from the Issues sheet for individual cycle times
  var issuesSheet = ss.getSheetByName('Issues');
  if (issuesSheet) {
    buildCycleTimeScatter_(ss, issuesSheet);
  }

  // Line chart of avg cycle time per week
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, numRows, 1)) // Week
    .addRange(sheet.getRange(1, 4, numRows, 1)) // Avg cycle time
    .setPosition(54, 6, 0, 0)
    .setOption('title', 'Average Cycle Time per Week')
    .setOption('hAxis.title', 'Week')
    .setOption('vAxis.title', 'Days')
    .setOption('series', { 0: { color: '#30a46c' } })
    .setOption('trendlines', { 0: { type: 'linear', color: '#e5484d', lineWidth: 2, opacity: 0.5 } })
    .setOption('legend', { position: 'none' })
    .setOption('width', 800)
    .setOption('height', 400)
    .build();

  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert('Cycle time charts created.');
}

/**
 * Scatter plot of individual issue cycle times over time.
 */
function buildCycleTimeScatter_(ss, issuesSheet) {
  var data = issuesSheet.getDataRange().getValues();
  if (data.length < 2) return;

  // Find column indices
  var headers = data[0];
  var completedCol = headers.indexOf('Completed') + 1;
  var cycleTimeCol = headers.indexOf('Cycle Time (days)') + 1;
  var pointsCol = headers.indexOf('Points') + 1;

  if (!completedCol || !cycleTimeCol) return;

  // Create a helper sheet for scatter data
  var scatterSheetName = '_CycleTimeScatter';
  var scatterSheet = ss.getSheetByName(scatterSheetName);
  if (scatterSheet) {
    scatterSheet.clear();
  } else {
    scatterSheet = ss.insertSheet(scatterSheetName);
    scatterSheet.hideSheet();
  }

  // Write completed date + cycle time pairs
  var scatterHeaders = ['Completed Date', 'Cycle Time (days)', 'Points'];
  var scatterRows = [];
  for (var i = 1; i < data.length; i++) {
    var completed = data[i][completedCol - 1];
    var cycleTime = data[i][cycleTimeCol - 1];
    var points = data[i][pointsCol - 1];
    if (completed && cycleTime) {
      scatterRows.push([new Date(completed), cycleTime, points || 1]);
    }
  }

  if (!scatterRows.length) return;

  var allData = [scatterHeaders].concat(scatterRows);
  scatterSheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);

  var numRows = allData.length;

  var chart = issuesSheet
    .newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(scatterSheet.getRange(1, 1, numRows, 1)) // Date
    .addRange(scatterSheet.getRange(1, 2, numRows, 1)) // Cycle time
    .setPosition(2, 18, 0, 0)
    .setOption('title', 'Issue Cycle Times')
    .setOption('hAxis.title', 'Completion Date')
    .setOption('vAxis.title', 'Cycle Time (days)')
    .setOption('series', {
      0: { color: '#5e6ad2', pointSize: 6, pointShape: 'circle' },
    })
    .setOption('trendlines', {
      0: { type: 'linear', color: '#e5484d', lineWidth: 2, opacity: 0.5 },
    })
    .setOption('legend', { position: 'none' })
    .setOption('width', 800)
    .setOption('height', 400)
    .build();

  issuesSheet.insertChart(chart);
}

/**
 * Builds a horizontal bar chart showing average hours spent in each status.
 */
function buildStatusBreakdownChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Status Breakdown');
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No "Status Breakdown" sheet found. Import issues first.'
    );
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data to chart.');
    return;
  }

  removeChartsFromSheet_(sheet);

  var numRows = data.length;

  // Bar chart: status name vs avg hours
  var chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(1, 1, numRows, 1)) // Status names
    .addRange(sheet.getRange(1, 2, numRows, 1)) // Avg hours
    .addRange(sheet.getRange(1, 3, numRows, 1)) // Median hours
    .setPosition(2, 7, 0, 0)
    .setOption('title', 'Average Time in Each Status')
    .setOption('hAxis.title', 'Hours')
    .setOption('vAxis.title', 'Status')
    .setOption('series', {
      0: { color: '#5e6ad2' },
      1: { color: '#30a46c' },
    })
    .setOption('legend', { position: 'top' })
    .setOption('width', 700)
    .setOption('height', 400)
    .build();

  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert(
    'Status breakdown chart created on the "Status Breakdown" sheet.'
  );
}

// --------------- Helpers ---------------

function removeChartsFromSheet_(sheet) {
  var charts = sheet.getCharts();
  charts.forEach(function (c) {
    sheet.removeChart(c);
  });
}
