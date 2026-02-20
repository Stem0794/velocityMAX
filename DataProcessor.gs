// ============================================================
// DataProcessor.gs – Turn raw Linear data into velocity metrics
// ============================================================

/**
 * Takes issues (with _statusHistory attached) and returns an array of
 * flat objects ready to be written to sheets.
 */
function processIssues(issues) {
  return issues.map(function (issue) {
    var timeByStatus = computeTimeByStatus_(issue);
    var cycleTime = computeCycleTime_(issue);
    var leadTime = computeLeadTime_(issue);

    return {
      id: issue.identifier,
      title: issue.title,
      points: issue.estimate || 0,
      priority: issue.priorityLabel || '',
      assignee: issue.assignee ? issue.assignee.name : '',
      project: issue.project ? issue.project.name : '',
      labels: issue.labels
        ? issue.labels.nodes.map(function (l) { return l.name; }).join(', ')
        : '',
      currentStatus: issue.state ? issue.state.name : '',
      currentStatusType: issue.state ? issue.state.type : '',
      cycleNumber: issue.cycle ? issue.cycle.number : '',
      createdAt: issue.createdAt || '',
      startedAt: issue.startedAt || '',
      completedAt: issue.completedAt || '',
      canceledAt: issue.canceledAt || '',
      cycleTimeDays: cycleTime,
      leadTimeDays: leadTime,
      timeByStatus: timeByStatus,
      _raw: issue,
    };
  });
}

// --------------- Time-by-status ---------------

/**
 * Returns an object { "Status Name": hours, ... }
 * representing how long the issue spent in each status.
 */
function computeTimeByStatus_(issue) {
  var history = issue._statusHistory || [];
  var result = {};

  if (!history.length) return result;

  for (var i = 0; i < history.length; i++) {
    var entry = history[i];
    var statusName = entry.fromState.name;
    var start;

    if (i === 0) {
      // Time from issue creation to first transition
      start = new Date(issue.createdAt);
    } else {
      start = new Date(history[i - 1].createdAt);
    }

    var end = new Date(entry.createdAt);
    var hours = Math.max(0, (end - start) / (1000 * 60 * 60));

    if (!result[statusName]) result[statusName] = 0;
    result[statusName] += hours;
  }

  // Time in the current (last) state up to now or completedAt
  var lastTransition = history[history.length - 1];
  var currentStatus = lastTransition.toState.name;
  var sinceLastTransition = new Date(lastTransition.createdAt);
  var until = issue.completedAt
    ? new Date(issue.completedAt)
    : issue.canceledAt
    ? new Date(issue.canceledAt)
    : new Date();

  var hoursInCurrent = Math.max(0, (until - sinceLastTransition) / (1000 * 60 * 60));
  if (!result[currentStatus]) result[currentStatus] = 0;
  result[currentStatus] += hoursInCurrent;

  // Round to 1 decimal
  Object.keys(result).forEach(function (k) {
    result[k] = Math.round(result[k] * 10) / 10;
  });

  return result;
}

// --------------- Cycle time ---------------

/**
 * Cycle time = time from first "started" state to completion (in days).
 * Returns null if not completed.
 */
function computeCycleTime_(issue) {
  if (!issue.startedAt || !issue.completedAt) return null;
  var start = new Date(issue.startedAt);
  var end = new Date(issue.completedAt);
  return Math.round(((end - start) / (1000 * 60 * 60 * 24)) * 10) / 10;
}

/**
 * Lead time = time from creation to completion (in days).
 */
function computeLeadTime_(issue) {
  if (!issue.completedAt) return null;
  var start = new Date(issue.createdAt);
  var end = new Date(issue.completedAt);
  return Math.round(((end - start) / (1000 * 60 * 60 * 24)) * 10) / 10;
}

// --------------- Write to sheets ---------------

/**
 * Write the main issues table to an "Issues" sheet.
 */
function writeIssuesToSheet(ss, processed) {
  var sheetName = 'Issues';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Collect all unique status names for dynamic columns
  var allStatuses = collectAllStatuses_(processed);

  // Header row
  var headers = [
    'ID',
    'Title',
    'Points',
    'Priority',
    'Assignee',
    'Project',
    'Labels',
    'Current Status',
    'Status Type',
    'Cycle #',
    'Created',
    'Started',
    'Completed',
    'Canceled',
    'Cycle Time (days)',
    'Lead Time (days)',
  ];
  allStatuses.forEach(function (s) {
    headers.push('Hours in: ' + s);
  });

  // Data rows
  var rows = processed.map(function (p) {
    var row = [
      p.id,
      p.title,
      p.points,
      p.priority,
      p.assignee,
      p.project,
      p.labels,
      p.currentStatus,
      p.currentStatusType,
      p.cycleNumber,
      p.createdAt,
      p.startedAt,
      p.completedAt,
      p.canceledAt,
      p.cycleTimeDays !== null ? p.cycleTimeDays : '',
      p.leadTimeDays !== null ? p.leadTimeDays : '',
    ];
    allStatuses.forEach(function (s) {
      row.push(p.timeByStatus[s] || 0);
    });
    return row;
  });

  var allData = [headers].concat(rows);
  sheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);

  formatAsTable_(sheet, allData.length, headers.length);
}

/**
 * Write a "Weekly Velocity" sheet with points completed per week.
 */
function writeWeeklyVelocity(ss, processed) {
  var sheetName = 'Weekly Velocity';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Only completed issues
  var completed = processed.filter(function (p) {
    return p.completedAt;
  });

  if (!completed.length) {
    sheet.getRange('A1').setValue('No completed issues found.');
    return;
  }

  // Group by ISO week
  var weekMap = {};
  completed.forEach(function (p) {
    var week = getISOWeekLabel_(new Date(p.completedAt));
    if (!weekMap[week]) {
      weekMap[week] = { points: 0, count: 0, cycleTimes: [] };
    }
    weekMap[week].points += p.points;
    weekMap[week].count += 1;
    if (p.cycleTimeDays !== null) {
      weekMap[week].cycleTimes.push(p.cycleTimeDays);
    }
  });

  var weeks = Object.keys(weekMap).sort();

  var headers = [
    'Week',
    'Points Completed',
    'Tickets Completed',
    'Avg Cycle Time (days)',
  ];
  var rows = weeks.map(function (w) {
    var d = weekMap[w];
    var avgCycle =
      d.cycleTimes.length > 0
        ? Math.round(
            (d.cycleTimes.reduce(function (a, b) { return a + b; }, 0) /
              d.cycleTimes.length) *
              10
          ) / 10
        : '';
    return [w, d.points, d.count, avgCycle];
  });

  var allData = [headers].concat(rows);
  sheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);

  formatAsTable_(sheet, allData.length, headers.length);

  // Number formatting for numeric columns
  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');   // Points
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0');   // Tickets
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.0');     // Avg Cycle Time
  }
}

/**
 * Write a "Status Breakdown" sheet with average hours in each status.
 */
function writeStatusBreakdown(ss, processed) {
  var sheetName = 'Status Breakdown';
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var allStatuses = collectAllStatuses_(processed);
  if (!allStatuses.length) {
    sheet.getRange('A1').setValue('No status transition data found.');
    return;
  }

  // Filter for user-specified statuses
  var allowedStatuses = [
    'In Progress',
    'In Review',
    'Done',
    'In Dev',
    'In Staging',
    // Add other statuses here if needed
  ];
  var filteredStatuses = allStatuses.filter(function(status) {
    return allowedStatuses.includes(status);
  });

  if (!filteredStatuses.length) {
    sheet.getRange('A1').setValue('No relevant status data found for the selected statuses.');
    return;
  }

  // Compute averages
  var headers = ['Status', 'Avg Hours', 'Median Hours', 'Total Hours', 'Issue Count'];
  var rows = filteredStatuses.map(function (status) {
    var values = [];
    processed.forEach(function (p) {
      if (p.timeByStatus[status]) values.push(p.timeByStatus[status]);
    });
    var total = values.reduce(function (a, b) { return a + b; }, 0);
    var avg = values.length ? Math.round((total / values.length) * 10) / 10 : 0;
    var median = values.length ? computeMedian_(values) : 0;
    return [status, avg, median, Math.round(total * 10) / 10, values.length];
  });

  var allData = [headers].concat(rows);
  sheet
    .getRange(1, 1, allData.length, allData[0].length)
    .setValues(allData);

  formatAsTable_(sheet, allData.length, headers.length);

  // Number formatting for numeric columns
  if (rows.length) {
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('0.0');  // Avg Hours
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0.0');  // Median Hours
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('#,##0.0'); // Total Hours
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0');   // Issue Count
  }
}

// --------------- Table formatting ---------------

/**
 * Formats a sheet range as a styled Google Sheet table:
 * - Bold white header on dark background
 * - Alternating row colors
 * - Thin borders around all cells
 * - Frozen header row
 * - Auto-resized columns
 */
function formatAsTable_(sheet, numRows, numCols) {
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setFontColor('#ffffff');
  headerRange.setBackground('#5e6ad2');
  headerRange.setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // Alternating row colors
  if (numRows > 1) {
    for (var r = 2; r <= numRows; r++) {
      var rowRange = sheet.getRange(r, 1, 1, numCols);
      if (r % 2 === 0) {
        rowRange.setBackground('#f3f4f6');
      } else {
        rowRange.setBackground('#ffffff');
      }
    }
  }

  // Borders around the full table
  var tableRange = sheet.getRange(1, 1, numRows, numCols);
  tableRange.setBorder(
    true, true, true, true, true, true,
    '#d1d5db', SpreadsheetApp.BorderStyle.SOLID
  );

  // Auto-resize columns (cap at 20 to avoid slow resizing on wide sheets)
  for (var c = 1; c <= Math.min(numCols, 20); c++) {
    sheet.autoResizeColumn(c);
  }
}

// --------------- Utility functions ---------------

function collectAllStatuses_(processed) {
  var statusSet = {};
  processed.forEach(function (p) {
    Object.keys(p.timeByStatus).forEach(function (s) {
      statusSet[s] = true;
    });
  });
  return Object.keys(statusSet).sort();
}

function getISOWeekLabel_(date) {
  // Returns "YYYY-Www" format
  var d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  var yearStart = new Date(d.getFullYear(), 0, 1);
  var weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  return d.getFullYear() + '-W' + (weekNo < 10 ? '0' : '') + weekNo;
}

function computeMedian_(arr) {
  var sorted = arr.slice().sort(function (a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  var val =
    sorted.length % 2 !== 0
      ? sorted[mid]
      : (sorted[mid - 1] + sorted[mid]) / 2;
  return Math.round(val * 10) / 10;
}
