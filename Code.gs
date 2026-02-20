// ============================================================
// VelocityMAX - Linear Velocity Tracking Dashboard
// Main entry point: menu, settings, UI dialogs
// ============================================================

var SETTINGS_SHEET = '_Settings';

/**
 * Runs on spreadsheet open – adds the VelocityMAX menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VelocityMAX')
    .addItem('Set API Key', 'showApiKeyDialog')
    .addSeparator()
    .addItem('Select Team', 'showTeamPicker')
    .addItem('Select Project', 'showProjectPicker')
    .addSeparator()
    .addItem('Import Issues & Build Dashboard', 'importIssues')
    .addItem('Refresh Data', 'refreshAllData')
    .addItem('Rebuild Dashboard Charts', 'buildAllCharts')
    .addSeparator()
    .addItem('Enable Auto-Refresh (hourly)', 'enableAutoRefresh')
    .addItem('Disable Auto-Refresh', 'disableAutoRefresh')
    .addSeparator()
    .addItem('Select Statuses for Breakdown', 'showStatusPicker')
    .addSeparator()
    .addItem('Set Date Range Filter', 'showDateRangeDialog')
    .addToUi();
}

// --------------- Settings helpers ---------------

function getSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.hideSheet();
    // headers
    sheet.getRange('A1').setValue('key');
    sheet.getRange('B1').setValue('value');
  }
  return sheet;
}

function saveSetting_(key, value) {
  var sheet = getSettingsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getSetting_(key) {
  var sheet = getSettingsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// --------------- API Key ---------------

function showApiKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var current = getSetting_('linearApiKey');
  var msg = current
    ? 'An API key is already saved. Enter a new one to replace it, or cancel.'
    : 'Enter your Linear API key (Personal API key from Linear Settings > API).';
  var result = ui.prompt('Linear API Key', msg, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() === ui.Button.OK) {
    var key = result.getResponseText().trim();
    if (key) {
      saveSetting_('linearApiKey', key);
      ui.alert('API key saved.');
    }
  }
}

function getApiKey_() {
  var key = getSetting_('linearApiKey');
  if (!key) {
    var msg = 'No Linear API key found. Use VelocityMAX > Set API Key first.';
    logError_('getApiKey_', msg, 'N/A'); // No stack trace for this simple check
    throw new Error(msg);
  }
  return key;
}

// --------------- Team picker ---------------

function showTeamPicker() {
  var apiKey = getApiKey_();
  var teams = fetchTeams(apiKey);
  if (!teams.length) {
    SpreadsheetApp.getUi().alert('No teams found in your Linear workspace.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPickerHtml_('team', teams))
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a Team');
}

function onTeamSelected(teamId, teamName) {
  saveSetting_('teamId', teamId);
  saveSetting_('teamName', teamName);
  // clear project selection when team changes
  saveSetting_('projectId', '');
  saveSetting_('projectName', '');
  // clear status selection when team changes
  saveSetting_('filteredStatuses', '');
}

// --------------- Project picker ---------------

function showProjectPicker() {
  var apiKey = getApiKey_();
  var teamId = getSetting_('teamId');
  if (!teamId) {
    SpreadsheetApp.getUi().alert('Please select a team first.');
    return;
  }
  var projects = fetchProjects(apiKey, teamId);
  if (!projects.length) {
    SpreadsheetApp.getUi().alert('No projects found for this team.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPickerHtml_('project', projects))
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a Project');
}

function onProjectSelected(projectIds, projectNames) {
  saveSetting_('projectId', JSON.stringify(projectIds));
  saveSetting_('projectName', JSON.stringify(projectNames));
}

// --------------- Status picker ---------------

function showStatusPicker() {
  var apiKey = getApiKey_();
  var teamId = getSetting_('teamId');
  if (!teamId) {
    SpreadsheetApp.getUi().alert('Please select a team first.');
    return;
  }
  var allStatuses = fetchAllStatuses(apiKey, teamId);

  if (!allStatuses.length) {
    SpreadsheetApp.getUi().alert('No statuses found for this team.');
    return;
  }

  var html = HtmlService.createHtmlOutput(buildPickerHtml_('status', allStatuses))
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Statuses for Breakdown');
}

function onStatusesSelected(statusNames) {
  saveSetting_('filteredStatuses', JSON.stringify(statusNames));
  // Re-import issues to apply the new status filter and rebuild charts
  importIssues();
}

// --------------- Date Range Picker ---------------

function showDateRangeDialog() {
  var ui = SpreadsheetApp.getUi();
  var startDate = getSetting_('startDate') || '';
  var endDate = getSetting_('endDate') || '';

  var htmlOutput = HtmlService.createHtmlOutput(
    '<p>Enter Start and End Dates (YYYY-MM-DD):</p>' +
    '<input type="date" id="startDate" value="' + startDate + '"><br><br>' +
    '<input type="date" id="endDate" value="' + endDate + '"><br><br>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).onDateRangeSelected(document.getElementById(\'startDate\').value, document.getElementById(\'endDate\').value)">Save</button>'
  )
  .setWidth(300)
  .setHeight(200);

  ui.showModalDialog(htmlOutput, 'Set Date Range Filter');
}

function onDateRangeSelected(startDate, endDate) {
  saveSetting_('startDate', startDate);
  saveSetting_('endDate', endDate);
  // Re-import issues to apply the new date range filter and rebuild charts
  importIssues();
}

// --------------- Shared picker HTML builder ---------------

function buildPickerHtml_(type, items) {
  var isProjectPicker = type === 'project';
  var isStatusPicker = type === 'status';

  var selectedItemsJson = null;
  if (isProjectPicker) {
    selectedItemsJson = getSetting_('projectId');
  } else if (isStatusPicker) {
    selectedItemsJson = getSetting_('filteredStatuses');
  }
  var selectedItems = selectedItemsJson ? JSON.parse(selectedItemsJson) : [];

  var rows = items
    .map(function (item) {
      var escapedName = item.name.replace(/'/g, "\\'").replace(/"/g, '&quot;');
      var isChecked = selectedItems.some(function(p) { return p === item.name; }) ? 'checked' : ''; // For statuses, item.id is actually item.name

      // For project picker, we pass id and name. For status picker, we only have name (which acts as id)
      var valueAttribute = isProjectPicker ? 'value="' + item.id + '"' : 'value="' + escapedName + '"';
      var dataNameAttribute = isProjectPicker ? 'data-name="' + escapedName + '"' : '';

      return (
        '<tr>' +
        '<td style="padding:6px 12px;width:30px;">' +
        (isProjectPicker || isStatusPicker ?
          '<input type="checkbox" class="' + type + '-checkbox" ' + isChecked + ' ' + valueAttribute + ' ' + dataNameAttribute + '>' :
          '<button onclick="pick(\'' + item.id + "','" + escapedName + '\')" ' +
          'style="cursor:pointer;padding:6px 16px;border:1px solid #5e6ad2;' +
          'background:#5e6ad2;color:#fff;border-radius:4px;">Select</button>') +
        '</td>' +
        '<td style="padding:6px 12px;">' +
        item.name +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  var callbackFn;
  if (isProjectPicker) {
    callbackFn = 'onProjectSelected';
  } else if (isStatusPicker) {
    callbackFn = 'onStatusesSelected';
  } else {
    callbackFn = 'onTeamSelected';
  }

  var headerText = '';
  var checkboxClass = '';
  if (isProjectPicker) {
    headerText = '<h2>Select Projects</h2>';
    checkboxClass = 'project-checkbox';
  } else if (isStatusPicker) {
    headerText = '<h2>Select Statuses for Breakdown</h2>';
    checkboxClass = 'status-checkbox';
  }

  var header = (isProjectPicker || isStatusPicker) ?
    headerText +
    '<div style="margin-bottom: 10px;">' +
    '<button onclick="selectAll()" style="margin-right: 10px;cursor:pointer;padding:6px 16px;border:1px solid #5e6ad2;' +
    'background:#5e6ad2;color:#fff;border-radius:4px;">Select All</button>' +
    '<button onclick="deselectAll()" style="cursor:pointer;padding:6px 16px;border:1px solid #5e6ad2;' +
    'background:#5e6ad2;color:#fff;border-radius:4px;">Deselect All</button>' +
    '<button onclick="saveSelection()" style="float:right;cursor:pointer;padding:6px 16px;border:1px solid #28a745;' +
    'background:#28a745;color:#fff;border-radius:4px;">Save Selection</button>' +
    '</div>' : '';

  var script = (isProjectPicker || isStatusPicker) ?
    '<script>' +
    'function selectAll() { Array.from(document.querySelectorAll(".' + checkboxClass + '")).forEach(cb => cb.checked = true); }' +
    'function deselectAll() { Array.from(document.querySelectorAll(".' + checkboxClass + '")).forEach(cb => cb.checked = false); }' +
    'function saveSelection() {' +
    'var selectedIds = [];' +
    'var selectedNames = [];' +
    'Array.from(document.querySelectorAll(".' + checkboxClass + ':checked")).forEach(function(cb){' +
    'selectedIds.push(cb.value);' +
    'if (cb.dataset.name) selectedNames.push(cb.dataset.name); else selectedNames.push(cb.value);' + // For statuses, value is name
    '});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).' +
    callbackFn +
    '(selectedIds, selectedNames);' +
    '}' +
    '</script>' :
    '<script>' +
    'function pick(id,name){' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).' +
    callbackFn +
    '(id,name);}' +
    '</script>';

  return (
    '<style>body{font-family:sans-serif;}' +
    'table{width:100%;border-collapse:collapse;} ' +
    'tr:hover{background:#f5f5f5;}</style>' +
    header +
    '<table>' +
    rows +
    '</table>' +
    script
  );
}

// --------------- Import / Refresh ---------------

function importIssues() {
  var ui = SpreadsheetApp.getUi();
  try {
    var apiKey = getApiKey_();
    var teamId = getSetting_('teamId');
    var projectIdsJson = getSetting_('projectId');
    var projectIds = projectIdsJson ? JSON.parse(projectIdsJson) : [];
    var startDate = getSetting_('startDate');
    var endDate = getSetting_('endDate');

    if (!teamId) {
      ui.alert(
        'Please select a team first (VelocityMAX > Select Team).'
      );
      return;
    }
    if (!projectIds || projectIds.length === 0) {
      ui.alert(
        'Please select at least one project first (VelocityMAX > Select Project).'
      );
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    ui.alert(
      'Import started',
      'Fetching issues from Linear… This may take a moment.',
      ui.ButtonSet.OK
    );

    // 1. Fetch raw issues
    var issues = fetchIssuesForProject(apiKey, teamId, projectIds, startDate, endDate);
    if (!issues.length) {
      ui.alert('No issues found for the selected team/project(s) in the specified date range.');
      return;
    }

    // 2. Fetch status history for every issue
    var issuesWithHistory = fetchStatusHistories(apiKey, issues);

    // 3. Compute metrics
    var processed = processIssues(issuesWithHistory);

    // 4. Write to sheets
    writeIssuesToSheet(ss, processed);
    writeWeeklyVelocity(ss, processed);
    writeStatusBreakdown(ss, processed);

    // 5. Compute and Write Burnup/Burndown data
    var burnupBurndownData = computeBurnupBurndownData_(processed);
    writeBurnupBurndownDataToSheet(ss, burnupBurndownData);

    // 6. Auto-build Dashboard with all charts
    buildAllCharts();

    ui.alert(
      'Done! ' +
        processed.length +
        ' issues imported and Dashboard updated.'
    );
  } catch (e) {
    logError_('importIssues', e.message, e.stack);
    ui.alert('Error during import', 'An error occurred during data import. Please check the _ErrorLog sheet for details: ' + e.message, ui.ButtonSet.OK);
  }
}

function refreshAllData() {
  importIssues();
}

// --------------- Auto-Refresh Trigger ---------------

/**
 * Creates an hourly time-driven trigger to re-import data from Linear.
 * Only one trigger is created; duplicates are prevented.
 */
function enableAutoRefresh() {
  // Remove any existing VelocityMAX triggers first
  removeAutoRefreshTriggers_();

  ScriptApp.newTrigger('autoRefresh_')
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert(
    'Auto-refresh enabled. Data will sync from Linear every hour.'
  );
}

/**
 * Removes the hourly auto-refresh trigger.
 */
function disableAutoRefresh() {
  var removed = removeAutoRefreshTriggers_();
  SpreadsheetApp.getUi().alert(
    removed
      ? 'Auto-refresh disabled.'
      : 'No auto-refresh trigger was active.'
  );
}

/**
 * The function invoked by the time-driven trigger.
 * Runs without UI prompts so it works unattended.
 */
function autoRefresh_() {
  try {
    var apiKey = getSetting_('linearApiKey');
    var teamId = getSetting_('teamId');
    if (!apiKey || !teamId) return; // nothing configured yet

    var projectIdsJson = getSetting_('projectId');
    var projectIds = projectIdsJson ? JSON.parse(projectIdsJson) : [];
    if (!projectIds || projectIds.length === 0) return; // no projects selected

    var startDate = getSetting_('startDate');
    var endDate = getSetting_('endDate');

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var issues = fetchIssuesForProject(apiKey, teamId, projectIds, startDate, endDate);
    if (!issues.length) return;

    var issuesWithHistory = fetchStatusHistories(apiKey, issues);
    var processed = processIssues(issuesWithHistory);

    writeIssuesToSheet(ss, processed);
    writeWeeklyVelocity(ss, processed);
    writeStatusBreakdown(ss, processed);

    // Compute and Write Burnup/Burndown data
    var burnupBurndownData = computeBurnupBurndownData_(processed);
    writeBurnupBurndownDataToSheet(ss, burnupBurndownData);

    buildAllCharts();
  } catch (e) {
    logError_('autoRefresh_', e.message, e.stack);
    // Silent failure for auto-refresh, as it runs unattended
  }
}

function removeAutoRefreshTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = false;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'autoRefresh_') {
      ScriptApp.deleteTrigger(t);
      removed = true;
    }
  });
  return removed;
}

// --------------- Error Logging ---------------

/**
 * Logs an error message to a hidden _ErrorLog sheet.
 */
function logError_(functionName, errorMessage, stack) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '_ErrorLog';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
    sheet.appendRow(['Timestamp', 'Function', 'Error Message', 'Stack Trace']);
  }
  sheet.appendRow([new Date(), functionName, errorMessage, stack]);
}
