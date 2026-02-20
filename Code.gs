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
    throw new Error(
      'No Linear API key found. Use VelocityMAX > Set API Key first.'
    );
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

function onProjectSelected(projectId, projectName) {
  saveSetting_('projectId', projectId);
  saveSetting_('projectName', projectName);
}

// --------------- Shared picker HTML builder ---------------

function buildPickerHtml_(type, items) {
  var rows = items
    .map(function (item) {
      var escapedName = item.name.replace(/'/g, "\\'").replace(/"/g, '&quot;');
      return (
        '<tr>' +
        '<td style="padding:6px 12px;">' +
        '<button onclick="pick(\'' +
        item.id +
        "','" +
        escapedName +
        '\')" ' +
        'style="cursor:pointer;padding:6px 16px;border:1px solid #5e6ad2;' +
        'background:#5e6ad2;color:#fff;border-radius:4px;">Select</button>' +
        '</td>' +
        '<td style="padding:6px 12px;">' +
        item.name +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  var callbackFn =
    type === 'team' ? 'onTeamSelected' : 'onProjectSelected';

  return (
    '<style>body{font-family:sans-serif;} table{width:100%;border-collapse:collapse;} ' +
    'tr:hover{background:#f5f5f5;}</style>' +
    '<table>' +
    rows +
    '</table>' +
    '<script>' +
    'function pick(id,name){' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).' +
    callbackFn +
    '(id,name);}' +
    '</script>'
  );
}

// --------------- Import / Refresh ---------------

function importIssues() {
  var apiKey = getApiKey_();
  var teamId = getSetting_('teamId');
  var projectId = getSetting_('projectId');
  if (!teamId) {
    SpreadsheetApp.getUi().alert(
      'Please select a team first (VelocityMAX > Select Team).'
    );
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ui.alert(
    'Import started',
    'Fetching issues from Linear… This may take a moment.',
    ui.ButtonSet.OK
  );

  // 1. Fetch raw issues
  var issues = fetchIssuesForProject(apiKey, teamId, projectId);
  if (!issues.length) {
    ui.alert('No issues found for the selected team/project.');
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

  // 5. Auto-build Dashboard with all charts
  buildAllCharts();

  ui.alert(
    'Done! ' +
      processed.length +
      ' issues imported and Dashboard updated.'
  );
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
  var apiKey = getSetting_('linearApiKey');
  var teamId = getSetting_('teamId');
  if (!apiKey || !teamId) return; // nothing configured yet

  var projectId = getSetting_('projectId');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var issues = fetchIssuesForProject(apiKey, teamId, projectId);
  if (!issues.length) return;

  var issuesWithHistory = fetchStatusHistories(apiKey, issues);
  var processed = processIssues(issuesWithHistory);

  writeIssuesToSheet(ss, processed);
  writeWeeklyVelocity(ss, processed);
  writeStatusBreakdown(ss, processed);
  buildAllCharts();
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
