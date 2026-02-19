// ============================================================
// LinearAPI.gs – All communication with the Linear GraphQL API
// ============================================================

var LINEAR_API_URL = 'https://api.linear.app/graphql';

// --------------- Low-level GraphQL helper ---------------

function linearQuery_(apiKey, query, variables) {
  var payload = { query: query };
  if (variables) payload.variables = variables;

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(LINEAR_API_URL, options);
  var json = JSON.parse(response.getContentText());

  if (json.errors) {
    throw new Error(
      'Linear API error: ' + json.errors.map(function (e) { return e.message; }).join(', ')
    );
  }
  return json.data;
}

// --------------- Teams ---------------

function fetchTeams(apiKey) {
  var query = '{ teams { nodes { id name } } }';
  var data = linearQuery_(apiKey, query);
  return data.teams.nodes;
}

// --------------- Projects ---------------

function fetchProjects(apiKey, teamId) {
  var query =
    'query($teamId: String!) {' +
    '  team(id: $teamId) {' +
    '    projects { nodes { id name state } }' +
    '  }' +
    '}';
  var data = linearQuery_(apiKey, query, { teamId: teamId });
  return data.team.projects.nodes;
}

// --------------- Issues ---------------

/**
 * Fetches all issues for a team, optionally filtered by project.
 * Handles pagination via cursor.
 */
function fetchIssuesForProject(apiKey, teamId, projectId) {
  var allIssues = [];
  var hasMore = true;
  var cursor = null;

  while (hasMore) {
    var variables = { teamId: teamId, first: 100 };
    if (cursor) variables.after = cursor;

    var filterClause = '';
    if (projectId) {
      variables.projectId = projectId;
      filterClause = ', filter: { project: { id: { eq: $projectId } } }';
    }

    var query =
      'query($teamId: String!, $first: Int!, $after: String' +
      (projectId ? ', $projectId: ID!' : '') +
      ') {' +
      '  team(id: $teamId) {' +
      '    issues(first: $first, after: $after' +
      filterClause +
      ', orderBy: createdAt) {' +
      '      pageInfo { hasNextPage endCursor }' +
      '      nodes {' +
      '        id identifier title' +
      '        estimate' +
      '        priority priorityLabel' +
      '        createdAt completedAt canceledAt startedAt' +
      '        state { id name type }' +
      '        assignee { name }' +
      '        project { id name }' +
      '        labels { nodes { name } }' +
      '        cycle { id number startsAt endsAt }' +
      '      }' +
      '    }' +
      '  }' +
      '}';

    var data = linearQuery_(apiKey, query, variables);
    var connection = data.team.issues;
    allIssues = allIssues.concat(connection.nodes);
    hasMore = connection.pageInfo.hasNextPage;
    cursor = connection.pageInfo.endCursor;
  }

  return allIssues;
}

// --------------- Issue status history ---------------

/**
 * For each issue, fetches its history entries where the state changed.
 * Returns the same array with a `_statusHistory` array attached to each issue.
 *
 * Each history entry: { fromState, toState, createdAt }
 */
function fetchStatusHistories(apiKey, issues) {
  // Process in batches to avoid hitting rate limits
  var BATCH_SIZE = 10;
  for (var i = 0; i < issues.length; i += BATCH_SIZE) {
    var batch = issues.slice(i, i + BATCH_SIZE);
    batch.forEach(function (issue) {
      issue._statusHistory = fetchSingleIssueHistory_(apiKey, issue.id);
    });
    // Small pause between batches to stay within rate limits
    if (i + BATCH_SIZE < issues.length) {
      Utilities.sleep(200);
    }
  }
  return issues;
}

function fetchSingleIssueHistory_(apiKey, issueId) {
  var history = [];
  var hasMore = true;
  var cursor = null;

  while (hasMore) {
    var variables = { issueId: issueId, first: 50 };
    if (cursor) variables.after = cursor;

    var query =
      'query($issueId: String!, $first: Int!, $after: String) {' +
      '  issue(id: $issueId) {' +
      '    history(first: $first, after: $after) {' +
      '      pageInfo { hasNextPage endCursor }' +
      '      nodes {' +
      '        createdAt' +
      '        fromState { id name type }' +
      '        toState { id name type }' +
      '      }' +
      '    }' +
      '  }' +
      '}';

    var data = linearQuery_(apiKey, query, variables);
    var connection = data.issue.history;

    // Only keep entries where state actually changed
    var stateChanges = connection.nodes.filter(function (entry) {
      return entry.fromState && entry.toState;
    });

    history = history.concat(stateChanges);
    hasMore = connection.pageInfo.hasNextPage;
    cursor = connection.pageInfo.endCursor;
  }

  // Sort chronologically
  history.sort(function (a, b) {
    return new Date(a.createdAt) - new Date(b.createdAt);
  });

  return history;
}
