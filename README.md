# VelocityMAX – Linear Velocity Tracking for Google Sheets

A Google Apps Script that imports your Linear issues and builds velocity dashboards directly in Google Sheets.

## Features

- **Import issues** from Linear by team and **multiple projects**
- **Track time in each status** (hours spent in Backlog, In Progress, In Review, Done, etc.), with configurable status filtering
- **Story points** and priority per ticket
- **Cycle time** (started → completed) and **lead time** (created → completed)
- **Weekly velocity chart** – points and tickets completed per week with trend line
- **Cycle time chart** – scatter plot + weekly average with trend
- **Status breakdown chart** – average/median hours per status

## Setup

### 1. Create a Google Sheet

1. Open [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete the default `Code.gs` content

### 2. Add the script files

Create the following files in the Apps Script editor (use **+** next to Files):

| File | Source |
|------|--------|
| `Code.gs` | Copy from `Code.gs` |
| `LinearAPI.gs` | Copy from `LinearAPI.gs` |
| `DataProcessor.gs` | Copy from `DataProcessor.gs` |
| `ChartBuilder.gs` | Copy from `ChartBuilder.gs` |
| `appsscript.json` | Copy from `appsscript.json` (enable "Show manifest file" in Project Settings) |

### 3. Get a Linear API Key

1. Go to **Linear > Settings > API** (or My Account > API)
2. Create a **Personal API key**
3. Copy the key

### 4. Configure & Run

1. Reload your Google Sheet — you'll see a **VelocityMAX** menu
2. **VelocityMAX > Set API Key** — paste your Linear API key
3. **VelocityMAX > Select Team** — pick your team
4. **VelocityMAX > Select Project** (optional) — select one or multiple projects
5. **VelocityMAX > Import Issues** — fetches all issues and builds data sheets
6. Use the chart menu items to generate visualizations

## Sheets Created

| Sheet | Description |
|-------|-------------|
| **Issues** | Full issue table with points, status, cycle/lead time, hours per status |
| **Weekly Velocity** | Points and tickets completed per ISO week + avg cycle time |
| **Status Breakdown** | Average, median, total hours per status across all issues (filtered by selected statuses) |

## Charts

- **Weekly Velocity** – combo chart (bars = points, line = ticket count) + trend line
- **Cycle Time Scatter** – individual issue cycle times over time with trend
- **Avg Cycle Time per Week** – line chart with trend
- **Status Breakdown** – horizontal bar chart (avg vs median hours), configurable status filtering

## Notes

- The script handles Linear API pagination automatically
- Status history is fetched per-issue to compute time-in-status accurately
- Rate limiting: batches of 10 issues with 200ms pauses between batches
- All settings (API key, team, project) are stored in a hidden `_Settings` sheet
