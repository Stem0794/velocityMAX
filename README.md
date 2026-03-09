# VelocityMAX – Linear Velocity Tracking for Google Sheets

A Google Apps Script that imports your Linear issues and builds velocity dashboards directly in Google Sheets.

## Features

- **Import issues** from Linear by team and **multiple projects**
- **Track time in each status** (days spent in Backlog, In Progress, In Review, Done, etc.), with configurable status filtering
- **Filter issues by a custom date range** (e.g., issues created or completed within a specific period)
- **Story points** and priority per ticket
- **Cycle time** (started → completed) and **lead time** (created → completed)
- **Weekly velocity chart** – points and tickets completed per week with trend line
- **Cycle time chart** – scatter plot + weekly average with trend
- **Status breakdown chart** – average/median days per status
- **Burn-up and Burn-down charts** for project progress visualization
- **Enhanced error reporting** with an error log sheet

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

### 3. Get a Linear API Key

1. Go to **Linear > Settings > API** (or My Account > API)
2. Create a **Personal API key**
3. Copy the key

### 4. Configure & Run

1. Reload your Google Sheet — you'll see a **VelocityMAX** menu
2. **VelocityMAX > Set API Key** — paste your Linear API key
3. **VelocityMAX > Select Team** — pick your team
4. **VelocityMAX > Select Project** (optional) — select one or multiple projects
5. **VelocityMAX > Select Statuses for Breakdown** (optional) — select which statuses to include in the breakdown chart
6. **VelocityMAX > Set Date Range Filter** (optional) — set a date range to filter issues
7. **VelocityMAX > Import Issues** — fetches all issues and builds data sheets
8. Use the chart menu items to generate visualizations

## Sheets Created

| Sheet | Description |
|-------|-------------|
| **Issues** | Full issue table with points, status, cycle/lead time, days per status |
| **Weekly Velocity** | Points and tickets completed per ISO week + avg cycle time |
| **Status Breakdown** | Average, median, total days per status across all issues (filtered by selected statuses) |
| **Burnup Burndown Data** | Daily cumulative data for burn-up/burn-down charts |
| **_ErrorLog** | Log of script errors for troubleshooting |

## Charts

- **Weekly Velocity** – combo chart (bars = points, line = ticket count) + trend line
- **Cycle Time Scatter** – individual issue cycle times over time with trend
- **Avg Cycle Time per Week** – line chart with trend
- **Status Breakdown** – horizontal bar chart (avg vs median days), configurable status filtering
- **Burn-up Chart** – cumulative scope and completed work over time (Points)
- **Burn-down Chart** – remaining work over time (Points Remaining)

## Notes

- The script handles Linear API pagination automatically
- Status history is fetched per-issue to compute time-in-status accurately
- Rate limiting: batches of 10 issues with 200ms pauses between batches
- All settings (API key, team, project, statuses, date range) are stored in a hidden `_Settings` sheet

# User Guide

Welcome to VelocityMAX! This guide will help you use the powerful features of this tool to track your Linear project velocity directly within Google Sheets.

## Getting Started

After the VelocityMAX script has been set up in your Google Sheet, you'll find a new menu item called **VelocityMAX**. This menu is your control center for the entire tool.

Here’s how to get started:

1.  **Set API Key**: Before you can import any data, you need to connect VelocityMAX to your Linear account.
    *   Go to **VelocityMAX > Set API Key**.
    *   A prompt will appear asking for your Linear API key.
    *   You can generate a new key in your **Linear settings under "API"**.
    *   Paste the key and click "OK".

2.  **Select Your Team**: Choose the Linear team you want to analyze.
    *   Go to **VelocityMAX > Select Team**.
    *   A sidebar will appear with a list of your teams. Select one.

3.  **Select Projects (Optional)**: By default, VelocityMAX will import issues from all projects within the selected team. If you want to focus on specific projects:
    *   Go to **VelocityMAX > Select Project**.
    *   Choose one or more projects from the sidebar.

4.  **Import Your Issues**: Once configured, you can import your issues.
    *   Go to **VelocityMAX > Import Issues**.
    *   The script will fetch all the issues from your selected team and projects. This may take a few moments.

## Understanding the Data

VelocityMAX organizes your data into several sheets:

*   **Issues**: This is the raw data for all your imported issues. It includes details like story points, current status, and cycle/lead times.
*   **Weekly Velocity**: This sheet summarizes how many story points and tickets were completed each week.
*   **Status Breakdown**: Shows how long, on average, issues spend in each status (e.g., "In Progress", "In Review").
*   **Burnup Burndown Data**: This sheet contains the data used to generate the burn-up and burn-down charts, showing progress over time.

## Visualizing Your Velocity

VelocityMAX also creates several charts to help you visualize your team's performance:

*   **Weekly Velocity Chart**: A bar chart showing the number of story points and tickets completed each week. This helps you understand your team's weekly output.
*   **Cycle Time Scatter Chart**: This scatter plot shows the cycle time (the time from when an issue is started to when it's completed) for individual issues. It's useful for spotting outliers.
*   **Average Cycle Time Chart**: A line chart showing the average cycle time per week. This helps you track trends in how quickly your team is completing work.
*   **Status Breakdown Chart**: A bar chart that shows the average and median time issues spend in each status. This can help you identify bottlenecks in your workflow.
*   **Burn-up and Burn-down Charts**: These charts visualize your project's progress over time, showing how much work is remaining and how much has been completed.

## Advanced Configuration

You can further customize your reports using these menu options:

*   **VelocityMAX > Select Statuses for Breakdown**: Choose which statuses to include in the "Status Breakdown" chart. This is useful for focusing on specific parts of your workflow.
*   **VelocityMAX > Set Date Range Filter**: Filter the imported issues by a specific date range (e.g., only issues created in the last quarter).

## Troubleshooting

If you encounter any errors, check the **_ErrorLog** sheet. This sheet logs any problems the script runs into, which can help with debugging.
