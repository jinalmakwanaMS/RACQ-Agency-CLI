---
name: racq-status-report
description: >
  Generates a daily RACQ support ticket status report by reading the RACQ_Support_Tickets.xlsx
  spreadsheet and Shruti's latest status email, then categorizes all high/critical IcMs into
  three priority buckets and emails the report. Use this skill when asked about RACQ status,
  IcM triage, or daily support ticket reporting.
allowed-tools: shell
---

# RACQ Daily Status Report Skill

## Overview

This skill reads RACQ support ticket data, correlates it with the latest status updates,
and produces a prioritized daily report for the 11 high and critical IcM items.

## Data Sources

1. **Excel Spreadsheet**: RACQ Support Tickets
   - URL: `https://microsoftapc-my.sharepoint.com/personal/shgulat_microsoft_com2/Documents/Microsoft%20Teams%20Chat%20Files/RACQ_Support_Tickets.xlsx?web=1`
   - Use the Work IQ / M365 Copilot tools to access this file.

2. **Daily Status Email from Shruti (shgulat@microsoft.com)**
   - Search for the most recent daily status report email from Shruti.
   - Use Work IQ to query: "What is the latest daily status report email from shgulat@microsoft.com about RACQ support tickets?"

3. **Team Chat / IcM Details**
   - For each IcM, look up team chat discussions and IcM details to understand current status.
   - Use Work IQ to query about specific IcM numbers for context from chats and emails.

## Report Format

### Step 1: Gather Data

1. Read the RACQ_Support_Tickets.xlsx file to get the full list of high and critical IcMs.
2. Query for Shruti's latest status report email to get current status updates.
3. For each of the 11 high/critical items, gather:
   - IcM ID/Number
   - Title/Description
   - Severity (High/Critical)
   - Date Created
   - Age (days since creation, calculated from today's date)
   - Current Status (from Shruti's email and team chat)
   - Workaround availability
   - ETA for fix (if available)

### Step 2: Categorize into Priority Buckets

Categorize each IcM into one of three buckets based on the gathered information:

#### 🔴 Priority 1 – No Workaround / Go-Live Blockers
Criteria:
- No workaround exists
- Blocking go-live or critical customer functionality
- Still open/active with no resolution path

For each item include:
| IcM ID | Title | Severity | Age (days) | Status | Notes |
|--------|-------|----------|------------|--------|-------|

#### 🟡 Priority 2 – Workaround Exists but Insufficient / ETA Pending
Criteria:
- A workaround exists but is not sustainable or sufficient
- Fix ETA is pending or unclear
- Still open/active but not immediately blocking

For each item include:
| IcM ID | Title | Severity | Age (days) | Workaround | ETA | Notes |
|--------|-------|----------|------------|------------|-----|-------|

#### 🟢 Priority 3 – Closed / Resolved / Not a Bug
Criteria:
- Issue has been resolved or closed
- Determined to be not a bug
- No further action required

For each item include:
| IcM ID | Title | Severity | Age (days) | Resolution | Closed Date |
|--------|-------|----------|------------|------------|-------------|

### Step 3: Generate Summary

At the top of the report, include:
- **Report Date**: Today's date
- **Total Items**: 11
- **P1 Count**: X items (🔴 No Workaround / Go-Live Blockers)
- **P2 Count**: X items (🟡 Workaround Exists / ETA Pending)
- **P3 Count**: X items (🟢 Closed / Resolved)
- **Average Age**: X days
- **Oldest Open Item**: IcM #XXXX (X days)

### Step 4: Send Email

Send the report via email to:
- **To**: jinalmakwana@microsoft.com
- **CC**: Toby.James@microsoft.com
- **Subject**: `RACQ Daily Status Report - {Today's Date}`

The email body should contain:
1. Executive summary (counts per bucket)
2. Priority 1 table (🔴)
3. Priority 2 table (🟡)
4. Priority 3 table (🟢)
5. Trend notes (any items that changed bucket since last report, if known)

## Invocation

This skill should run daily. When invoked, follow all steps above in order.

Example prompts that should trigger this skill:
- "Generate the RACQ daily status report"
- "What's the latest status on RACQ IcMs?"
- "Run the daily RACQ triage report"
- "Send the RACQ status email"
