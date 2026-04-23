---
name: racq-status-report
description: >
  Generates a daily RACQ support ticket status report by combining live IcM data (via IcM MCP Server),
  Shruti's latest status email, Teams chat context, and the SharePoint Excel tracker. Categorizes
  all high/critical IcMs into three priority buckets with AI-synthesized status summaries.
  Use this skill when asked about RACQ status, IcM triage, or daily support ticket reporting.
allowed-tools: shell
---

# RACQ Daily Status Report Skill

## Overview

This skill generates a comprehensive daily RACQ IcM status report by cross-referencing
four data sources to produce accurate, up-to-date status summaries for each incident.
The output is a formatted HTML email opened in Outlook for review and send.

## Data Sources (Query ALL four, then synthesize)

### Source 1: IcM MCP Server (Primary - live incident data)

Query the IcM MCP API for each IcM to get authoritative live data:
- **Endpoint**: `https://icm-mcp-prod.azure-api.net/v1/`
- **Auth**: Azure CLI token with scope `api://icmmcpapi-prod/mcp.tools`
- **Protocol**: JSON-RPC 2.0 over HTTPS POST. Response is SSE format: strip `event: message\ndata: ` prefix before parsing.
- **Key tools**:
  - `get_incident_details_by_id` (param: `incidentId` as integer) - state, severity, owner, created date, title
  - `get_ai_summary` (param: `incidentId` as string) - AI-generated summary if available
  - `get_incident_context` (param: `incidentId` as string) - full context metadata

From IcM, extract: **State** (ACTIVE/MITIGATED/RESOLVED/CLOSED), **Age** (days since createdDate),
**Owner** (contactAlias), **Title**, **Severity**.

### Source 2: Shruti's Daily Status Email (Latest engineer updates)

Use Work IQ to query:
> "What is the latest daily status report email from shgulat@microsoft.com about RACQ support tickets? Include the full body."

From this email, extract per-IcM: **Latest engineer actions**, **next steps**, **workaround details**,
**ETA information**, **any blockers mentioned**.

### Source 3: Teams Chat Context (Real-time discussion)

For each ACTIVE/MITIGATED IcM, use Work IQ to query:
> "What are the latest Teams chat messages about IcM {IcM_ID} or {short_issue_description} in the RACQ channel?"

From chat, extract: **Real-time updates not yet in email**, **PG responses**, **escalation status**.

### Source 4: SharePoint Excel Tracker (SR numbers and baseline)

Use Work IQ to read:
> "Read the file RACQ_Support_Tickets.xlsx from shgulat's OneDrive and list all rows with columns"

URL: `https://microsoftapc-my.sharepoint.com/personal/shgulat_microsoft_com2/Documents/Microsoft%20Teams%20Chat%20Files/RACQ_Support_Tickets.xlsx?web=1`

From Excel, extract: **SR numbers**, **IcM-to-SR mapping**, **any new IcMs added**.

## IcM Registry

Current tracked IcMs (update this list as items are added/removed):

| IcM ID | Sev | SR # | Short Title |
|--------|-----|------|-------------|
| 21000000890308 | High | 2601270030006343 | IVR Welcome Message Stutters |
| 51000000978035 | High | 2604070030006100 | Missing Conversation Summary |
| 51000000910153 | High | 2602130030002880 | Custom column data not visible after transfer |
| 21000000968169 | High | 2603270030005523 | Script Errors on Active Conversation |
| 51000000902213 | High | 2602120030007445 | Conversation form disappears on navigation |
| 51000000958196 | High | 2603120030002536 | Automated Messages not working for Voice |
| 21000000863530 | High | 2601130030007340 | Inconsistent Call/Chat Notifications |
| 51000000899493 | High | 2601300030001832 | Automated Messages EN-AU limitation |
| 21000000930073 | Critical | 2603040030006553 | Solution import failure |
| 21000000895436 | Critical | 2601150030006239 | Copilot Studio Agent Publish Fails |

### Known Child/Linked IcMs
- **51000000969890** (child of 51000000910153) - OOB Save button saves zero data. Always query this too.

## Step-by-Step Execution

### Step 1: Gather Data from All Sources

1. **IcM MCP API**: For each IcM in the registry above, call `get_incident_details_by_id`. Also query child IcMs.
2. **Shruti's email**: Use Work IQ to get the latest status email body.
3. **Teams chat**: For each ACTIVE/MITIGATED item, search Teams chat for recent discussion.
4. **Excel**: Use Work IQ to read the tracker for any new items or SR number changes.

### Step 2: Synthesize Status per IcM

For each IcM, combine all four sources to produce a **Latest Status** summary:
- Lead with the most recent factual update (from whichever source is newest)
- Include workaround status (working/not working/none)
- Include ETA if mentioned in any source
- Note if PG has responded recently (from chat/email)
- Flag if status has changed since last report

### Step 3: Categorize into Priority Buckets

**Priority 1 - No Workaround / Go-Live Blockers** (Red):
- No workaround exists, OR
- Workaround exists but is NOT working (e.g., 51000000910153 where child IcM proves OOB save fails), OR
- Blocking go-live or critical customer functionality

**Priority 2 - Workaround Exists but Insufficient / ETA Pending** (Orange):
- A workaround exists and is functioning (even if not ideal)
- Fix ETA is pending or known
- Not immediately blocking go-live

**Priority 3 - Resolved / Closed** (Green):
- IcM state is RESOLVED or CLOSED

### Step 4: Build Key Asks

Generate Key Asks list at the TOP of the report:
- Include ALL P1 items (always)
- Include P2 items that do NOT have BOTH a workaround AND an ETA
- EXCLUDE P2 items that have a working workaround AND a confirmed ETA
- For each Key Ask: state the issue, age, what's needed from PG

### Step 5: Generate HTML Email

Create an HTML email with this structure:
1. Greeting: "Hi Team,"
2. Intro line with report date and data source note
3. **Key Asks** section (numbered list, bold titles with IcM links)
4. Horizontal rule
5. **Priority 1 table** - Red header (#f8d7da), columns: Sev, SR#, IcM (linked), Issue, Age (days), State, Owner, Latest Status
6. **Priority 2 table** - Orange header (#fff3cd), same columns
7. **Priority 3 table** - Green header (#d4edda), columns: Sev, SR#, IcM (linked), Issue, Age (days), State, Latest Status
8. Footer note: "Data sourced live from IcM portal on {date} using IcM MCP Server"
9. Sign-off and signature block

IcM links format: `https://portal.microsofticm.com/imp/v5/incidents/details/{IcM_ID}`

### Step 6: Create .eml and Open in Outlook

Save as `.eml` file with headers:
```
From: jinalmakwana@microsoft.com
To: jinalmakwana@microsoft.com
Cc: Toby.James@microsoft.com
Subject: RACQ Support Ticket Summary - {date} (Generated by Agency)
MIME-Version: 1.0
Content-Type: text/html; charset=utf-8
X-Unsent: 1
```

Save to: `C:\Users\jinalmakwana\OneDrive - Microsoft\Agency\RACQ-Agency-CLI\`
Then open with `Start-Process` to launch in Outlook compose window.

## Signature Block

```html
<p style="font-size:10pt; color:#555;">
<b>Jinal Makwana</b><br/>
Senior FastTrack Solution Architect<br/>
Dynamics 365 Apps and Common Data Service R&amp;D<br/><br/>
Mobile +61 430407653<br/>
Office +61 (3) 93204331<br/>
LinkedIn: <a href="https://www.linkedin.com/in/d365lady/">https://www.linkedin.com/in/d365lady/</a>
</p>
```

## Invocation

Example prompts that trigger this skill:
- "Generate the RACQ daily status report"
- "Run the daily RACQ report"
- "What's the latest status on RACQ IcMs?"
- "Send the RACQ status email"
