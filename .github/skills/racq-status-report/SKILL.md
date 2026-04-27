---
name: racq-status-report
description: >
  Generates a daily RACQ support ticket status report by combining SharePoint Excel data,
  Shruti's latest status email, live IcM data with discussion/bridge notes, and Teams chat context.
  Uses SharePoint MCP, Outlook MCP, IcM MCP, and Teams MCP directly — no local files or Work IQ needed.
  Categorizes all high/critical IcMs into three priority buckets with AI-synthesized status
  summaries and creates a draft email in Outlook for review.
  Also supports IcM delay analysis — identifies IcMs with no ISD ETA from Shruti's email,
  investigates root cause of delays via IcM MCP discussion/context, and drafts an escalation email.
  Use this skill when asked about RACQ status, IcM triage, daily support ticket reporting,
  or IcM delay analysis.
allowed-tools: shell, sharepoint, outlook, mail, icm, teams, calendar
---

# RACQ Daily Status Report Skill

## Overview

This skill generates a comprehensive daily RACQ IcM status report by cross-referencing
four data sources via MCP servers to produce accurate, up-to-date status summaries for each incident.
The output is a draft email created in Outlook (via Outlook MCP) ready for review and send.

**Execution order**: Excel (base data) → Shruti's email (engineer updates) → IcM per-item (live state + discussion) → Teams per-item (latest chat) → Create draft email.

## Data Sources (in execution order)

### Source 1: SharePoint Excel Tracker (Base data — read FIRST)

Read the RACQ_Support_Tickets.xlsx file directly from SharePoint using the **SharePoint MCP server**.
This is the base data source that provides the list of all tracked IcMs, SR numbers, descriptions,
severity, workaround info, ETAs, and current status notes.

1. **Get file metadata** using `getFileOrFolderMetadataByUrl` with URL:
   `https://microsoftapc-my.sharepoint.com/personal/shgulat_microsoft_com2/Documents/Microsoft%20Teams%20Chat%20Files/RACQ_Support_Tickets.xlsx?web=1`
   - Known values (may change if file is re-created):
     - **driveId**: `b!qLRFeBDYX02Df3ceNu6FMyCRtqj_ciJBgTYSrfF87ACQTYYT6zFNQIvRuSkVgMoG`
     - **fileId**: `01FJXTEBCXTCYNSCZ5YNBL7XWKAZGVCNY7`

2. **Download the file** using `readSmallBinaryFile` with the driveId and fileId above.
   The file is ~275KB (well under the 5MB limit). Content is returned as base64.

3. **Parse the Excel content**: Save base64 to a temp .xlsx file, then parse using
   `Import-Excel` (PowerShell) or `openpyxl` (Python). Read the "Support tickets" worksheet.

4. **Extract all High/Critical rows** — for each row, capture:
   - IcM # (incident ID)
   - SR # (support request number)
   - Description (issue title)
   - Updated Severity / Old Severity
   - Status, ETA, Workaround, Next Steps, Comments

This gives you the **complete IcM registry** — do NOT rely on a hardcoded list.

### Source 2: Shruti's Daily Status Email (Latest engineer updates)

Use the **Outlook MCP server** `mail_search_messages` to find the latest email:
- **Query**: `subject:"URGENT : RACQ | Open Product Issues" from:shgulat@microsoft.com`
- This returns the most recent status email from Shruti.

Then use `mail_get_message` with the returned message ID to read the full body.

From this email, extract per-IcM: **Latest engineer actions**, **next steps**, **workaround details**,
**ETA information**, **any blockers mentioned**.

### Source 3: IcM MCP Server (Live state + discussion per IcM)

For **each IcM** found in the Excel tracker (Source 1), query the **IcM MCP server** using
THREE tools to get the complete picture — state, discussion history, and AI summary:

1. **`get_incident_details_by_id`** (param: `incidentId` as integer)
   - Extract: **State** (ACTIVE/MITIGATED/RESOLVED/CLOSED), **Age** (days since createdDate),
     **Owner** (contactAlias), **Title**, **Severity**, **AssignedTo**.

2. **`get_incident_context`** (param: `incidentId` as string)
   - This returns **all detailed context** including **bridge summaries** (`IncidentSummaryLists`
     with `BridgeSummaryList`) that contain discussion and summary data from incident bridges.
   - Extract: **Latest discussion points**, **PG responses**, **action items from bridges**,
     **timeline of updates**, **any root cause analysis or investigation notes**.

3. **`get_ai_summary`** (param: `incidentId` as string)
   - Returns an **AI-generated summary** of the incident including key events and current state.
   - Use this to cross-reference and fill gaps from the bridge discussion.

Also query any known **child/linked IcMs** with the same three tools
(e.g., 51000000969890 is a child of 51000000910153).

**Important**: The IcM discussion (from `get_incident_context`) is critical for understanding
the latest engineering status — workaround attempts, PG fix progress, deployment timelines,
and any blockers. Do NOT skip this step.

### Source 4: Teams Chat Context (Real-time discussion per IcM)

For each **ACTIVE or MITIGATED** IcM, search Teams for recent discussion using the **Teams MCP server**:
- Use `SearchTeamsMessages` with a natural language query like:
  `"latest updates on IcM {IcM_ID} {short_issue_description} RACQ"`
- OR use `SearchTeamMessagesQueryParameters` with KQL for more precise results:
  `"{IcM_ID}"` or `"RACQ AND {keyword_from_title}"`

From chat, extract: **Real-time updates not yet in email or IcM discussion**, **PG responses**,
**escalation status**, **any recent decisions or workaround confirmations**.

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

### Step 1: Read Excel Base Data (SharePoint MCP)

1. Use SharePoint MCP `readSmallBinaryFile` with driveId and fileId above to download the Excel.
2. Save base64 content to a temp .xlsx file, parse the "Support tickets" worksheet.
3. Extract all High/Critical rows — this becomes your working IcM list for the report.
4. Note any new IcMs not previously tracked.

### Step 2: Read Shruti's Latest Status Email (Outlook MCP)

1. Use Outlook MCP `mail_search_messages` with query: `subject:"URGENT : RACQ | Open Product Issues" from:shgulat@microsoft.com`
2. Use `mail_get_message` on the first (most recent) result to get the full body.
3. Parse the email body to extract per-IcM status notes, next steps, workarounds, and ETAs.
4. Map each update back to the corresponding IcM from Step 1.

### Step 3: Query Each IcM for Live State + Discussion (IcM MCP)

For **each IcM** from Step 1, call all three IcM MCP tools:

1. **`get_incident_details_by_id`** → live state, age, owner, severity
2. **`get_incident_context`** → bridge summaries, discussion history, investigation notes
3. **`get_ai_summary`** → AI-synthesized incident summary

Also query child/linked IcMs with the same three tools.

From the IcM discussion, extract:
- Latest engineering updates and PG responses
- Workaround validation (confirmed working or not)
- Fix deployment timelines or ETAs mentioned in bridges
- Any blockers or escalation actions

### Step 4: Search Teams Chat per IcM (Teams MCP)

For each **ACTIVE or MITIGATED** IcM:

1. Use Teams MCP `SearchTeamsMessages` with: `"IcM {IcM_ID} RACQ {short_title_keywords}"`
2. Review returned messages for any updates newer than what's in the email or IcM discussion.
3. Extract: real-time PG responses, escalation decisions, workaround confirmations.

Skip this step for RESOLVED/CLOSED IcMs.

### Step 5: Synthesize Status per IcM

For each IcM, combine **all four sources** to produce a **Latest Status** summary:
- Lead with the most recent factual update (from whichever source is newest)
- Cross-reference IcM bridge discussion with Shruti's email — use the more recent/detailed one
- Include workaround status (working/not working/none) — validated against IcM discussion
- Include ETA if mentioned in any source (IcM discussion, email, or Teams chat)
- Note if PG has responded recently (from IcM bridges, chat, or email)
- Flag if status has changed since last report

### Step 6: Categorize into Priority Buckets

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

### Step 7: Build Key Asks

Generate Key Asks list at the TOP of the report:
- Include ALL P1 items (always)
- Include P2 items that do NOT have BOTH a workaround AND an ETA
- EXCLUDE P2 items that have a working workaround AND a confirmed ETA
- For each Key Ask: state the issue, age, what's needed from PG

### Step 8: Generate HTML Email Body

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

### Step 9: Create Draft Email in Outlook (Outlook MCP)

Use the **Outlook MCP server** `mail_create_draft` to create the email as a draft:
- **to_recipients**: `jinalmakwana@microsoft.com`
- **cc_recipients**: `Toby.James@microsoft.com`
- **subject**: `RACQ Support Ticket Summary - {date} (Generated by Agency)`
- **body**: The full HTML content generated in Step 8
- **content_type**: `html`
- **importance**: `high`

The draft appears in the user's Outlook Drafts folder for review, editing, and manual send.
This is safer than auto-sending — the user can verify the synthesized status before distributing.

> **Fallback**: If Outlook MCP is unavailable, fall back to saving a `.eml` file with
> `X-Unsent: 1` header and opening it with `Start-Process` to launch in Outlook compose window.

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
- "Analyze RACQ IcMs with no ISD ETA"
- "Look at Shruti's email and find IcMs with delays"
- "Draft an IcM delay analysis email for RACQ"

---

## IcM Delay Analysis (Escalation Report)

This sub-skill analyzes IcMs from Shruti's "RACQ Open Product Issues" email that have **no ISD ETA**,
investigates root causes of delays via IcM MCP, and drafts an escalation email.

### When to Trigger

Trigger this when the user asks about:
- IcMs with no ISD ETA or missing ETAs
- Root cause of IcM delays
- Why PG hasn't responded on RACQ tickets
- Drafting an escalation email about delayed IcMs

### IcM Exclusion Rules

Before reporting, apply these exclusion rules:

1. **IcM 51000000910153** (Custom column data not visible after transfer):
   - **SKIP** — this IcM is soft-closed. The actual issue is tracked in **IcM 51000000969890**
     (Save Button in Active Conversation form not saving notes). Always investigate 51000000969890 instead.

2. **IcM 51000000902213** (Active conversation form visibility / proactive engagement subgrid):
   - **SKIP** — this has a fix ETA (PG ETA for OCE region) and does not need escalation reporting,
     even if Shruti's email lists it without an ISD ETA.

### Step-by-Step Execution

#### Step 1: Read Shruti's Latest "Open Product Issues" Email (Mail MCP)

1. Use Mail MCP `SearchMessagesQueryParameters` with:
   `?$search="from:shruti subject:RACQ Open Product issues"&$top=3&$select=id,subject,from,receivedDateTime,bodyPreview`
2. Use `GetMessage` on the most recent result to get the full HTML body.
3. Parse the HTML table in the email body. The table has columns:
   - Sr., ADO#, Title, Severity, ICM, SR Raised Date, PG ETA, **ISD ETA**, ISD Comments
4. Extract all rows where **ISD ETA is empty/blank** — these are the IcMs to investigate.
5. From each row, capture: IcM ID (from URL), Title, Severity, ISD Comments.

#### Step 2: Apply Exclusion Rules

Filter out excluded IcMs per the rules above:
- Remove 51000000910153 → replace with 51000000969890
- Remove 51000000902213 entirely

#### Step 3: Query Each IcM for Details and Root Cause (IcM MCP)

For **each remaining IcM**, call these IcM MCP tools to build the delay picture:

1. **`get_incident_details_by_id`** (param: `incidentId` as integer)
   - Extract: **State**, **Owner** (contactAlias), **Owning Team**, **Created Date** (to calculate age),
     **Severity**, **AssignedTo**, **HowFixed**, **Mitigation data**.
   - Check custom field `PG_Review_Requested` and `PG_Review_Reason` for PG engagement status.

2. **`get_incident_context`** (param: `incidentId` as string)
   - Extract bridge summaries, discussion history, investigation notes.
   - Look for: PG responses, team transfers, escalation notes, blockers.

3. **`get_ai_summary`** (param: `incidentId` as string)
   - Use AI summary to fill gaps from discussion.

**Parallelize**: Call `get_incident_details_by_id` for ALL IcMs simultaneously, then
call `get_incident_context` and `get_ai_summary` for all simultaneously.

#### Step 4: Analyze Root Cause of Delay for Each IcM

For each IcM, determine the delay root cause by checking:

| Check | Indicates |
|-------|-----------|
| `PG_Review_Reason` = "Request assistance ignored" | PG not engaging |
| Multiple team transfers (check `redirectToTeamPublicId` history) | Ticket bounced across teams |
| `state` = ACTIVE but `isAcknowledged` = false after >7 days | Triage delays |
| `acknowledgeTime` >> `createdDate` (weeks/months gap) | Late acknowledgment |
| Dependency on another IcM (from email ISD Comments) | Blocked by upstream ticket |
| `mitigateData` present but no resolution | Mitigated but not fixed |
| No `get_incident_context` data available | Limited investigation visibility |

Synthesize a concise root cause statement for each IcM.

#### Step 5: Generate HTML Email Draft

Create an HTML email with per-IcM tables containing:
- **Status** (from IcM state)
- **Owner** (contactAlias + owning team name)
- **Open Since** (created date + calculated age in days)
- **Root Cause of Delay** (synthesized from Step 4)

Include a **Summary of Delay Patterns** table at the bottom grouping common issues:
- Pending PG Response (no action)
- Dependency chain blocking resolution
- Ticket bounced across teams
- PG Review request ignored
- Regional deployment lag

End with a **Key Takeaway** paragraph highlighting the core theme.

#### Step 6: Create Draft Email (Mail MCP)

Use Mail MCP `CreateDraftMessage` to create the email:
- **to**: `shgulat@microsoft.com`, `Toby.James@microsoft.com`
- **subject**: `RACQ | IcM Delay Analysis — Tickets with No ISD ETA`
- **contentType**: `HTML`
- **body**: The full HTML content from Step 5

The draft appears in Outlook Drafts for review before sending.
