# Power Automate Flow: RACQ Excel to JSON Cache
# Step-by-step guide to create the daily flow

## Flow Name: "RACQ Excel Cache Refresh"

## Quick Setup (5 minutes)

### Step 1: Go to Power Automate
Open: https://make.powerautomate.com

### Step 2: Create new Scheduled Flow
- Click **+ Create** > **Scheduled cloud flow**
- Name: `RACQ Excel Cache Refresh`
- Start: Tomorrow at **8:00 AM**
- Repeat every: **1 day**
- On these days: Mon, Tue, Wed, Thu, Fri
- Click **Create**

### Step 3: Add "List rows present in a table" action
- Search for **Excel Online (Business)**
- Select **List rows present in a table**
- **Location**: `shgulat@microsoft.com` OneDrive (or the SharePoint location)
- **Document Library**: Documents
- **File**: `/Microsoft Teams Chat Files/RACQ_Support_Tickets.xlsx`
- **Table**: Select the table from the "Support tickets" sheet
  (If no table exists, Shruti needs to format the data range as a table first)

### Step 4: Add "Compose" action to build JSON
- Add action: **Data Operations > Compose**
- In the **Inputs** field, paste this expression:

```json
{
  "lastUpdated": "@{utcNow()}",
  "source": "RACQ_Support_Tickets.xlsx (Power Automate)",
  "incidents": @{body('List_rows_present_in_a_table')?['value']}
}
```

Or for more control, use a **Select** action first to map columns:
- From: `body('List_rows_present_in_a_table')?['value']`
- Map:
  - icmId: `@{item()?['IcM #']}`
  - sr: `@{item()?['SR #']}`
  - severity: `@{item()?['Updated Severity']}`
  - title: `@{item()?['Description']}`
  - status: `@{item()?['Status']}`
  - eta: `@{item()?['ETA']}`
  - workaround: `@{item()?['Workaround']}`
  - nextSteps: `@{item()?['Next Steps']}`
  - comments: `@{item()?['Comments']}`

### Step 5: Add "Create file" action to save JSON
- Search for **OneDrive for Business**
- Select **Create file** (or **Update file** if it already exists)
- **Folder path**: `/Agency/RACQ-Agency-CLI/`
- **File name**: `racq-icm-cache.json`
- **File content**: Output from the Compose/Select step

> **Important**: Use "Update file" instead of "Create file" after the first run,
> or add a "Delete file" step before "Create file" to avoid duplicates.

### Step 6: Save and Test
- Click **Save**
- Click **Test** > **Manually** > **Run flow**
- Check that `racq-icm-cache.json` appears in your local OneDrive sync folder:
  `C:\Users\jinalmakwana\OneDrive - Microsoft\Agency\RACQ-Agency-CLI\`

## Alternative: Simpler Approach (just save the raw Excel)

If building JSON is too complex, the flow can simply:
1. **Get file content** from Shruti's OneDrive (RACQ_Support_Tickets.xlsx)
2. **Create/Update file** in YOUR OneDrive (`/Agency/RACQ-Agency-CLI/RACQ_Support_Tickets.xlsx`)

Then the PowerShell script reads the local Excel directly using the ImportExcel module.

### To install ImportExcel:
```powershell
Install-Module ImportExcel -Force -Scope CurrentUser
```

### Script reads Excel directly:
```powershell
$data = Import-Excel ".\RACQ_Support_Tickets.xlsx" -WorksheetName "Support tickets"
$highCrit = $data | Where-Object { $_."Updated Severity" -in @("High", "Critical") }
```

This is simpler and keeps the Excel as the single source of truth.

## Flow Timing
- **Power Automate**: 8:00 AM (reads Excel, saves to OneDrive)
- **OneDrive sync**: ~1-5 minutes to sync locally
- **Task Scheduler**: 8:30 AM (reads local file + queries IcM MCP API)
