# RACQ Daily Report Generator
# Runs daily before 9AM Melbourne time via Task Scheduler
# Queries IcM MCP API for live data, generates HTML email, opens in Outlook

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# ── CONFIG ──────────────────────────────────────────────────────────────
$ReportDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$AzCliPath = "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"
$To = "jinalmakwana@microsoft.com"
$Cc = "Toby.James@microsoft.com"
$Scope = "api://icmmcpapi-prod/mcp.tools"
$Endpoint = "https://icm-mcp-prod.azure-api.net/v1/"

# IcM mapping: IcM ID -> [Severity, SR#, Short Title]
$IcmMap = [ordered]@{
    21000000890308 = @("High",     "2601270030006343", "IVR Welcome Message Stutters in Audio Playback")
    51000000978035 = @("High",     "2604070030006100", "Missing Conversation Summary with custom Omnichannel Agent role")
    51000000910153 = @("High",     "2602130030002880", "Custom column data not visible after conversation transfer")
    21000000968169 = @("High",     "2603270030005523", "Script Errors - LiveWorkItemId session context error on Voice conversations")
    51000000902213 = @("High",     "2602120030007445", "Active conversation form disappears on navigation")
    51000000958196 = @("High",     "2603120030002536", "Automated Messages not working for Voice Channels")
    21000000863530 = @("High",     "2601130030007340", "Inconsistent Call and Chat Notifications for Agents")
    51000000899493 = @("High",     "2601300030001832", "Automated Messages EN-AU - only one record exists")
    21000000930073 = @("Critical", "2603040030006553", "Solution import failure - ISV code reduced open transaction count")
    21000000895436 = @("Critical", "2601150030006239", "Copilot Studio Agent Publish Fails (IdentifierNotRecognized)")
}

# Child IcM for 51000000910153 (OOB Save button issue)
$ChildIcm = @{
    Id    = 51000000969890
    Title = "OOB Save button saves zero data"
}

# ── FUNCTIONS ───────────────────────────────────────────────────────────
function Get-AzToken {
    $token = & $AzCliPath account get-access-token --scope $Scope --query "accessToken" -o tsv 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Azure CLI auth failed. Run 'az login' first. Error: $token" }
    return $token.Trim()
}

function Call-IcmTool($Token, $ToolName, $ToolArgs) {
    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
        "Accept"        = "application/json, text/event-stream"
    }
    $body = @{
        jsonrpc = "2.0"
        method  = "tools/call"
        id      = [int](Get-Random -Maximum 999999)
        params  = @{ name = $ToolName; arguments = $ToolArgs }
    } | ConvertTo-Json -Depth 5

    $raw = Invoke-WebRequest -Uri $Endpoint -Method POST -Headers $headers -Body $body -UseBasicParsing
    $content = $raw.Content -replace "^event: message\s*data: ", ""
    $parsed = $content | ConvertFrom-Json
    if ($parsed.result.content) {
        return ($parsed.result.content | Where-Object { $_.type -eq "text" } | ForEach-Object { $_.text }) -join "`n"
    }
    return $null
}

function Get-AgeDays($CreatedDate) {
    $created = [DateTime]::Parse($CreatedDate)
    return [math]::Floor(([DateTime]::UtcNow - $created).TotalDays)
}

# ── MAIN ────────────────────────────────────────────────────────────────
Write-Host "$(Get-Date) - RACQ Daily Report Generator starting..." -ForegroundColor Cyan

# 1. Authenticate
Write-Host "Authenticating with Azure..."
$token = Get-AzToken
Write-Host "  [OK] Token acquired"

# 2. Query all IcMs
Write-Host "Querying IcM portal for live data..."
$incidents = @{}
foreach ($icmId in $IcmMap.Keys) {
    Write-Host "  Querying IcM $icmId..."
    try {
        $raw = Call-IcmTool $token "get_incident_details_by_id" @{ incidentId = [long]$icmId }
        $data = $raw | ConvertFrom-Json
        $incidents[$icmId] = @{
            State       = $data.state
            Age         = Get-AgeDays $data.createdDate
            Owner       = if ($data.contactAlias) { $data.contactAlias } else { "-" }
            AssignedTo  = if ($data.assignedTo) { $data.assignedTo } else { "-" }
            CreatedDate = $data.createdDate
            Title       = $data.title -replace "^\[Premier\]\s*", ""
        }
        Write-Host "    [OK] $($data.state) | Age: $($incidents[$icmId].Age) days"
    } catch {
        Write-Host "    [FAIL] Failed: $_" -ForegroundColor Red
        $incidents[$icmId] = @{ State = "UNKNOWN"; Age = "?"; Owner = "?"; Title = $IcmMap[$icmId][2] }
    }
}

# Also query child IcM
Write-Host "  Querying child IcM $($ChildIcm.Id)..."
try {
    $childRaw = Call-IcmTool $token "get_incident_details_by_id" @{ incidentId = [long]$ChildIcm.Id }
    $childData = $childRaw | ConvertFrom-Json
    $ChildIcm.State = $childData.state
    $ChildIcm.Age = Get-AgeDays $childData.createdDate
    Write-Host "    [OK] $($childData.state) | Age: $($ChildIcm.Age) days"
} catch {
    $ChildIcm.State = "UNKNOWN"
    $ChildIcm.Age = "?"
}

# 3. Categorize into priority buckets
# P1: No workaround / Go-live blockers / Workaround not working
# P2: Workaround exists but insufficient / Has ETA
# P3: Resolved / Closed

$P1 = @(); $P2 = @(); $P3 = @()

foreach ($icmId in $IcmMap.Keys) {
    $inc = $incidents[$icmId]
    $sev = $IcmMap[$icmId][0]
    $sr  = $IcmMap[$icmId][1]
    $title = $IcmMap[$icmId][2]

    $entry = @{
        IcmId = $icmId; Sev = $sev; SR = $sr; Title = $title
        State = $inc.State; Age = $inc.Age; Owner = $inc.Owner
    }

    if ($inc.State -in @("RESOLVED", "CLOSED")) {
        $P3 += $entry
    }
    elseif ($icmId -eq 21000000890308) {
        $entry.Status = 'Root cause confirmed - 2 IVR bot instances joining same call. Transferred to Skype MediaPaaS / Champs team. **No workaround. No ETA.**'
        $P1 += $entry
    }
    elseif ($icmId -eq 51000000978035) {
        $entry.Status = 'Copilot Summary init failures and org setting read errors with custom agent role. **No workaround. No ETA.**'
        $P1 += $entry
    }
    elseif ($icmId -eq 21000000968169) {
        $entry.Status = 'Opening any voice conversation throws LiveWorkItemId error; conversation does not load. PG validating script fix. **No workaround. Pending PG confirmation.**'
        $P1 += $entry
    }
    elseif ($icmId -eq 51000000910153) {
        $childLink = "https://portal.microsofticm.com/imp/v5/incidents/details/$($ChildIcm.Id)"
        $entry.Status = "OOB Save button not working - child IcM [$($ChildIcm.Id)]($childLink) ($($ChildIcm.State), $($ChildIcm.Age) days) confirms save writes zero data. **No effective workaround.**"
        $P1 += $entry
    }
    elseif ($icmId -eq 51000000902213) {
        $entry.Status = 'Temp workaround: avoid form dropdown. Permanent fix **ETA 15 May**.'
        $P2 += $entry
    }
    elseif ($icmId -eq 51000000958196) {
        $entry.Status = 'Stale Dataverse record root cause. Manual Dataverse cleanup workaround in place but operationally heavy. **ETA pending PG prioritization.**'
        $P2 += $entry
    }
    else {
        $entry.Status = "Status from IcM: $($inc.State)"
        $P2 += $entry
    }
}

# 4. Build Key Asks (exclude items with workaround + ETA)
$keyAsks = @()
foreach ($item in $P1) {
    $keyAsks += $item
}
# From P2, only include items WITHOUT (workaround + ETA)
foreach ($item in $P2) {
    # 51000000902213 has workaround + ETA - exclude
    if ($item.IcmId -ne 51000000902213) {
        $keyAsks += $item
    }
}

# 5. Generate HTML
Write-Host "Generating HTML report..."
$today = (Get-Date).ToString("dd MMMM yyyy")
$todayShort = (Get-Date).ToString("ddMMMyyyy")

function Convert-ToHtml($Text) {
    # Convert **bold** to <b>bold</b>
    $result = $Text -replace '\*\*([^*]+)\*\*', '<b>$1</b>'
    # Convert [text](url) to <a href="url">text</a>
    $result = $result -replace '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>'
    # Convert & to &amp; (but not already-encoded)
    $result = $result -replace '&(?!amp;|lt;|gt;|#)', '&amp;'
    return $result
}

function Build-TableRow($Item, $ColorState) {
    $stateColor = switch -Wildcard ($Item.State) {
        "ACTIVE*"    { "#c00" }
        "MITIGATED*" { "#e67e00" }
        "RESOLVED*"  { "#28a745" }
        default      { "#333" }
    }
    return @"
  <tr>
    <td>$($Item.Sev)</td>
    <td>$($Item.SR)</td>
    <td><a href="https://portal.microsofticm.com/imp/v5/incidents/details/$($Item.IcmId)">$($Item.IcmId)</a></td>
    <td>$($Item.Title)</td>
    <td style="text-align:center;"><b>$($Item.Age)</b></td>
    <td style="color:$stateColor;"><b>$($Item.State)</b></td>
    <td>$(Convert-ToHtml $Item.Owner)</td>
    <td>$(Convert-ToHtml $Item.Status)</td>
  </tr>
"@
}

function Build-ResolvedRow($Item) {
    return @"
  <tr>
    <td>$($Item.Sev)</td>
    <td>$($Item.SR)</td>
    <td><a href="https://portal.microsofticm.com/imp/v5/incidents/details/$($Item.IcmId)">$($Item.IcmId)</a></td>
    <td>$($Item.Title)</td>
    <td style="text-align:center;">$($Item.Age)</td>
    <td style="color:#28a745;"><b>$($Item.State)</b></td>
    <td>&#x2705; $($Item.State).</td>
  </tr>
"@
}

# Key Asks HTML
$keyAsksHtml = ""
foreach ($ka in $keyAsks) {
    $statusClean = (Convert-ToHtml $ka.Status) -replace '<b>','' -replace '</b>','' -replace '<a [^>]+>','(' -replace '</a>',')'
    $keyAsksHtml += "  <li><b>$($ka.Title) (<a href=`"https://portal.microsofticm.com/imp/v5/incidents/details/$($ka.IcmId)`">$($ka.IcmId)</a>)</b> - $($ka.Age) days open, $($ka.State). $statusClean</li>`n"
}

# P1 rows
$p1Rows = ($P1 | ForEach-Object { Build-TableRow $_ }) -join "`n"
# P2 rows
$p2Rows = ($P2 | ForEach-Object { Build-TableRow $_ }) -join "`n"
# P3 rows
$p3Rows = ($P3 | ForEach-Object { Build-ResolvedRow $_ }) -join "`n"

$html = @"
<html>
<head><meta charset="utf-8"></head>
<body style="font-family: Segoe UI, Calibri, Arial, sans-serif; font-size: 11pt; color: #333;">

<p>Hi Team,</p>

<p>Please find below the latest status of the RACQ High &amp; Critical IcMs (sourced live from IcM portal on $today), re-prioritized based on workaround availability and go-live impact.</p>

<!-- KEY ASKS -->
<h3 style="color:#c00; margin-bottom:6px;">&#x1F4CC; Key Asks</h3>
<ol>
$keyAsksHtml</ol>

<hr style="border:1px solid #ccc;"/>

<!-- PRIORITY 1 -->
<h3 style="color:#c00; margin-bottom:6px;">&#x1F534; Priority 1 – No Workaround / Go-Live Blockers</h3>

<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-size:10pt; width:100%;">
<thead style="background:#f8d7da;">
  <tr><th>Sev</th><th>SR #</th><th>IcM</th><th>Issue</th><th>Age (days)</th><th>State</th><th>Owner</th><th>Latest Status</th></tr>
</thead>
<tbody>
$p1Rows
</tbody>
</table>

<br/>

<!-- PRIORITY 2 -->
<h3 style="color:#e67e00; margin-bottom:6px;">&#x1F7E0; Priority 2 – Workaround Exists but Insufficient / ETA Pending</h3>

<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-size:10pt; width:100%;">
<thead style="background:#fff3cd;">
  <tr><th>Sev</th><th>SR #</th><th>IcM</th><th>Issue</th><th>Age (days)</th><th>State</th><th>Owner</th><th>Latest Status</th></tr>
</thead>
<tbody>
$p2Rows
</tbody>
</table>

<br/>

<!-- PRIORITY 3 -->
<h3 style="color:#28a745; margin-bottom:6px;">&#x2705; Priority 3 – Resolved / Closed</h3>

<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-size:10pt; width:100%;">
<thead style="background:#d4edda;">
  <tr><th>Sev</th><th>SR #</th><th>IcM</th><th>Issue</th><th>Age (days)</th><th>State</th><th>Latest Status</th></tr>
</thead>
<tbody>
$p3Rows
</tbody>
</table>

<br/>

<p><i>Data sourced live from IcM portal on $today using IcM MCP Server. Report auto-generated at $(Get-Date -Format 'HH:mm') AEST.</i></p>

<p>Please let me know if there are any questions or updates.</p>

<p>Thanks,<br/>Jinal</p>

<br/>
<p style="font-size:10pt; color:#555;">
<b>Jinal Makwana</b><br/>
Senior FastTrack Solution Architect<br/>
Dynamics 365 Apps and Common Data Service R&amp;D<br/><br/>
Mobile +61 430407653<br/>
Office +61 (3) 93204331<br/>
LinkedIn: <a href="https://www.linkedin.com/in/d365lady/">https://www.linkedin.com/in/d365lady/</a>
</p>

</body>
</html>
"@

# 6. Save HTML and generate .eml
$htmlPath = Join-Path $ReportDir "racq-report.html"
$emlPath = Join-Path $ReportDir "RACQ-Daily-Report-$todayShort.eml"

[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($false))

$eml = @"
From: $To
To: $To
Cc: $Cc
Subject: RACQ Support Ticket Summary - $today (Generated by Agency)
Date: $((Get-Date).ToUniversalTime().ToString("ddd, dd MMM yyyy HH:mm:ss +0000"))
MIME-Version: 1.0
Content-Type: text/html; charset=utf-8
X-Unsent: 1

$html
"@

[System.IO.File]::WriteAllText($emlPath, $eml, [System.Text.UTF8Encoding]::new($false))
Write-Host "  Done! Report saved: $emlPath"

# 7. Open in Outlook
Write-Host "Opening in Outlook..."
Start-Process $emlPath
Write-Host "  Done! Email opened in Outlook compose window."
Write-Host "$(Get-Date) - RACQ Daily Report Generator completed." -ForegroundColor Green
