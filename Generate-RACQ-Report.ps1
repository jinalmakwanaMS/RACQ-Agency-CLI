# RACQ Daily Report Generator
# Runs daily before 9AM Melbourne time via Task Scheduler
# Queries IcM MCP API for live data, generates HTML email, opens in Outlook

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# ── CONFIG ──────────────────────────────────────────────────────────────
$ReportDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$AzCliPath = "C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"
$ExcelPath = Join-Path $ReportDir "RACQ_Support_Tickets.xlsx"
$CachePath = Join-Path $ReportDir "racq-icm-cache.json"
$To = "jinalmakwana@microsoft.com"
$Cc = "Toby.James@microsoft.com"
$Scope = "api://icmmcpapi-prod/mcp.tools"
$Endpoint = "https://icm-mcp-prod.azure-api.net/v1/"

# ── LOAD EXCEL DATA ────────────────────────────────────────────────────
# Priority: Live Excel (synced by Power Automate) > JSON cache fallback
$cache = $null

if (Test-Path $ExcelPath) {
    Write-Host "Reading live Excel file: $ExcelPath"
    Write-Host "  Last modified: $((Get-Item $ExcelPath).LastWriteTime)"
    Import-Module ImportExcel -ErrorAction Stop

    $rows = Import-Excel $ExcelPath -WorksheetName "Support tickets" -ErrorAction Stop
    $highCrit = $rows | Where-Object {
        $_."Updated Severity" -in @("High", "Critical") -or
        $_."Old Severity" -in @("High", "Critical")
    }

    # Build cache object from live Excel
    $cacheIncidents = @{}
    foreach ($row in $highCrit) {
        $icmId = "$($row.'IcM #')".Trim()
        if (-not $icmId -or $icmId -eq "") { continue }

        $sev = if ($row."Updated Severity") { $row."Updated Severity" } else { $row."Old Severity" }
        $cacheIncidents[$icmId] = @{
            sr          = "$($row.'SR #')".Trim()
            severity    = $sev
            title       = "$($row.'Description')".Trim()
            excelStatus = "$($row.'Status')".Trim()
            eta         = "$($row.'ETA')".Trim()
            workaround  = if ($row.'Workaround') { "$($row.'Workaround')".Trim() } else { "" }
            nextSteps   = if ($row.'Next Steps') { "$($row.'Next Steps')".Trim() } else { "" }
            comments    = if ($row.'Comments') { "$($row.'Comments')".Trim() } else { "" }
            notes       = ""
        }
    }

    # Build cache structure
    $cache = [PSCustomObject]@{
        lastUpdated    = (Get-Date).ToString("o")
        source         = "RACQ_Support_Tickets.xlsx (live)"
        incidents      = [PSCustomObject]$cacheIncidents
        childIncidents = [PSCustomObject]@{}
    }

    # Auto-save as JSON cache for reference
    $cache | ConvertTo-Json -Depth 5 | Set-Content $CachePath -Encoding UTF8
    Write-Host "  [OK] Loaded $($cacheIncidents.Count) High/Critical items from Excel"

} elseif (Test-Path $CachePath) {
    Write-Host "No Excel file found. Using JSON cache: $CachePath"
    $cache = Get-Content $CachePath -Raw | ConvertFrom-Json
    Write-Host "  Cache from: $($cache.lastUpdated)"

} else {
    throw "No data source available. Place RACQ_Support_Tickets.xlsx in $ReportDir or ensure racq-icm-cache.json exists."
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

# 2. Build IcM list from cache
$icmIds = @()
$childIcmIds = @()
if ($cache) {
    foreach ($prop in $cache.incidents.PSObject.Properties) {
        $icmIds += $prop.Name
    }
    foreach ($prop in $cache.childIncidents.PSObject.Properties) {
        $childIcmIds += $prop.Name
    }
} else {
    throw "No cache file found. Cannot proceed without IcM registry."
}

# 3. Query all IcMs from IcM MCP API
Write-Host "Querying IcM portal for live data ($($icmIds.Count) incidents + $($childIcmIds.Count) child)..."
$incidents = @{}
foreach ($icmId in $icmIds) {
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
        $cachedItem = $cache.incidents.$icmId
        $incidents[$icmId] = @{ State = "UNKNOWN"; Age = "?"; Owner = "?"; Title = if ($cachedItem) { $cachedItem.title } else { "Unknown" } }
    }
}

# Also query child IcMs
$childIncidents = @{}
foreach ($childId in $childIcmIds) {
    Write-Host "  Querying child IcM $childId..."
    try {
        $childRaw = Call-IcmTool $token "get_incident_details_by_id" @{ incidentId = [long]$childId }
        $childData = $childRaw | ConvertFrom-Json
        $childIncidents[$childId] = @{
            State = $childData.state
            Age   = Get-AgeDays $childData.createdDate
        }
        Write-Host "    [OK] $($childData.state) | Age: $($childIncidents[$childId].Age) days"
    } catch {
        $childIncidents[$childId] = @{ State = "UNKNOWN"; Age = "?" }
        Write-Host "    [FAIL] Failed: $_" -ForegroundColor Red
    }
}

# 4. Categorize into priority buckets using live IcM state + cached Excel data
# P1: No workaround / Go-live blockers / Workaround not working
# P2: Workaround exists but insufficient / Has ETA
# P3: Resolved / Closed

$P1 = @(); $P2 = @(); $P3 = @()

foreach ($icmId in $icmIds) {
    $inc = $incidents[$icmId]
    $cached = $cache.incidents.$icmId

    $entry = @{
        IcmId = $icmId
        Sev   = $cached.severity
        SR    = $cached.sr
        Title = $cached.title
        State = $inc.State
        Age   = $inc.Age
        Owner = $inc.Owner
    }

    # P3: Resolved/Closed (from live IcM state OR Excel status)
    if ($inc.State -in @("RESOLVED", "CLOSED") -or $cached.excelStatus -eq "Closed") {
        $entry.Status = $cached.nextSteps
        $P3 += $entry
        continue
    }

    # Build status from Excel cache data
    $statusParts = @()
    if ($cached.nextSteps) { $statusParts += $cached.nextSteps }
    if ($cached.notes -and $cached.notes -ne "Resolved.") { $statusParts += $cached.notes }

    $hasWorkaround = ($cached.workaround -and $cached.workaround -ne "None")
    $workaroundFailed = ($cached.workaround -match "NOT working|not working|failed|zero data|unsustainable")
    $hasEta = ($cached.eta -and $cached.eta -ne "Not specified")

    # Check for child IcM info
    if ($cached.childIcm) {
        $childId = $cached.childIcm
        $childInc = $childIncidents[$childId]
        $childCached = $cache.childIncidents.$childId
        if ($childInc) {
            $childLink = "https://portal.microsofticm.com/imp/v5/incidents/details/$childId"
            $statusParts += "Child IcM [$childId]($childLink) ($($childInc.State), $($childInc.Age) days): $($childCached.title)"
        }
    }

    # Determine priority bucket
    # P1: No workaround, or workaround failed/not working
    if (-not $hasWorkaround -or $workaroundFailed) {
        $woText = if ($workaroundFailed) { "**Workaround not working.**" } else { "**No workaround.**" }
        $etaText = if ($hasEta) { "ETA: $($cached.eta)" } else { "**No ETA.**" }
        $entry.Status = ($statusParts -join " ") + " $woText $etaText"
        $P1 += $entry
    }
    # P2: Has workaround (working) or has ETA
    else {
        $woText = "Workaround: $($cached.workaround)."
        $etaText = if ($hasEta) { "**ETA: $($cached.eta).**" } else { "**ETA pending.**" }
        $entry.Status = ($statusParts -join " ") + " $woText $etaText"
        $P2 += $entry
    }
}

# 5. Build Key Asks (exclude items with working workaround + confirmed ETA)
$keyAsks = @()
foreach ($item in $P1) {
    $keyAsks += $item
}
foreach ($item in $P2) {
    $cachedItem = $cache.incidents.($item.IcmId)
    $hasWorkingWorkaround = ($cachedItem.workaround -and $cachedItem.workaround -ne "None" -and $cachedItem.workaround -notmatch "NOT working|not working|failed|unsustainable")
    $hasConfirmedEta = ($cachedItem.eta -and $cachedItem.eta -ne "Not specified")
    # Exclude from Key Asks only if BOTH workaround is working AND ETA is confirmed
    if (-not ($hasWorkingWorkaround -and $hasConfirmedEta)) {
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
