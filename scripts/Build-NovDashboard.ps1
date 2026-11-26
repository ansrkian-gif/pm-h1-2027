# Build the phone dashboard from the TeleOWS Excel extract.
# Counting rule: each completed PM work order = 1. Minimum is pm-config.json dailyTarget.
# Friday is a rest day. A working day is scored after 8:00 PM Kuwait time.
param(
    [string]$PeriodStart = "",
    [int]$DailyTarget = 0,
    [string]$ExtractPath = "",
    [string]$ThroughDate = "",
    [string]$DashName = "dashboard",
    [switch]$NoWrite
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Dash = Join-Path $Root $DashName
$ConfigPath = Join-Path $Root "pm-config.json"
$RosterPath = Join-Path $Root "fme_roster.txt"

$cfg = @{
    periodStart = "2026-11-01"
    dailyTarget = 4
    updateTime  = "20:00"
}
if (Test-Path -LiteralPath $ConfigPath) {
    $fileCfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($fileCfg.periodStart) { $cfg.periodStart = [string]$fileCfg.periodStart }
    if ($fileCfg.dailyTarget) { $cfg.dailyTarget = [int]$fileCfg.dailyTarget }
    if ($fileCfg.updateTime) { $cfg.updateTime = [string]$fileCfg.updateTime }
}
if ($PeriodStart) { $cfg.periodStart = $PeriodStart }
if ($DailyTarget -gt 0) { $cfg.dailyTarget = $DailyTarget }

$Category = "OGK Active and Passive Routine Maintenance"
$FmeAlias = @{
    "Mohd Yasin Mohd Matin Ansari" = "Nabijohn Piyarjan Piyarjan"
}

function Get-KuwaitNow {
    return [DateTimeOffset]::UtcNow.ToOffset([TimeSpan]::FromHours(3))
}

function Get-ShortName([string]$name) {
    $parts = @($name -split "\s+" | Where-Object { $_ })
    if ($parts.Count -ge 2) { return ($parts[0] + " " + $parts[1]) }
    return $name
}

function Read-Roster {
    $names = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $RosterPath)) { return $names }
    foreach ($line in [IO.File]::ReadAllLines($RosterPath)) {
        $line = $line.Trim()
        if (-not $line -or $line.StartsWith("#")) { continue }
        if ($FmeAlias.ContainsKey($line)) { $line = $FmeAlias[$line] }
        if (-not $names.Contains($line)) { $names.Add($line) }
    }
    return $names
}

function Find-Extract {
    if ($ExtractPath -and (Test-Path -LiteralPath $ExtractPath)) { return (Resolve-Path -LiteralPath $ExtractPath).Path }
    $envPath = $env:EXTRACT_PATH
    if ($envPath -and (Test-Path -LiteralPath $envPath)) { return (Resolve-Path -LiteralPath $envPath).Path }
    $cloud = Join-Path $Root "cloud\extract.xlsx"
    if (Test-Path -LiteralPath $cloud) { return $cloud }
    $files = @(Get-ChildItem -LiteralPath $Root -Filter "*.xlsx" -File | Where-Object { $_.Name -notlike "~$*" })
    if ($files.Count -eq 0) { return $null }
    return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

function Get-ColNumber([string]$cellRef) {
    $letters = ($cellRef -replace "\d", "")
    $n = 0
    foreach ($ch in $letters.ToCharArray()) {
        $n = $n * 26 + ([int][char]$ch - 64)
    }
    return $n
}

function Read-SharedStrings([xml]$ssXml) {
    $list = New-Object System.Collections.Generic.List[string]
    if (-not $ssXml -or -not $ssXml.sst) { return $list }
    foreach ($si in @($ssXml.sst.si)) {
        if ($null -eq $si) { $list.Add(""); continue }
        if ($si.t) {
            $list.Add([string]$si.t)
        } elseif ($si.r) {
            $text = ""
            foreach ($run in @($si.r)) { $text += [string]$run.t }
            $list.Add($text)
        } else {
            $list.Add("")
        }
    }
    return $list
}

function Read-WorkOrders([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempRoot = [IO.Path]::GetTempPath()
    $tmp = Join-Path $tempRoot ("pm-nov2027-" + [guid]::NewGuid().ToString("n") + ".xlsx")
    Copy-Item -LiteralPath $path -Destination $tmp
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
    try {
        function Read-Entry([string]$name) {
            $entry = $zip.GetEntry($name.Replace("\", "/"))
            if (-not $entry) { return $null }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        $ssText = Read-Entry "xl/sharedStrings.xml"
        $sheetText = Read-Entry "xl/worksheets/sheet1.xml"
        if (-not $sheetText) { throw "Excel file has no sheet1." }
        function ConvertTo-SimpleXml([string]$text) {
            $clean = [regex]::Replace($text, '\sxmlns="[^"]*"', "")
            return [xml]$clean
        }
        function Get-CellText($cell, $strings) {
            if (-not $cell) { return "" }
            $kind = [string]$cell.t
            if ($kind -eq "s") {
                $val = [string]$cell.v
                if ($val -eq "") { return "" }
                return [string]$strings[[int]$val]
            }
            if ($kind -eq "inlineStr") {
                if ($cell.is.t) { return [string]$cell.is.t }
                $text = ""
                foreach ($run in @($cell.is.r)) { $text += [string]$run.t }
                return $text
            }
            return [string]$cell.v
        }
        $strings = @()
        if ($ssText) { $strings = @(Read-SharedStrings (ConvertTo-SimpleXml $ssText)) }
        $sheet = ConvertTo-SimpleXml $sheetText
    } finally {
        $zip.Dispose()
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }

    $rows = @($sheet.worksheet.sheetData.row)
    if ($rows.Count -eq 0) { return @() }

    $headerRowIndex = 0
    for ($h = 0; $h -lt [Math]::Min(8, $rows.Count); $h++) {
        $blob = ""
        foreach ($cell in @($rows[$h].c)) { $blob += " " + (Get-CellText $cell $strings) }
        if ($blob -match "Task\s*Id" -and $blob -match "Complete") { $headerRowIndex = $h; break }
    }
    $header = @{}
    foreach ($cell in @($rows[$headerRowIndex].c)) {
        $text = (Get-CellText $cell $strings).Trim()
        if ($text) { $header[$text] = Get-ColNumber ([string]$cell.r) }
    }
    $aliases = @{
        "Task Id" = @("Task Id", "Task ID", "TaskID")
        "Task Type" = @("Task Type", "TaskType")
        "Task Status" = @("Task Status", "Status")
        "Task Category" = @("Task Category")
        "Assign To FME Full Name" = @("Assign To FME Full Name", "Assign To")
        "Complete Time" = @("Complete Time", "Completion Time")
        "Accept Operator" = @("Accept Operator")
        "Arrive Operator" = @("Arrive Operator")
        "Complete Operator" = @("Complete Operator")
    }
    $colOf = @{}
    foreach ($canon in $aliases.Keys) {
        foreach ($name in $aliases[$canon]) {
            if ($header.ContainsKey($name)) { $colOf[$canon] = $header[$name]; break }
        }
        if (-not $colOf.ContainsKey($canon)) {
            $found = ($header.Keys | Select-Object -First 25) -join ", "
            throw "Excel is missing the '$canon' column. Found: $found"
        }
    }
    $header = $colOf

    $orders = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    for ($i = $headerRowIndex + 1; $i -lt $rows.Count; $i++) {
        $vals = @{}
        foreach ($cell in @($rows[$i].c)) {
            $ref = [string]$cell.r
            if (-not $ref) { continue }
            $vals[(Get-ColNumber $ref)] = Get-CellText $cell $strings
        }
        function Cell([string]$name) {
            $col = $header[$name]
            if ($vals.ContainsKey($col)) { return [string]$vals[$col] }
            return ""
        }
        $taskType = (Cell "Task Type").Trim()
        $category = (Cell "Task Category").Trim()
        $status = (Cell "Task Status").Trim().ToLowerInvariant()
        if ($taskType -ne "PM") { continue }
        if ($category -ne $Category) { continue }
        if ($status -ne "completed" -and $status -ne "closed") { continue }

        $taskId = (Cell "Task Id").Trim()
        if ($taskId -and $seen.ContainsKey($taskId)) { continue }
        if ($taskId) { $seen[$taskId] = $true }

        $ops = "$(Cell 'Accept Operator')|$(Cell 'Arrive Operator')|$(Cell 'Complete Operator')"
        if ($ops -notmatch "NTE") { continue }

        $completeRaw = (Cell "Complete Time").Trim()
        $day = $null
        if ($completeRaw -match "^(\d{4}-\d{2}-\d{2})") {
            $day = $Matches[1]
        } elseif ($completeRaw -match "^\d+(\.\d+)?$") {
            $serial = [double]$completeRaw
            if ($serial -gt 20000 -and $serial -lt 80000) {
                $day = [datetime]::FromOADate($serial).ToString("yyyy-MM-dd")
            }
        }
        if (-not $day) { continue }

        $fme = (Cell "Assign To FME Full Name").Trim()
        if (-not $fme) { continue }
        if ($FmeAlias.ContainsKey($fme)) { $fme = $FmeAlias[$fme] }

        $orders.Add([pscustomobject]@{ TaskId = $taskId; Fme = $fme; Day = $day })
    }
    return $orders
}

function Add-DaysIso([string]$iso, [int]$days) {
    return [datetime]::ParseExact($iso, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture).AddDays($days).ToString("yyyy-MM-dd")
}

function Test-Friday([string]$iso) {
    $d = [datetime]::ParseExact($iso, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    return $d.DayOfWeek -eq [DayOfWeek]::Friday
}

function Format-DayLabel([string]$iso) {
    $d = [datetime]::ParseExact($iso, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    return $d.ToString("ddd d MMM", [Globalization.CultureInfo]::InvariantCulture)
}

$kuwait = Get-KuwaitNow
$today = $kuwait.ToString("yyyy-MM-dd")
# 20:00 Kuwait time is the evening update. Today joins the totals then.
$pastEight = $kuwait.Hour -ge 20
$period = $cfg.periodStart
$target = [int]$cfg.dailyTarget
$roster = @(Read-Roster)

function Read-DashboardWorkOrders([string]$path) {
    $raw = [IO.File]::ReadAllText($path).Trim()
    $eq = $raw.IndexOf("=")
    if ($raw.StartsWith("window.DASHBOARD_DATA") -and $eq -ge 0) {
        $raw = $raw.Substring($eq + 1).Trim()
    }
    if ($raw.EndsWith(";")) { $raw = $raw.Substring(0, $raw.Length - 1) }
    $data = $raw | ConvertFrom-Json
    $orders = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($w in @($data.workOrders)) {
        $taskId = [string]$w.taskId
        if ($taskId -and $seen.ContainsKey($taskId)) { continue }
        if ($taskId) { $seen[$taskId] = $true }
        $status = ([string]$w.status).ToLowerInvariant()
        if ($status -ne "completed" -and $status -ne "closed") { continue }
        if (([string]$w.completeOperator) -notmatch "NTE") { continue }
        $day = [string]$w.completeDate
        if ($day -notmatch "^\d{4}-\d{2}-\d{2}$") { continue }
        $fme = ([string]$w.assignTo).Trim()
        if (-not $fme) { $fme = ([string]$w.fmeShort).Trim() }
        if (-not $fme) { continue }
        if ($FmeAlias.ContainsKey($fme)) { $fme = $FmeAlias[$fme] }
        $orders.Add([pscustomobject]@{ TaskId = $taskId; Fme = $fme; Day = $day })
    }
    return $orders
}

$sourceName = ""
$orders = @()
$extract = Find-Extract
$owsJs = Join-Path $Root "cloud\h1-dashboard-data.js"
if ($extract) {
    $sourceName = [IO.Path]::GetFileName($extract)
    Write-Host "Reading $extract"
    $orders = @(Read-WorkOrders $extract)
    Write-Host ("Work orders in file (NTE PM, completed): {0}" -f $orders.Count)
} elseif (Test-Path -LiteralPath $owsJs) {
    $sourceName = "PM H1 2027 OWS extract"
    Write-Host "Reading $owsJs"
    $orders = @(Read-DashboardWorkOrders $owsJs)
    Write-Host ("Work orders in OWS output (NTE PM, completed): {0}" -f $orders.Count)
}

$inPeriod = @($orders | Where-Object { $_.Day -ge $period -and $_.Day -le $today })
$phase = if ($today -lt $period) { "waiting" } else { "live" }
$latestInFile = $null
if ($inPeriod.Count -gt 0) {
    $latestInFile = [string]($inPeriod | ForEach-Object { $_.Day } | Sort-Object | Select-Object -Last 1)
}

# People who completed at least one work order in this period.
# A name with no November work orders is not scored as a daily miss.
$names = New-Object System.Collections.Generic.List[string]
foreach ($row in $inPeriod) {
    if (-not $names.Contains($row.Fme)) { $names.Add($row.Fme) }
}
if ($names.Count -eq 0) {
    foreach ($n in $roster) { if (-not $names.Contains($n)) { $names.Add($n) } }
}

# Closed working days are scored. Today stays "in progress" until 8:00 PM.
$scoreThrough = $null
if ($ThroughDate) {
    $scoreThrough = $ThroughDate
} elseif ($phase -eq "live") {
    if ($pastEight -or (Test-Friday $today)) {
        $scoreThrough = $today
    } else {
        $y = Add-DaysIso $today -1
        if ($y -ge $period) { $scoreThrough = $y }
    }
}
if ($scoreThrough -and $latestInFile -and $scoreThrough -gt $latestInFile) {
    $scoreThrough = $latestInFile
}

$closedDays = New-Object System.Collections.Generic.List[string]
if ($scoreThrough) {
    $cursor = [datetime]::ParseExact($period, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    $end = [datetime]::ParseExact($scoreThrough, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    while ($cursor -le $end) {
        $iso = $cursor.ToString("yyyy-MM-dd")
        if ($cursor.DayOfWeek -ne [DayOfWeek]::Friday) { $closedDays.Add($iso) }
        $cursor = $cursor.AddDays(1)
    }
}

$counts = @{}
foreach ($row in $inPeriod) {
    $key = $row.Fme + "|" + $row.Day
    if (-not $counts.ContainsKey($key)) { $counts[$key] = 0 }
    $counts[$key] = $counts[$key] + 1
}

function Get-Count([string]$fme, [string]$day) {
    $key = $fme + "|" + $day
    if ($counts.ContainsKey($key)) { return [int]$counts[$key] }
    return 0
}

$todayFriday = Test-Friday $today
$todayInPeriod = ($today -ge $period)
$todayClosed = $todayInPeriod -and $pastEight -and -not $todayFriday

$fmeOut = New-Object System.Collections.Generic.List[hashtable]
$todayRows = New-Object System.Collections.Generic.List[hashtable]
$teamActual = 0
$teamMissed = 0
$teamDaysMissed = 0
$teamDaysHit = 0
$todayActual = 0

foreach ($name in $names) {
    $actual = 0
    $missed = 0
    $daysMissed = 0
    $daysHit = 0
    $daily = New-Object System.Collections.Generic.List[hashtable]
    foreach ($day in $closedDays) {
        $n = Get-Count $name $day
        $shortfall = [Math]::Max(0, $target - $n)
        $hit = $n -ge $target
        $actual += $n
        $missed += $shortfall
        if ($hit) { $daysHit++ } else { $daysMissed++ }
        $daily.Add(@{
            day     = $day
            label   = (Format-DayLabel $day)
            actual  = $n
            target  = $target
            missed  = $shortfall
            hit     = $hit
        })
    }
    $closedN = $closedDays.Count
    $expected = $closedN * $target
    $efficiency = $null
    $average = $null
    if ($closedN -gt 0) {
        $efficiency = [int][Math]::Round((100.0 * $actual) / $expected)
        $average = [Math]::Round(($actual / $closedN), 1)
    }
    $shape = "NONE"
    if ($null -ne $efficiency) {
        if ($efficiency -ge 100) { $shape = "GOOD" }
        elseif ($efficiency -ge 75) { $shape = "WATCH" }
        else { $shape = "BAD" }
    }

    $todayCount = $null
    if ($todayInPeriod -and -not $todayFriday) {
        $todayCount = Get-Count $name $today
        $todayActual += $todayCount
        $todayMiss = [Math]::Max(0, $target - $todayCount)
        $todayRows.Add(@{
            name     = $name
            short    = (Get-ShortName $name)
            actual   = $todayCount
            target   = $target
            missed   = $todayMiss
            hit      = ($todayCount -ge $target)
            closed   = $todayClosed
        })
    } elseif ($todayInPeriod -and $todayFriday) {
        $extra = Get-Count $name $today
        if ($extra -gt 0) { $todayActual += $extra }
        $todayRows.Add(@{
            name     = $name
            short    = (Get-ShortName $name)
            actual   = $extra
            target   = $null
            missed   = 0
            hit      = $true
            closed   = $true
            rest     = $true
        })
    }

    $teamActual += $actual
    $teamMissed += $missed
    $teamDaysMissed += $daysMissed
    $teamDaysHit += $daysHit

    $fmeOut.Add(@{
        name        = $name
        short       = (Get-ShortName $name)
        actual      = $actual
        expected    = $expected
        average     = $average
        efficiency  = $efficiency
        daysHit     = $daysHit
        daysMissed  = $daysMissed
        totalMissed = $missed
        shape       = $shape
        today       = $todayCount
        daily       = @($daily.ToArray())
    })
}

$fmeSorted = @($fmeOut | Sort-Object { -$_.totalMissed }, { $_.name })
$todaySorted = @($todayRows | Sort-Object { -$_.missed }, { $_.name })

$fmeCount = $names.Count
$closedN = $closedDays.Count
$teamExpected = $fmeCount * $closedN * $target
$teamEfficiency = $null
$teamAverage = $null
if ($teamExpected -gt 0) {
    $teamEfficiency = [int][Math]::Round((100.0 * $teamActual) / $teamExpected)
    $teamAverage = [Math]::Round(($teamActual / ($fmeCount * $closedN)), 1)
}

$nextUpdate = "today at 8:00 PM"
if ($pastEight) { $nextUpdate = "tomorrow at 8:00 PM" }

$rules = @(
    "Counts completed PM work orders from 1 November 2026 through today.",
    "Each completed PM work order counts as 1. The same task id is counted once.",
    "Minimum 4 work orders for each FME on every working day.",
    "Friday is a rest day, same as the PM H1 2027 dashboard. It is not a miss.",
    "NTE accounts only. Mohd Yasin is counted under Nabijohn.",
    "Today stays in progress until 8:00 PM Kuwait time, then it is added to the totals.",
    "Efficiency = work orders done ÷ (working days × 4). Average = work orders per working day.",
    "Days missed = working days under 4. WOs missed = how many short of 4, added up."
)

$payload = @{
    title          = "PM from Nov 2027"
    periodStart    = $period
    dailyTarget    = $target
    generatedAt    = $kuwait.ToString("yyyy-MM-dd HH:mm")
    timezone       = "Kuwait"
    updateTime     = "20:00"
    nextUpdate     = $nextUpdate
    phase          = $phase
    today          = $today
    todayLabel     = (Format-DayLabel $today)
    todayFriday    = [bool]$todayFriday
    todayClosed    = [bool]$todayClosed
    sourceFile     = $sourceName
    rules          = $rules
    kpis           = @{
        fmes             = $fmeCount
        closedDays       = $closedN
        workOrders       = $teamActual
        expected         = $teamExpected
        efficiency       = $teamEfficiency
        average          = $teamAverage
        daysMissed       = $teamDaysMissed
        daysOnTarget     = $teamDaysHit
        totalMissed      = $teamMissed
        todayWorkOrders  = $todayActual
        todayExpected    = $(if ($todayInPeriod -and -not $todayFriday) { $fmeCount * $target } else { 0 })
    }
    fme            = $fmeSorted
    todayRows      = $todaySorted
}

$json = $payload | ConvertTo-Json -Depth 8 -Compress

$effText = if ($null -eq $teamEfficiency) { "n/a" } else { "$teamEfficiency%" }
$avgText = if ($null -eq $teamAverage) { "n/a" } else { "$teamAverage" }
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("PM from 1 Nov  $($payload.todayLabel)")
if ($phase -eq "waiting") {
    $lines.Add("Tracking starts 1 Nov 2026.")
    $lines.Add("Each FME: minimum $target work orders a day.")
} elseif ($todayFriday) {
    $lines.Add("Friday — rest day, no target.")
} else {
    $lines.Add("Today: $todayActual / $($fmeCount * $target) work orders")
    if (-not $todayClosed) { $lines.Add("Today is still open until 8:00 PM.") }
}
$lines.Add("Efficiency $effText")
$lines.Add("Average $avgText per FME per day")
$lines.Add("Days missed $teamDaysMissed")
$lines.Add("WOs missed $teamMissed")
$behind = @($todaySorted | Where-Object { $_.target -and -not $_.hit })
if ($behind.Count -gt 0 -and -not $todayFriday -and $phase -eq "live") {
    $lines.Add("Behind today:")
    foreach ($row in $behind) {
        $lines.Add("$($row.short) $($row.actual)/$($row.target)")
    }
}
$pages = $env:PAGES_URL
if ($pages) { $lines.Add($pages.TrimEnd("/")) }
$ntfy = ($lines -join "`n") + "`n"

foreach ($person in $fmeSorted) {
    Write-Host ("{0}: {1} done, avg {2}, efficiency {3}, days missed {4}, WOs missed {5}" -f $person.short, $person.actual, $person.average, $person.efficiency, $person.daysMissed, $person.totalMissed)
}
Write-Host $ntfy

if ($NoWrite) { return }

New-Item -ItemType Directory -Force -Path $Dash | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding $false
[IO.File]::WriteAllText((Join-Path $Dash "data.json"), $json, $utf8)
[IO.File]::WriteAllText((Join-Path $Dash "dashboard-data.js"), "window.DASHBOARD_DATA = $json;`n", $utf8)
[IO.File]::WriteAllText((Join-Path $Dash "ntfy.txt"), $ntfy, $utf8)

$pin = $env:DASHBOARD_PIN
if (-not $pin) {
    $pinFile = Join-Path $Dash ".pin-local.txt"
    if (Test-Path -LiteralPath $pinFile) { $pin = ([IO.File]::ReadAllText($pinFile)).Trim() }
}
if ($pin) {
    $pinJs = "window.DASHBOARD_PIN = " + (($pin | ConvertTo-Json -Compress)) + ";`n"
    [IO.File]::WriteAllText((Join-Path $Dash "pin.js"), $pinJs, $utf8)
}

Write-Host "Wrote dashboard\dashboard-data.js"
