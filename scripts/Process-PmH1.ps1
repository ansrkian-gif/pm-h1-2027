# PM H1 2027 - process OWS extract, apply pairing/NTE/FME rules, write mobile dashboard data.
# ASCII only. Friday is not a working day.
param(
    [string]$ExtractPath = "",
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
if (-not $ExtractPath) { $ExtractPath = Join-Path $Root "PM H1 2027.xlsx" }
if (-not $OutDir) { $OutDir = Join-Path $Root "dashboard" }

$ACTIVE = "OGK Active General"
$PASSIVE = "OGK Passive General"
$SMALL = "OGK Active Small Cell /Book RRU/Easy Macro"
$CATEGORY = "OGK Active and Passive Routine Maintenance"

# Nabi replaced Mohd: score both as one FME.
$FmeAlias = @{
    "Mohd Yasin Mohd Matin Ansari" = "Nabijohn Piyarjan Piyarjan"
}

function Resolve-FmeName([string]$name) {
    $t = ([string]$name).Trim()
    if ($FmeAlias.ContainsKey($t)) { return $FmeAlias[$t] }
    return $t
}

function Get-WeekStart([datetime]$d) {
    return $d.Date.AddDays(-([int]$d.DayOfWeek))
}

function Convert-ColToIndex([string]$col) {
    $n = 0
    foreach ($ch in $col.ToUpper().ToCharArray()) {
        $n = (26 * $n) + ([int]$ch - 64)
    }
    return $n
}

function Import-OwsXlsx([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $tmp = Join-Path $env:TEMP ("pmh1_" + [guid]::NewGuid().ToString("N") + ".xlsx")
    Copy-Item -LiteralPath $path -Destination $tmp -Force
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
    try {
        function Read-ZipEntry([string]$name) {
            $e = $zip.GetEntry($name)
            if (-not $e) { return "" }
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { return $sr.ReadToEnd() } finally { $sr.Close() }
        }

        $ssXml = Read-ZipEntry "xl/sharedStrings.xml"
        $strings = New-Object System.Collections.Generic.List[string]
        if ($ssXml) {
            $matches = [regex]::Matches($ssXml, "(?s)<si>(.*?)</si>")
            foreach ($m in $matches) {
                $tMatches = [regex]::Matches($m.Groups[1].Value, "(?s)<t[^>]*>(.*?)</t>")
                $sb = New-Object System.Text.StringBuilder
                foreach ($t in $tMatches) {
                    $val = [System.Net.WebUtility]::HtmlDecode($t.Groups[1].Value)
                    [void]$sb.Append($val)
                }
                [void]$strings.Add($sb.ToString())
            }
        }

        $sheet = Read-ZipEntry "xl/worksheets/sheet1.xml"
        $cellMatches = [regex]::Matches($sheet, '<c r="([A-Z]+)(\d+)"([^>]*)>(?:<v>([^<]*)</v>)?')
        $grid = @{}
        foreach ($c in $cellMatches) {
            $col = $c.Groups[1].Value
            $row = [int]$c.Groups[2].Value
            $attrs = $c.Groups[3].Value
            $v = $c.Groups[4].Value
            $isStr = $attrs -match 't="s"'
            $text = ""
            if ($isStr -and $v -ne "") {
                $idx = [int]$v
                if ($idx -ge 0 -and $idx -lt $strings.Count) { $text = $strings[$idx] }
            } else {
                $text = $v
            }
            $key = "$row|$col"
            $grid[$key] = $text
        }

        $headerRow = 1
        $headers = @{}
        foreach ($key in @($grid.Keys)) {
            $parts = $key.Split("|")
            if ([int]$parts[0] -ne 1) { continue }
            $headers[(Convert-ColToIndex $parts[1])] = $grid[$key]
        }
        $nameToCol = @{}
        foreach ($k in $headers.Keys) { $nameToCol[$headers[$k]] = $k }

        function ColLetter([int]$n) {
            $s = ""
            while ($n -gt 0) {
                $n--
                $s = [char](65 + ($n % 26)) + $s
                $n = [math]::Floor($n / 26)
            }
            return $s
        }

        $maxRow = 1
        foreach ($key in $grid.Keys) {
            $r = [int]($key.Split("|")[0])
            if ($r -gt $maxRow) { $maxRow = $r }
        }

        $rows = New-Object System.Collections.ArrayList
        for ($r = 2; $r -le $maxRow; $r++) {
            $obj = [ordered]@{}
            foreach ($name in $nameToCol.Keys) {
                $letter = ColLetter ([int]$nameToCol[$name])
                $val = $grid["$r|$letter"]
                if ($null -eq $val) { $val = "" }
                $obj[$name] = $val
            }
            [void]$rows.Add([pscustomobject]$obj)
        }
        return @($rows)
    }
    finally {
        $zip.Dispose()
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Get-Day([string]$ct) {
    if ([string]::IsNullOrWhiteSpace($ct) -or $ct.Length -lt 10) { return $null }
    return [datetime]::ParseExact($ct.Substring(0, 10), "yyyy-MM-dd", $null)
}

Write-Host "Reading $ExtractPath"
$raw = Import-OwsXlsx $ExtractPath
$pm = @($raw | Where-Object {
    $_.("Task Type") -eq "PM" -and $_.("Task Category") -eq $CATEGORY
})
Write-Host "PM OGK rows: $($pm.Count) / $($raw.Count)"

$rows = New-Object System.Collections.ArrayList
foreach ($r in $pm) {
    $op = "$($r.'Accept Operator')|$($r.'Arrive Operator')|$($r.'Complete Operator')"
    $day = Get-Day $r."Complete Time"
    if ($null -eq $day) { continue }
    [void]$rows.Add([pscustomobject]@{
        TaskId      = $r."Task Id"
        SiteId      = ([string]$r."Site ID").Trim()
        Subcat      = ([string]$r."Task Subcategory").Trim()
        Fme         = Resolve-FmeName $r."Assign To FME Full Name"
        Day         = $day
        WeekStart   = Get-WeekStart $day
        Complete    = $r."Complete Time"
        IsNte       = [bool]($op -match "NTE")
        CompleteOp  = $r."Complete Operator"
    })
}

$exceptions = New-Object System.Collections.ArrayList
$counted = New-Object System.Collections.ArrayList
$pairUnits = New-Object System.Collections.ArrayList

$ap = @($rows | Where-Object { $_.Subcat -eq $ACTIVE -or $_.Subcat -eq $PASSIVE })
$pairGroups = $ap | Group-Object SiteId, { $_.WeekStart.ToString("yyyy-MM-dd") }
foreach ($g in $pairGroups) {
    $items = @($g.Group)
    $act = @($items | Where-Object { $_.Subcat -eq $ACTIVE } | Sort-Object Complete, TaskId)
    $pas = @($items | Where-Object { $_.Subcat -eq $PASSIVE } | Sort-Object Complete, TaskId)
    $site = $items[0].SiteId
    $week = $items[0].WeekStart
    if ($act.Count -ge 1 -and $pas.Count -ge 1) {
        $keepA = $act[0]; $keepP = $pas[0]
        [void]$pairUnits.Add([pscustomobject]@{
            Kind = "Active+Passive"; SiteId = $site; WeekStart = $week
            CompleteDay = $keepP.Day; Complete = $keepP.Complete
            NteFme = $keepP.Fme; ActiveTask = $keepA.TaskId; PassiveTask = $keepP.TaskId
            ActiveDay = $keepA.Day; PassiveDay = $keepP.Day
        })
        if ($act.Count -gt 1) {
            for ($i = 1; $i -lt $act.Count; $i++) {
                $d = $act[$i]
                [void]$exceptions.Add([pscustomobject]@{ Reason = "Duplicate Active row same site + same week"; SiteId = $site; Day = $d.Day.ToString("yyyy-MM-dd"); Week = $week.ToString("yyyy-MM-dd"); Subcat = $d.Subcat; TaskId = $d.TaskId; Fme = $d.Fme })
            }
        }
        if ($pas.Count -gt 1) {
            for ($i = 1; $i -lt $pas.Count; $i++) {
                $d = $pas[$i]
                [void]$exceptions.Add([pscustomobject]@{ Reason = "Duplicate Passive row same site + same week"; SiteId = $site; Day = $d.Day.ToString("yyyy-MM-dd"); Week = $week.ToString("yyyy-MM-dd"); Subcat = $d.Subcat; TaskId = $d.TaskId; Fme = $d.Fme })
            }
        }
    } else {
        foreach ($x in $items) {
            [void]$exceptions.Add([pscustomobject]@{ Reason = "Unpaired Active/Passive in the same week (ignored)"; SiteId = $site; Day = $x.Day.ToString("yyyy-MM-dd"); Week = $week.ToString("yyyy-MM-dd"); Subcat = $x.Subcat; TaskId = $x.TaskId; Fme = $x.Fme })
        }
    }
}

$pairBySite = $pairUnits | Sort-Object Complete, SiteId | Group-Object SiteId
foreach ($g in $pairBySite) {
    $units = @($g.Group)
    [void]$counted.Add($units[0])
    if ($units.Count -gt 1) {
        for ($i = 1; $i -lt $units.Count; $i++) {
            $d = $units[$i]
            [void]$exceptions.Add([pscustomobject]@{ Reason = "Duplicate Active+Passive visit later week/site (ignored)"; SiteId = $d.SiteId; Day = $d.CompleteDay.ToString("yyyy-MM-dd"); Week = $d.WeekStart.ToString("yyyy-MM-dd"); Subcat = "Active+Passive pair"; TaskId = "$($d.ActiveTask) / $($d.PassiveTask)"; Fme = $d.NteFme })
        }
    }
}

$sc = @($rows | Where-Object { $_.Subcat -eq $SMALL } | Sort-Object Complete, TaskId)
foreach ($g in ($sc | Group-Object SiteId)) {
    $items = @($g.Group)
    $keep = $items[0]
    [void]$counted.Add([pscustomobject]@{
        Kind = "Small Cell"; SiteId = $keep.SiteId; WeekStart = $keep.WeekStart
        CompleteDay = $keep.Day; Complete = $keep.Complete
        NteFme = $keep.Fme; ActiveTask = ""; PassiveTask = $keep.TaskId
        ActiveDay = $null; PassiveDay = $keep.Day
    })
    if ($items.Count -gt 1) {
        for ($i = 1; $i -lt $items.Count; $i++) {
            $d = $items[$i]
            [void]$exceptions.Add([pscustomobject]@{ Reason = "Duplicate Small Cell for same site (ignored)"; SiteId = $d.SiteId; Day = $d.Day.ToString("yyyy-MM-dd"); Week = $d.WeekStart.ToString("yyyy-MM-dd"); Subcat = $d.Subcat; TaskId = $d.TaskId; Fme = $d.Fme })
        }
    }
}

$perfRows = New-Object System.Collections.ArrayList
$byFmeDay = $counted | Group-Object { "$($_.NteFme.Trim())|$($_.CompleteDay.ToString('yyyy-MM-dd'))" }
foreach ($g in $byFmeDay) {
    $items = @($g.Group)
    $fme = $items[0].NteFme.Trim()
    $day = $items[0].CompleteDay
    if ($day.DayOfWeek -eq "Friday") { continue }
    $apN = @($items | Where-Object { $_.Kind -eq "Active+Passive" }).Count
    $scN = @($items | Where-Object { $_.Kind -eq "Small Cell" }).Count
    $actual = $apN + $scN
    if ($apN -gt 0) { $mode = "Active+Passive"; $target = 3 } else { $mode = "Small Cell only"; $target = 5 }
    $hit = $actual -ge $target
    [void]$perfRows.Add([pscustomobject]@{
        Fme = $fme; Day = $day.ToString("yyyy-MM-dd"); Dow = $day.DayOfWeek.ToString()
        AP = $apN; SC = $scN; Actual = $actual; Mode = $mode; Target = $target
        Gap = $actual - $target; Hit = $hit
    })
}

$perf = @($perfRows)
$fmeNames = @($perf | Select-Object -ExpandProperty Fme -Unique | Sort-Object)
$minDay = ($counted | ForEach-Object { $_.CompleteDay } | Measure-Object -Minimum).Minimum
$maxDay = ($counted | ForEach-Object { $_.CompleteDay } | Measure-Object -Maximum).Maximum
$days = @()
for ($d = $minDay.Date; $d -le $maxDay.Date; $d = $d.AddDays(1)) {
    if ($d.DayOfWeek -ne "Friday") { $days += $d.ToString("yyyy-MM-dd") }
}

$fmeSum = foreach ($f in $fmeNames) {
    $daysW = @($perf | Where-Object { $_.Fme -eq $f })
    $actual = [int]($daysW | Measure-Object Actual -Sum).Sum
    $expected = [int]($daysW | Measure-Object Target -Sum).Sum
    $apN = [int]($daysW | Measure-Object AP -Sum).Sum
    $scN = [int]($daysW | Measure-Object SC -Sum).Sum
    $on = @($daysW | Where-Object { $_.Hit }).Count
    $below = @($daysW | Where-Object { -not $_.Hit }).Count
    $hitPct = if ($daysW.Count -gt 0) { [math]::Round(100.0 * $on / $daysW.Count, 0) } else { 0 }
    $shape = if ($hitPct -ge 80) { "GOOD" } elseif ($hitPct -ge 50) { "WATCH" } else { "BAD" }
    $daily = foreach ($day in $days) {
        $row = @($perf | Where-Object { $_.Fme -eq $f -and $_.Day -eq $day }) | Select-Object -First 1
        if ($null -eq $row) {
            [pscustomobject]@{ day = $day; actual = $null; target = $null; hit = $null }
        } else {
            [pscustomobject]@{ day = $day; actual = [int]$row.Actual; target = [int]$row.Target; hit = [bool]$row.Hit }
        }
    }
    [pscustomobject]@{
        name = $f; daysWorked = $daysW.Count; ap = $apN; sc = $scN
        actual = $actual; expected = $expected
        avg = [math]::Round($actual / [math]::Max($daysW.Count, 1), 1)
        onTarget = $on; below = $below; hitPct = $hitPct; shape = $shape
        daily = @($daily)
    }
}
$fmeSum = @($fmeSum | Sort-Object @{ Expression = "hitPct"; Descending = $true }, name)

$teamAct = [int]($perf | Measure-Object Actual -Sum).Sum
$teamExp = [int]($perf | Measure-Object Target -Sum).Sum
$teamOn = @($perf | Where-Object { $_.Hit }).Count
$teamBelow = @($perf | Where-Object { -not $_.Hit }).Count
$teamHit = if ($perf.Count -gt 0) { [math]::Round(100.0 * $teamOn / $perf.Count, 0) } else { 0 }
$teamShape = if ($teamHit -ge 80) { "GOOD" } elseif ($teamHit -ge 50) { "WATCH" } else { "BAD" }
$pairs = @($counted | Where-Object { $_.Kind -eq "Active+Passive" }).Count
$scs = @($counted | Where-Object { $_.Kind -eq "Small Cell" }).Count
$ws = @($counted | Select-Object -ExpandProperty SiteId -Unique).Count

$payload = [pscustomobject]@{
    generatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    sourceFile  = [IO.Path]::GetFileName($ExtractPath)
    owsUrl      = "https://106d-sg.teleows.com/"
    kpis = [pscustomobject]@{
        tasks = $counted.Count; sites = $ws; fmes = $fmeNames.Count
        onTarget = $teamOn; fmeDays = $perf.Count; hitPct = $teamHit
        shape = $teamShape; below = $teamBelow; expected = $teamExp; actual = $teamAct
        pairs = $pairs; smallCell = $scs
    }
    days = $days
    fme = $fmeSum
    below = @($perf | Where-Object { -not $_.Hit } | Sort-Object Fme, Day | ForEach-Object {
        [pscustomobject]@{ fme = $_.Fme; day = $_.Day; mode = $_.Mode; actual = $_.Actual; target = $_.Target; ap = $_.AP; sc = $_.SC }
    })
    exceptions = @($exceptions | ForEach-Object {
        [pscustomobject]@{ reason = $_.Reason; siteId = $_.SiteId; day = $_.Day; week = $_.Week; subcat = $_.Subcat; taskId = $_.TaskId; fme = $_.Fme }
    })
    rules = @(
        "PM only - OGK Active and Passive Routine Maintenance",
        "Active + Passive, same site, same week (Sun-Sat) = 1 task",
        "Small Cell = 1 task by itself",
        "NTE FME accounts only",
        "Mohd Yasin counted under Nabijohn (replacement)",
        "Daily target: 3 paired tasks, or 5 if Small Cell only",
        "Friday is not a working day"
    )
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$jsonPath = Join-Path $OutDir "data.json"
$json = $payload | ConvertTo-Json -Depth 8 -Compress
[IO.File]::WriteAllText($jsonPath, $json, [Text.UTF8Encoding]::new($false))

$templatePath = Join-Path $OutDir "template.html"
if (-not (Test-Path $templatePath)) { throw "Missing dashboard template: $templatePath" }
$html = [IO.File]::ReadAllText($templatePath, [Text.Encoding]::UTF8)
$html = $html.Replace("__DASHBOARD_DATA__", $json)
$html = $html.Replace("__DASHBOARD_PIN__", [string]$env:DASHBOARD_PIN)
$html = $html.Replace("__GENERATED_AT__", $payload.generatedAt)
$indexPath = Join-Path $OutDir "index.html"
[IO.File]::WriteAllText($indexPath, $html, [Text.UTF8Encoding]::new($false))

Write-Host "Wrote $jsonPath"
Write-Host "Wrote $indexPath"
Write-Host "Tasks=$($counted.Count) FME=$($fmeNames.Count) hit=$teamHit% shape=$teamShape"
