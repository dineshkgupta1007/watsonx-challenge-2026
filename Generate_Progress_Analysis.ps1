# =============================================================================
# IBM Industrial Sector – Compliance Progress Analysis Generator
# =============================================================================
# Called by WFM_Launcher.hta "Open Progress Analysis" button.
# Steps:
#   1. Read per-rule KPI counts from the current IND_July2026_Compliance_by_FS.html
#   2. Read the same KPI counts for every previous daily snapshot from git history
#   3. Rebuild IND_July2026_Compliance_Progress_Analysis.html with all versions
#   4. git add + commit + push both HTML files
#   5. Exit 0 — launcher then opens the Progress Analysis in the browser
# =============================================================================

$root        = $PSScriptRoot
$byFsPath    = Join-Path $root "IND_July2026_Compliance_by_FS.html"
$progPath    = Join-Path $root "IND_July2026_Compliance_Progress_Analysis.html"

# --------------------------------------------------------------------------
# STEP 1 – Extract KPI numbers from an HTML string
# Returns ordered array: R1,R2,R3,R4,R5a,R5b,R5c,R6,R7  (9 values)
# The KPI row has exactly 9 .num divs in that order.
# --------------------------------------------------------------------------
function Get-KpiCounts([string]$html) {
    $nums = [regex]::Matches($html, "class='kpi [^']*'><div class='num'>(\d+)</div>") |
            ForEach-Object { [int]$_.Groups[1].Value }
    if ($nums.Count -lt 9) { return $null }
    return $nums[0..8]   # R1..R7 (indices 0-8)
}

function Get-TotalFromHtml([string]$html) {
    $m = [regex]::Match($html, "(\d+) flagged records")
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return 0
}

function Get-DateFromHtml([string]$html) {
    $m = [regex]::Match($html, "Generated:\s*([\d]+ \w+ \d{4})")
    if ($m.Success) { return $m.Groups[1].Value }
    return ""
}

# --------------------------------------------------------------------------
# STEP 2 – Read today's report
# --------------------------------------------------------------------------
if (-not (Test-Path $byFsPath)) {
    Write-Error "Compliance report not found: $byFsPath  — run 'Generate Report' first."
    exit 1
}
$todayHtml   = [System.IO.File]::ReadAllText($byFsPath, [System.Text.Encoding]::UTF8)
$todayCounts = Get-KpiCounts $todayHtml
$todayTotal  = Get-TotalFromHtml $todayHtml
$todayDateRaw= Get-DateFromHtml $todayHtml
$todayDate   = if ($todayDateRaw) { $todayDateRaw } else { (Get-Date).ToString("dd MMM yyyy") }
# Short label e.g. "31 Jul"
$todayLabel  = (Get-Date).ToString("d MMM")

Write-Host "Today's report: $todayTotal flags, date=$todayDate"

# --------------------------------------------------------------------------
# STEP 3 – Build history from git log
# We read every distinct commit that touched IND_July2026_Compliance_by_FS.html
# --------------------------------------------------------------------------
Push-Location $root

$gitLog = git log --pretty=format:"%H %ai %s" -- "IND_July2026_Compliance_by_FS.html" 2>&1
$allCommits = $gitLog | Where-Object { $_ -match '^[0-9a-f]{40}' }

# Build a list: @{ Hash; DateLabel; Total; Counts[9] }
$snapshots = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenDates = @{}

foreach ($line in $allCommits) {
    $parts = $line -split ' ', 3
    $hash  = $parts[0]
    # $parts[1] is ISO datetime e.g. 2026-07-31T14:21:00+05:30
    $isoDate = ($parts[1] -split 'T')[0]
    if ($seenDates.ContainsKey($isoDate)) { continue }   # take latest commit per day only

    $rawHtml = git show "$hash`:IND_July2026_Compliance_by_FS.html" 2>$null
    if (-not $rawHtml) { continue }
    $htmlStr = $rawHtml -join ""

    $counts = Get-KpiCounts $htmlStr
    $total  = Get-TotalFromHtml $htmlStr
    if ($total -eq 0 -or $null -eq $counts) { continue }

    # Short date label: "13 Jul"
    $dt     = [datetime]::ParseExact($isoDate, "yyyy-MM-dd", $null)
    $label  = $dt.ToString("d MMM")

    $seenDates[$isoDate] = $true
    $snapshots.Add([PSCustomObject]@{
        Hash      = $hash
        IsoDate   = $isoDate
        DateLabel = $label
        Total     = $total
        Counts    = $counts
    })
}

# Sort ascending by date
$snapshots = @($snapshots | Sort-Object IsoDate)

# Replace or add today's entry (the working-tree version is the most current)
$todayIso = (Get-Date).ToString("yyyy-MM-dd")
$snapshots = @($snapshots | Where-Object { $_.IsoDate -ne $todayIso })
if ($null -ne $todayCounts) {
    $snapshots += [PSCustomObject]@{
        Hash      = "HEAD"
        IsoDate   = $todayIso
        DateLabel = $todayLabel
        Total     = $todayTotal
        Counts    = $todayCounts
    }
    $snapshots = @($snapshots | Sort-Object IsoDate)
}

$vCount = $snapshots.Count
Write-Host "Total snapshots found: $vCount"

Pop-Location

# --------------------------------------------------------------------------
# STEP 4 – Compute per-rule v1→vN delta
# --------------------------------------------------------------------------
function Delta([int]$from, [int]$to) {
    $d = $to - $from
    if ($d -lt 0) { return "<span class='delta-good'>&#9660; $([Math]::Abs($d))</span>" }
    if ($d -gt 0) { return "<span class='delta-bad'>&#9650; $d</span>" }
    return "<span class='delta-neu'>&#8212;</span>"
}
function DeltaTotal([int]$from, [int]$to) {
    $d = $to - $from
    if ($d -lt 0) { return "<span class='delta-good' style='color:#6ee7b7'>&#9660; $([Math]::Abs($d))</span>" }
    if ($d -gt 0) { return "<span class='delta-bad' style='color:#f87171'>&#9650; $d</span>" }
    return "<span class='delta-neu'>&#8212;</span>"
}

$peakTotal = ($snapshots | Measure-Object Total -Maximum).Maximum
$firstTotal = $snapshots[0].Total
$lastTotal  = $snapshots[-1].Total
$bestTotal  = ($snapshots | Measure-Object Total -Minimum).Minimum
$bestSnap   = $snapshots | Where-Object { $_.Total -eq $bestTotal } | Select-Object -First 1
$pctDrop    = [Math]::Round((($peakTotal - $lastTotal) / $peakTotal) * 100)

$ruleNames = @(
    "R1 — EST Non Compliant",
    "R2 — Comment Non Compliant",
    "R3 — Track / Fieldglass Mismatch",
    "R4 — Track Type / FA Mismatch",
    "R5a — Recruiting w/o Ticket",
    "R5b — FA Bench/Rolloff + Ticket",
    "R5c — SubK / FG Mismatch",
    "R6 — Backfill N + Ranking 999",
    "R7 — Confirmed &gt; 2 Working Days"
)

# --------------------------------------------------------------------------
# STEP 5 – Build HTML
# --------------------------------------------------------------------------
$sb = [System.Text.StringBuilder]::new()

# ---- CSS ----
$css = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"Segoe UI",system-ui,sans-serif;font-size:14px;line-height:1.6;background:#fff;color:#1f2328}
.wrap{max-width:820px;margin:0 auto;padding:28px 16px}
h1{font-size:18px;font-weight:700;margin-bottom:4px}
.sub{color:#57606a;font-size:12px;margin-bottom:24px}
h2{font-size:14px;font-weight:700;margin:28px 0 10px;border-bottom:1px solid #e5e7eb;padding-bottom:5px}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:24px}
.kpi{background:#f7f8fa;border:1px solid #e5e7eb;border-radius:6px;padding:12px 14px}
.kpi .num{font-size:26px;font-weight:700}.kpi .lbl{font-size:11px;color:#57606a;margin-top:2px}
.green .num{color:#166534}.red .num{color:#b91c1c}.blue .num{color:#1d4ed8}.amber .num{color:#b45309}
table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px}
th{background:#f7f8fa;border:1px solid #e5e7eb;padding:7px 8px;text-align:center;font-weight:600;white-space:nowrap}
th.left{text-align:left}
td{border:1px solid #e5e7eb;padding:5px 8px;text-align:center;vertical-align:middle}
td.label{text-align:left;font-weight:600;font-size:11px;background:#f7f8fa;white-space:nowrap}
tr.total-row td{font-weight:700;background:#1f2328;color:#fff}
tr.total-row td.label{background:#1f2328;color:#fff}
.bar-section{margin-bottom:24px}
.bar-row{display:flex;align-items:center;margin-bottom:6px;gap:8px}
.bar-label{width:70px;font-size:11px;font-weight:600;color:#57606a;flex-shrink:0;text-align:right}
.bar-wrap{flex:1;background:#f0f0f0;border-radius:3px;height:18px;position:relative}
.bar-fill{height:100%;border-radius:3px;display:flex;align-items:center;padding-left:6px;font-size:11px;font-weight:700;color:#fff;min-width:24px}
.bar-num{font-size:11px;color:#57606a;flex-shrink:0;width:36px;text-align:right}
.delta-good{color:#166534;font-weight:700}.delta-bad{color:#b91c1c;font-weight:700}.delta-neu{color:#57606a}
.ver-links{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:22px}
.ver-chip{border:1px solid #e5e7eb;border-radius:20px;padding:3px 12px;font-size:11px;color:#3b82d4;text-decoration:none;background:#f7f8fa}
.ver-chip:hover{background:#e5e7eb}
.latest{border-color:#166534;color:#166534;font-weight:700}
.insights{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px}
.card{background:#f7f8fa;border:1px solid #e5e7eb;border-radius:6px;padding:12px 14px;font-size:12px}
.card strong{display:block;font-size:12px;margin-bottom:4px}
footer{margin-top:32px;padding-top:10px;border-top:1px solid #e5e7eb;text-align:center;font-size:11px;color:#57606a}
'@

$null = $sb.AppendLine('<!DOCTYPE html>')
$null = $sb.AppendLine('<html lang="en"><head><meta charset="UTF-8"/>')
$null = $sb.AppendLine('<title>IBM Industrial July 2026 – Compliance Progress Analysis</title>')
$null = $sb.AppendLine("<style>$css</style></head><body><div class='wrap'>")

$genDate = (Get-Date).ToString("dd MMM yyyy")
$null = $sb.AppendLine('<h1>IBM Industrial Sector — July 2026 Demand Compliance</h1>')
$null = $sb.AppendLine("<p class='sub'>Comprehensive Progress Analysis &nbsp;|&nbsp; $vCount report snapshots &nbsp;|&nbsp; $($snapshots[0].DateLabel) – $($snapshots[-1].DateLabel) 2026 &nbsp;|&nbsp; 9 compliance rules tracked &nbsp;|&nbsp; Generated: $genDate</p>")

# ---- Version chips ----
$null = $sb.AppendLine("<h2>Report Versions — Click to Open</h2>")
$null = $sb.AppendLine("<div class='ver-links'>")
$vn = 1
foreach ($snap in $snapshots) {
    $isLatest = ($snap -eq $snapshots[-1])
    $cls   = if ($isLatest) { "ver-chip latest" } else { "ver-chip" }
    $label = "v$vn &middot; $($snap.DateLabel) &middot; $($snap.Total) flags"
    if ($isLatest) { $label += " &#10003; Latest" }

    if ($snap.Hash -eq "HEAD") {
        $href = "https://htmlpreview.github.io/?https://raw.githubusercontent.com/dineshkgupta1007/watsonx-challenge-2026/main/IND_July2026_Compliance_by_FS.html"
    } else {
        $shortHash = $snap.Hash.Substring(0,7)
        $href = "https://htmlpreview.github.io/?https://raw.githubusercontent.com/dineshkgupta1007/watsonx-challenge-2026/$($shortHash)/IND_July2026_Compliance_by_FS.html"
    }
    $null = $sb.AppendLine("  <a class='$cls' href='$href' target='_blank'>$label</a>")
    $vn++
}
$null = $sb.AppendLine("</div>")

# ---- KPI summary ----
$null = $sb.AppendLine("<h2>Overall Progress Summary</h2>")
$null = $sb.AppendLine("<div class='kpi-row'>")
$null = $sb.AppendLine("  <div class='kpi green'><div class='num'>$pctDrop%</div><div class='lbl'>Net reduction from peak ($peakTotal &rarr; $lastTotal)</div></div>")
$null = $sb.AppendLine("  <div class='kpi blue'><div class='num'>$($peakTotal - $lastTotal)</div><div class='lbl'>Issues resolved peak-to-latest</div></div>")
$null = $sb.AppendLine("  <div class='kpi amber'><div class='num'>$bestTotal</div><div class='lbl'>Best day: $($bestSnap.DateLabel) (lowest flags)</div></div>")
$null = $sb.AppendLine("  <div class='kpi red'><div class='num'>$lastTotal</div><div class='lbl'>Remaining flags as of $($snapshots[-1].DateLabel) 2026</div></div>")
$null = $sb.AppendLine("</div>")

# ---- Bar chart ----
$null = $sb.AppendLine("<h2>Total Flagged Records — All $vCount Snapshots</h2>")
$null = $sb.AppendLine("<div class='bar-section'>")

# Bar colours cycle: red-shades for high, blue/green for low
$barColors = @("#e57373","#b91c1c","#f59e0b","#f59e0b","#6366f1","#22c55e","#f59e0b","#3b82d4","#0ea5e9","#f59e0b","#f59e0b","#3b82d4","#f59e0b","#10b981","#059669","#166534")

$vn = 1
foreach ($snap in $snapshots) {
    $pct   = [Math]::Round(($snap.Total / $peakTotal) * 100, 1)
    $color = $barColors[($vn - 1) % $barColors.Count]
    # Override: best day = blue, latest = green
    if ($snap.Total -eq $bestTotal)   { $color = "#3b82d4" }
    if ($snap -eq $snapshots[-1])     { $color = "#10b981" }
    $null = $sb.AppendLine("  <div class='bar-row'>")
    $null = $sb.AppendLine("    <span class='bar-label'>v$vn $($snap.DateLabel)</span>")
    $null = $sb.AppendLine("    <div class='bar-wrap'><div class='bar-fill' style='width:$pct%;background:$color'>$($snap.Total)</div></div>")
    $null = $sb.AppendLine("    <span class='bar-num'>$($snap.Total)</span>")
    $null = $sb.AppendLine("  </div>")
    $vn++
}
$null = $sb.AppendLine("</div>")

# ---- Rule-by-rule table ----
$null = $sb.AppendLine("<h2>Rule-by-Rule Comparison — Full Month</h2>")
$null = $sb.AppendLine("<div style='overflow-x:auto'>")
$null = $sb.Append("<table class='rule-tbl'><thead><tr><th class='left' style='width:175px'>Rule</th>")
$vn = 1
foreach ($snap in $snapshots) {
    $null = $sb.Append("<th>v$vn<br><span style='font-weight:400;font-size:10px'>$($snap.DateLabel)</span></th>")
    $vn++
}
$null = $sb.AppendLine("<th>v1&rarr;v$($vCount)<br><span style='font-weight:400;font-size:10px'>Change</span></th></tr></thead><tbody>")

for ($ri = 0; $ri -lt 9; $ri++) {
    $null = $sb.Append("<tr><td class='label'>$($ruleNames[$ri])</td>")
    foreach ($snap in $snapshots) {
        if ($null -ne $snap.Counts) {
            $null = $sb.Append("<td>$($snap.Counts[$ri])</td>")
        } else {
            $null = $sb.Append("<td>&#8212;</td>")
        }
    }
    $from = if ($null -ne $snapshots[0].Counts) { $snapshots[0].Counts[$ri] } else { 0 }
    $to   = if ($null -ne $snapshots[-1].Counts) { $snapshots[-1].Counts[$ri] } else { 0 }
    $null = $sb.AppendLine("<td>$(Delta $from $to)</td></tr>")
}

# Total row
$null = $sb.Append("<tr class='total-row'><td class='label'>TOTAL FLAGGED</td>")
foreach ($snap in $snapshots) { $null = $sb.Append("<td>$($snap.Total)</td>") }
$null = $sb.AppendLine("<td>$(DeltaTotal $firstTotal $lastTotal)</td></tr>")
$null = $sb.AppendLine("</tbody></table></div>")

# ---- Day-over-day change table ----
$null = $sb.AppendLine("<h2>Day-over-Day Change</h2>")
$null = $sb.AppendLine("<table><thead><tr><th class='left'>Version Transition</th><th>Before</th><th>After</th><th>Change</th></tr></thead><tbody>")
for ($i = 1; $i -lt $snapshots.Count; $i++) {
    $prev = $snapshots[$i-1]; $curr = $snapshots[$i]
    $chg  = $curr.Total - $prev.Total
    if ($chg -lt 0) {
        $chgHtml = "<span class='delta-good'>$chg</span>"
    } elseif ($chg -gt 0) {
        $chgHtml = "<span class='delta-bad'>+$chg</span>"
    } else {
        $chgHtml = "<span class='delta-neu'>0</span>"
    }
    $null = $sb.AppendLine("<tr><td class='label'>v$i &rarr; v$($i+1) ($($prev.DateLabel) &rarr; $($curr.DateLabel))</td><td>$($prev.Total)</td><td>$($curr.Total)</td><td>$chgHtml</td></tr>")
}
$null = $sb.AppendLine("</tbody></table>")

$null = $sb.AppendLine("<footer>IBM Industrial WFM &mdash; Sector Collaboration &nbsp;|&nbsp; Made with IBM Bob</footer>")
$null = $sb.AppendLine("</div></body></html>")

$htmlOut = $sb.ToString()
[System.IO.File]::WriteAllText($progPath, $htmlOut, [System.Text.Encoding]::UTF8)
Write-Host "Progress Analysis written: $($htmlOut.Length / 1KB) KB"

# --------------------------------------------------------------------------
# STEP 6 – Git commit + push both files
# --------------------------------------------------------------------------
Push-Location $root
git add "IND_July2026_Compliance_Progress_Analysis.html" "IND_July2026_Compliance_by_FS.html" 2>&1 | Write-Host
$msg = "Progress Analysis refresh - $(Get-Date -Format 'dd MMM yyyy HH:mm') | $vCount snapshots | $lastTotal flags"
$gitOut = git commit -m $msg 2>&1
Write-Host $gitOut
if ($LASTEXITCODE -eq 0 -or ($gitOut -match "nothing to commit")) {
    git push origin main 2>&1 | Write-Host
    Write-Host "GitHub push complete."
} else {
    git push origin main 2>&1 | Write-Host
}
Pop-Location

Write-Host "Done."
exit 0
