# =============================================================================
# IBM Industrial Sector – Demand Compliance Report Generator
# =============================================================================
# USAGE:  Double-click this file  OR  run in PowerShell:
#         powershell -ExecutionPolicy Bypass -File "Generate_Compliance_Report.ps1"
#
# INPUT:  Reads the pre-dumped JSON from the xlsx dump cache (auto-created by Bob)
#         JSON path: .bob\tmp\xlsx-dumps\Ind July month 30 days demands-...\30days.json
#
# OUTPUT: IND_July2026_Compliance_by_FS.html  (same folder as this script)
#
# STEPS TO RE-RUN TOMORROW:
#   1. Place the new/updated xlsx file in this folder (same name)
#   2. Open Bob chat and type: "regenerate compliance report"
#      Bob will re-dump the xlsx and run this script automatically.
#   OR manually:
#   3. Run this script directly (see USAGE above) if the JSON dump is still current.
# =============================================================================

$root = $PSScriptRoot
$raw = Get-Content "$root\.bob\tmp\xlsx-dumps\Ind July month 30 days demands-2c09300b65411b62\30days.json" -Raw | ConvertFrom-Json
$headers = $raw.headers
$rows    = $raw.rows

$idx = @{}
for ($i = 0; $i -lt $headers.Count; $i++) { $idx[$headers[$i]] = $i }

function cell($row, $col) {
    $i = $idx[$col]; if ($null -eq $i) { return "" }
    $v = $row[$i];   if ($null -eq $v)  { return "" }
    return "$v".Trim()
}

function parseDate($s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $d = New-Object datetime
    $fmts = "yyyy-MM-dd","M/d/yyyy","MM/dd/yyyy","d/M/yyyy","dd/MM/yyyy","yyyy-MM-ddTHH:mm:ss"
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    foreach ($f in $fmts) {
        if ([datetime]::TryParseExact($s.Trim(), $f, $ci, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    }
    if ([datetime]::TryParse($s.Trim(), $ci, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function latestCommentDate($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $pat = '\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{4}[-/]\d{1,2}[-/]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{1,2},?\s*\d{4}|\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4})\b'
    $ms = [regex]::Matches($text, $pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $best = $null
    foreach ($m in $ms) {
        $d = parseDate $m.Value
        if ($null -ne $d -and ($null -eq $best -or $d -gt $best)) { $best = $d }
    }
    return $best
}

$refDate    = [datetime]"2026-07-10"
$commentCut = [datetime]"2026-06-25"

# Returns the number of working days (Mon-Fri) between two dates (exclusive of $start, inclusive of $end)
function workingDaysBetween($start, $end) {
    if ($end -le $start) { return 0 }
    $days = 0; $cur = $start.Date.AddDays(1)
    while ($cur -le $end.Date) {
        if ($cur.DayOfWeek -ne [DayOfWeek]::Saturday -and $cur.DayOfWeek -ne [DayOfWeek]::Sunday) { $days++ }
        $cur = $cur.AddDays(1)
    }
    return $days
}

$contractorTracks = @(
    "contractor being pursued - core skill",
    "contractor being pursued - non-core skill",
    "contractor identified - awaiting start",
    "temporarily mitigated - contractor being pursued"
)
$r4bTracks = @(
    "candidate identified","project team reviewing candidates",
    "actively searching","gr being pursued",
    "not urgent - ongoing search","roll-off being pursued"
)

$allRecords = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $rows) {
    $openSeatId = cell $row "Open Seat ID"
    $client     = cell $row "Client Name"
    $industry   = cell $row "Industry"
    $title      = cell $row "Open Seat Title"
    $estDt      = cell $row "Est Strt Dt"
    $comments   = cell $row "Additional Comments"
    $track      = cell $row "Candidate Track Type"
    $fa         = cell $row "Fulfillment Action"
    $fg         = cell $row "Fieldglass Request Flag"
    $ticket     = cell $row "Hiring Ticket Number AskFile"
    $askDetail  = cell $row "Ask File Details"
    $fsId       = cell $row "FS Intranet ID"
    $backfilled  = cell $row "Backfilled"
    $priority    = cell $row "Priority Ranking"
    $candDisp    = cell $row "Candidate Disposition"
    $confirmedDt = cell $row "Confirmed Date"

    $trackL = $track.ToLower(); $faL = $fa.ToLower(); $fgL = $fg.ToLower()
    $flags = [System.Collections.Generic.List[string]]::new()

    $estParsed = parseDate $estDt
    if ([string]::IsNullOrWhiteSpace($estDt) -or ($null -ne $estParsed -and $estParsed -le $refDate)) {
        $flags.Add("be:EST Non Compliant")
    }

    $lcd = latestCommentDate $comments
    if ([string]::IsNullOrWhiteSpace($comments) -or ($null -eq $lcd) -or ($lcd -le $commentCut)) {
        $flags.Add("bc:Comment Non Compliant")
    }

    if ($trackL -like "*contractor*" -and ($fgL -eq "" -or $fgL -eq "n")) { $flags.Add("bf:Track Type/Fieldglass mismatch") }

    $priorityNum = 0
    $isPriority999 = ([int]::TryParse($priority.Trim().Split('.')[0], [ref]$priorityNum) -and $priorityNum -eq 999)
    $isBackfillN   = ($backfilled.ToUpper() -eq "N" -or $backfilled.ToLower() -eq "no")
    if ($isBackfillN -and $isPriority999) { $flags.Add("r6:Backfill N but Ranking 999") }

    $r4 = $false
    if (($trackL -eq "actively recruiting" -or $trackL -eq "new hire identified - awaiting start") -and $faL -ne "external hire") { $r4 = $true }
    if (!$r4 -and ($r4bTracks -contains $trackL) -and $faL -ne "bench/rolloff") { $r4 = $true }
    if (!$r4 -and $trackL -eq "rotation being pursued" -and $faL -ne "rotation") { $r4 = $true }
    if (!$r4 -and ($contractorTracks -contains $trackL) -and $faL -ne "contractor") { $r4 = $true }
    if ($r4) { $flags.Add("ba:Track Type/FA mismatch") }

    if ($trackL -eq "actively recruiting" -and [string]::IsNullOrWhiteSpace($ticket)) { $flags.Add("r5a:No Hiring Ticket") }
    if ($faL -eq "bench/rolloff" -and -not [string]::IsNullOrWhiteSpace($ticket))     { $flags.Add("r5b:FA Bench/Rolloff with Hiring Ticket") }
    if (($contractorTracks -contains $trackL) -and ($fgL -eq "" -or $fgL -eq "n"))    { $flags.Add("r5c:SubK/FG Mismatch") }

    if ($candDisp.ToLower() -eq "confirmed" -and -not [string]::IsNullOrWhiteSpace($confirmedDt)) {
        $confParsed = parseDate $confirmedDt
        if ($null -ne $confParsed -and (workingDaysBetween $confParsed ([datetime]::Today)) -gt 2) {
            $flags.Add("r7:Candidate confirmed more than 2 days")
        }
    }

    if ($flags.Count -gt 0) {
        $allRecords.Add([PSCustomObject]@{
            FS        = if ([string]::IsNullOrWhiteSpace($fsId)) { "(blank)" } else { $fsId }
            SeatId    = $openSeatId; Client = $client; Industry = $industry; Title = $title
            EstDt     = $estDt; EstParsed = $estParsed; Comments = $comments
            Track     = $track; FA = $fa; FG = if ([string]::IsNullOrWhiteSpace($fg)) { "" } else { $fg }
            Ticket    = $ticket; AskDetail = $askDetail; Flags = $flags
        })
    }
}

$cR1  = @($allRecords | Where-Object { $_.Flags -like "be:*" }).Count
$cR2  = @($allRecords | Where-Object { $_.Flags -like "bc:*" }).Count
$cR3  = @($allRecords | Where-Object { $_.Flags -like "bf:*" }).Count
$cR4  = @($allRecords | Where-Object { $_.Flags -like "ba:*" }).Count
$cR5a = @($allRecords | Where-Object { $_.Flags -like "r5a:*" }).Count
$cR5b = @($allRecords | Where-Object { $_.Flags -like "r5b:*" }).Count
$cR5c = @($allRecords | Where-Object { $_.Flags -like "r5c:*" }).Count
$cR6  = @($allRecords | Where-Object { $_.Flags -like "r6:*" }).Count
$cR7  = @($allRecords | Where-Object { $_.Flags -like "r7:*" }).Count
$total = $allRecords.Count

$groups = $allRecords | Group-Object FS | Sort-Object Name

function hesc($s) { $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' }

function estCell($estDt, $estParsed) {
    if ([string]::IsNullOrWhiteSpace($estDt)) { return "<span style='color:#57606a;font-style:italic'>&#8212;</span>" }
    if ($null -ne $estParsed -and $estParsed -le $refDate) {
        return "<span style='color:#b91c1c;font-weight:600'>$($estParsed.ToString('yyyy-MM-dd'))</span>"
    }
    if ($null -ne $estParsed) { return $estParsed.ToString('yyyy-MM-dd') }
    return hesc $estDt
}

function commentCell($c) {
    if ([string]::IsNullOrWhiteSpace($c)) { return "<span style='color:#57606a;font-style:italic'>&#8212;</span>" }
    $full  = hesc $c
    $short = if ($c.Length -gt 50) { (hesc $c.Substring(0,50)) + "&#8230;" } else { $full }
    return "<span title='$full' style='cursor:default'>$short</span>"
}

function flagBadges($flags) {
    $out = ""
    foreach ($f in $flags) { $p = $f -split ":",2; $out += "<span class='badge $($p[0])'>$($p[1])</span>" }
    return $out
}

$pmp = "https://w3.ibm.com/services/tools/marketplace/newOpenSeatProfile.spr?openSeatId="

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine('<!DOCTYPE html>')
$null = $sb.AppendLine('<html lang="en"><head><meta charset="UTF-8"/>')
$null = $sb.AppendLine('<title>IBM Industrial July 2026 - Compliance by FS Intranet ID</title>')
$null = $sb.AppendLine('<style>')
$css = @'
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"Segoe UI",system-ui,sans-serif;font-size:13px;line-height:1.5;background:#fff;color:#1f2328}
.wrap{max-width:1380px;margin:0 auto;padding:20px 16px}
h1{font-size:18px;font-weight:700;margin-bottom:3px}
.sub{color:#57606a;font-size:12px;margin-bottom:16px}
.kpi-row{display:grid;grid-template-columns:repeat(9,1fr);gap:8px;margin-bottom:18px}
.kpi{background:#f7f8fa;border:1px solid #e5e7eb;border-radius:6px;padding:8px 10px}
.kpi .num{font-size:22px;font-weight:700}.kpi .lbl{font-size:10px;color:#57606a;margin-top:1px}
.c1 .num{color:#b91c1c}.c2 .num{color:#b45309}.c3 .num{color:#7c5cd8}.c4 .num{color:#1d4ed8}
.c5 .num{color:#047857}.c6 .num{color:#b45309}.c7 .num{color:#7c5cd8}.c8 .num{color:#0e7490}
.rules{background:#f7f8fa;border:1px solid #e5e7eb;border-radius:6px;padding:10px 14px;margin-bottom:16px;font-size:11px}
.rules h3{font-size:12px;font-weight:600;margin-bottom:6px}
.rules table{font-size:11px;width:100%;border-collapse:collapse}
.rules th{background:#e5e7eb;padding:4px 8px;text-align:left}
.rules td{padding:3px 8px;border:1px solid #e5e7eb;vertical-align:top}
.toc{background:#f7f8fa;border:1px solid #e5e7eb;border-radius:6px;padding:10px 14px;margin-bottom:16px;font-size:12px;line-height:2.2}
.toc a{color:#3b82d4;text-decoration:none;margin-right:2px}.toc a:hover{text-decoration:underline}
.toc-cnt{background:#e5e7eb;border-radius:9px;padding:0 6px;font-size:11px;font-weight:600;margin-right:8px}
.legend{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px;font-size:11px}
.legend span{display:flex;align-items:center;gap:3px}
.fs-block{margin-bottom:26px}
.fs-header{display:flex;align-items:center;background:#1f2328;color:#fff;padding:8px 12px;border-radius:6px 6px 0 0}
.fs-name{font-weight:600;font-size:13px}
.fs-count{background:#3b82d4;color:#fff;border-radius:9px;padding:1px 8px;font-size:11px;font-weight:600;margin-left:auto}
.tbl-wrap{overflow-x:auto;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 6px 6px}
table{width:100%;border-collapse:collapse;font-size:11px}
th{background:#f7f8fa;border:1px solid #e5e7eb;padding:6px 8px;text-align:left;font-weight:600;white-space:nowrap}
td{border:1px solid #e5e7eb;padding:4px 7px;vertical-align:top;max-width:180px;word-break:break-word}
td.cmt{max-width:230px;font-size:10px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
tbody tr:nth-child(even) td{background:#fafbfc}
tbody tr:hover td{background:#f0f6ff}
.badge{display:inline-block;padding:1px 5px;border-radius:9px;font-size:10px;font-weight:600;white-space:nowrap;margin:1px 1px 1px 0;line-height:1.7}
.be{background:#fee2e2;color:#991b1b}.bc{background:#fef3c7;color:#92400e}
.bf{background:#ede9fe;color:#5b21b6}.ba{background:#dbeafe;color:#1e40af}
.r5a{background:#dcfce7;color:#166534}.r5b{background:#ffedd5;color:#9a3412}.r5c{background:#fce7f3;color:#9d174d}
.r6{background:#cffafe;color:#155e75}
.r7{background:#fef9c3;color:#713f12}
.seat-link{color:#3b82d4;text-decoration:none;font-variant-numeric:tabular-nums;white-space:nowrap}
.seat-link:hover{text-decoration:underline}
footer{margin-top:28px;padding-top:10px;border-top:1px solid #e5e7eb;text-align:center;font-size:11px;color:#57606a}
'@
$null = $sb.AppendLine($css)
$null = $sb.AppendLine('</style></head><body><div class="wrap">')
$null = $sb.AppendLine('<h1>IBM Industrial - July 2026 Demand Compliance by FS Intranet ID</h1>')
$null = $sb.AppendLine("<p class='sub'>Source: Ind July month 30 days demands.xlsx &nbsp;|&nbsp; Ref date: 10 Jul 2026 &nbsp;|&nbsp; $total flagged records &nbsp;|&nbsp; Open Seat IDs link to IBM PMP &nbsp;|&nbsp; Additional Comments truncated to 50 chars (hover for full text)</p>")

$null = $sb.AppendLine("<div class='kpi-row'>
  <div class='kpi c1'><div class='num'>$cR1</div><div class='lbl'>EST Non Compliant</div></div>
  <div class='kpi c2'><div class='num'>$cR2</div><div class='lbl'>Comment Non Compliant</div></div>
  <div class='kpi c3'><div class='num'>$cR3</div><div class='lbl'>Track Type/Fieldglass mismatch</div></div>
  <div class='kpi c4'><div class='num'>$cR4</div><div class='lbl'>Track Type/FA mismatch</div></div>
  <div class='kpi c5'><div class='num'>$cR5a</div><div class='lbl'>Recruiting w/o Ticket</div></div>
  <div class='kpi c6'><div class='num'>$cR5b</div><div class='lbl'>FA Bench/Rolloff with Hiring Ticket</div></div>
  <div class='kpi c7'><div class='num'>$cR5c</div><div class='lbl'>SubK / FG Mismatch</div></div>
  <div class='kpi c8'><div class='num'>$cR6</div><div class='lbl'>Backfill N but Ranking 999</div></div>
  <div class='kpi c1'><div class='num'>$cR7</div><div class='lbl'>Confirmed &gt; 2 Working Days</div></div>
</div>")

$null = $sb.AppendLine("<div class='rules'><h3>Rule Definitions</h3><table><thead><tr><th>Rule</th><th>Condition</th><th>Flag</th></tr></thead><tbody>
<tr><td>R1</td><td>Est Strt Dt blank OR &le; 10 Jul 2026</td><td><span class='badge be'>EST Non Compliant</span></td></tr>
<tr><td>R2</td><td>Additional Comments blank OR last date &lt; 25 Jun 2026</td><td><span class='badge bc'>Comment Non Compliant</span></td></tr>
<tr><td>R3</td><td>Track contains &quot;contractor&quot; AND Fieldglass flag blank/N</td><td><span class='badge bf'>Track Type/Fieldglass mismatch</span></td></tr>
<tr><td>R4a</td><td>Track = Actively recruiting / New hire identified &rarr; FA must be External hire</td><td><span class='badge ba'>Track Type/FA mismatch</span></td></tr>
<tr><td>R4b</td><td>Track = Candidate identified / Proj team reviewing / Actively searching / GR being pursued / Not urgent / Roll-off &rarr; FA must be Bench/Rolloff</td><td><span class='badge ba'>Track Type/FA mismatch</span></td></tr>
<tr><td>R4c</td><td>Track = Rotation being pursued &rarr; FA must be Rotation</td><td><span class='badge ba'>Track Type/FA mismatch</span></td></tr>
<tr><td>R4d</td><td>Track = Contractor being pursued (core/non-core) / Contractor identified / Temporarily mitigated &rarr; FA must be Contractor</td><td><span class='badge ba'>Track Type/FA mismatch</span></td></tr>
<tr><td>R5a</td><td>Track = Actively recruiting AND Hiring Ticket (AskFile) is blank</td><td><span class='badge r5a'>No Hiring Ticket</span></td></tr>
<tr><td>R5b</td><td>FA = Bench/Rolloff AND Hiring Ticket (AskFile) is non-blank</td><td><span class='badge r5b'>FA Bench/Rolloff with Hiring Ticket</span></td></tr>
<tr><td>R5c</td><td>Track = Contractor track types AND Fieldglass Request Flag = N or blank</td><td><span class='badge r5c'>SubK/FG Mismatch</span></td></tr>
<tr><td>R6</td><td>Backfilled = N AND Priority Ranking = 999</td><td><span class='badge r6'>Backfill N but Ranking 999</span></td></tr>
<tr><td>R7</td><td>Candidate Disposition = Confirmed AND Confirmed Date is more than 2 working days before today</td><td><span class='badge r7'>Candidate confirmed more than 2 days</span></td></tr>
</tbody></table></div>")

$null = $sb.AppendLine("<div class='legend'>
  <span><span class='badge be'>EST Non Compliant</span></span>
  <span><span class='badge bc'>Comment Non Compliant</span></span>
  <span><span class='badge bf'>Track Type/Fieldglass mismatch</span></span>
  <span><span class='badge ba'>Track Type/FA mismatch</span></span>
  <span><span class='badge r5a'>No Hiring Ticket</span> Recruiting w/o AskFile ticket</span>
  <span><span class='badge r5b'>FA Bench/Rolloff with Hiring Ticket</span> Bench/Rolloff FA but has AskFile ticket</span>
  <span><span class='badge r5c'>SubK/FG Mismatch</span> Contractor track but FG blank/N</span>
  <span><span class='badge r6'>Backfill N but Ranking 999</span> Backfilled=N with Priority Ranking 999</span>
  <span><span class='badge r7'>Candidate confirmed more than 2 days</span> Confirmed Disposition but Confirmed Date &gt; 2 working days ago</span>
</div>")

# TOC
$null = $sb.Append("<div class='toc'><strong>FS Staff ($($groups.Count))</strong><br>")
$fi = 1
foreach ($g in $groups) {
    $null = $sb.Append("<a href='#fs$fi'>$(hesc $g.Name)</a> <span class='toc-cnt'>$($g.Count)</span>&nbsp;&nbsp;")
    $fi++
}
$null = $sb.AppendLine("</div>")

# FS blocks
$thRow = "<tr><th>#</th><th>Open Seat ID</th><th>Client Name</th><th>Industry</th><th>Open Seat Title</th><th>Est Strt Dt</th><th>Additional Comments</th><th>Candidate Track Type</th><th>Fulfillment Action</th><th>FG</th><th>Hiring Ticket (AskFile)</th><th>Compliance Flags</th></tr>"
$fi = 1
foreach ($g in $groups) {
    $cnt = $g.Count; $noun = if ($cnt -eq 1) {"demand"} else {"demands"}
    $null = $sb.AppendLine("<div class='fs-block' id='fs$fi'><div class='fs-header'><span class='fs-name'>$(hesc $g.Name)</span><span class='fs-count'>$cnt $noun</span></div><div class='tbl-wrap'><table><thead>$thRow</thead><tbody>")
    $rn = 1
    foreach ($rec in $g.Group) {
        $sl  = "<a href='$pmp$($rec.SeatId)' target='_blank' class='seat-link'>$($rec.SeatId)</a>"
        $ed  = estCell $rec.EstDt $rec.EstParsed
        $cmt = commentCell $rec.Comments
        $fgd = if ([string]::IsNullOrWhiteSpace($rec.FG)) {"&#8212;"} else {hesc $rec.FG}
        if ([string]::IsNullOrWhiteSpace($rec.Ticket)) {
            $tkd = "&#8212;"
        } elseif (-not [string]::IsNullOrWhiteSpace($rec.AskDetail)) {
            $tkd = (hesc $rec.Ticket) + " : " + (hesc $rec.AskDetail)
        } else {
            $tkd = hesc $rec.Ticket
        }
        $bdg = flagBadges $rec.Flags
        $null = $sb.AppendLine("<tr><td>$rn</td><td>$sl</td><td>$(hesc $rec.Client)</td><td>$(hesc $rec.Industry)</td><td>$(hesc $rec.Title)</td><td>$ed</td><td class='cmt'>$cmt</td><td>$(hesc $rec.Track)</td><td>$(hesc $rec.FA)</td><td>$fgd</td><td>$tkd</td><td>$bdg</td></tr>")
        $rn++
    }
    $null = $sb.AppendLine("</tbody></table></div></div>")
    $fi++
}
$null = $sb.AppendLine("<footer>Made with IBM Bob</footer>")
$null = $sb.AppendLine("</div></body></html>")

$out = $sb.ToString()
$outPath = Join-Path $root "IND_July2026_Compliance_by_FS.html"
[System.IO.File]::WriteAllText($outPath, $out, [System.Text.Encoding]::UTF8)
Write-Host "Done. Size: $([math]::Round($out.Length/1KB,1)) KB | Records: $total"

# =============================================================================
# SEND COMPLIANCE EMAILS — one per FS Intranet ID, CC smouttou@in.ibm.com + additional recipients
# Set $sendEmails = $true only when you explicitly want to dispatch emails.
# =============================================================================
$sendEmails = $false   # <-- change to $true when ready to send

if ($sendEmails) {

$reportLink = "https://htmlpreview.github.io/?https://raw.githubusercontent.com/dineshkgupta1007/watsonx-challenge-2026/main/IND_July2026_Compliance_by_FS.html"
$cc         = "smouttou@in.ibm.com;rakmukhe@in.ibm.com;jayghosh@in.ibm.com;manjhans@in.ibm.com;Dineshkgupta@in.ibm.com"

$fsList = $groups | Select-Object -ExpandProperty Name | Where-Object { $_ -ne "(blank)" }

$outlook = New-Object -ComObject Outlook.Application
foreach ($fs in $fsList) {
    $body = @"
Hi,

Your open seats have been flagged for non-compliance items in the IBM Industrial Sector July 2026 Demand Compliance Report.

Please review your section in the report at the link below and action all flagged compliance items at the earliest - latest by end of day tomorrow.

Report Link:
$reportLink

Regards,
Dinesh Gupta
IBM Industrial Sector - Staffing
"@
    $mail = $outlook.CreateItem(0)
    $mail.To      = $fs
    $mail.CC      = $cc
    $mail.Subject = "Action Required: IBM Industrial Demand Compliance - July 2026"
    $mail.Body    = $body
    $mail.Send()
    Write-Host "Sent to: $fs (CC: $cc)"
    Start-Sleep -Milliseconds 500
}
Write-Host "All $($fsList.Count) emails sent."

} else {
    Write-Host "Email sending skipped. Set `$sendEmails = `$true to dispatch."
}

