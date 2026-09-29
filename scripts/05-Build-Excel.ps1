<#
.SYNOPSIS
  Builds Excel workbooks (via Excel COM, so Excel must be installed):
    workbook.xlsx               shareable: no Agents sheet, no agent names/coaching notes
    workbook_MANAGER_ONLY.xlsx  everything
.EXAMPLE
  .\scripts\05-Build-Excel.ps1 -RunDir .\run            # manager edition
  .\scripts\05-Build-Excel.ps1 -RunDir .\run -Shareable # shareable edition
#>
param([string]$RunDir = '.\run', [switch]$Shareable)
$Company = $Shareable
$ErrorActionPreference = 'Stop'
$RunDir = (Resolve-Path $RunDir).Path
$an  = "$RunDir\output"
$out = if ($Company) { "$an\workbook.xlsx" } else { "$an\workbook_MANAGER_ONLY.xlsx" }
$m = [IO.File]::ReadAllText("$an\metrics.json", (New-Object Text.UTF8Encoding $false)) | ConvertFrom-Json
$agents  = Import-Csv "$an\agents.csv"
$tickets = Import-Csv "$an\tickets_tagged.csv"
function Write-Sheet($wb, [string]$name, [string[]]$headers, $rows, [int[]]$widths) {
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = $name
    $r = @($rows).Count; $c = $headers.Count
    $arr = New-Object 'object[,]' ($r + 1), $c
    for ($j = 0; $j -lt $c; $j++) { $arr[0, $j] = $headers[$j] }
    $i = 1
    foreach ($row in $rows) {
        for ($j = 0; $j -lt $c; $j++) {
            $v = $row[$j]
            if ($v -is [string] -and $v.Length -gt 32000) { $v = $v.Substring(0, 32000) }
            if ($v -is [string] -and $v -match '^-?\d+(\.\d+)?$' -and $v.Length -lt 12) { $v = [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture) }
            if ($v -is [string] -and $v.StartsWith('=')) { $v = "'" + $v }
            $arr[$i, $j] = $v
        }
        $i++
    }
    $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($r + 1, $c))
    $rng.Value2 = $arr
    $hdr = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $c))
    $hdr.Font.Bold = $true; $hdr.Interior.Color = 0x3B2A1F; $hdr.Font.Color = 0xFFFFFF
    $ws.ListObjects.Add(1, $rng, $null, 1) | Out-Null
    for ($j = 0; $j -lt $c; $j++) {
        $w = if ($widths -and $j -lt $widths.Count) { $widths[$j] } else { 14 }
        $ws.Columns.Item($j + 1).ColumnWidth = $w
    }
    $ws.Rows.Item(1).WrapText = $true
    $ws.Activate(); $ws.Application.ActiveWindow.SplitRow = 1; $ws.Application.ActiveWindow.FreezePanes = $true
    return $ws
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
    $first = $wb.Worksheets.Item(1)

    # 1. Summary
    $S = $m.scopes
    $metricRows = @(
        @('Tickets (total)', 'tickets_total'), @('Customer tickets (excl. noise)', 'tickets_real'), @('Noise (spam, auto-replies, etc.)', 'noise'),
        @('Median first response, excl. weekends (h)', 'frt_median'), @('90th pct first response, excl. weekends (h)', 'frt_p90'), @('First response <= 4h, excl. weekends (%)', 'frt_within_4h'),
        @('First response <= 8h, excl. weekends (%)', 'frt_within_8h'), @('First response <= 24h, excl. weekends (%)', 'frt_within_24h'), @('Median resolution (h)', 'res_median'),
        @('90th pct resolution (h)', 'res_p90'), @('Resolved <= 24h (%)', 'res_within_24h'), @('Resolved <= 72h (%)', 'res_within_72h'),
        @('Resolved <= 7 days (%)', 'res_within_7d'), @('Median closure (h)', 'close_median'), @('90th pct closure (h)', 'close_p90'),
        @('Agent public replies per ticket', 'avg_agent_replies'), @('Messages per ticket (incl. notes)', 'avg_messages'),
        @('First contact resolution (%)', 'first_reply_resolved_pct'), @('Customer need met (%)', 'resolved_pct'), @('Ended positive or neutral (%)', 'pos_neutral_pct'), @('Customer wrote back after resolution (%)', 'reopen_pct'),
        @('Clarity (1-5)', 'clarity'), @('Empathy (1-5)', 'empathy'), @('Ownership (1-5)', 'ownership'), @('Closed after customer stopped responding to follow-ups', 'auto_closed'), @('Median resolution excl. no customer reply (h)', 'res_median_excl_auto'), @('Resolved <= 72h excl. no customer reply (%)', 'res_within_72h_excl_auto'), @('Open time waiting on customer (%)', 'wait_customer_pct'), @('Open / pending now', 'open_backlog'),
        @('Product-preventable tickets', 'prod_n')
    )
    $rows = foreach ($mr in $metricRows) { , @(@($mr[0]) + @($m.scopeOrder | ForEach-Object { $S.$_.($mr[1]) })) }
    $ws = Write-Sheet $wb 'Summary' (@('Metric') + @($m.scopeOrder)) $rows @(50, 16, 16, 16, 16, 16)

    # 2. Themes (per scope)
    $rows = foreach ($sc in $m.scopeOrder) {
        foreach ($t in $S.$sc.by_theme) {
            , @($sc, $t.theme, $t.count, $t.share, $t.frt_median, $t.res_median, $t.avg_messages, $t.frr_pct, $t.posneu_pct, (@($t.top_sub_themes) -join ' | '))
        }
    }
    Write-Sheet $wb 'Contact reasons' @('Scope', 'Contact reason', 'Tickets', 'Share of customer tickets (%)', 'Median 1st response, excl. weekends (h)', 'Median resolution (h)', 'Messages / ticket', 'FCR (%)', 'Positive / neutral end (%)', 'Top specific issues') $rows @(16, 26, 9, 12, 12, 12, 10, 10, 10, 90) | Out-Null

    if (-not $Company) {
    # 3. Agents
    $cols = 'group', 'agent', 'tickets', 'score', 'frt_median', 'res_median', 'avg_agent_replies', 'frr_pct', 'resolved_pct', 'reopen_pct', 'clarity', 'empathy', 'ownership', 'quality', 'negative_pct'
    $rows = foreach ($a in $agents) { , @($cols | ForEach-Object { $a.$_ }) }
    Write-Sheet $wb 'Agents' @('Group', 'Agent', 'Tickets', 'Score', 'Median 1st response, excl. weekends (h)', 'Median resolution (h)', 'Replies / ticket', 'FCR (%)', 'Need met (%)', 'Reopened (%)', 'Clarity', 'Empathy', 'Ownership', 'Quality avg', 'Negative end (%)') $rows @(18, 26, 9, 10, 12, 12, 10, 10, 10, 10, 9, 9, 10, 10, 10) | Out-Null

    }
    # 3b. Product recommendations
    if (Test-Path "$RunDir\product_recommendations.json") {
        $pr = [IO.File]::ReadAllText("$RunDir\product_recommendations.json", (New-Object Text.UTF8Encoding $false)) | ConvertFrom-Json
        $i = 0
        $rows = foreach ($p in $pr) { $i++; , @($i, $p.title, $p.product_area, $p.type, $p.tickets, $p.problem, $p.evidence, ((@($p.recommendation) | ForEach-Object { "- $_" }) -join "`n"), (@($p.ticket_ids) -join ', ')) }
        $ws = Write-Sheet $wb 'Product recommendations' @('#', 'Recommendation', 'Product area', 'Type', 'Tickets', 'What customers run into', 'Evidence', 'Suggested changes', 'Example tickets') $rows @(4, 45, 24, 14, 8, 70, 50, 80, 30)
        $ws.UsedRange.WrapText = $true; $ws.UsedRange.VerticalAlignment = -4160
    }
    # 4. Product feedback
    $pf = @($tickets | Where-Object { $_.theme -ne 'NOISE' -and ($_.preventable_by -eq 'Product' -or $_.root_cause -in 'Product defect/bug', 'Product usability/UX gap', 'Missing feature') })
    $rows = foreach ($t in ($pf | Sort-Object product_area, sub_theme)) { , @($t.ticket_id, $t.group, $t.product_area, $t.component, $t.theme, $t.sub_theme, $t.root_cause, $t.product_fix, $t.ai_summary, $t.resolution_hours, $t.message_count, $t.escalated) }
    Write-Sheet $wb 'Product feedback' @('Ticket', 'Group', 'Product area', 'Component', 'Contact reason', 'Specific issue', 'Root cause', 'Suggested product fix', 'Summary', 'Resolution (h)', 'Messages', 'Escalated') $rows @(10, 16, 22, 12, 20, 30, 20, 60, 60, 11, 9, 9) | Out-Null

    # 5. All tickets
    $tc = 'ticket_id', 'group', 'agent', 'created_at', 'status', 'source', 'tier', 'subject', 'theme', 'issue_cluster', 'sub_theme', 'product_area', 'component', 'root_cause', 'preventable_by', 'product_fix', 'ai_summary', 'first_response_hours', 'first_response_hours_calendar', 'resolution_hours', 'close_hours', 'reopened', 'message_count', 'agent_public_replies', 'private_notes', 'resolved', 'fcr', 'ai_first_reply_resolved', 'auto_closed', 'wait_customer_hours', 'wait_team_hours', 'clarity', 'empathy', 'ownership', 'sentiment_end', 'escalated', 'notable', 'fd_category', 'fd_request_type'
    if ($Company) { $tc = @($tc | Where-Object { $_ -notin 'agent', 'notable', 'clarity', 'empathy', 'ownership' }) }
    $rows = foreach ($t in $tickets) { , @($tc | ForEach-Object { $t.$_ }) }
    Write-Sheet $wb 'All tickets' @($tc) $rows $null | Out-Null

    # 6. Freshdesk links note
    $first.Delete()
    $wb.Worksheets.Item('Summary').Activate()
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    $wb.Close($false)
    Write-Host "Wrote $out"
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
