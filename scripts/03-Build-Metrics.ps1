<#
.SYNOPSIS
  Merges the export with the AI tags and computes all report metrics.
  Writes <RunDir>\output\metrics.json, tickets_tagged.csv, agents.csv
.EXAMPLE
  .\scripts\03-Build-Metrics.ps1 -RunDir .\run
#>
param(
    [string]$RunDir = '.\run',
    [string]$Config = (Join-Path $PSScriptRoot '..\config.json')
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Config)) { $Config = Join-Path $PSScriptRoot '..\config.example.json' }
$cfg = Get-Content $Config -Raw | ConvertFrom-Json
$RunDir = (Resolve-Path $RunDir).Path
$out = "$RunDir\output"; New-Item -ItemType Directory -Force $out | Out-Null
$inv = [Globalization.CultureInfo]::InvariantCulture

$tickets = Import-Csv (Get-ChildItem "$RunDir\tickets_*.csv" | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
$convAll = Import-Csv (Get-ChildItem "$RunDir\conversations_*.csv" | Sort-Object LastWriteTime | Select-Object -Last 1).FullName

# ---- AI tags (JSONL) and optional issue clusters
$tags = @{}
foreach ($f in Get-ChildItem "$RunDir\tags\*.jsonl" -ErrorAction SilentlyContinue) {
    foreach ($line in [IO.File]::ReadAllLines($f.FullName)) {
        if (-not $line.Trim()) { continue }
        try { $o = $line | ConvertFrom-Json; $tags[[string]$o.ticket_id] = $o } catch { Write-Warning "Bad JSON in $($f.Name)" }
    }
}
$clusters = @{}
if (Test-Path "$RunDir\issue_clusters.csv") { foreach ($c in (Import-Csv "$RunDir\issue_clusters.csv")) { $clusters[[string]$c.ticket_id] = $c.issue_cluster } }
Write-Host "Tags loaded: $($tags.Count) / $($tickets.Count)"

function D($s) { if ($s) { [datetime]::ParseExact($s, 'yyyy-MM-dd HH:mm:ss', $inv) } }
function N($s) { if ($s -ne $null -and $s -ne '') { [double]::Parse($s, $inv) } }
function Pct($arr, $p) {
    $a = @($arr | Where-Object { $_ -ne $null } | Sort-Object)
    if ($a.Count -eq 0) { return $null }
    return [math]::Round($a[[math]::Max(0, [math]::Ceiling($p * $a.Count) - 1)], 1)
}
function Share($arr, $pred) {
    $a = @($arr | Where-Object { $_ -ne $null })
    if ($a.Count -eq 0) { return $null }
    return [math]::Round(100 * @($a | Where-Object $pred).Count / $a.Count, 1)
}
function Avg($arr) {
    $a = @($arr | Where-Object { $_ -ne $null -and $_ -gt 0 })
    if ($a.Count -eq 0) { return $null }
    return [math]::Round(($a | Measure-Object -Average).Average, 2)
}
function CountBy($rows, $prop, $top = 50) {
    $o = [ordered]@{}
    $rows | Group-Object $prop | Sort-Object Count -Descending | Select-Object -First $top | ForEach-Object {
        $k = if ($_.Name) { $_.Name } else { '(blank)' }; $o[$k] = $_.Count }
    return $o
}
# hours from $start for $hours, not counting Saturday/Sunday (UTC)
function WeekdayHours([datetime]$start, $hours) {
    if ($hours -eq $null) { return $null }
    $end = $start.AddHours($hours); $h = 0.0; $x = $start
    while ($x -lt $end) {
        $nx = $x.Date.AddDays(1); if ($nx -gt $end) { $nx = $end }
        if ($x.DayOfWeek -ne 'Saturday' -and $x.DayOfWeek -ne 'Sunday') { $h += ($nx - $x).TotalHours }
        $x = $nx
    }
    return [math]::Round($h, 2)
}

# ---- per-ticket facts from the conversation log
$convBy = @{}
foreach ($m in $convAll) {
    if (-not $convBy.ContainsKey($m.ticket_id)) { $convBy[$m.ticket_id] = New-Object System.Collections.Generic.List[object] }
    $convBy[$m.ticket_id].Add($m)
}
$facts = @{}
foreach ($t in $tickets) {
    $msgs = if ($convBy.ContainsKey($t.ticket_id)) { $convBy[$t.ticket_id].ToArray() } else { @() }
    $notes = @($msgs | Where-Object { $_.is_private -eq 'True' })
    $autoClosed = @($notes | Where-Object { $_.body_text -match $cfg.autoClose.closedNotePattern }).Count -gt 0
    $autoNotes = @($notes | Where-Object { $_.body_text -match $cfg.autoClose.followUpNotePattern }).Count
    # split open time: after an agent's public reply we wait on the customer, otherwise on the team
    $waitCust = 0.0; $waitTeam = 0.0
    $c0 = D $t.created_at; $end = D $t.resolved_at
    if ($end) {
        $state = 'team'; $prev = $c0
        foreach ($m in ($msgs | Where-Object { $_.is_private -ne 'True' -and $_.seq -ne '0' } | Sort-Object { D $_.created_at })) {
            $mt = D $m.created_at; if ($mt -gt $end) { break }
            $d = ($mt - $prev).TotalHours
            if ($d -gt 0) { if ($state -eq 'team') { $waitTeam += $d } else { $waitCust += $d } }
            $state = if ($m.author_role -eq 'Agent') { 'cust' } else { 'team' }; $prev = $mt
        }
        $d = ($end - $prev).TotalHours
        if ($d -gt 0) { if ($state -eq 'team') { $waitTeam += $d } else { $waitCust += $d } }
    }
    $facts[$t.ticket_id] = @{ auto_closed = $autoClosed; auto_notes = $autoNotes; wait_cust = $waitCust; wait_team = $waitTeam }
}

# ---- merged rows
$fcrMax = [int]$cfg.fcrMaxCustomerReplies
$rows = foreach ($t in $tickets) {
    $g = $tags[[string]$t.ticket_id]; $f = $facts[$t.ticket_id]
    $created = D $t.created_at; $closed = D $t.closed_at
    $frtCal = N $t.first_response_hours
    [pscustomobject][ordered]@{
        ticket_id = $t.ticket_id; group = $t.group; agent = $(if ($t.agent) { $t.agent } else { '(unassigned)' })
        subject = $t.subject; status = $t.status; source = $(if ($t.source) { $t.source } else { $cfg.blankSourceLabel })
        tier = $(if ($t.($cfg.fields.tier)) { $t.($cfg.fields.tier) } else { '(none)' })
        fd_category = $t.($cfg.fields.category); fd_request_type = $t.($cfg.fields.requestType)
        created_at = $t.created_at; created_date = $created.ToString('yyyy-MM-dd'); weekday = $created.DayOfWeek.ToString(); hour_utc = $created.Hour
        first_response_hours = $(if ($cfg.excludeWeekendsFromFirstResponse) { WeekdayHours $created $frtCal } else { $frtCal })
        first_response_hours_calendar = $frtCal
        resolution_hours = N $t.resolution_hours
        close_hours = $(if ($closed) { [math]::Round(($closed - $created).TotalHours, 2) })
        reopened = ($t.reopened -eq 'True')
        message_count = [int]$t.message_count; customer_messages = [int]$t.customer_messages
        agent_public_replies = [int]$t.agent_public_replies; private_notes = [int]$t.private_notes - $f.auto_notes
        automated_notes = $f.auto_notes; auto_closed = $f.auto_closed
        wait_customer_hours = [math]::Round($f.wait_cust, 2); wait_team_hours = [math]::Round($f.wait_team, 2)
        theme = $(if ($g) { ([string]$g.theme).ToUpper() -replace '[^A-Z_]', '' -replace '^_+|_+$', '' } else { 'UNTAGGED' })
        issue_cluster = $(if ($clusters.ContainsKey([string]$t.ticket_id)) { $clusters[[string]$t.ticket_id] } else { $g.sub_theme })
        sub_theme = $g.sub_theme; product_area = $g.product_area
        component = $(if ($g.component) { $g.component } else { $g.sensor_type })
        root_cause = $g.root_cause
        preventable_by = $(if ($g.preventable_by -eq 'Documentation/KB gap') { 'Self-service/KB' } else { $g.preventable_by })
        product_fix = $g.product_fix; ai_summary = $g.summary; resolved = $g.resolved
        fcr = $(if ($t.resolved_at) { if (([int]$t.customer_messages - 1) -le $fcrMax) { 'Yes' } else { 'No' } } else { '' })
        ai_first_reply_resolved = $g.first_reply_resolved
        clarity = $(if ($g) { [int]$g.clarity }); empathy = $(if ($g) { [int]$g.empathy }); ownership = $(if ($g) { [int]$g.ownership })
        sentiment_end = $g.sentiment_end; escalated = $g.escalated; notable = $g.notable
    }
}
$rows | Export-Csv "$out\tickets_tagged.csv" -NoTypeInformation -Encoding UTF8

$prodPred = { $_.preventable_by -eq 'Product' -or $_.root_cause -in 'Product defect/bug', 'Product usability/UX gap', 'Missing feature' }
function ScopeStats($all) {
    $real = @($all | Where-Object { $_.theme -ne 'NOISE' })
    $q = @($real | Where-Object { $_.clarity -gt 0 })
    $notAuto = @($real | Where-Object { -not $_.auto_closed })
    $known = @($real | Where-Object { $_.sentiment_end -in 'Positive', 'Neutral', 'Negative' })
    $wc = ($real | Measure-Object wait_customer_hours -Sum).Sum; $wt = ($real | Measure-Object wait_team_hours -Sum).Sum
    $s = [ordered]@{
        tickets_total = @($all).Count; noise = @($all).Count - $real.Count; tickets_real = $real.Count
        status = CountBy $all 'status'
        frt_n = @($real | Where-Object { $_.first_response_hours -ne $null }).Count
        frt_median = Pct ($real.first_response_hours) 0.5; frt_p90 = Pct ($real.first_response_hours) 0.9
        frt_within_4h = Share ($real.first_response_hours) { $_ -le 4 }
        frt_within_8h = Share ($real.first_response_hours) { $_ -le 8 }
        frt_within_24h = Share ($real.first_response_hours) { $_ -le 24 }
        frt_within_24h_calendar = Share ($real.first_response_hours_calendar) { $_ -le 24 }
        res_n = @($real | Where-Object { $_.resolution_hours -ne $null }).Count
        res_median = Pct ($real.resolution_hours) 0.5; res_p90 = Pct ($real.resolution_hours) 0.9
        res_within_24h = Share ($real.resolution_hours) { $_ -le 24 }
        res_within_72h = Share ($real.resolution_hours) { $_ -le 72 }
        res_within_7d = Share ($real.resolution_hours) { $_ -le 168 }
        auto_closed = @($real | Where-Object { $_.auto_closed }).Count
        res_median_excl_auto = Pct ($notAuto.resolution_hours) 0.5
        res_p90_excl_auto = Pct ($notAuto.resolution_hours) 0.9
        res_within_72h_excl_auto = Share ($notAuto.resolution_hours) { $_ -le 72 }
        wait_customer_pct = $(if (($wc + $wt) -gt 0) { [math]::Round(100 * $wc / ($wc + $wt)) })
        close_median = Pct ($real.close_hours) 0.5; close_p90 = Pct ($real.close_hours) 0.9
        avg_messages = [math]::Round((@($real.message_count) | Measure-Object -Average).Average, 1)
        avg_agent_replies = [math]::Round((@($real.agent_public_replies) | Measure-Object -Average).Average, 1)
        reopen_pct = Share ($real | ForEach-Object { $_.reopened }) { $_ }
        open_backlog = @($all | Where-Object { $_.status -notin 'Closed', 'Resolved' }).Count
        first_reply_resolved_pct = Share (@($real | Where-Object { $_.fcr }).fcr) { $_ -eq 'Yes' }
        resolved_pct = Share (@($real | Where-Object { $_.resolved -in 'Yes', 'No' }).resolved) { $_ -eq 'Yes' }
        clarity = Avg $q.clarity; empathy = Avg $q.empathy; ownership = Avg $q.ownership
        sentiment = CountBy $real 'sentiment_end'
        pos_neutral_pct = Share ($known.sentiment_end) { $_ -ne 'Negative' }
        pos_neutral_by_week = [ordered]@{}
        by_day = [ordered]@{}; by_weekday = CountBy $all 'weekday'; by_hour = [ordered]@{}
        by_source = CountBy $real 'source'; by_tier = CountBy $real 'tier'
        by_theme = @(); by_product_area = CountBy $real 'product_area'
        by_root_cause = CountBy $real 'root_cause'; by_preventable = CountBy $real 'preventable_by'
        by_sensor = CountBy @($real | Where-Object { $_.component } | Where-Object $prodPred) 'component'
        prod_n = @($real | Where-Object $prodPred).Count
        by_prod_area = CountBy @($real | Where-Object $prodPred | Where-Object { $_.product_area -and $_.product_area -ne 'None' }) 'product_area'
    }
    $fb = [ordered]@{ '<1h' = 0; '1-4h' = 0; '4-8h' = 0; '8-24h' = 0; '1-2d' = 0; '>2d' = 0 }
    foreach ($v in @($real.first_response_hours | Where-Object { $_ -ne $null })) {
        if ($v -lt 1) { $fb['<1h']++ } elseif ($v -lt 4) { $fb['1-4h']++ } elseif ($v -lt 8) { $fb['4-8h']++ } elseif ($v -lt 24) { $fb['8-24h']++ } elseif ($v -lt 48) { $fb['1-2d']++ } else { $fb['>2d']++ } }
    $rb = [ordered]@{ '<1d' = 0; '1-3d' = 0; '3-7d' = 0; '7-14d' = 0; '14-30d' = 0 }
    foreach ($v in @($notAuto.resolution_hours | Where-Object { $_ -ne $null })) {
        if ($v -lt 24) { $rb['<1d']++ } elseif ($v -lt 72) { $rb['1-3d']++ } elseif ($v -lt 168) { $rb['3-7d']++ } elseif ($v -lt 336) { $rb['7-14d']++ } else { $rb['14-30d']++ } }
    $s['frt_buckets'] = $fb; $s['res_buckets'] = $rb
    $known | Group-Object { $d = [datetime]::ParseExact($_.created_date, 'yyyy-MM-dd', $inv); $d.AddDays(-(([int]$d.DayOfWeek + 6) % 7)).ToString('yyyy-MM-dd') } |
        Sort-Object Name | Where-Object { $_.Count -ge 20 } |
        ForEach-Object { $s.pos_neutral_by_week[$_.Name] = [math]::Round(100 * @($_.Group | Where-Object { $_.sentiment_end -ne 'Negative' }).Count / $_.Count, 1) }
    $all | Group-Object created_date | Sort-Object Name | ForEach-Object { $s.by_day[$_.Name] = $_.Count }
    0..23 | ForEach-Object { $h = $_; $s.by_hour["$h"] = @($all | Where-Object { $_.hour_utc -eq $h }).Count }
    $s.by_theme = @($real | Group-Object theme | Sort-Object Count -Descending | ForEach-Object {
        $grp = @($_.Group); $gk = @($grp | Where-Object { $_.sentiment_end -in 'Positive', 'Neutral', 'Negative' })
        [ordered]@{
            theme = $_.Name; count = $_.Count; share = [math]::Round(100 * $_.Count / [math]::Max(1, $real.Count), 1)
            frt_median = Pct ($grp.first_response_hours) 0.5; res_median = Pct ($grp.resolution_hours) 0.5
            avg_messages = [math]::Round((@($grp.message_count) | Measure-Object -Average).Average, 1)
            frr_pct = Share (@($grp | Where-Object { $_.fcr }).fcr) { $_ -eq 'Yes' }
            posneu_pct = Share ($gk.sentiment_end) { $_ -ne 'Negative' }
            top_sub_themes = @($grp | Where-Object { $_.issue_cluster } | Group-Object issue_cluster | Sort-Object Count -Descending | Select-Object -First 6 | ForEach-Object { "$($_.Name) ($($_.Count))" })
            preventable = CountBy $grp 'preventable_by'
        }
    })
    return $s
}

$scopeOrder = @($cfg.allLabel) + @($cfg.groups)
$scopes = [ordered]@{}
$scopes[$cfg.allLabel] = ScopeStats @($rows | Where-Object { $_.group -in $cfg.groups })
foreach ($g in $cfg.groups) { $scopes[$g] = ScopeStats @($rows | Where-Object { $_.group -eq $g }) }

# ---- agents (within their own team)
$w = $cfg.agentScoreWeights; $minN = [int]$cfg.agentMinTickets
$agentStats = foreach ($grpName in $cfg.groups) {
    $real = @($rows | Where-Object { $_.group -eq $grpName -and $_.theme -ne 'NOISE' -and $_.agent -ne '(unassigned)' })
    foreach ($a in ($real | Group-Object agent)) {
        $g = @($a.Group); $q = @($g | Where-Object { $_.clarity -gt 0 })
        [pscustomobject][ordered]@{
            group = $grpName; agent = $a.Name; tickets = $g.Count
            frt_median = Pct ($g.first_response_hours) 0.5
            res_median = Pct ($g.resolution_hours) 0.5
            res_median_excl_auto = Pct (@($g | Where-Object { -not $_.auto_closed }).resolution_hours) 0.5
            avg_agent_replies = [math]::Round((@($g.agent_public_replies) | Measure-Object -Average).Average, 1)
            reopen_pct = Share ($g | ForEach-Object { $_.reopened }) { $_ }
            frr_pct = Share (@($g | Where-Object { $_.fcr }).fcr) { $_ -eq 'Yes' }
            resolved_pct = Share (@($g | Where-Object { $_.resolved -in 'Yes', 'No' }).resolved) { $_ -eq 'Yes' }
            clarity = Avg $q.clarity; empathy = Avg $q.empathy; ownership = Avg $q.ownership
            quality = Avg (@($q | ForEach-Object { ($_.clarity + $_.empathy + $_.ownership) / 3 }))
            negative_pct = Share ($g.sentiment_end) { $_ -eq 'Negative' }
            score = $null
        }
    }
}
function Rank($list, $prop) {
    $vals = @($list | Where-Object { $_.$prop -ne $null } | Sort-Object $prop)
    $r = @{}; for ($i = 0; $i -lt $vals.Count; $i++) { $r[$vals[$i].agent] = if ($vals.Count -gt 1) { 1 - $i / ($vals.Count - 1) } else { 0.5 } }
    return $r
}
foreach ($grpName in $cfg.groups) {
    $el = @($agentStats | Where-Object { $_.group -eq $grpName -and $_.tickets -ge $minN })
    if ($el.Count -lt 2) { continue }
    $rF = Rank $el 'frt_median'; $rR = Rank $el 'res_median_excl_auto'
    foreach ($a in $el) {
        $qn = if ($a.quality) { ($a.quality - 1) / 4 } else { 0.5 }
        $fr = if ($a.frr_pct -ne $null) { $a.frr_pct / 100 } else { 0.5 }
        $f = if ($rF.ContainsKey($a.agent)) { $rF[$a.agent] } else { 0.5 }
        $r = if ($rR.ContainsKey($a.agent)) { $rR[$a.agent] } else { 0.5 }
        $a.score = [math]::Round(100 * ($w.quality * $qn + $w.fcr * $fr + $w.firstResponse * $f + $w.resolution * $r), 1)
    }
}
$agentStats | Sort-Object group, @{ e = { if ($_.score -ne $null) { $_.score } else { -1 } }; Descending = $true } |
    Export-Csv "$out\agents.csv" -NoTypeInformation -Encoding UTF8

$dates = @($rows.created_date | Sort-Object)
$result = [ordered]@{
    generated = (Get-Date).ToString('yyyy-MM-dd HH:mm'); period = "$($dates[0]) to $($dates[-1])"
    orgName = $cfg.orgName; scopeOrder = $scopeOrder; scopes = $scopes; agents = @($agentStats)
}
[IO.File]::WriteAllText("$out\metrics.json", (ConvertTo-Json $result -Depth 8), (New-Object Text.UTF8Encoding $false))
Write-Host "Wrote $out\metrics.json, tickets_tagged.csv, agents.csv"
