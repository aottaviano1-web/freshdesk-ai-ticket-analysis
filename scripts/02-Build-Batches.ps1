<#
.SYNOPSIS
  Turns the export into compact per-ticket digests, split into N batch files for parallel AI tagging.
.EXAMPLE
  .\scripts\02-Build-Batches.ps1 -RunDir .\run -Batches 10
#>
param(
    [string]$RunDir = '.\run',
    [int]$Batches = 10,
    [string]$Config = (Join-Path $PSScriptRoot '..\config.json'),
    [int]$MaxCharsPerTicket = 2200
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Config)) { $Config = Join-Path $PSScriptRoot '..\config.example.json' }
$cfg = Get-Content $Config -Raw | ConvertFrom-Json
$RunDir = (Resolve-Path $RunDir).Path

$t = Import-Csv (Get-ChildItem "$RunDir\tickets_*.csv" | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
$c = Import-Csv (Get-ChildItem "$RunDir\conversations_*.csv" | Sort-Object LastWriteTime | Select-Object -Last 1).FullName
New-Item -ItemType Directory -Force "$RunDir\batches", "$RunDir\tags" | Out-Null

$byT = @{}
foreach ($m in $c) {
    if (-not $byT.ContainsKey($m.ticket_id)) { $byT[$m.ticket_id] = New-Object System.Collections.Generic.List[object] }
    $byT[$m.ticket_id].Add($m)
}
function Cut($s, $n) { $s = ($s -replace '\s+', ' ').Trim(); if ($s.Length -gt $n) { $s.Substring(0, $n) + '...' } else { $s } }

$blocks = New-Object System.Collections.Generic.List[string]
foreach ($r in ($t | Sort-Object { [long]$_.ticket_id })) {
    $msgs = @($byT[$r.ticket_id] | Sort-Object { [int]$_.seq })
    # first 6 + last 2 messages keeps the opening problem and the outcome
    $sel = if ($msgs.Count -le 8) { $msgs } else { @($msgs[0..5]) + @($msgs[-2..-1]) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("### T$($r.ticket_id) | group=$($r.group) | agent=$($r.agent) | cat=$($r.($cfg.fields.category)) | req=$($r.($cfg.fields.requestType)) | tier=$($r.($cfg.fields.tier)) | status=$($r.status) | msgs=$($r.message_count)")
    [void]$sb.AppendLine("SUBJECT: $(Cut $r.subject 200)")
    $budget = $MaxCharsPerTicket
    foreach ($m in $sel) {
        $tag = if ($m.is_private -eq 'True') { 'NOTE' } elseif ($m.author_role -eq 'Agent') { 'AGENT' } else { 'CUST' }
        $lim = if ($m.seq -eq '0') { 800 } elseif ($tag -eq 'NOTE') { 200 } else { 380 }
        $line = "[$tag] " + (Cut $m.body_text $lim)
        if ($budget - $line.Length -lt 0) { break }; $budget -= $line.Length
        [void]$sb.AppendLine($line)
    }
    if ($msgs.Count -gt 8) { [void]$sb.AppendLine("(... $($msgs.Count - 8) middle messages omitted)") }
    $blocks.Add($sb.ToString())
}

$size = [math]::Ceiling($blocks.Count / $Batches)
$utf8 = New-Object Text.UTF8Encoding $false
for ($i = 0; $i -lt $Batches; $i++) {
    $from = $i * $size; $to = [math]::Min($blocks.Count, ($i + 1) * $size) - 1
    if ($from -gt $to) { break }
    [IO.File]::WriteAllText("$RunDir\batches\batch_$('{0:D2}' -f ($i + 1)).txt", ($blocks[$from..$to] -join "`n"), $utf8)
}
Get-ChildItem "$RunDir\batches" | ForEach-Object {
    "{0}  {1:N0} KB  tickets={2}" -f $_.Name, ($_.Length / 1KB), ([regex]::Matches([IO.File]::ReadAllText($_.FullName), '(?m)^### T')).Count
}
