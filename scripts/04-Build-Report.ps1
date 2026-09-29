<#
.SYNOPSIS
  Builds two self-contained HTML reports from output\metrics.json:
    report.html               shareable: team-level metrics, no per-agent data embedded
    report_MANAGER_ONLY.html  adds sortable per-agent scorecards
  Optional inputs in RunDir: narrative.json (your written insights), product_recommendations.json
.EXAMPLE
  .\scripts\04-Build-Report.ps1 -RunDir .\run
#>
param(
    [string]$RunDir = '.\run',
    [string]$Taxonomy = (Join-Path $PSScriptRoot '..\taxonomy.json')
)
$ErrorActionPreference = 'Stop'
$RunDir = (Resolve-Path $RunDir).Path
$utf8 = New-Object Text.UTF8Encoding $false
$tpl  = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\report\report_template.html'), $utf8)
$m    = [IO.File]::ReadAllText("$RunDir\output\metrics.json", $utf8)
$tax  = (Get-Content $Taxonomy -Raw | ConvertFrom-Json).themes | ConvertTo-Json -Compress
$prod = if (Test-Path "$RunDir\product_recommendations.json") { [IO.File]::ReadAllText("$RunDir\product_recommendations.json", $utf8) } else { '[]' }
$narr = if (Test-Path "$RunDir\narrative.json") { [IO.File]::ReadAllText("$RunDir\narrative.json", $utf8) } else { '{}' }

# shareable edition: strip agent-level data from the embedded JSON, not just hide it
$mObj = $m | ConvertFrom-Json; $mObj.agents = @()
$mShare = ConvertTo-Json $mObj -Depth 10 -Compress

function Fill($metrics, $extra) {
    $tpl.Replace('__METRICS__', $metrics).Replace('__TAXONOMY__', $tax).Replace('__PRODUCT__', $prod).Replace('__NARRATIVE__', "Object.assign($narr, $extra)")
}
[IO.File]::WriteAllText("$RunDir\output\report.html", (Fill $mShare '{showAgentTables:false}'), $utf8)
[IO.File]::WriteAllText("$RunDir\output\report_MANAGER_ONLY.html", (Fill $m "{showAgentTables:true, titleSuffix:' (Manager edition, internal)'}"), $utf8)
Write-Host "Wrote $RunDir\output\report.html and report_MANAGER_ONLY.html"
