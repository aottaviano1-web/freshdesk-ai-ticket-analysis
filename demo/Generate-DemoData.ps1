<#
.SYNOPSIS
  Generates a fully synthetic dataset (fictional company "Contoso CloudWatch") in the same format
  as the Freshdesk export + AI tags, so the whole pipeline can be run without real data or an LLM.
.EXAMPLE
  .\demo\Generate-DemoData.ps1 -RunDir .\demo\run -Tickets 400
#>
param([string]$RunDir = (Join-Path $PSScriptRoot 'run'), [int]$Tickets = 400, [int]$Seed = 42)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $RunDir, "$RunDir\tags" | Out-Null
$RunDir = (Resolve-Path $RunDir).Path
$rnd = New-Object System.Random $Seed
function Rand([double]$a, [double]$b) { $a + $rnd.NextDouble() * ($b - $a) }
function Pick($arr) { $arr[$rnd.Next($arr.Count)] }
function Chance([double]$p) { $rnd.NextDouble() -lt $p }
function LogN([double]$median, [double]$spread) { $median * [math]::Exp((Rand -1 1) * $spread) }

# theme -> weight, subjects (sub_theme), product area, root cause, preventable, fix
$catalog = @{
  'Customer Care' = @(
    @{ t='ACCOUNT_ACCESS'; w=22; subs=@('Password reset link fails','Invitation link expired','Cannot log in to customer portal'); area='Customer Portal & Billing'; rc='Portal/self-service gap'; prev='Self-service/KB'; fix='Let users re-send their own invitation and show a specific error when a reset link fails.' },
    @{ t='LICENSE_ACTIVATION'; w=16; subs=@('Move license to new server','License certificate request','Add-on license activation'); area='Licensing & Activation'; rc='Licensing/commercial process'; prev='Self-service/KB'; fix='Self-service license transfer in the portal.' },
    @{ t='BILLING_INVOICE'; w=14; subs=@('Invoice copy request','Update billing address','Payment confirmation'); area='Customer Portal & Billing'; rc='Licensing/commercial process'; prev='Process/Automation'; fix='' },
    @{ t='RENEWAL_CANCEL'; w=10; subs=@('Cancel auto-renewal','Renewal quote request'); area='None'; rc='Licensing/commercial process'; prev='Not preventable'; fix='' },
    @{ t='SALES_QUOTE_TRIAL'; w=10; subs=@('Trial extension request','Pricing question','Upgrade quote'); area='None'; rc='Customer how-to question'; prev='Not preventable'; fix='' },
    @{ t='ACCOUNT_DATA'; w=8; subs=@('Company name change','Primary contact change'); area='Customer Portal & Billing'; rc='Internal process/routing'; prev='Process/Automation'; fix='' },
    @{ t='CALLBACK_CHASE'; w=6; subs=@('Callback request','Status update request'); area='None'; rc='Internal process/routing'; prev='Process/Automation'; fix='' },
    @{ t='VENDOR_COMPLIANCE'; w=4; subs=@('Supplier registration form','Security questionnaire'); area='None'; rc='Licensing/commercial process'; prev='Self-service/KB'; fix='' },
    @{ t='NOISE'; w=12; subs=@('Out-of-office reply','Unsolicited sales email','Bounce notification'); area='None'; rc='Noise'; prev='Not preventable'; fix='' }
  )
  'Technical Support' = @(
    @{ t='BUG_ERROR'; w=18; subs=@('Dashboard widget shows no data','Export fails with error 500','Agent crashes after update'); area='Core Product'; rc='Product defect/bug'; prev='Product'; fix='Fix the regression and add an automated upgrade test for existing configurations.' },
    @{ t='CONFIG_HOWTO'; w=18; subs=@('How to set thresholds','Configure custom check','Set up maintenance window'); area='Core Product'; rc='Customer how-to question'; prev='Self-service/KB'; fix='' },
    @{ t='INSTALL_UPGRADE'; w=12; subs=@('Upgrade path from old version','Service fails to start after upgrade'); area='Installation & Upgrade'; rc='Product usability/UX gap'; prev='Product'; fix='Add a pre-flight check and automatic rollback to the installer.' },
    @{ t='PERFORMANCE_STABILITY'; w=10; subs=@('Web UI slow','High memory usage on server'); area='Core Product'; rc='Customer environment/config'; prev='Product'; fix='Warn admins in-product before sizing limits are reached.' },
    @{ t='INTEGRATION_API'; w=9; subs=@('API token returns 401','Webhook not firing'); area='Integrations & API'; rc='Product defect/bug'; prev='Product'; fix='Show clear API error messages with a link to the fix.' },
    @{ t='USERS_SSO'; w=9; subs=@('SSO login loop','User permissions reset'); area='Users, SSO & Access'; rc='Product usability/UX gap'; prev='Product'; fix='Show specific SSO errors instead of a generic failure page.' },
    @{ t='NOTIFICATIONS'; w=7; subs=@('Email alerts not delivered','Duplicate notifications'); area='Notifications'; rc='Customer environment/config'; prev='Self-service/KB'; fix='' },
    @{ t='SECURITY'; w=6; subs=@('Certificate renewal','Vulnerability scan finding'); area='Security'; rc='Documentation/KB gap'; prev='Self-service/KB'; fix='' },
    @{ t='MIGRATION'; w=5; subs=@('Move server to new VM'); area='Installation & Upgrade'; rc='Customer how-to question'; prev='Self-service/KB'; fix='' },
    @{ t='NOISE'; w=6; subs=@('Internal test ticket','Auto-reply loop'); area='None'; rc='Noise'; prev='Not preventable'; fix='' }
  )
}
$agentsBy = @{ 'Customer Care' = @('Alex Sample', 'Jordan Sample', 'Casey Sample', 'Riley Sample'); 'Technical Support' = @('Morgan Sample', 'Taylor Sample', 'Jamie Sample', 'Drew Sample', 'Quinn Sample', 'Avery Sample') }
$agentSkill = @{}; foreach ($g in $agentsBy.Keys) { foreach ($a in $agentsBy[$g]) { $agentSkill[$a] = Rand 0.75 1.3 } }
$components = @('Web UI', 'Backend/Server', 'Agent/Connector', 'API', 'Email', 'Database')
$start = [datetime]'2026-08-31'; $days = 28
$inv = [Globalization.CultureInfo]::InvariantCulture
function F($d) { if ($d) { $d.ToString('yyyy-MM-dd HH:mm:ss', $inv) } }

$tRows = New-Object System.Collections.Generic.List[object]; $cRows = New-Object System.Collections.Generic.List[object]; $tagLines = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $Tickets; $i++) {
    $id = 900000 + $i
    $grp = if (Chance 0.5) { 'Customer Care' } else { 'Technical Support' }
    $cat = $catalog[$grp]; $tot = ($cat | ForEach-Object { $_.w } | Measure-Object -Sum).Sum; $x = $rnd.Next($tot); $th = $null
    foreach ($c in $cat) { if ($x -lt $c.w) { $th = $c; break }; $x -= $c.w }
    $sub = Pick $th.subs
    do { $day = $start.AddDays($rnd.Next($days)) } while ($day.DayOfWeek -in 'Saturday', 'Sunday' -and (Chance 0.8))
    $created = $day.AddHours((Rand 6 19)).AddMinutes($rnd.Next(60))
    $agent = Pick $agentsBy[$grp]; $sk = $agentSkill[$agent]
    $noise = $th.t -eq 'NOISE'
    $tech = $grp -eq 'Technical Support'
    $frt = $null; if (-not ($noise -and (Chance 0.6))) { $frt = LogN ($(if ($tech) { 5 } else { 3.5 }) * $sk) 1.9; if (Chance 0.07) { $frt += Rand 24 72 } }
    $custReplies = if ($noise) { 0 } elseif ($tech) { [int][math]::Floor([math]::Pow($rnd.NextDouble(), 1.6) * 6) } else { [int][math]::Floor([math]::Pow($rnd.NextDouble(), 2.5) * 4) }
    $autoClose = $tech -and -not $noise -and (Chance 0.28)
    $msgs = New-Object System.Collections.Generic.List[object]
    $msgs.Add(@{ role = 'Customer'; t = $created; priv = $false; type = 'Original request'; body = "Hello, we have an issue: $($sub.ToLower()). Can you help?" })
    $tcur = $created
    if ($frt -ne $null) {
        $tcur = $created.AddHours($frt); $msgs.Add(@{ role = 'Agent'; t = $tcur; priv = $false; type = 'Public reply'; body = "Thanks for reaching out about $($sub.ToLower()). Here are the next steps..." })
        for ($k = 0; $k -lt $custReplies; $k++) {
            $tcur = $tcur.AddHours((LogN $(if ($tech) { 30 } else { 8 }) 1.0)); $msgs.Add(@{ role = 'Customer'; t = $tcur; priv = $false; type = 'Inbound reply'; body = 'Thanks, we tried that. Here are the logs / details.' })
            $tcur = $tcur.AddHours((LogN (4 * $sk) 1.0)); $msgs.Add(@{ role = 'Agent'; t = $tcur; priv = $false; type = 'Public reply'; body = 'Thanks, based on the logs please try the following...' })
            if ($tech -and (Chance 0.3)) { $msgs.Add(@{ role = 'Agent'; t = $tcur.AddMinutes(5); priv = $true; type = 'Private note'; body = 'Checked with engineering, known issue.' }) }
        }
        if ($autoClose) {
            foreach ($d in 3, 6) { $msgs.Add(@{ role = 'Agent'; t = $tcur.AddDays($d); priv = $true; type = 'Private note'; body = 'Subject: Following up on your ticket - we have not heard back from you.' }) }
            $tcur = $tcur.AddDays(10); $msgs.Add(@{ role = 'Agent'; t = $tcur; priv = $true; type = 'Private note'; body = 'Subject: Your ticket has been closed - since we haven''t heard from you, we closed this ticket.' })
        }
    }
    $end = $start.AddDays($days + 2)
    $resolved = if ($frt -ne $null -and $tcur -lt $end -and (Chance 0.9)) { $tcur.AddMinutes($rnd.Next(5, 120)) } else { $null }
    $status = if ($resolved) { 'Closed' } elseif ($noise) { 'Closed' } else { Pick @('Open', 'Pending') }
    if ($noise -and -not $resolved) { $resolved = $created.AddHours((Rand 1 48)) }
    $pub = @($msgs | Where-Object { -not $_.priv })
    $seq = 0
    foreach ($m in ($msgs | Sort-Object { $_.t })) {
        $cRows.Add([pscustomobject][ordered]@{ ticket_id = $id; subject = $sub; ticket_status = $status; group = $grp; ticket_agent = $agent; seq = $seq; created_at = F $m.t
            author_role = $m.role; author_name = $(if ($m.role -eq 'Agent') { $agent } else { 'Demo Customer' }); author_email = ''; message_type = $m.type; is_private = $m.priv; body_chars = $m.body.Length; body_text = $m.body })
        $seq++
    }
    $tRows.Add([pscustomobject][ordered]@{
        ticket_id = $id; subject = $sub; status = $status; priority = 'Medium'; source = Pick @('Email', 'Email', 'Portal', 'Chat', 'Phone'); type = ''; group = $grp; agent = $agent
        requester_name = 'Demo Customer'; requester_email = "customer$i@example.com"; company = "Example Corp $($rnd.Next(1,60))"; tags = ''
        created_at = F $created; updated_at = F $tcur; due_by = ''
        first_responded_at = $(if ($frt -ne $null) { F $created.AddHours($frt) }); resolved_at = F $resolved; closed_at = F $resolved
        first_response_hours = $(if ($frt -ne $null) { [math]::Round($frt, 2) }); resolution_hours = $(if ($resolved) { [math]::Round(($resolved - $created).TotalHours, 2) })
        reopened = (Chance 0.08); message_count = $msgs.Count; customer_messages = @($pub | Where-Object { $_.role -eq 'Customer' }).Count
        agent_public_replies = @($pub | Where-Object { $_.role -eq 'Agent' }).Count; private_notes = @($msgs | Where-Object { $_.priv }).Count
        'Ticket Category' = $(if (Chance 0.6) { $th.t }); 'Request Type' = ''; 'Support Tier' = Pick @('Standard', 'Premium', '')
    })
    $hasReply = $frt -ne $null -and -not $noise
    $q = { param($b) if (-not $hasReply) { 0 } else { [int][math]::Max(1, [math]::Min(5, [math]::Round($b + (Rand -0.8 0.8) + (1.1 - $sk)))) } }
    $sent = if ((-not $hasReply) -or $autoClose -or (($custReplies -eq 0) -and (Chance 0.5))) { 'Unknown' } else { $p = $rnd.NextDouble(); if ($p -lt 0.45) { 'Positive' } elseif ($p -lt 0.88) { 'Neutral' } else { 'Negative' } }
    $tag = [ordered]@{ ticket_id = $id; theme = $th.t; sub_theme = $sub; product_area = $th.area
        component = $(if ($th.prev -eq 'Product') { Pick $components } else { '' }); root_cause = $th.rc; preventable_by = $th.prev; product_fix = $th.fix
        summary = "Customer needed help with $($sub.ToLower())."; resolved = $(if ($noise) { 'Unclear' } elseif ($resolved) { 'Yes' } else { 'Unclear' })
        first_reply_resolved = $(if (-not $hasReply) { 'N/A' } elseif ($custReplies -eq 0) { 'Yes' } else { 'No' })
        clarity = (& $q 3.9); empathy = (& $q 3.2); ownership = (& $q 3.6); sentiment_end = $sent
        escalated = $(if ($tech -and $th.rc -eq 'Product defect/bug' -and (Chance 0.4)) { 'Yes' } else { 'No' }); notable = '' }
    $tagLines.Add(($tag | ConvertTo-Json -Compress))
}
$tRows | Export-Csv "$RunDir\tickets_demo.csv" -NoTypeInformation -Encoding UTF8
$cRows | Export-Csv "$RunDir\conversations_demo.csv" -NoTypeInformation -Encoding UTF8
[IO.File]::WriteAllLines("$RunDir\tags\demo_tags.jsonl", $tagLines, (New-Object Text.UTF8Encoding $false))
Copy-Item (Join-Path $PSScriptRoot 'narrative.json'), (Join-Path $PSScriptRoot 'product_recommendations.json') $RunDir -Force
Write-Host "Generated $Tickets synthetic tickets, $($cRows.Count) messages in $RunDir"
