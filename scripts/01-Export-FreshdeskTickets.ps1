<#
.SYNOPSIS
  Export Freshdesk tickets + full conversation threads to JSON and CSV for analysis.

.DESCRIPTION
  Pulls tickets created in a date range, every conversation (replies, inbound
  emails, private notes) per ticket, and resolves IDs to names (agents, groups,
  companies, statuses, custom field labels). Also joins CSAT ratings if available.

  Outputs (in -OutDir):
    tickets_full_<ts>.json     nested: one object per ticket with a conversations[] array
    tickets_<ts>.csv           one row per ticket, key fields + metrics + full transcript
    conversations_<ts>.csv     one row per message (good for agent-communication analysis)

  The API key is read from $env:FRESHDESK_API_KEY, or prompted for (hidden input).

.EXAMPLE
  $env:FRESHDESK_API_KEY = "<your key>"
  .\Export-FreshdeskTickets.ps1 -Domain yourcompany -Since 2026-07-01 -MaxTickets 25 -Groups "Customer Care,Technical Support"

.EXAMPLE
  .\Export-FreshdeskTickets.ps1 -Domain yourcompany.freshdesk.com -Since 2026-06-01 -Until 2026-09-01
#>
param(
    # Your Freshdesk subdomain ("yourcompany") or full host ("yourcompany.freshdesk.com")
    [Parameter(Mandatory = $true)][string]$Domain,
    # Only tickets CREATED on/after this date (UTC). Default: last 30 days.
    [string]$Since = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd'),
    # Only tickets CREATED before this date (UTC). Optional.
    [string]$Until,
    # Stop after this many tickets (0 = no limit). Use a small number for a test run.
    [int]$MaxTickets = 0,
    # Only tickets assigned to these groups (by name), e.g. -Groups "Technical Support","Customer Care"
    [string[]]$Groups,
    [string]$OutDir = '.\run',
    # Drop internal private notes from the output.
    [switch]$ExcludePrivateNotes,
    # Keep quoted email history ("On ... wrote:") in message bodies. Off by default to avoid duplicated text.
    [switch]$KeepQuotedText
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------- auth / setup
$apiKey = $env:FRESHDESK_API_KEY
if (-not $apiKey) {
    $secure = Read-Host "Freshdesk API key" -AsSecureString
    $apiKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
}
$apiKey = ($apiKey -replace '\s', '').Trim('"', "'")
Write-Host "Using API key of length $($apiKey.Length) (Freshdesk keys are usually 20 characters)"
$auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
$headers = @{ Authorization = "Basic $auth" }

$Domain = ($Domain -replace '^https?://', '') -replace '/.*$', ''
if ($Domain -notmatch '\.') { $Domain = "$Domain.freshdesk.com" }
$base = "https://$Domain/api/v2"

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'

# ---------------------------------------------------------------- helpers
function Invoke-Fd([string]$path) {
    $url = "$base$path"
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        try {
            $r = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
            if ($null -eq $r) { return }
            return $r
        } catch [System.Net.WebException] {
            $resp = $_.Exception.Response
            $code = 0
            if ($resp) { $code = [int]$resp.StatusCode }
            if ($code -eq 429) {
                $wait = 60
                $ra = $resp.Headers['Retry-After']
                if ($ra) { $wait = [int]$ra + 1 }
                Write-Host "  Rate limited - waiting $wait s..." -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
            } elseif ($code -ge 500 -or $code -eq 0) {
                $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Host "  Transient error ($code) - retrying in $wait s..." -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
            } else {
                $body = ''
                try { $body = (New-Object IO.StreamReader($resp.GetResponseStream())).ReadToEnd() } catch {}
                if ($code -eq 401) {
                    Write-Host "`nFreshdesk rejected the API key (401). Check that you copied the full key from" -ForegroundColor Red
                    Write-Host "https://$Domain -> profile picture -> Profile settings -> 'View API key', with no extra characters." -ForegroundColor Red
                    exit 1
                }
                throw "Freshdesk API error $code on $url : $body"
            }
        }
    }
    throw "Giving up on $url after repeated failures"
}

function Get-FdAll([string]$path) {
    $sep = '?'
    if ($path -match '\?') { $sep = '&' }
    $all = New-Object System.Collections.Generic.List[object]
    $page = 1
    while ($true) {
        $items = @(Invoke-Fd "$path${sep}per_page=100&page=$page")
        foreach ($i in $items) { $all.Add($i) }
        if ($items.Count -lt 100) { break }
        $page++
    }
    return $all.ToArray()
}

function ToUtc($v) {
    if (-not $v) { return $null }
    return ([datetime]$v).ToUniversalTime()
}

function Fmt($v) {
    $d = ToUtc $v
    if ($d) { return $d.ToString('yyyy-MM-dd HH:mm:ss') }
    return $null
}

function Hours($from, $to) {
    if ($from -and $to) { return [math]::Round(((ToUtc $to) - (ToUtc $from)).TotalHours, 2) }
    return $null
}

# Cut quoted email history (English + German reply headers) and tidy whitespace.
$quoteRx = '(?m)^\s*(>?\s*On .{5,250}wrote:|>?\s*Am .{5,250}schrieb.{0,120}:|-{2,}\s*Original Message\s*-{2,}|-{2,}\s*Urspr.ngliche Nachricht\s*-{2,}|From:\s.+\r?\n\s*(Sent|Date):\s|Von:\s.+\r?\n\s*(Gesendet|Datum):\s)'
function Clean-Text([string]$s) {
    if (-not $s) { return '' }
    if (-not $KeepQuotedText) {
        $m = [regex]::Match($s, $quoteRx)
        if ($m.Success -and $m.Index -gt 0) { $s = $s.Substring(0, $m.Index) }
    }
    $s = $s -replace "`r", ''
    $s = $s -replace "[ \t]+`n", "`n"
    $s = $s -replace "`n{3,}", "`n`n"
    return $s.Trim()
}

# ---------------------------------------------------------------- lookups
Write-Host "Connecting to $base ..."
$ticketFields = @(Invoke-Fd '/ticket_fields')

$statusMap   = @{ '2' = 'Open'; '3' = 'Pending'; '4' = 'Resolved'; '5' = 'Closed' }
$priorityMap = @{ '1' = 'Low'; '2' = 'Medium'; '3' = 'High'; '4' = 'Urgent' }
$sourceMap   = @{ '1' = 'Email'; '2' = 'Portal'; '3' = 'Phone'; '5' = 'Twitter'; '6' = 'Facebook'; '7' = 'Chat';
                 '8' = 'Mobihelp'; '9' = 'Feedback Widget'; '10' = 'Outbound Email'; '11' = 'Ecommerce'; '12' = 'Bot'; '13' = 'WhatsApp' }

$statusField = $ticketFields | Where-Object { $_.name -eq 'status' } | Select-Object -First 1
if ($statusField -and $statusField.choices) {
    foreach ($p in $statusField.choices.PSObject.Properties) { $statusMap[[string]$p.Name] = @($p.Value)[0] }
}
$sourceField = $ticketFields | Where-Object { $_.name -eq 'source' } | Select-Object -First 1
if ($sourceField -and $sourceField.choices) {
    foreach ($p in $sourceField.choices.PSObject.Properties) {
        if ($p.Value -is [int] -or $p.Value -is [long]) { $sourceMap[[string]$p.Value] = $p.Name }
    }
}
$customFields = @($ticketFields | Where-Object { -not $_.default } | Sort-Object position)

$agents = @{}
try {
    foreach ($a in @(Get-FdAll '/agents')) { $agents[[string]$a.id] = $a.contact.name }
    Write-Host "  $($agents.Count) agents"
} catch { Write-Warning "Could not load agents (needs admin-level key?). Agent names will be blank. $_" }

$groupNames = @{}
try {
    foreach ($g in @(Get-FdAll '/groups')) { $groupNames[[string]$g.id] = $g.name }
    Write-Host "  $($groupNames.Count) groups"
} catch { Write-Warning "Could not load groups. $_" }

$groupFilter = $null
if ($Groups) {
    # "powershell -File" passes "A","B" as one string "A,B" - split it back apart
    $Groups = @($Groups | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $groupFilter = @{}
    foreach ($name in $Groups) {
        $match = @($groupNames.Keys | Where-Object { $groupNames[$_] -eq $name.Trim() })
        if ($match.Count -eq 0) {
            Write-Host "Group '$name' not found. Available groups:" -ForegroundColor Red
            $groupNames.Values | Sort-Object | ForEach-Object { Write-Host "  $_" }
            return
        }
        foreach ($id in $match) { $groupFilter[$id] = $true }
    }
    Write-Host "  Filtering to groups: $($Groups -join ', ')"
}

$companies = @{}
function Get-CompanyName($id) {
    if (-not $id) { return $null }
    $k = [string]$id
    if (-not $companies.ContainsKey($k)) {
        try { $companies[$k] = (Invoke-Fd "/companies/$k").name } catch { $companies[$k] = $null }
    }
    return $companies[$k]
}

$sinceDt = [datetime]::SpecifyKind([datetime]$Since, 'Utc')
$untilDt = [datetime]::MaxValue
if ($Until) { $untilDt = [datetime]::SpecifyKind([datetime]$Until, 'Utc') }
$sinceIso = $sinceDt.ToString('yyyy-MM-ddTHH:mm:ssZ')

$csatLabels = @{ '103' = 'Extremely happy'; '102' = 'Very happy'; '101' = 'Happy'; '100' = 'Neutral';
                 '-101' = 'Unhappy'; '-102' = 'Very unhappy'; '-103' = 'Extremely unhappy' }
$csat = @{}
try {
    foreach ($s in @(Get-FdAll "/surveys/satisfaction_ratings?created_since=$sinceIso")) {
        $score = $null
        if ($s.ratings) { $score = [string](@($s.ratings.PSObject.Properties)[0].Value) }
        $csat[[string]$s.ticket_id] = @{ score = $score; label = $csatLabels[$score]; feedback = $s.feedback }
    }
    Write-Host "  $($csat.Count) CSAT ratings"
} catch { Write-Warning "Could not load CSAT ratings (feature/permission may be unavailable). $_" }

# ---------------------------------------------------------------- tickets
Write-Host "Fetching tickets created since $($sinceDt.ToString('yyyy-MM-dd'))$(if ($Until) { " until $($untilDt.ToString('yyyy-MM-dd'))" }) ..."
$tickets = New-Object System.Collections.Generic.List[object]
$page = 1
:pages while ($true) {
    $batch = @(Invoke-Fd "/tickets?updated_since=$sinceIso&order_by=created_at&order_type=asc&include=requester,stats,description&per_page=100&page=$page")
    foreach ($t in $batch) {
        $c = ToUtc $t.created_at
        if ($c -lt $sinceDt) { continue }
        if ($c -ge $untilDt) { break pages }   # sorted ascending, nothing later qualifies
        if ($groupFilter -and -not $groupFilter.ContainsKey([string]$t.group_id)) { continue }
        $tickets.Add($t)
        if ($MaxTickets -gt 0 -and $tickets.Count -ge $MaxTickets) { break pages }
    }
    Write-Host "  page $page -> $($tickets.Count) tickets"
    if ($batch.Count -lt 100) { break }
    if ($page -ge 300) { Write-Warning "Hit 300 pages; Freshdesk may cap listing here. Split the date range with -Since/-Until."; break }
    $page++
}

$total = $tickets.Count
if ($total -eq 0) { Write-Host "No tickets found for that range."; return }

# ---------------------------------------------------------------- conversations
$records    = New-Object System.Collections.Generic.List[object]
$ticketRows = New-Object System.Collections.Generic.List[object]
$convRows   = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($t in $tickets) {
    $i++
    Write-Progress -Activity 'Fetching conversations' -Status "$i / $total  (ticket #$($t.id))" -PercentComplete (100 * $i / $total)

    $convs = @(Get-FdAll "/tickets/$($t.id)/conversations")
    if ($ExcludePrivateNotes) { $convs = @($convs | Where-Object { -not $_.private }) }
    $convs = @($convs | Sort-Object { ToUtc $_.created_at })

    $requesterName  = $t.requester.name
    $requesterEmail = $t.requester.email
    $agentName   = $agents[[string]$t.responder_id]
    $groupName   = $groupNames[[string]$t.group_id]
    $companyName = Get-CompanyName $t.company_id
    $status   = $statusMap[[string]$t.status]
    $priority = $priorityMap[[string]$t.priority]
    $source   = $sourceMap[[string]$t.source]
    $desc     = Clean-Text $t.description_text

    # Build message list: original request first, then conversations.
    $messages = New-Object System.Collections.Generic.List[object]
    $messages.Add([ordered]@{
        seq = 0; conversation_id = $null; created_at = Fmt $t.created_at
        author_role = 'Customer'; author_name = $requesterName; author_email = $requesterEmail
        message_type = 'Original request'; is_private = $false; incoming = $true; body_text = $desc
    })
    $seq = 0
    foreach ($cv in $convs) {
        $seq++
        $uid = [string]$cv.user_id
        if ($uid -and $uid -eq [string]$t.requester_id) { $role = 'Customer'; $name = $requesterName }
        elseif ($agents.ContainsKey($uid))              { $role = 'Agent';    $name = $agents[$uid] }
        elseif ($cv.incoming)                           { $role = 'Customer'; $name = $cv.from_email }
        else                                            { $role = 'Agent';    $name = $cv.from_email }

        if ($cv.private)      { $type = 'Private note' }
        elseif ($cv.incoming) { $type = 'Inbound reply' }
        else                  { $type = 'Public reply' }

        $messages.Add([ordered]@{
            seq = $seq; conversation_id = $cv.id; created_at = Fmt $cv.created_at
            author_role = $role; author_name = $name; author_email = $cv.from_email
            message_type = $type; is_private = [bool]$cv.private; incoming = [bool]$cv.incoming
            body_text = Clean-Text $cv.body_text
        })
    }

    # Readable transcript
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in $messages) {
        [void]$sb.AppendLine("[$($m.created_at) UTC] $($m.author_role.ToUpper()) - $($m.author_name) ($($m.message_type)):")
        [void]$sb.AppendLine($m.body_text)
        [void]$sb.AppendLine()
    }
    $transcript = $sb.ToString().Trim()

    $custom = [ordered]@{}
    foreach ($cf in $customFields) {
        $v = $null
        if ($t.custom_fields) { $v = $t.custom_fields.($cf.name) }
        if ($v -is [array]) { $v = $v -join '; ' }
        $custom[$cf.label] = $v
    }

    $customerMsgs = @($messages | Where-Object { $_.author_role -eq 'Customer' }).Count
    $agentPublic  = @($messages | Where-Object { $_.author_role -eq 'Agent' -and -not $_.is_private }).Count
    $privateNotes = @($messages | Where-Object { $_.is_private }).Count
    $rating = $csat[[string]$t.id]

    $row = [ordered]@{
        ticket_id            = $t.id
        subject              = $t.subject
        status               = $status
        priority             = $priority
        source               = $source
        type                 = $t.type
        group                = $groupName
        agent                = $agentName
        requester_name       = $requesterName
        requester_email      = $requesterEmail
        company              = $companyName
        tags                 = (@($t.tags) -join '; ')
        created_at           = Fmt $t.created_at
        updated_at           = Fmt $t.updated_at
        due_by               = Fmt $t.due_by
        first_responded_at   = Fmt $t.stats.first_responded_at
        resolved_at          = Fmt $t.stats.resolved_at
        closed_at            = Fmt $t.stats.closed_at
        first_response_hours = Hours $t.created_at $t.stats.first_responded_at
        resolution_hours     = Hours $t.created_at $t.stats.resolved_at
        reopened             = [bool]$t.stats.reopened_at
        message_count        = $messages.Count
        customer_messages    = $customerMsgs
        agent_public_replies = $agentPublic
        private_notes        = $privateNotes
        csat_score           = $(if ($rating) { $rating.score })
        csat_label           = $(if ($rating) { $rating.label })
        csat_feedback        = $(if ($rating) { $rating.feedback })
    }
    foreach ($k in $custom.Keys) {
        $col = $k
        if ($row.Contains($col)) { $col = "custom_$k" }
        $row[$col] = $custom[$k]
    }

    $record = [ordered]@{}
    foreach ($k in $row.Keys) { $record[$k] = $row[$k] }
    $record['description']   = $desc
    $record['conversations'] = $messages.ToArray()
    $records.Add($record)

    $row['description'] = $desc
    $row['transcript']  = $transcript
    $ticketRows.Add([pscustomobject]$row)

    foreach ($m in $messages) {
        $convRows.Add([pscustomobject]([ordered]@{
            ticket_id = $t.id; subject = $t.subject; ticket_status = $status; group = $groupName; ticket_agent = $agentName
            seq = $m.seq; created_at = $m.created_at; author_role = $m.author_role; author_name = $m.author_name
            author_email = $m.author_email; message_type = $m.message_type; is_private = $m.is_private
            body_chars = $m.body_text.Length; body_text = $m.body_text
        }))
    }
}
Write-Progress -Activity 'Fetching conversations' -Completed

# ---------------------------------------------------------------- write output
$jsonPath = Join-Path $OutDir "tickets_full_$ts.json"
$tixPath  = Join-Path $OutDir "tickets_$ts.csv"
$convPath = Join-Path $OutDir "conversations_$ts.csv"

$json = ConvertTo-Json -InputObject $records.ToArray() -Depth 8
[IO.File]::WriteAllText($jsonPath, $json, (New-Object Text.UTF8Encoding $false))
$ticketRows | Export-Csv -Path $tixPath  -NoTypeInformation -Encoding UTF8
$convRows   | Export-Csv -Path $convPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Done. $total tickets, $($convRows.Count) messages." -ForegroundColor Green
Write-Host "  $jsonPath"
Write-Host "  $tixPath"
Write-Host "  $convPath"
