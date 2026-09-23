<#
.SYNOPSIS
    Find and open a Claude desktop app conversation by name using Windows UI Automation.

.DESCRIPTION
    Searches the running Claude desktop app for conversations matching the given name.
    Checks both the sidebar (recent/pinned chats) and all Projects. When multiple
    matches share the same timestamp, opens the topmost one (most recently active).

.PARAMETER Name
    Partial or full conversation title to search for (case-insensitive).

.PARAMETER Project
    Optional project name to scope the search (e.g. "Work - WordPress").
    If provided, navigates into that project and lists its chats with timestamps.

.PARAMETER List
    If specified, lists all available conversation titles without opening any.

.PARAMETER ListProjects
    If specified, lists all visible projects.

.PARAMETER Code
    Search local Claude Code session transcripts (~/.claude/projects/**/*.jsonl)
    instead of the running desktop app. Multi-word -Name is treated as AND: a chat
    matches when every word appears in the actual conversation (your prompts +
    Claude's replies), not in injected CLAUDE.md/memory context. Results are ranked
    full-match-first then most-recently-active, each showing project, date range,
    snippet, session id, and a ready-to-run `claude --resume` command.

.PARAMETER Cloud
    Full-text search of EVERY claude.ai chat and Claude Code session (local and
    cloud, including sessions started from the phone app) by driving the desktop
    app's own Search palette (server-side index). Needs the app running. Use
    -Type to restrict to Sessions / Code / Projects. -Open opens the top hit.
    Cowork sessions are NOT in this index (verified 2026-09-23) - see -Cowork.

.PARAMETER Cowork
    Cowork sessions (https://claude.ai/cowork/cse_*) have no local transcript and
    the desktop app's palette does not index them, so the only search surface is
    the web UI. -Cowork runs the -Cloud search first (in case a future app build
    indexes them), then opens https://claude.ai/cowork in your default browser
    where you can search/scroll the session list yourself.

.PARAMETER Type
    With -Cloud: palette filter tab - All (default), Sessions (claude.ai chats),
    Code (Claude Code sessions), Projects, Artifacts, Scheduled.

.PARAMETER Open
    With -Cloud -Name: open the top matching result in the app.

.PARAMETER Resume
    With -Code -Name, immediately resume the top matching Claude Code session
    (runs `claude --resume <id>` in that session's working directory).

.PARAMETER IncludeAgents
    Also include programmatic agent/SDK sessions (entrypoint 'sdk-cli', e.g. the
    Hermes agent's own automated runs). By default -Code searches only interactive
    chats you typed in, which hides the high-volume agent transcripts.

.PARAMETER Limit
    Max Claude Code results to show / sessions to list (default 20).

.EXAMPLE
    .\Find-ClaudeChat.ps1 -Name "mental health"
    .\Find-ClaudeChat.ps1 -Name "wordpress" -Project "Work - WordPress"
    .\Find-ClaudeChat.ps1 -List
    .\Find-ClaudeChat.ps1 -ListProjects
    .\Find-ClaudeChat.ps1 -Code -Name "hermes desktop"
    .\Find-ClaudeChat.ps1 -Code -Name "hermes dashboard" -Project hermes-deployment
    .\Find-ClaudeChat.ps1 -Code -Name "hermes desktop" -Resume
    .\Find-ClaudeChat.ps1 -Code -List
    .\Find-ClaudeChat.ps1 -Cloud -Name "chatgpt subscription"           # all claude.ai chats + Code sessions
    .\Find-ClaudeChat.ps1 -Cloud -Name "kling credits" -Type Sessions -Open
    .\Find-ClaudeChat.ps1 -Cowork -Name "chatgpt subscription"

.NOTES
    Desktop-app modes require the Claude desktop app running and use Windows UI
    Automation (UIAutomationClient/.NET). -Code mode reads local transcript files
    and does NOT need the app running.
    Conversation order = most recently active first.
#>

param(
    [string]$Name,
    [string]$Project,
    [switch]$List,
    [switch]$ListProjects,
    [switch]$Code,
    [switch]$Cloud,
    [switch]$Cowork,
    [switch]$Scheduled,
    [string]$Task,
    [ValidateSet('All','Sessions','Code','Projects','Artifacts','Scheduled')]
    [string]$Type = 'All',
    [switch]$Open,
    [switch]$Resume,
    [switch]$IncludeAgents,
    [int]$Limit = 20
)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinFocusClaude {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

function Get-ClaudeWindow {
    # The MAIN window is named "Claude". Popped-out session windows carry the
    # chat title instead, so match by process AND name, not name alone.
    $claudePids = @(Get-Process claude -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if (-not $claudePids.Count) { return $null }
    $tops = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition
    ) | Where-Object { $_.Current.ProcessId -in $claudePids }
    $main = $tops | Where-Object { $_.Current.Name -eq 'Claude' } | Select-Object -First 1
    if ($main) { return $main }
    return $tops | Select-Object -First 1
}

function Get-ByAutomationId($window, [string]$id) {
    return $window.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $id
        )
    )
}

# ── Cloud search via the app's command palette (Ctrl+K "Search") ───────────────
# The palette is a server-side full-text search over every claude.ai chat AND
# every Claude Code session (local + cloud). Result ids: "local_<guid>" = Code
# session, plain guid = claude.ai chat. NOTE: Cowork sessions (claude.ai/cowork/
# cse_*) were NOT indexed by the palette as of 2026-09-23 — see -Cowork.

function Open-CommandPalette($window) {
    $input = Get-ByAutomationId $window 'command-palette-input'
    if ($input) { return $input }
    $searchBtn = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    ) | Where-Object { $_.Current.Name -eq 'Search' } | Select-Object -First 1
    if (-not $searchBtn) { return $null }
    Invoke-UiaElement $searchBtn
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 150
        $input = Get-ByAutomationId $window 'command-palette-input'
        if ($input) { return $input }
    }
    return $null
}

function Close-CommandPalette($window) {
    $close = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    ) | Where-Object { $_.Current.Name -eq 'Close' -and $_.Current.AutomationId -like '_r_*' } | Select-Object -First 1
    if ($close) { try { Invoke-UiaElement $close } catch { } }
}

function Search-CloudSessions($window, [string]$query, [string]$typeTab) {
    $input = Open-CommandPalette $window
    if (-not $input) { throw "Could not open the app's Search palette (no 'Search' button / palette input found)." }
    $input.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($query)
    Start-Sleep -Milliseconds 3000
    if ($typeTab) {
        $tab = $window.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.PropertyCondition]::new(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::TabItem
            )
        ) | Where-Object { $_.Current.Name -eq $typeTab } | Select-Object -First 1
        if ($tab) {
            $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
            Start-Sleep -Milliseconds 2500
        } else { Write-Warning "Palette has no '$typeTab' filter tab; showing all types." }
    }
    $list = Get-ByAutomationId $window 'command-palette-results'
    if ($env:FINDCLAUDECHAT_DEBUG) { Write-Host "   [debug] window='$($window.Current.Name)' input=$($null -ne $input) value='$($input.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value)' list=$($null -ne $list)" -ForegroundColor DarkGray }
    if (-not $list) { return @() }
    $items = $list.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::ListItem
        )
    )
    $out = @()
    foreach ($it in $items) {
        $aid = [string]$it.Current.AutomationId
        if ($aid -notlike 'command-palette-item-*') { continue }
        $id = $aid.Substring('command-palette-item-'.Length)
        # skip actions / filters (new_code_session, send_message, delete_chat, filter-value:..)
        $kind = $null
        if ($id -match '^local_[0-9a-f]{8}-') { $kind = 'code' }
        elseif ($id -match '^cse_') { $kind = 'cowork' }          # Cowork / scheduled-task session
        elseif ($id -match '^trig_') { $kind = 'task' }           # a scheduled task (its runs are cse_ sessions)
        elseif ($id -match '^[0-9a-f]{8}-[0-9a-f]{4}-') { $kind = 'chat' }
        else { if ($env:FINDCLAUDECHAT_DEBUG) { Write-Host "   (skip $id)" -ForegroundColor DarkGray }; continue }
        $out += [PSCustomObject]@{ Kind = $kind; Id = $id; Label = [string]$it.Current.Name; Element = $it }
    }
    return $out
}

# ── Scheduled-task runs (the Cowork sessions the palette can't see) ───────────
# Each scheduled task (e.g. "Morning brief") has a page listing its runs as
# hyperlinks ("Today at 9:12 AM", "Sep 21 at 9:06 AM Awaiting input"). Each run
# IS a Cowork session (https://claude.ai/cowork/cse_*). Run bodies are not
# full-text indexed anywhere on the desktop, so this lists/opens runs by date.

function Get-ScheduledRunLinks($window) {
    $links = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Hyperlink
        )
    )
    $out = @()
    foreach ($l in $links) {
        $c = $l.Current; $r = $c.BoundingRectangle
        if ([double]::IsInfinity($r.X)) { continue }
        # run rows: "Today at 9:12 AM", "Yesterday at 9:05 AM", "Sep 21 at 9:06 AM Awaiting input"
        if ($c.Name -notmatch '^(Today|Yesterday|[A-Z][a-z]{2} \d{1,2}(, \d{4})?) at \d{1,2}:\d{2} [AP]M') { continue }
        $href = $null
        try { $href = [string]$l.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { }
        $out += [PSCustomObject]@{ Label = $c.Name; Href = $href; Y = $r.Y; Element = $l }
    }
    return $out | Sort-Object Y
}

function Invoke-UiaElement($el) {
    $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
}

function Focus-ClaudeWindow {
    $proc = Get-Process claude | Where-Object { $_.MainWindowTitle -eq "Claude" } | Select-Object -First 1
    if ($proc) {
        [WinFocusClaude]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null
        [WinFocusClaude]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 400
    }
}

function Get-SidebarChatButtons($window) {
    $buttons = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    )
    return $buttons | Where-Object {
        $n = $_.Current.Name
        $n -and $n.Length -gt 5 -and
        $n -notlike "More options for *" -and
        $n -notlike "Copy*" -and $n -notlike "Retry*" -and $n -notlike "Edit*" -and
        $n -notlike "Give*" -and $n -notlike "Show*" -and $n -notlike "Download*" -and
        $n -notlike "Start a task*" -and $n -notlike "New*" -and $n -notlike "Project*" -and
        $n -notin @("Collapse","Expand","Send","Close","Settings","Menu","Search","Back","Forward",
                    "Minimize","Maximize","Artifacts","Customize","Pinned","Recents","Projects",
                    "Get apps and extensions","Press and hold to record","Collapse sidebar",
                    "Open sidebar","Share chat","View all","Add files, connectors, and more",
                    "Relaunch to update","Chat","Cowork","Code","Sort by","Search projects","New project")
    } | Sort-Object { $_.Current.BoundingRectangle.Y }
}

function Get-ProjectLinks($window) {
    $links = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Hyperlink
        )
    )
    return $links | Where-Object {
        $n = $_.Current.Name
        $n -and $n.Length -gt 3 -and $n -ne "Skip to content" -and $n -ne "All projects"
    } | Sort-Object { $_.Current.BoundingRectangle.Y }
}

function Get-ProjectChatLinks($window) {
    # Project page chat list — rendered as Hyperlinks in main content area (X > sidebar width)
    $allEls = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition
    )
    return $allEls | Where-Object {
        $r = $_.Current.BoundingRectangle
        $n = $_.Current.Name
        $_.Current.ControlType.Id -eq 50005 -and  # Hyperlink
        $r.X -gt 3250 -and $r.Y -gt 1850 -and $r.Y -lt 2700 -and
        $n -and $n.Length -gt 10
    } | Sort-Object { $_.Current.BoundingRectangle.Y }
}

# ── Claude Code transcript search ───────────────────────────────────────────────

function Format-ChatTimestamp($ts) {
    if (-not $ts) { return '????-??-?? ??:??' }
    if ($ts -is [datetime]) { return $ts.ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    return ([string]$ts -replace 'T', ' ').Substring(0, [Math]::Min(16, ([string]$ts).Length))
}

function Get-SessionHead([string]$path) {
    $cwd = $null; $entry = $null
    try { $r = [System.IO.File]::OpenText($path) } catch { return [PSCustomObject]@{ Cwd = $null; Entry = $null } }
    try {
        for ($n = 0; $n -lt 8 -and $null -ne ($l = $r.ReadLine()); $n++) {
            if ($l.Length -lt 2) { continue }
            $o = $null; try { $o = $l | ConvertFrom-Json } catch { continue }
            if (-not $cwd -and $o.cwd) { $cwd = [string]$o.cwd }
            if (-not $entry -and $o.entrypoint) { $entry = [string]$o.entrypoint }
            if ($cwd -and $entry) { break }
        }
    } finally { $r.Close() }
    return [PSCustomObject]@{ Cwd = $cwd; Entry = $entry }
}

function Get-ChatSnippet([string]$text, [string]$term, [string]$who) {
    if (-not $text) { return $null }
    $flat = ($text -replace '\s+', ' ').Trim()
    if ($term) {
        $i = $flat.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase)
        if ($i -ge 0) {
            $start = [Math]::Max(0, $i - 50)
            $len = [Math]::Min(170, $flat.Length - $start)
            $snip = $flat.Substring($start, $len)
            if ($start -gt 0) { $snip = "...$snip" }
            if (($start + $len) -lt $flat.Length) { $snip = "$snip..." }
            return "[$who] $snip"
        }
    }
    if ($flat.Length -gt 170) { $flat = $flat.Substring(0, 170) + '...' }
    return "[$who] $flat"
}

function Get-CodeSessionInfo([string]$path, [string[]]$terms) {
    $cwd = $null; $firstTs = $null; $lastTs = $null; $firstPrompt = $null; $title = $null; $entry = $null
    $found = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $msgHits = New-Object System.Collections.Generic.List[object]
    try { $reader = [System.IO.File]::OpenText($path) } catch { return $null }
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line.Length -lt 2) { continue }
            $o = $null; try { $o = $line | ConvertFrom-Json } catch { continue }
            if (-not $cwd -and $o.cwd) { $cwd = [string]$o.cwd }
            if (-not $entry -and $o.entrypoint) { $entry = [string]$o.entrypoint }
            if ($o.timestamp) {
                $dt = $null
                if ($o.timestamp -is [datetime]) { $dt = $o.timestamp.ToUniversalTime() }
                else { try { $dt = [datetime]::Parse([string]$o.timestamp, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal) } catch { } }
                if ($dt) { if (-not $firstTs) { $firstTs = $dt }; $lastTs = $dt }
            }
            $type = [string]$o.type
            if ($type -eq 'summary' -and $o.summary) { if (-not $title) { $title = [string]$o.summary }; continue }
            $texts = @(); $who = $null
            if ($type -eq 'user') {
                $who = 'you'; $c = $o.message.content
                if ($c -is [string]) { $texts = @($c) }
                elseif ($c) { foreach ($b in $c) { if ($b.type -eq 'text' -and $b.text) { $texts += [string]$b.text } } }
            } elseif ($type -eq 'assistant') {
                $who = 'claude'; $c = $o.message.content
                if ($c) { foreach ($b in $c) { if ($b.type -eq 'text' -and $b.text) { $texts += [string]$b.text } } }
            } else { continue }
            foreach ($t in $texts) {
                if (-not $t) { continue }
                if ($who -eq 'you') {
                    if ($t.StartsWith('This session is being continued') -or $t.StartsWith('<') -or $t.StartsWith('Caveat')) { continue }
                    if ($t.Length -gt 4000 -or $t -match '<system-reminder>' -or $t -match 'These instructions OVERRIDE' -or $t -match 'Codebase and user instructions') { continue }
                    if (-not $firstPrompt) { $firstPrompt = $t }
                }
                if ($terms -and $terms.Count) {
                    $hits = 0
                    foreach ($term in $terms) {
                        if ($t.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { [void]$found.Add($term); $hits++ }
                    }
                    if ($hits -gt 0 -and $msgHits.Count -lt 200) { $msgHits.Add([PSCustomObject]@{ Who = $who; Text = $t; Hits = $hits }) }
                }
            }
        }
    } finally { $reader.Close() }
    $snips = @()
    foreach ($m in ($msgHits | Sort-Object Hits -Descending | Select-Object -First 3)) {
        $firstTerm = $null
        foreach ($term in $terms) { if ($m.Text.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $firstTerm = $term; break } }
        $snips += (Get-ChatSnippet $m.Text $firstTerm $m.Who)
    }
    return [PSCustomObject]@{
        Id = [System.IO.Path]::GetFileNameWithoutExtension($path); Path = $path; Cwd = $cwd
        Project = $(if ($cwd) { Split-Path $cwd -Leaf } else { '(unknown)' })
        Title = $title; Entry = $entry; FirstTs = $firstTs; LastTs = $lastTs; FirstPrompt = $firstPrompt
        TermsFound = $found.Count; TermsTotal = @($terms).Count
        AllPresent = ($terms -and $terms.Count -and $found.Count -eq @($terms).Count)
        MsgMatches = $msgHits.Count; Snippets = $snips
    }
}

function Show-CodeSession($info, [int]$idx) {
    $range = (Format-ChatTimestamp $info.FirstTs) + ' -> ' + (Format-ChatTimestamp $info.LastTs)
    $tag = ''
    if ($info.TermsTotal -gt 0) {
        if ($info.AllPresent) { $tag = "   (all terms, $($info.MsgMatches) msgs)" }
        else { $tag = "   ($($info.TermsFound)/$($info.TermsTotal) terms)" }
    }
    $head = $(if ($info.Title) { "$($info.Project)  -  $($info.Title)" } else { $info.Project })
    Write-Host ("[{0}] {1}   {2}{3}" -f $idx, $head, $range, $tag) -ForegroundColor Green
    if ($info.Cwd) { Write-Host "     $($info.Cwd)" -ForegroundColor DarkGray }
    if ($info.Snippets -and $info.Snippets.Count) {
        foreach ($s in $info.Snippets) { Write-Host "     $s" -ForegroundColor Gray }
    } elseif ($info.FirstPrompt) {
        Write-Host "     $(Get-ChatSnippet $info.FirstPrompt $null 'you')" -ForegroundColor Gray
    }
    $dir = $(if ($info.Cwd) { $info.Cwd } else { '<project-dir>' })
    Write-Host "     resume:  cd `"$dir`"; claude --resume $($info.Id)" -ForegroundColor Cyan
}

# ── Main ──────────────────────────────────────────────────────────────────────

if ($Code) {
    $base = Join-Path $env:USERPROFILE '.claude\projects'
    if (-not (Test-Path $base)) { Write-Error "No Claude Code projects directory at $base"; exit 1 }
    $sessionFiles = @(Get-ChildItem -Path (Join-Path $base '*\*.jsonl') -File -ErrorAction SilentlyContinue)
    if (-not $sessionFiles.Count) { Write-Error "No Claude Code session transcripts found under $base"; exit 1 }

    if ($List) {
        Write-Host "Recent Claude Code chats (newest first):" -ForegroundColor Cyan
        $shown = 0
        foreach ($f in ($sessionFiles | Sort-Object LastWriteTime -Descending)) {
            if (-not $IncludeAgents -and (Get-SessionHead $f.FullName).Entry -eq 'sdk-cli') { continue }
            $info = Get-CodeSessionInfo $f.FullName @()
            if (-not $info) { continue }
            if ($Project -and ($info.Cwd -notlike "*$Project*")) { continue }
            $shown++
            $label = $(if ($info.Title) { $info.Title } elseif ($info.FirstPrompt) { ($info.FirstPrompt -replace '\s+', ' ').Trim() } else { '(no prompt)' })
            if ($label.Length -gt 64) { $label = $label.Substring(0, 64) + '...' }
            Write-Host ("  {0}  {1,-22}  {2}  {3}" -f (Format-ChatTimestamp $info.LastTs), $info.Project, $info.Id.Substring(0, 8), $label)
            if ($shown -ge $Limit) { break }
        }
        if (-not $shown) { Write-Warning "No sessions matched." }
        exit 0
    }

    if (-not $Name) { Write-Error "Provide -Name to search, or -List."; exit 1 }

    $terms = @($Name -split '\s+' | Where-Object { $_ })
    Write-Host "Searching Claude Code transcripts for: $($terms -join ' + ')" -ForegroundColor Cyan

    # AND prefilter: keep only files that contain every term (necessary condition)
    $hitPaths = @($sessionFiles | Select-Object -ExpandProperty FullName)
    foreach ($term in $terms) {
        if (-not $hitPaths.Count) { break }
        $hitPaths = @(Select-String -Path $hitPaths -Pattern $term -SimpleMatch -List -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path)
    }
    if (-not $hitPaths.Count) { Write-Warning "No Claude Code chats contain all of: $($terms -join ', ')"; exit 1 }

    # Default: only interactive chats you typed in. Programmatic agent/SDK runs
    # (entrypoint 'sdk-cli', e.g. the Hermes agent's own sessions) carry the same
    # terms in their injected CLAUDE.md and would swamp results. -IncludeAgents keeps them.
    if (-not $IncludeAgents) {
        $interactive = @($hitPaths | Where-Object { (Get-SessionHead $_).Entry -ne 'sdk-cli' })
        if ($interactive.Count) { $hitPaths = $interactive }
        else { Write-Host "  (only agent/SDK sessions matched; showing them -- pass -IncludeAgents to keep this)" -ForegroundColor DarkGray }
    }

    # Scope by project BEFORE the recency cap so older in-project chats aren't cut.
    if ($Project) {
        $hitPaths = @($hitPaths | Where-Object { (Get-SessionHead $_).Cwd -like "*$Project*" })
        if (-not $hitPaths.Count) { Write-Warning "No matching chats in a project like '*$Project*'."; exit 1 }
    }

    $origCount = $hitPaths.Count
    $cap = [Math]::Max($Limit * 4, 60)
    if ($origCount -gt $cap) {
        $hitPaths = @($hitPaths | Sort-Object { (Get-Item -LiteralPath $_).LastWriteTime } -Descending | Select-Object -First $cap)
        Write-Host "  (scanning $cap most-recent of $origCount candidates; use -Project to refine)" -ForegroundColor DarkGray
    }

    $infos = @()
    foreach ($p in $hitPaths) {
        $info = Get-CodeSessionInfo $p $terms
        if (-not $info) { continue }
        if ($Project -and ($info.Cwd -notlike "*$Project*")) { continue }
        if ($info.TermsFound -gt 0) { $infos += $info }
    }
    if (-not $infos.Count) { Write-Warning "No conversational matches for '$Name'."; exit 1 }

    $full = @($infos | Where-Object { $_.AllPresent })
    $show = $(if ($full.Count) { $full } else { $infos })
    if (-not $full.Count) { Write-Host "(no chat had all terms in-conversation; showing closest partial matches)" -ForegroundColor DarkYellow }
    $show = @($show | Sort-Object LastTs -Descending | Select-Object -First $Limit)

    Write-Host ""
    $i = 0
    foreach ($info in $show) { $i++; Show-CodeSession $info $i; Write-Host "" }

    if ($Resume) {
        $best = $show[0]
        if ($best.Cwd -and (Test-Path $best.Cwd)) {
            Write-Host "Resuming top match $($best.Id) in $($best.Cwd)..." -ForegroundColor Green
            Push-Location $best.Cwd
            try { claude --resume $best.Id } finally { Pop-Location }
        } else {
            Write-Warning "Top match's working directory '$($best.Cwd)' isn't accessible here; can't auto-resume."
        }
    }
    exit 0
}

$win = Get-ClaudeWindow
if (-not $win) { Write-Error "Claude desktop app not running."; exit 1 }

# ── -Scheduled: list / open the runs of a scheduled task (Cowork sessions) ────
if ($Scheduled) {
    Write-Host "Scheduled tasks (via palette):" -ForegroundColor Cyan
    $tasks = @()
    try { $tasks = @(Search-CloudSessions $win '' 'Scheduled' | Where-Object { $_.Kind -eq 'task' }) }
    catch { Write-Error $_; exit 1 }
    if (-not $tasks.Count) { Close-CommandPalette $win; Write-Warning "No scheduled tasks listed in the palette."; exit 1 }
    $i = 0
    foreach ($t in $tasks) { $i++; Write-Host ("  [{0}] {1}   ({2})" -f $i, $t.Label, $t.Id) }
    $pick = $null
    if ($Task) { $pick = $tasks | Where-Object { $_.Label -like "*$Task*" } | Select-Object -First 1 }
    elseif ($tasks.Count -eq 1) { $pick = $tasks[0] }
    if (-not $pick) {
        Close-CommandPalette $win
        if ($Task) { Write-Warning "No scheduled task matching '*$Task*'." } else { Write-Host "Pass -Task <name> to list a task's runs." -ForegroundColor Yellow }
        exit $(if ($Task) { 1 } else { 0 })
    }
    Write-Host "Opening task page: $($pick.Label)" -ForegroundColor Cyan
    Focus-ClaudeWindow
    Invoke-UiaElement $pick.Element
    $runs = @()
    for ($n = 0; $n -lt 25 -and -not $runs.Count; $n++) { Start-Sleep -Milliseconds 200; $runs = @(Get-ScheduledRunLinks $win) }
    if (-not $runs.Count) { Write-Warning "Task page opened but no run rows were found."; exit 1 }
    Write-Host "Runs (newest first) - each is a Cowork session:" -ForegroundColor Cyan
    $i = 0
    foreach ($r in ($runs | Select-Object -First $Limit)) {
        $i++
        $hit = $(if ($Name -and $r.Label -like "*$Name*") { '  <== match' } else { '' })
        Write-Host ("  [{0}] {1}{2}" -f $i, $r.Label, $hit) -ForegroundColor $(if ($hit) { 'Green' } else { 'White' })
        if ($r.Href) { Write-Host "       $($r.Href)" -ForegroundColor DarkGray }
    }
    $target = $(if ($Name) { $runs | Where-Object { $_.Label -like "*$Name*" } | Select-Object -First 1 } else { $runs[0] })
    if ($Open -and $target) {
        Write-Host "Opening run: $($target.Label)" -ForegroundColor Green
        Invoke-UiaElement $target.Element
    } elseif ($Name -and -not $target) {
        Write-Warning "No run matching '*$Name*' (match on the date label, e.g. 'Today', 'Sep 21')."
        exit 1
    }
    exit 0
}

# ── -Cloud / -Cowork: server-side search through the app's palette ────────────
if ($Cloud -or $Cowork) {
    if (-not $Name) { Write-Error "Provide -Name to search."; exit 1 }
    $tab = $(if ($Type -eq 'All') { $null } else { $Type })
    Write-Host "Searching claude.ai + Claude Code sessions (palette, type=$Type) for: $Name" -ForegroundColor Cyan
    $hits = @()
    try { $hits = @(Search-CloudSessions $win $Name $tab) }
    catch { Write-Error $_; exit 1 }
    if (-not $hits.Count) {
        Write-Warning "No palette results for '$Name'."
    } else {
        $i = 0
        foreach ($h in ($hits | Select-Object -First $Limit)) {
            $i++
            $url = switch ($h.Kind) {
                'code'   { "claude --resume $($h.Id.Substring(6))" }
                'cowork' { "https://claude.ai/cowork/$($h.Id)" }
                'task'   { "scheduled task - runs: .\Find-ClaudeChat.ps1 -Scheduled -Task `"$(($h.Label -split ' (Weekdays|Daily|Every|Weekly|Monthly|Scheduled)')[0])`"" }
                default  { "https://claude.ai/chat/$($h.Id)" }
            }
            Write-Host ("[{0}] {1,-4} {2}" -f $i, $h.Kind, $h.Label) -ForegroundColor Green
            Write-Host "     $url" -ForegroundColor DarkGray
        }
    }
    if ($Open -and $hits.Count) {
        Write-Host "Opening: $($hits[0].Label)" -ForegroundColor Green
        Focus-ClaudeWindow
        Invoke-UiaElement $hits[0].Element
    } else {
        Close-CommandPalette $win
    }
    if ($Cowork) {
        Write-Host ""
        Write-Host "Cowork session bodies are not in the desktop app's index. If the chat was a scheduled-task run (e.g. Morning brief), list those with:" -ForegroundColor Yellow
        Write-Host "    .\Find-ClaudeChat.ps1 -Scheduled -Task 'Morning brief' [-Name 'Today' -Open]" -ForegroundColor Yellow
        Write-Host "Otherwise search the web UI: https://claude.ai/cowork" -ForegroundColor Yellow
    }
    exit $(if ($hits.Count) { 0 } else { 1 })
}

# Navigate to Projects page first if needed
if ($ListProjects -or $Project) {
    $projectsBtn = $win.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    ) | Where-Object { $_.Current.Name -eq "Projects" } | Select-Object -First 1

    if ($projectsBtn) {
        Focus-ClaudeWindow
        Invoke-UiaElement $projectsBtn
        Start-Sleep -Milliseconds 1000
        $win = Get-ClaudeWindow  # re-query after navigation
    }
}

if ($ListProjects) {
    Write-Host "Projects:" -ForegroundColor Cyan
    Get-ProjectLinks $win | ForEach-Object { Write-Host "  - $($_.Current.Name)" }
    exit 0
}

if ($Project) {
    $projLink = Get-ProjectLinks $win | Where-Object { $_.Current.Name -like "*$Project*" } | Select-Object -First 1
    if (-not $projLink) { Write-Error "Project '$Project' not found."; exit 1 }

    Write-Host "Opening project: $($projLink.Current.Name)" -ForegroundColor Cyan
    Focus-ClaudeWindow
    Invoke-UiaElement $projLink
    Start-Sleep -Milliseconds 1200
    $win = Get-ClaudeWindow

    $chatLinks = Get-ProjectChatLinks $win
    Write-Host "Chats in project (newest first):" -ForegroundColor Cyan
    $chatLinks | ForEach-Object { Write-Host "  - $($_.Current.Name)" }

    if ($Name) {
        $match = $chatLinks | Where-Object { $_.Current.Name -like "*$Name*" } | Select-Object -First 1
    } else {
        $match = $chatLinks | Select-Object -First 1
    }

    if (-not $match) { Write-Warning "No matching chat found in project."; exit 1 }
    Write-Host "Opening: $($match.Current.Name)" -ForegroundColor Green
    Focus-ClaudeWindow
    Invoke-UiaElement $match
    exit 0
}

# Default: search sidebar chats
$chats = Get-SidebarChatButtons $win

if ($List) {
    Write-Host "Sidebar chats ($($chats.Count)):" -ForegroundColor Cyan
    $chats | ForEach-Object { Write-Host "  - $($_.Current.Name)" }
    exit 0
}

if (-not $Name) { Write-Error "Provide -Name or use -List / -ListProjects."; exit 1 }

$match = $chats | Where-Object { $_.Current.Name -like "*$Name*" } | Select-Object -First 1
if (-not $match) {
    Write-Warning "No sidebar chat matching '$Name'. Checking projects..."
    # Fall back: open Projects and search all project chats
    $projectsBtn = $win.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    ) | Where-Object { $_.Current.Name -eq "Projects" } | Select-Object -First 1
    if ($projectsBtn) { Focus-ClaudeWindow; Invoke-UiaElement $projectsBtn; Start-Sleep -Milliseconds 1000 }
    Write-Host "Available sidebar chats:" -ForegroundColor Yellow
    $chats | Select-Object -First 20 | ForEach-Object { Write-Host "  - $($_.Current.Name)" }
    exit 1
}

Write-Host "Opening: $($match.Current.Name)" -ForegroundColor Green
Focus-ClaudeWindow
Invoke-UiaElement $match
Write-Host "Done." -ForegroundColor Green
