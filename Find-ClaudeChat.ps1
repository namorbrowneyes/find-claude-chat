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
    return [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty, "Claude"
        )
    )
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
