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

.EXAMPLE
    .\Find-ClaudeChat.ps1 -Name "mental health"
    .\Find-ClaudeChat.ps1 -Name "wordpress" -Project "Work - WordPress"
    .\Find-ClaudeChat.ps1 -List
    .\Find-ClaudeChat.ps1 -ListProjects

.NOTES
    Requires the Claude desktop app to be running.
    Uses Windows UI Automation (UIAutomationClient/.NET).
    Conversation order = most recently active first (Claude's own ordering).
#>

param(
    [string]$Name,
    [string]$Project,
    [switch]$List,
    [switch]$ListProjects
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

# ── Main ──────────────────────────────────────────────────────────────────────

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
