<#
.SYNOPSIS
    Find and open a Claude desktop app conversation by name using Windows UI Automation.

.DESCRIPTION
    Enumerates all conversation buttons in the running Claude desktop app window,
    searches for a partial or full name match, and opens the matching chat.

.PARAMETER Name
    Partial or full conversation title to search for (case-insensitive).

.PARAMETER List
    If specified, lists all available conversation titles without opening any.

.EXAMPLE
    .\Find-ClaudeChat.ps1 -Name "mental health"
    .\Find-ClaudeChat.ps1 -Name "charlie health"
    .\Find-ClaudeChat.ps1 -List

.NOTES
    Requires the Claude desktop app to be running.
    Uses Windows UI Automation (UIAutomationClient/.NET).
#>

param(
    [string]$Name,
    [switch]$List
)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinFocus {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

function Get-ClaudeWindow {
    $win = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::NameProperty, "Claude"
        )
    )
    return $win
}

function Get-ConversationButtons($window) {
    $buttons = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
    )
    # Filter to conversation buttons (exclude "More options for ..." and short UI buttons)
    return $buttons | Where-Object {
        $n = $_.Current.Name
        $n -and $n.Length -gt 5 -and $n -notlike "More options for *" -and
        $n -notin @("New chat","Collapse","Expand","Send","Close","Settings","Menu")
    }
}

$claudeWindow = Get-ClaudeWindow
if (-not $claudeWindow) {
    Write-Error "Claude desktop app is not running or window not found."
    exit 1
}

$convButtons = Get-ConversationButtons $claudeWindow

if ($List) {
    Write-Host "Claude conversations ($($convButtons.Count) found):" -ForegroundColor Cyan
    foreach ($btn in $convButtons) {
        Write-Host "  - $($btn.Current.Name)"
    }
    exit 0
}

if (-not $Name) {
    Write-Error "Provide -Name <search term> or use -List to see all conversations."
    exit 1
}

$match = $convButtons | Where-Object { $_.Current.Name -like "*$Name*" } | Select-Object -First 1

if (-not $match) {
    Write-Warning "No conversation matching '$Name' found."
    Write-Host "Available conversations:" -ForegroundColor Yellow
    foreach ($btn in $convButtons | Select-Object -First 20) {
        Write-Host "  - $($btn.Current.Name)"
    }
    exit 1
}

Write-Host "Opening: $($match.Current.Name)" -ForegroundColor Green

# Bring Claude to foreground
$proc = Get-Process claude | Where-Object { $_.MainWindowTitle -eq "Claude" } | Select-Object -First 1
if ($proc) {
    [WinFocus]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null
    [WinFocus]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 300
}

# Click the conversation
try {
    $invoke = $match.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $invoke.Invoke()
    Write-Host "Done." -ForegroundColor Green
} catch {
    Write-Error "Failed to open conversation: $_"
    exit 1
}
