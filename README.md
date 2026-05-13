# find-claude-chat

PowerShell script to find and open Claude desktop app conversations by name using Windows UI Automation.

## Why this exists

The Claude desktop app stores conversation history server-side on claude.ai — there's no exported local file to search. But the running app exposes its sidebar through Windows Accessibility (UIA), so we can enumerate all conversation buttons and click the right one programmatically.

## Usage

```powershell
# Open a conversation by partial name (case-insensitive)
.\Find-ClaudeChat.ps1 -Name "mental health"
.\Find-ClaudeChat.ps1 -Name "charlie health"

# List all visible conversations
.\Find-ClaudeChat.ps1 -List
```

## Requirements

- Windows 10/11
- Claude desktop app running
- PowerShell 5.1+ or PowerShell 7+
- No external dependencies — uses built-in .NET UIAutomationClient

## How it works

1. Finds the Claude window via `AutomationElement` (window title = "Claude")
2. Enumerates all `ControlType.Button` descendants in the window
3. Filters to conversation items (excludes "More options for ...", UI chrome buttons)
4. Finds the first button whose name contains the search term
5. Brings Claude to the foreground and fires `InvokePattern.Invoke()` on the button

## Notes

- Only conversations visible in the sidebar can be found (the app loads recent chats on startup)
- If a conversation isn't showing, scroll the sidebar in the Claude app first
- The script matches the first result — use `-List` to see all options if the name is ambiguous
