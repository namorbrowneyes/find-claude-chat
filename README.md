# find-claude-chat

PowerShell script to find Claude conversations by name. Searches **two** stores:

1. **Claude desktop app** — the running app's sidebar/projects, via Windows UI Automation.
2. **Claude Code** (`-Code`) — your local CLI transcripts under `~/.claude/projects/**/*.jsonl`.

## Why this exists

The Claude **desktop app** stores history server-side on claude.ai — no local file to grep. But the running app exposes its sidebar through Windows Accessibility (UIA), so we enumerate the conversation buttons and click the right one.

**Claude Code** is the opposite: every session is a local JSONL transcript, but there's no built-in "find the chat about X" across projects. `-Code` mode greps those transcripts by topic, searching the *actual conversation* (your prompts + Claude's replies) and ignoring injected `CLAUDE.md`/memory context and the high-volume programmatic agent runs.

## Usage

```powershell
# ── Desktop app (UIA) ──
.\Find-ClaudeChat.ps1 -Name "mental health"     # open a conversation by partial name
.\Find-ClaudeChat.ps1 -List                     # list visible sidebar chats
.\Find-ClaudeChat.ps1 -ListProjects             # list projects

# ── Claude Code transcripts (-Code) ──
.\Find-ClaudeChat.ps1 -Code -Name "hermes desktop"          # multi-word = AND (all terms in-conversation)
.\Find-ClaudeChat.ps1 -Code -Name "dashboard" -Project hermes-deployment
.\Find-ClaudeChat.ps1 -Code -Name "hermes desktop" -Resume  # resume the top match in its cwd
.\Find-ClaudeChat.ps1 -Code -List                           # recent Claude Code chats
.\Find-ClaudeChat.ps1 -Code -Name "agent" -IncludeAgents    # also include programmatic SDK/agent runs
```

Each `-Code` result shows the project, date range, match count, a snippet, and a ready-to-run `claude --resume <id>` command.

## Requirements

- Windows 10/11
- Claude desktop app running
- PowerShell 5.1+ or PowerShell 7+
- No external dependencies — uses built-in .NET UIAutomationClient

## How it works

**Desktop app (UIA):**
1. Finds the Claude window via `AutomationElement` (window title = "Claude")
2. Enumerates all `ControlType.Button` descendants in the window
3. Filters to conversation items (excludes "More options for ...", UI chrome buttons)
4. Finds the first button whose name contains the search term
5. Brings Claude to the foreground and fires `InvokePattern.Invoke()` on the button

**Claude Code (`-Code`):**
1. Enumerates session transcripts at `~/.claude/projects/*/*.jsonl`
2. AND-prefilters to files containing every search word (fast `Select-String`)
3. By default keeps only **interactive** chats (`entrypoint != 'sdk-cli'`), hiding programmatic agent/SDK runs that swamp results with injected-context matches
4. Deep-parses each candidate, scoring matches in real conversation turns only (skips `<system-reminder>` / `CLAUDE.md` / memory blocks)
5. Ranks full-term matches first, then most-recently-active, and prints a `claude --resume` line per hit

## Notes

- **Desktop:** only sidebar-visible conversations can be found; scroll the sidebar first if a chat isn't showing. Matches the first result — use `-List` to disambiguate.
- **-Code:** multi-word `-Name` is an AND across the conversation, not an exact phrase. Use `-Project <substr>` to scope by working directory, `-IncludeAgents` to include the agent's own SDK sessions, and `-Limit` to widen the recency cap.
