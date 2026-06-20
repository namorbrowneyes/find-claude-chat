# find-claude-chat

Find a past Claude conversation by topic. Two tools, one per platform:

- **Windows** — `Find-ClaudeChat.ps1` (PowerShell). Searches the running desktop app's
  sidebar via UI Automation, and Claude Code transcripts via `-Code`.
- **macOS** — `find-claude-chat-mac.py` (Python). Searches the desktop app's local HTTP
  cache, since macOS has no UIA. Decompresses the cached conversation bodies and greps the
  real message text by topic.

The PowerShell tool searches **two** stores:

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

### macOS — desktop app cache (`find-claude-chat-mac.py`)

```bash
python3 find-claude-chat-mac.py "kvm"                  # find by topic
python3 find-claude-chat-mac.py "macbook windows switch"  # multi-word = AND across the convo
python3 find-claude-chat-mac.py --list                 # all cached conversations, newest first
python3 find-claude-chat-mac.py "kvm" --open           # open the top match in the browser
python3 find-claude-chat-mac.py "kvm" --json           # machine-readable output
```

Each result shows the title, a `https://claude.ai/chat/<uuid>` link, match count, last-active
date, and a snippet. Dependency: `zstandard` (`pip3 install zstandard`); falls back to the
`zstd` CLI if absent.

## Requirements

**macOS (`find-claude-chat-mac.py`):**

- macOS with the Claude desktop app installed
- Python 3 + `zstandard` (or the `zstd` CLI on PATH)
- Read-only; no network calls

**Windows (`Find-ClaudeChat.ps1`):**

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

**macOS desktop cache (`find-claude-chat-mac.py`):**
1. Enumerates Chromium Simple Cache entries under `~/Library/Application Support/Claude/Cache/Cache_Data/*_0`
2. Parses each file's Simple Cache header (magic `0xfcfb6d1ba7725c30`) to read the request URL (the cache key); keeps only `chat_conversations/<uuid>` responses
3. Locates the **zstd** frame in the body (claude.ai serves `content-encoding: zstd`) and decompresses it — a plain grep fails because the JSON is compressed
4. Whole-word AND-matches the search terms against the decoded message text, dedupes to the richest cache entry per conversation uuid
5. Ranks by match count then last-active, and prints the title, a `https://claude.ai/chat/<uuid>` link, and a snippet

## Notes

- **Desktop:** only sidebar-visible conversations can be found; scroll the sidebar first if a chat isn't showing. Matches the first result — use `-List` to disambiguate.
- **-Code:** multi-word `-Name` is an AND across the conversation, not an exact phrase. Use `-Project <substr>` to scope by working directory, `-IncludeAgents` to include the agent's own SDK sessions, and `-Limit` to widen the recency cap.
- **macOS:** only conversations **opened in the desktop app on this Mac** are cached, so an unopened chat won't be found — open/scroll it once to cache it. The script is read-only and never mutates the cache. On this Mac it's also wired up as a `/find-claude-chat` Claude Code skill.
