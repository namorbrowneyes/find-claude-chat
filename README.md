# find-claude-chat

Find a past Claude conversation by topic. Two tools, one per platform:

- **Windows** — `Find-ClaudeChat.ps1` (PowerShell). Searches the running desktop app via
  UI Automation (sidebar, the app's own Search palette, scheduled-task runs) and Claude
  Code transcripts via `-Code`.
- **macOS** — `find-claude-chat-mac.py` (Python). Searches the desktop app's local HTTP
  cache, since macOS has no UIA. Decompresses the cached conversation bodies and greps the
  real message text by topic.

The PowerShell tool searches **four** stores:

1. **Claude desktop app sidebar** — the running app's sidebar/projects, via Windows UI Automation.
2. **Claude Code** (`-Code`) — your local CLI transcripts under `~/.claude/projects/**/*.jsonl`.
3. **Everything on claude.ai** (`-Cloud`) — full-text search of every claude.ai chat and every Claude Code session (local *and* cloud, incl. ones started from the phone app), by driving the desktop app's own Search palette (server-side index).
4. **Scheduled-task runs / Cowork sessions** (`-Scheduled`) — the `https://claude.ai/cowork/cse_*` sessions that a scheduled task (e.g. *Morning brief*) produces. These have no local transcript and are **not** in the palette's index, so they're listed by date from the task's page and opened from there.

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

# ── Everything on claude.ai (-Cloud, needs the app running) ──
.\Find-ClaudeChat.ps1 -Cloud -Name "chatgpt subscription"             # chats + Code sessions, full text
.\Find-ClaudeChat.ps1 -Cloud -Name "kling credits" -Type Sessions -Open # -Type: All|Sessions|Code|Projects|Artifacts|Scheduled
.\Find-ClaudeChat.ps1 -Cowork -Name "chatgpt subscription"             # -Cloud + pointers for Cowork sessions

# ── Scheduled-task runs = Cowork sessions (-Scheduled) ──
.\Find-ClaudeChat.ps1 -Scheduled                                       # list scheduled tasks
.\Find-ClaudeChat.ps1 -Scheduled -Task "Morning brief"                 # list that task's runs + cse_ URLs
.\Find-ClaudeChat.ps1 -Scheduled -Task "Morning brief" -Name "Today" -Open   # open today's run in the app
.\Find-ClaudeChat.ps1 -Scheduled -Task "Morning brief" -Find "chatgpt pro subscription" -Open   # search INSIDE the runs
.\Find-ClaudeChat.ps1 -Scheduled -Task "Morning brief" -Find "kling" -Limit 40 -Refresh         # re-read, ignore cache
```

`-Cloud` prints, per hit, the kind (`chat` / `code` / `cowork` / `task`) and a URL or `claude --resume` line. `-Scheduled` prints each run's `https://claude.ai/cowork/cse_…` link; `-Name` matches the run's date label (`Today`, `Yesterday`, `Sep 21`). `-Find` reads the runs' actual content: it opens each run in the app (newest first, up to `-Limit`, ~2 s each), extracts the rendered conversation, caches it in `%LOCALAPPDATA%\find-claude-chat\runs\<cse_id>.txt`, and AND-matches whole words with snippets. Cached runs are instant on later searches; runs still marked *Awaiting input* / *Unread* are re-read each time.

Each `-Code` result shows the project, date range, match count, a snippet, and a ready-to-run `claude --resume <id>` command.

### macOS — desktop app cache (`find-claude-chat-mac.py`)

```bash
pip3 install -r requirements.txt                       # one-time setup (see requirements.txt)
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
- Python 3.8+ and the packages in [`requirements.txt`](requirements.txt) (`pip3 install -r requirements.txt`) — `zstandard` is required (or the `zstd` CLI on PATH; `brew install zstd`); `brotli` is optional
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

**Cloud (`-Cloud`):**
1. Finds the app's **main** window (named "Claude"; popped-out session windows carry the chat title and are skipped)
2. Invokes the sidebar **Search** button → command palette (`AutomationId command-palette-input`), sets the query via `ValuePattern`
3. Optionally selects a filter tab (`Sessions`, `Code`, …), then reads `command-palette-results` list items. Item ids tell the kind: `local_<guid>` = Claude Code session, `<guid>` = claude.ai chat, `cse_…` = Cowork session, `trig_…` = scheduled task
4. `-Open` invokes the top item; otherwise the palette is closed again

**Scheduled runs (`-Scheduled`):**
1. Palette → `Scheduled` tab lists the tasks (`trig_…`); the chosen task is invoked, which navigates to its page
2. The page lists runs as hyperlinks (`Today at 9:12 AM`, …); each hyperlink's `ValuePattern` is the `https://claude.ai/cowork/cse_…` URL
3. `-Name` matches the date label, `-Open` invokes that hyperlink
4. `-Find`: for each run, invoke its link → wait until the "Primary pane" text stops growing → collect every Text/ListItem/Hyperlink name inside that pane (sidebar excluded) → cache → press **Back** and re-bind the run links (the page re-renders). Matching is whole-word, case-insensitive, AND across terms

**macOS desktop cache (`find-claude-chat-mac.py`):**
1. Enumerates Chromium Simple Cache entries under `~/Library/Application Support/Claude/Cache/Cache_Data/*_0`
2. Parses each file's Simple Cache header (magic `0xfcfb6d1ba7725c30`) to read the request URL (the cache key); keeps only `chat_conversations/<uuid>` responses
3. Locates the **zstd** frame in the body (claude.ai serves `content-encoding: zstd`) and decompresses it — a plain grep fails because the JSON is compressed
4. Whole-word AND-matches the search terms against the decoded message text, dedupes to the richest cache entry per conversation uuid
5. Ranks by match count then last-active, and prints the title, a `https://claude.ai/chat/<uuid>` link, and a snippet

## Notes

- **Cowork sessions** (`claude.ai/cowork/cse_*`, incl. every scheduled-task run) are cloud-only: no local file, and as of 2026-09-23 the desktop palette does not full-text index them. `-Scheduled` is the only way to reach them from this script, and only by date. Ad-hoc Cowork sessions that aren't task runs: use the web UI.
- **Desktop:** only sidebar-visible conversations can be found; scroll the sidebar first if a chat isn't showing. Matches the first result — use `-List` to disambiguate.
- **-Code:** multi-word `-Name` is an AND across the conversation, not an exact phrase. Use `-Project <substr>` to scope by working directory, `-IncludeAgents` to include the agent's own SDK sessions, and `-Limit` to widen the recency cap.
- **macOS:** only conversations **opened in the desktop app on this Mac** are cached, so an unopened chat won't be found — open/scroll it once to cache it. The script is read-only and never mutates the cache. On this Mac it's also wired up as a `/find-claude-chat` Claude Code skill.
