#!/usr/bin/env python3
"""
find-claude-chat-mac — find Claude **desktop app** conversations on macOS by topic.

The Windows desktop app stores history server-side (no local file), so Find-ClaudeChat.ps1
reaches it through UI Automation. macOS has no UIA, BUT the desktop app is an Electron app
that caches every conversation it has opened as an HTTP response under:

    ~/Library/Application Support/Claude/Cache/Cache_Data/

Those cache entries are framed in Chromium's "Simple Cache" format and the JSON bodies are
**zstd-compressed** (content-encoding: zstd), so a plain `grep` finds only byte-coincidences.
This script parses the Simple Cache header, locates the zstd stream, decompresses each
conversation body, and searches the *actual* message text — then prints the title, a
claude.ai link, match count, last-active date, and a snippet.

Caveat: this only sees conversations the desktop app has cached locally (i.e. ones you've
opened on this Mac). It does not hit the network. Scroll/open a chat once and it gets cached.

Usage:
    python3 find-claude-chat-mac.py "kvm"                 # multi-word = AND across the convo
    python3 find-claude-chat-mac.py "macbook windows switch"
    python3 find-claude-chat-mac.py --list                # all cached conversations, most-recent first
    python3 find-claude-chat-mac.py "kvm" --open          # open the top match in your browser
    python3 find-claude-chat-mac.py "kvm" --json          # machine-readable output

Dependencies: see requirements.txt (`pip3 install -r requirements.txt`). zstandard is required;
auto-falls back to the `zstd` CLI if the module is missing. brotli/gzip/deflate bodies are also
handled if Claude ever changes encoding.
"""
import sys, os, re, struct, glob, argparse, subprocess, json, gzip, zlib, io

CACHE_DIR = os.path.expanduser(
    "~/Library/Application Support/Claude/Cache/Cache_Data"
)
SIMPLE_MAGIC = struct.pack("<Q", 0xFCFB6D1BA7725C30)  # Chromium Simple Cache file header
ZSTD_MAGIC = b"\x28\xb5\x2f\xfd"


# ---- zstd decode (module preferred, CLI fallback) ----------------------------
def _zstd_decompress(blob):
    try:
        import zstandard
        return zstandard.ZstdDecompressor().decompressobj().decompress(blob)
    except ImportError:
        pass
    except Exception:
        # streaming reader tolerates trailing cache bytes after the zstd frame
        try:
            import zstandard
            return zstandard.ZstdDecompressor().stream_reader(io.BytesIO(blob)).read()
        except Exception:
            pass
    # CLI fallback
    try:
        p = subprocess.run(["zstd", "-d", "-c"], input=blob,
                           capture_output=True, timeout=30)
        if p.returncode == 0 and p.stdout:
            return p.stdout
    except Exception:
        pass
    return None


def _decode_body(raw):
    """Return (key, decoded_body_bytes|None) for one Simple Cache file."""
    if raw[0:8] != SIMPLE_MAGIC or len(raw) < 20:
        return None, None
    _ver, key_len, _hash = struct.unpack("<III", raw[8:20])
    key = raw[20:20 + key_len].decode("utf8", "replace")
    if "chat_conversations" not in key:
        return key, None
    # Body is the HTTP response body stored right after the key. Find its encoding.
    region = raw[20 + key_len - 8:]
    body = None
    z = region.find(ZSTD_MAGIC)
    if z != -1:
        body = _zstd_decompress(region[z:])
    if body is None:
        tail = raw[20 + key_len:]
        for fn in (_try_brotli, gzip.decompress,
                   lambda b: zlib.decompress(b),
                   lambda b: zlib.decompress(b, -15)):
            try:
                d = fn(tail)
                if b'"name"' in d or b"chat_messages" in d:
                    body = d
                    break
            except Exception:
                pass
    return key, body


def _try_brotli(b):
    import brotli
    return brotli.decompress(b)


# ---- conversation extraction -------------------------------------------------
def _iter_conversations():
    """Yield dicts: {uuid,title,created,updated,text(lowercased),raw_text}."""
    for f in glob.glob(os.path.join(CACHE_DIR, "*_0")):
        try:
            raw = open(f, "rb").read()
        except Exception:
            continue
        key, body = _decode_body(raw)
        if not key or "chat_conversations" not in key or not body:
            continue
        m = re.search(r"chat_conversations/([0-9a-f-]{36})", key)
        if not m:
            continue
        uuid = m.group(1)
        name = re.search(rb'"name"\s*:\s*"((?:[^"\\]|\\.)*)"', body)
        created = re.search(rb'"created_at"\s*:\s*"([^"]+)"', body)
        updated = re.search(rb'"updated_at"\s*:\s*"([^"]+)"', body)
        text = body.decode("utf8", "replace")
        yield {
            "uuid": uuid,
            "title": (name.group(1).decode("utf8", "replace") if name else "(untitled)"),
            "created": created.group(1).decode() if created else "",
            "updated": updated.group(1).decode() if updated else "",
            "raw_text": text,
            "text": text.lower(),
            "file": f,
        }


def _dedupe(convs):
    """Keep the richest cache entry per conversation uuid."""
    best = {}
    for c in convs:
        cur = best.get(c["uuid"])
        if cur is None or len(c["raw_text"]) > len(cur["raw_text"]):
            best[c["uuid"]] = c
    return list(best.values())


def _snippet(raw_text, term):
    m = re.search(r".{0,70}" + re.escape(term) + r".{0,110}", raw_text, re.I | re.S)
    if not m:
        return ""
    return re.sub(r"\s+", " ", m.group(0)).strip()


def search(terms):
    words = [w.lower() for w in terms if w.strip()]
    results = []
    for c in _dedupe(_iter_conversations()):
        # AND across the conversation, whole-word match for each term
        if not all(re.search(r"\b" + re.escape(w) + r"\b", c["text"]) for w in words):
            continue
        count = sum(len(re.findall(r"\b" + re.escape(w) + r"\b", c["text"])) for w in words)
        c["count"] = count
        c["snippet"] = _snippet(c["raw_text"], words[0]) if words else ""
        results.append(c)
    # rank: most matches first, then most-recently updated
    results.sort(key=lambda c: (c["count"], c["updated"]), reverse=True)
    return results


def list_all():
    convs = _dedupe(_iter_conversations())
    convs.sort(key=lambda c: c["updated"], reverse=True)
    return convs


def _print(results, want_json):
    if want_json:
        print(json.dumps([
            {k: c[k] for k in ("uuid", "title", "created", "updated", "count", "snippet")
             if k in c} | {"link": f"https://claude.ai/chat/{c['uuid']}"}
            for c in results
        ], indent=2))
        return
    if not results:
        print("No matching cached conversations found.")
        print("(Only chats you've opened in the desktop app on this Mac are cached.)")
        return
    for c in results:
        print("=" * 66)
        print(f"TITLE : {c['title']}")
        print(f"LINK  : https://claude.ai/chat/{c['uuid']}")
        meta = []
        if c.get("count"):
            meta.append(f"matches={c['count']}")
        if c.get("updated"):
            meta.append(f"updated={c['updated'][:10]}")
        if meta:
            print("META  : " + "  ".join(meta))
        if c.get("snippet"):
            print(f"SNIPPET: {c['snippet'][:220]}")
    print("=" * 66)
    print(f"{len(results)} conversation(s).")


def main():
    ap = argparse.ArgumentParser(
        description="Find Claude desktop-app conversations cached locally on macOS.")
    ap.add_argument("terms", nargs="*", help="search words (AND, whole-word)")
    ap.add_argument("--list", action="store_true", help="list all cached conversations")
    ap.add_argument("--open", action="store_true", help="open the top match in browser")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    args = ap.parse_args()

    if not os.path.isdir(CACHE_DIR):
        sys.exit(f"Claude desktop cache not found at:\n  {CACHE_DIR}\n"
                 "Is the macOS Claude desktop app installed?")

    if args.list and not args.terms:
        results = list_all()
    elif args.terms:
        results = search(args.terms)
    else:
        ap.print_help()
        return

    _print(results, args.json)

    if args.open and results:
        subprocess.run(["open", f"https://claude.ai/chat/{results[0]['uuid']}"])


if __name__ == "__main__":
    main()
