"""Shared Claude Code CLI invocation helper (subscription seat, not the metered API)."""
import json
import shutil
import subprocess

CLAUDE_BIN = "claude"
MODEL = "claude-opus-5"


def resolve_claude_cmd():
    """Resolve `claude` to an invocable argv prefix.

    On Windows, npm installs global CLIs as .CMD shim scripts (e.g.
    C:\\nvm4w\\nodejs\\claude.CMD). subprocess.run can't execute those
    directly in list form (raises WinError 2 / FileNotFoundError) —
    cmd.exe has to be the one interpreting them.
    """
    resolved = shutil.which(CLAUDE_BIN)
    if not resolved:
        return None
    if resolved.lower().endswith((".cmd", ".bat")):
        return ["cmd", "/c", resolved]
    return [resolved]


def call_claude(cmd_prefix, prompt, schema):
    proc = subprocess.run(
        cmd_prefix + ["-p", "--model", MODEL,
                      "--output-format", "json", "--json-schema", json.dumps(schema),
                      "--tools", "", "--no-session-persistence"],
        input=prompt, capture_output=True, text=True, timeout=180,
        encoding="utf-8", errors="replace",
    )
    if proc.returncode != 0:
        raise RuntimeError(f"claude CLI exited {proc.returncode}: {proc.stderr[:300]}")
    data = json.loads(proc.stdout)
    if data.get("is_error"):
        raise RuntimeError(f"claude CLI error: {str(data.get('result', ''))[:300]}")
    return data["structured_output"]
