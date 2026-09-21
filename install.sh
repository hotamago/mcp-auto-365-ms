#!/usr/bin/env bash
# ==============================================================================
# mcp-auto-365-ms installer: uv-based, configures Claude Code, Zed and Oh My Pi.
# ==============================================================================
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN_DIR="$HOME/.local/bin"
CLAUDE_CONFIG="$HOME/.claude.json"
ZED_CONFIG="$HOME/.config/zed/settings.json"
OMP_CONFIG="$HOME/.omp/agent/mcp.json"

echo "========================================================"
echo "  Installing mcp-auto-365-ms (SharePoint & Teams MCP)   "
echo "========================================================"

# 1. uv is the only prerequisite; it manages Python and the dependencies.
UV="${UV:-$(command -v uv 2>/dev/null || echo "$HOME/.local/bin/uv")}"
if [ ! -x "$UV" ]; then
    echo "Error: 'uv' not found."
    echo "  Install it with: curl -LsSf https://astral.sh/uv/install.sh | sh"
    exit 1
fi
echo "✓ uv found: $("$UV" --version)"

# 2. Resolve the environment up front, so the first MCP spawn is not blocked
#    by a sync (an MCP host will time out waiting on it).
echo "Syncing dependencies (this also installs the right Python)..."
"$UV" sync --project "$SCRIPT_DIR" --frozen
echo "✓ Environment ready at $SCRIPT_DIR/.venv"

# 3. Sanity check: the server must start and list its tools.
echo "Verifying the server starts..."
TOOL_COUNT="$("$UV" run --project "$SCRIPT_DIR" --frozen --quiet python - <<'PY'
import sys
from pathlib import Path
sys.path.insert(0, str(Path("src").resolve()))
from mcp.server.mcpserver import MCPServer
from tools import register_all
mcp = MCPServer("verify")
register_all(mcp)
import anyio
print(len(anyio.run(mcp.list_tools)))
PY
)"
echo "✓ Server registers $TOOL_COUNT tools"

# 4. Install launchers.
mkdir -p "$BIN_DIR"
chmod +x "$SCRIPT_DIR/bin/"*
for launcher in mcp-auto-365-ms mcp-doc-reader mcp-teams-reader mcp-word-companion; do
    ln -sf "$SCRIPT_DIR/bin/$launcher" "$BIN_DIR/$launcher"
done
echo "✓ Launchers linked into $BIN_DIR"
if [[ ":$PATH:" != *":$BIN_DIR:"* ]]; then
    echo "  Note: $BIN_DIR is not on your PATH; add it to ~/.zshrc or ~/.bashrc."
fi

# 5. Register the unified server with each harness.
#    Only 'auto-365-ms' is registered: doc-reader and teams-reader expose
#    subsets of the same tools and would show up as duplicates.
register_harness() {
    local config_path="$1" json_key="$2" label="$3"
    [ -f "$config_path" ] || { echo "  - $label: not installed, skipped"; return 0; }
    CONFIG_PATH="$config_path" JSON_KEY="$json_key" LABEL="$label" \
    BIN_PATH="$BIN_DIR/mcp-auto-365-ms" python3 <<'PY'
import json, os, shutil

path, key, label, bin_path = (os.environ[k] for k in ("CONFIG_PATH", "JSON_KEY", "LABEL", "BIN_PATH"))
try:
    with open(path) as fh:
        data = json.load(fh)
except (json.JSONDecodeError, OSError) as exc:
    print(f"  - {label}: could not parse ({exc}); left untouched")
    raise SystemExit(0)

shutil.copy2(path, path + ".mcp365.bak")
servers = data.setdefault(key, {})
entry = {"command": {"path": bin_path, "args": []}} if key == "context_servers" else {"command": bin_path}
servers["auto-365-ms"] = entry
for stale in ("doc-reader", "teams-reader"):
    servers.pop(stale, None)

with open(path, "w") as fh:
    json.dump(data, fh, indent=2, ensure_ascii=False)
print(f"  - {label}: registered 'auto-365-ms' (backup at {os.path.basename(path)}.mcp365.bak)")
PY
}

echo "Configuring harnesses..."
register_harness "$CLAUDE_CONFIG" "mcpServers" "Claude Code"
register_harness "$ZED_CONFIG" "context_servers" "Zed Editor"
register_harness "$OMP_CONFIG" "mcpServers" "Oh My Pi"

# 6. Seed a user config file if there is none.
USER_CONFIG_DIR="${XDG_CONFIG_HOME:-$HOME/.config}/mcp-auto-365-ms"
if [ ! -f "$USER_CONFIG_DIR/config.toml" ]; then
    mkdir -p "$USER_CONFIG_DIR"
    cp "$SCRIPT_DIR/config.example.toml" "$USER_CONFIG_DIR/config.toml"
    echo "✓ Seeded config at $USER_CONFIG_DIR/config.toml"
fi

echo
echo "Done. Restart your editor/harness, then run the 'check_365_connection'"
echo "tool to verify every Microsoft 365 channel is authenticated."
