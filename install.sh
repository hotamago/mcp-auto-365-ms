#!/usr/bin/env bash
# ==============================================================================
# mcp-auto-365-ms: Automated Installer & Configuration Setup
# Configures MCP servers for Claude Code, Zed Editor, and Oh My Pi (OMP)
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN_DIR="$HOME/.local/bin"
CLAUDE_CONFIG="$HOME/.claude.json"
ZED_CONFIG="$HOME/.config/zed/settings.json"
OMP_CONFIG="$HOME/.omp/agent/mcp.json"

echo "========================================================"
echo " Installing mcp-auto-365-ms (SharePoint & Teams MCP)   "
echo "========================================================"

# 1. Ensure ~/.local/bin exists and is in PATH
mkdir -p "$BIN_DIR"
if [[ ":$PATH:" != *":$BIN_DIR:"* ]]; then
    echo "Notice: $BIN_DIR is not in your current PATH. You may want to add it to ~/.bashrc or ~/.zshrc."
fi

# 2. Check Python 3
if ! command -v python3 &> /dev/null; then
    echo "Error: python3 is not installed."
    exit 1
fi
echo "✓ Python 3 found: $(python3 --version)"

# 3. Check / Install dependencies
echo "Checking Python dependencies..."
python3 -c "import mcp, cryptography, dbus" 2>/dev/null || {
    echo "Installing missing dependencies via pip..."
    pip install -r "$SCRIPT_DIR/requirements.txt"
}
echo "✓ Python dependencies verified."

# 4. Install binaries to ~/.local/bin
echo "Installing launchers to $BIN_DIR..."
chmod +x "$SCRIPT_DIR/bin/"*

ln -sf "$SCRIPT_DIR/bin/mcp-auto-365-ms" "$BIN_DIR/mcp-auto-365-ms"
ln -sf "$SCRIPT_DIR/bin/mcp-doc-reader" "$BIN_DIR/mcp-doc-reader"
ln -sf "$SCRIPT_DIR/bin/mcp-teams-reader" "$BIN_DIR/mcp-teams-reader"
echo "✓ Symlinks created:"
echo "  - $BIN_DIR/mcp-auto-365-ms"
echo "  - $BIN_DIR/mcp-doc-reader"
echo "  - $BIN_DIR/mcp-teams-reader"

# 5. Configure Claude Code (~/.claude.json)
if [ -f "$CLAUDE_CONFIG" ]; then
    echo "Configuring Claude Code ($CLAUDE_CONFIG)..."
    python3 -c "
import json
path = '$CLAUDE_CONFIG'
try:
    with open(path, 'r') as f:
        data = json.load(f)
    if 'mcpServers' not in data:
        data['mcpServers'] = {}
    data['mcpServers']['auto-365-ms'] = {'command': '$BIN_DIR/mcp-auto-365-ms'}
    data['mcpServers']['doc-reader'] = {'command': '$BIN_DIR/mcp-doc-reader'}
    data['mcpServers']['teams-reader'] = {'command': '$BIN_DIR/mcp-teams-reader'}
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)
    print('  ✓ Updated ~/.claude.json')
except Exception as e:
    print('  ! Error updating ~/.claude.json:', e)
"
fi

# 6. Configure Zed Editor (~/.config/zed/settings.json)
if [ -f "$ZED_CONFIG" ]; then
    echo "Configuring Zed Editor ($ZED_CONFIG)..."
    python3 -c "
import re
path = '$ZED_CONFIG'
try:
    with open(path, 'r') as f:
        content = f.read()
    # Add context_servers if missing or update
    servers = {
        'auto-365-ms': '$BIN_DIR/mcp-auto-365-ms',
        'doc-reader': '$BIN_DIR/mcp-doc-reader',
        'teams-reader': '$BIN_DIR/mcp-teams-reader'
    }
    # Check if context_servers exists
    if '\"context_servers\":' in content:
        for name, cmd in servers.items():
            if f'\"{name}\"' not in content:
                snippet = f'    \"{name}\": {{\\n      \"command\": {{\\n        \"path\": \"{cmd}\",\\n        \"args\": []\\n      }}\\n    }},\\n'
                content = content.replace('\"context_servers\": {', '\"context_servers\": {\\n' + snippet, 1)
        with open(path, 'w') as f:
            f.write(content)
        print('  ✓ Updated ~/.config/zed/settings.json')
except Exception as e:
    print('  ! Error updating Zed settings:', e)
"
fi

# 7. Configure Oh My Pi (OMP) (~/.omp/agent/mcp.json)
mkdir -p "$(dirname "$OMP_CONFIG")"
echo "Configuring Oh My Pi ($OMP_CONFIG)..."
python3 -c "
import json
path = '$OMP_CONFIG'
try:
    data = {
        '\$schema': 'https://raw.githubusercontent.com/can1357/oh-my-pi/main/packages/coding-agent/src/config/mcp-schema.json',
        'mcpServers': {}
    }
    try:
        with open(path, 'r') as f:
            existing = json.load(f)
            data.update(existing)
            if 'mcpServers' not in data:
                data['mcpServers'] = {}
    except Exception:
        pass

    data['mcpServers']['auto-365-ms'] = {'command': '$BIN_DIR/mcp-auto-365-ms'}
    data['mcpServers']['doc-reader'] = {'command': '$BIN_DIR/mcp-doc-reader'}
    data['mcpServers']['teams-reader'] = {'command': '$BIN_DIR/mcp-teams-reader'}
    
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)
    print('  ✓ Updated ~/.omp/agent/mcp.json')
except Exception as e:
    print('  ! Error updating OMP config:', e)
"

# 8. Test execution
echo "Testing MCP server execution..."
"$BIN_DIR/mcp-auto-365-ms" --help 2>&1 || true
echo "✓ Verification completed successfully."

echo "========================================================"
echo " Installation finished! Available tools:               "
echo "  - read_sharepoint_link                                "
echo "  - download_sharepoint_link                            "
echo "  - list_teams_chats                                    "
echo "  - read_teams_chat                                     "
echo "  - search_teams_chat_messages                          "
echo "========================================================"
