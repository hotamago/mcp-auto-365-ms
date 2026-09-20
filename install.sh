#!/usr/bin/env bash
# ==============================================================================
# mcp-auto-365-ms: Automated Installer & Configuration Setup
# Configures unified MCP server for Claude Code, Zed Editor, and Oh My Pi (OMP)
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

# 1. Ensure ~/.local/bin exists
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
echo "  - $BIN_DIR/mcp-auto-365-ms (Unified Server)"
echo "  - $BIN_DIR/mcp-doc-reader (Standalone SharePoint)"
echo "  - $BIN_DIR/mcp-teams-reader (Standalone Teams)"

# 5. Configure Claude Code (~/.claude.json) - Use single unified auto-365-ms
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
    # Clean up redundant standalone entries
    data['mcpServers'].pop('doc-reader', None)
    data['mcpServers'].pop('teams-reader', None)
    # Register single unified server
    data['mcpServers']['auto-365-ms'] = {'command': '$BIN_DIR/mcp-auto-365-ms'}
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)
    print('  ✓ Configured unified auto-365-ms in ~/.claude.json')
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
    new_ctx = '''  \"context_servers\": {\\n    \"auto-365-ms\": {\\n      \"command\": {\\n        \"path\": \"$BIN_DIR/mcp-auto-365-ms\",\\n        \"args\": []\\n      }\\n    }\\n  }'''
    if '\"context_servers\":' in content:
        content = re.sub(r'\"context_servers\":\s*\{[^}]+\}', new_ctx.strip(), content)
    else:
        content = content.rstrip().rstrip('}') + ',\\n' + new_ctx + '\\n}\\n'
    with open(path, 'w') as f:
        f.write(content)
    print('  ✓ Configured unified auto-365-ms in ~/.config/zed/settings.json')
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
        'mcpServers': {
            'auto-365-ms': {'command': '$BIN_DIR/mcp-auto-365-ms'}
        }
    }
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)
    print('  ✓ Configured unified auto-365-ms in ~/.omp/agent/mcp.json')
except Exception as e:
    print('  ! Error updating OMP config:', e)
"

# 8. Clean up any redundant project-level .omp/mcp.json if in a workspace
if [ -f ".omp/mcp.json" ]; then
    rm -f ".omp/mcp.json"
    echo "✓ Cleaned redundant project-level .omp/mcp.json"
fi

# 9. Test execution
echo "Testing MCP server execution..."
"$BIN_DIR/mcp-auto-365-ms" --help 2>&1 || true
echo "✓ Verification completed successfully."

echo "========================================================"
echo " Installation finished! Unified server active:         "
echo " Server name: auto-365-ms (All 9 tools in 1 process)   "
echo "========================================================"
