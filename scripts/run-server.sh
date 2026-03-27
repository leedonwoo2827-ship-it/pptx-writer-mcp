#!/bin/bash
# pptx-writer MCP server launcher
# CLAUDE_PLUGIN_ROOT and CLAUDE_PLUGIN_DATA are exported as env vars

if [ -f "$CLAUDE_PLUGIN_DATA/.venv/Scripts/python.exe" ]; then
  exec "$CLAUDE_PLUGIN_DATA/.venv/Scripts/python.exe" "$CLAUDE_PLUGIN_ROOT/server.py"
elif [ -f "$CLAUDE_PLUGIN_DATA/.venv/bin/python" ]; then
  exec "$CLAUDE_PLUGIN_DATA/.venv/bin/python" "$CLAUDE_PLUGIN_ROOT/server.py"
else
  echo "[pptx-writer] Error: Virtual environment not found. Restart Claude Code to trigger setup." >&2
  exit 1
fi
