#!/bin/bash
# pptx-writer plugin auto-setup
# Installs Python dependencies into ${CLAUDE_PLUGIN_DATA}/.venv

set -e

PLUGIN_ROOT="${CLAUDE_PLUGIN_ROOT}"
PLUGIN_DATA="${CLAUDE_PLUGIN_DATA}"

# Check if already set up (requirements unchanged)
if diff -q "$PLUGIN_ROOT/requirements.txt" "$PLUGIN_DATA/requirements.txt" >/dev/null 2>&1; then
  exit 0
fi

echo "[pptx-writer] Installing dependencies..."

# Create venv
python -m venv "$PLUGIN_DATA/.venv"

# Detect pip path (Windows: Scripts/, Unix: bin/)
if [ -f "$PLUGIN_DATA/.venv/Scripts/pip.exe" ]; then
  PIP="$PLUGIN_DATA/.venv/Scripts/pip"
else
  PIP="$PLUGIN_DATA/.venv/bin/pip"
fi

# Install dependencies
"$PIP" install -q -r "$PLUGIN_ROOT/requirements.txt" || {
  rm -f "$PLUGIN_DATA/requirements.txt"
  echo "[pptx-writer] Failed to install dependencies" >&2
  exit 1
}

# Generate default template if not bundled
if [ ! -f "$PLUGIN_ROOT/templates/default.pptx" ]; then
  if [ -f "$PLUGIN_DATA/.venv/Scripts/python.exe" ]; then
    PYTHON="$PLUGIN_DATA/.venv/Scripts/python.exe"
  else
    PYTHON="$PLUGIN_DATA/.venv/bin/python"
  fi
  "$PYTHON" "$PLUGIN_ROOT/create_template.py"
fi

# Mark setup complete
cp "$PLUGIN_ROOT/requirements.txt" "$PLUGIN_DATA/requirements.txt"
echo "[pptx-writer] Setup complete."
