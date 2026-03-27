# -*- coding: utf-8 -*-
"""
PPTX Writer MCP Server Launcher

Automatically sets up venv + dependencies, then starts the MCP server.
Used by Claude Desktop plugin system.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

PLUGIN_ROOT = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", Path(__file__).parent.parent))
PLUGIN_DATA = Path(os.environ.get("CLAUDE_PLUGIN_DATA", PLUGIN_ROOT / ".venv_data"))

VENV_DIR = PLUGIN_DATA / ".venv"
REQ_FILE = PLUGIN_ROOT / "requirements.txt"
REQ_MARKER = PLUGIN_DATA / "requirements.txt"

if sys.platform == "win32":
    VENV_PYTHON = VENV_DIR / "Scripts" / "python.exe"
    VENV_PIP = VENV_DIR / "Scripts" / "pip"
else:
    VENV_PYTHON = VENV_DIR / "bin" / "python"
    VENV_PIP = VENV_DIR / "bin" / "pip"


def needs_install():
    """Check if dependencies need to be installed."""
    if not VENV_PYTHON.exists():
        return True
    if not REQ_MARKER.exists():
        return True
    return REQ_FILE.read_text(encoding="utf-8") != REQ_MARKER.read_text(encoding="utf-8")


def setup():
    """Create venv and install dependencies if needed."""
    if not needs_install():
        return

    print("[pptx-writer] Setting up dependencies...", file=sys.stderr)

    # Create venv
    if not VENV_PYTHON.exists():
        PLUGIN_DATA.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [sys.executable, "-m", "venv", str(VENV_DIR)],
            check=True,
        )

    # Install dependencies
    subprocess.run(
        [str(VENV_PIP), "install", "-q", "-r", str(REQ_FILE)],
        check=True,
    )

    # Mark as done
    shutil.copy2(REQ_FILE, REQ_MARKER)
    print("[pptx-writer] Setup complete.", file=sys.stderr)


def ensure_template():
    """Generate default template if not bundled."""
    template = PLUGIN_ROOT / "templates" / "default.pptx"
    if not template.exists():
        print("[pptx-writer] Generating default template...", file=sys.stderr)
        subprocess.run(
            [str(VENV_PYTHON), str(PLUGIN_ROOT / "create_template.py")],
            check=True,
        )


def main():
    setup()
    ensure_template()

    # Replace this process with the MCP server
    server_py = str(PLUGIN_ROOT / "server.py")
    if sys.platform == "win32":
        # os.execv doesn't work well on Windows; use subprocess
        result = subprocess.run([str(VENV_PYTHON), server_py])
        sys.exit(result.returncode)
    else:
        os.execv(str(VENV_PYTHON), [str(VENV_PYTHON), server_py])


if __name__ == "__main__":
    main()
