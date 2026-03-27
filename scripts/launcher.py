# -*- coding: utf-8 -*-
"""
PPTX Writer MCP Server Launcher

Installs dependencies on first run (pip install --target),
then runs the MCP server directly in the same process.
No venv needed - uses PYTHONPATH for package resolution.
"""

import importlib
import os
import subprocess
import sys
from pathlib import Path

PLUGIN_ROOT = Path(os.environ.get("CLAUDE_PLUGIN_ROOT", Path(__file__).parent.parent))
PLUGIN_DATA = Path(os.environ.get("CLAUDE_PLUGIN_DATA", PLUGIN_ROOT / ".plugin_data"))
LIB_DIR = PLUGIN_DATA / "lib"

# Ensure lib dir is on sys.path
sys.path.insert(0, str(LIB_DIR))
sys.path.insert(0, str(PLUGIN_ROOT / "src"))


def install_deps():
    """Install dependencies using pip --target (no venv needed)."""
    print("[pptx-writer] Installing dependencies (first run)...", file=sys.stderr)
    LIB_DIR.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            sys.executable, "-m", "pip", "install",
            "--target", str(LIB_DIR),
            "--quiet",
            "-r", str(PLUGIN_ROOT / "requirements.txt"),
        ],
        check=True,
    )
    # Write marker so we can skip next time
    marker = PLUGIN_DATA / ".installed"
    marker.write_text(
        (PLUGIN_ROOT / "requirements.txt").read_text(encoding="utf-8"),
        encoding="utf-8",
    )
    importlib.invalidate_caches()
    print("[pptx-writer] Dependencies installed.", file=sys.stderr)


def needs_install():
    """Fast check: are dependencies already installed?"""
    marker = PLUGIN_DATA / ".installed"
    if not marker.exists():
        return True
    try:
        current = (PLUGIN_ROOT / "requirements.txt").read_text(encoding="utf-8")
        installed = marker.read_text(encoding="utf-8")
        return current != installed
    except Exception:
        return True


def main():
    # Install deps if needed
    if needs_install():
        install_deps()

    # Generate default template if not bundled
    template = PLUGIN_ROOT / "templates" / "default.pptx"
    if not template.exists():
        print("[pptx-writer] Generating default template...", file=sys.stderr)
        subprocess.run(
            [sys.executable, str(PLUGIN_ROOT / "create_template.py")],
            env={**os.environ, "PYTHONPATH": str(LIB_DIR)},
            check=True,
        )

    # Run server.py directly in this process (no subprocess overhead)
    server_path = str(PLUGIN_ROOT / "server.py")
    server_code = compile(
        open(server_path, encoding="utf-8").read(),
        server_path,
        "exec",
    )
    exec(server_code, {"__name__": "__main__", "__file__": server_path})


if __name__ == "__main__":
    main()
