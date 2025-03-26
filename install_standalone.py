"""
Utility script to install the standalone MathCAD MCP server in Claude Desktop.
"""

import os
import json
import sys
import argparse
from pathlib import Path

def get_config_path():
    """Get the path to the Claude Desktop configuration file."""
    if sys.platform == "darwin":  # macOS
        return Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
    elif sys.platform == "win32":  # Windows
        return Path(os.environ.get("APPDATA")) / "Claude" / "claude_desktop_config.json"
    else:
        raise RuntimeError(f"Unsupported platform: {sys.platform}")

def install_server(show_mathcad=True):
    """Install the MathCAD MCP server in Claude Desktop."""
    config_path = get_config_path()
    
    # Create directory if it doesn't exist
    config_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Load existing config or create new one
    if config_path.exists():
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
        except json.JSONDecodeError:
            print(f"Warning: Existing config at {config_path} is invalid. Creating new config.")
            config = {}
    else:
        config = {}
    
    # Get absolute path to the standalone script
    server_dir = Path(__file__).resolve().parent
    standalone_script = server_dir / "standalone.py"
    
    # Set up MCP server configuration
    if "mcpServers" not in config:
        config["mcpServers"] = {}
    
    # Get Python executable from the virtual environment if active
    python_exe = sys.executable
    
    config["mcpServers"]["mathcad"] = {
        "command": python_exe,
        "args": [
            str(standalone_script)  # Direct path to standalone.py
        ],
        "env": {
            "MATHCAD_VISIBLE": "true" if show_mathcad else "false"
        }
    }
    
    # Write config back to file
    with open(config_path, 'w') as f:
        json.dump(config, f, indent=2)
    
    print(f"Successfully installed MathCAD MCP server to Claude Desktop config at: {config_path}")
    print(f"Using Python: {python_exe}")
    print(f"Using script: {standalone_script}")
    print("Please restart Claude Desktop for the changes to take effect.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Install MathCAD MCP server in Claude Desktop")
    parser.add_argument("--hide", action="store_true", help="Hide MathCAD window when running")
    args = parser.parse_args()
    
    install_server(not args.hide)
