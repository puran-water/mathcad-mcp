# MathCAD MCP Server

## Standalone Windows integration

This repository connects an MCP client to an installed, licensed MathCAD Prime application through Windows COM. It is a separate integration, not a portable replacement for MathCAD and not part of the six-server PuranOS calculation mirror. A client on another operating system still needs a supported Windows execution arrangement.

Current PuranOS [proposal and BEP workflows](https://github.com/puran-water/PuranOS-public/blob/main/docs/workflows.md) use shared engineering libraries and retained calculation evidence. Their operation does not depend on installing this MathCAD connector. See the [public package catalogue](https://github.com/puran-water/PuranOS-public/blob/main/docs/public-tools.md) for local engineering tools with their own supported runtimes.

Published by [Puran Water](https://puranwater.com/). The API and application requirements below apply to this integration.


A Model Context Protocol (MCP) server that provides integration with PTC MathCAD Prime, allowing AI assistants like Claude to interact with MathCAD worksheets through a standardized interface.

## Features

- **Worksheet Management**: Open, close, and save MathCAD worksheets
- **Input Control**: Set numerical values, strings, and matrices with units
- **Output Access**: Retrieve calculation results with proper units
- **Calculation Management**: Pause, resume, and recalculate worksheets
- **Export Options**: Save worksheets as PDF, RTF, and XPS files

## Requirements

- Windows operating system (required for MathCAD COM automation)
- Python 3.10 or higher
- PTC MathCAD Prime 3.0 or higher installed
- Dependencies: MathcadPy, pywin32, and MCP Python SDK (see requirements.txt)

## Installation

### Quick Installation with Claude Desktop

1. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the installation script:
   ```
   python install_standalone.py
   ```
   
   Options:
   - Use `--hide` flag to hide the MathCAD window when running: `python install_standalone.py --hide`

3. Restart Claude Desktop for the changes to take effect.

### Manual Configuration

Configure the Windows MCP client, for example `%APPDATA%\Claude\claude_desktop_config.json` for Claude Desktop. The server process must have access to Windows COM and the installed MathCAD application:

```json
{
  "mcpServers": {
    "mathcad": {
      "command": "python",
      "args": [
        "path/to/mathcad-mcp/standalone.py"
      ],
      "env": {
        "MATHCAD_VISIBLE": "true"
      }
    }
  }
}
```

## Usage

### Tools

The server exposes the following MCP tools:

- `list_mathcad_version`: Get the running MathCAD version
- `open_worksheet`: Open a MathCAD worksheet from a file path
- `close_worksheet`: Close a specific worksheet
- `close_all_worksheets`: Close all open worksheets
- `quit_mathcad`: Quit the MathCAD application entirely
- `save_worksheet`: Save a worksheet to a specified path and format
- `save_worksheet_inplace`: Save the worksheet in place
- `activate_worksheet`: Bring a worksheet to the front
- `set_real_input`: Set a numerical input with optional units
- `set_string_input`: Set a string input
- `set_matrix_input`: Set a matrix input with optional units
- `get_input`: Retrieve the current value of a designated input (real/string/matrix)
- `get_matrix_input`: Retrieve the current value of a designated matrix input
- `calculate_worksheet`: Recalculate a worksheet
- `sync_worksheet`: Alias for calculating the worksheet
- `pause_calculation`: Pause worksheet calculation
- `resume_calculation`: Resume worksheet calculation
- `get_output`: Get a designated output value (real, string, or matrix)
- `get_real_output`: Get a numerical output value from a worksheet
- `get_real_output_with_units`: Get a numerical output value with conversion to specified units
- `get_matrix_output`: Get a matrix output from a worksheet
- `get_matrix_output_with_units`: Get a matrix output with unit conversion
- `save_as_pdf`: Save the worksheet as a PDF file (Mathcad Prime 5+ only)
- `save_as_rtf`: Save the worksheet as an RTF file
- `save_as_xps`: Save the worksheet as an XPS file
- `is_worksheet_readonly`: Check if a worksheet is read-only
- `is_worksheet_modified`: Check if a worksheet has unsaved changes
- `set_worksheet_modified`: Force-set the 'modified' status of a worksheet

### Resources

The server exposes the following MCP resources:

- `mathcad://worksheets`: List all open worksheets
- `mathcad://worksheets/paths`: List worksheets with full file paths
- `mathcad://worksheet/{worksheet_name}/inputs`: List inputs in a worksheet
- `mathcad://worksheet/{worksheet_name}/outputs`: List outputs in a worksheet
- `mathcad://worksheet/{worksheet_name}/output/{output_name}`: Get a specific output value
- `mathcad://worksheet/{worksheet_name}/output/{output_name}/matrix`: Get a matrix output
- `mathcad://worksheet/{worksheet_name}/output/{output_name}/with_units/{units}`: Get an output with units conversion
- `mathcad://worksheet/{worksheet_name}/info`: Get information about a worksheet
- `mathcad://worksheet/{worksheet_name}/status`: Get the status of a worksheet

### Prompt Templates

The server provides the following prompt templates to assist in MathCAD-related tasks:

- `solve_equation`: Template for solving mathematical equations
- `plot_function`: Template for plotting functions
- `analyze_engineering_problem`: Template for analyzing engineering problems
- `unit_conversion`: Template for unit conversion tasks
- `create_worksheet_template`: Template for creating new worksheet templates

## Example Workflows

### Example 1: Solving a Physics Problem

```
1. Open an existing MathCAD worksheet:
   "Can you open my physics worksheet at C:\Users\username\Documents\Physics\projectile_motion.mcdx?"

2. Set input parameters:
   "Set the initial velocity to 50 m/s and the launch angle to 30 degrees."

3. Calculate the results:
   "Now calculate the trajectory and show me the maximum height and distance."

4. Extract output values:
   "What's the maximum height reached by the projectile?"
```

### Example 2: Engineering Analysis

```
1. Open or create a worksheet for analysis:
   "I need to analyze a beam deflection problem. Can you open the structural_analysis.mcdx worksheet?"

2. Configure the worksheet:
   "Set the load P to 5000 N, beam length L to 2 m, Young's modulus E to 200 GPa, and moment of inertia I to 4×10⁻⁶ m⁴."

3. Run the calculation:
   "Calculate the worksheet and tell me the maximum deflection."

4. Export the results:
   "Save the worksheet as a PDF so I can include it in my report."
```

## Environment Variables

- `MATHCAD_VISIBLE`: Set to "true" to show the MathCAD application window (default), or "false" to hide it.

## Architecture

The server is built using the MCP Python SDK's FastMCP framework. The main components are:

- `standalone.py`: Self-contained server script with MathCAD integration
- `install_standalone.py`: Utility script to install the server in Claude Desktop

## Troubleshooting

### Common Issues

#### COM Automation Errors

If you encounter COM automation errors:
- Ensure MathCAD is installed and properly licensed
- Check that you're running on Windows
- Try restarting MathCAD and the MCP server

#### Worksheet Not Found

If worksheets can't be found:
- Use absolute paths, not relative paths
- Ensure file paths are correctly escaped
- Verify the file exists and is a valid `.mcdx` file

#### Unit Conversion Issues

If unit conversion fails:
- Check for compatibility between units (e.g., length can't convert to mass)
- Ensure that the right unit symbols are used (e.g., 'm' for meters)
- Try using standard unit formats (e.g., 'm/s' instead of 'm*s^(-1)')

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [MathcadPy](https://github.com/MattWoodhead/MathcadPy) - Python wrapper for the MathCAD Prime automation API
- [Model Context Protocol](https://modelcontextprotocol.io) - Protocol specification and SDK
