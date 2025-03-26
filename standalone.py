"""
Standalone version of the MathCAD MCP Server,
updated to cover the full public MathcadPy API.
"""

import os
import logging
import json
import sys
import traceback
from dataclasses import dataclass
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from typing import Dict, List, Optional, Union, Any
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("mathcad-mcp")

# Import MathcadPy - this will fail if it's not installed
from MathcadPy import Mathcad, Worksheet
import win32com.client
logger.info("Successfully imported MathcadPy and COM components")

# MCP import
try:
    from mcp.server.fastmcp import FastMCP, Context, Image
except ImportError:
    sys.exit("Error: MCP package not installed. Please install with 'pip install mcp'")

@dataclass
class MathcadContext:
    app: Mathcad
    worksheets: Dict[str, Worksheet]

@asynccontextmanager
async def mathcad_lifespan(server: FastMCP) -> AsyncIterator[MathcadContext]:
    """Manage Mathcad application lifecycle."""
    logger.info("Starting MathCAD application...")
    try:
        visible = os.environ.get("MATHCAD_VISIBLE", "true").lower() == "true"
        mathcad_app = Mathcad(visible=visible)
        version = mathcad_app.get_version()
        logger.info(f"Successfully started MathCAD version {version}")

        worksheets = {}
        for i, name in enumerate(mathcad_app.worksheet_names()):
            if name in mathcad_app.open_worksheets:
                worksheets[name] = mathcad_app.open_worksheets[name]
                logger.info(f"Found open worksheet: {name}")

        yield MathcadContext(app=mathcad_app, worksheets=worksheets)
    except Exception as e:
        logger.error(f"Error initializing MathCAD: {e}")
        logger.error(traceback.format_exc())
        sys.exit(f"Fatal error: Unable to initialize MathCAD: {e}")
    finally:
        logger.info("Shutting down MathCAD application...")
        try:
            if 'mathcad_app' in locals():
                # Close all without forcibly quitting. Could also call `mathcad_app.quit(...)`.
                mathcad_app.close_all(save_option="Prompt")
                logger.info("MathCAD shutdown complete")
        except Exception as e:
            logger.error(f"Error during MathCAD shutdown: {e}")

mcp = FastMCP("MathcadServer", lifespan=mathcad_lifespan)

# --------------------------------------------------------------------------
# Helper functions
# --------------------------------------------------------------------------

def format_matrix(matrix: List[List[float]], precision: int = 4) -> str:
    result = []
    for row in matrix:
        formatted_row = " ".join([f"{val:.{precision}f}" for val in row])
        result.append(formatted_row)
    return "\n".join(result)

def resolve_path(path: str) -> Path:
    if path.startswith("~"):
        path = os.path.expanduser(path)
    path = os.path.expandvars(path)
    p = Path(path)
    if p.parent and not p.parent.exists():
        p.parent.mkdir(parents=True, exist_ok=True)
    return p

# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------

@mcp.tool()
def list_mathcad_version(ctx: Context) -> str:
    """Get the version of MathCAD currently running."""
    app = ctx.request_context.lifespan_context.app
    try:
        version = app.get_version()
        return f"MathCAD version: {version}"
    except Exception as e:
        return f"Error getting MathCAD version: {e}"

@mcp.tool()
def open_worksheet(ctx: Context, path: str) -> str:
    """Open a Mathcad worksheet at the specified path (.mcdx)."""
    app = ctx.request_context.lifespan_context.app
    worksheets = ctx.request_context.lifespan_context.worksheets

    try:
        if not path.lower().endswith('.mcdx'):
            return f"Error: Path must end with .mcdx extension. Got: {path}"

        full_path = resolve_path(path)
        if not full_path.exists():
            return f"Error: File not found at path: {full_path}"

        worksheet = app.open(full_path)
        if worksheet is None:
            return f"Error: Failed to open worksheet at {full_path}"

        name = worksheet.name()
        worksheets[name] = worksheet
        return f"Successfully opened worksheet: {name}"
    except Exception as e:
        return f"Error opening worksheet: {str(e)}"

@mcp.tool()
def close_worksheet(ctx: Context, worksheet_name: str, save_option: str = "Prompt") -> str:
    """
    Close a specific worksheet.

    Args:
        worksheet_name: Name of the worksheet to close
        save_option: One of 'Discard', 'Prompt', or 'Save'
    """
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"

    try:
        if save_option not in ["Discard", "Prompt", "Save"]:
            return f"Error: save_option must be one of: Discard, Prompt, or Save."

        worksheets[worksheet_name].close(save_option=save_option)
        del worksheets[worksheet_name]
        return f"Successfully closed worksheet '{worksheet_name}' with save_option='{save_option}'"
    except Exception as e:
        return f"Error closing worksheet: {str(e)}"

@mcp.tool()
def close_all_worksheets(ctx: Context, save_option: str = "Prompt") -> str:
    """
    Close all open worksheets.

    Args:
        save_option: One of 'Discard', 'Prompt', or 'Save'
    """
    app = ctx.request_context.lifespan_context.app
    if save_option not in ["Discard", "Prompt", "Save"]:
        return f"Error: save_option must be one of: Discard, Prompt, or Save."

    try:
        app.close_all(save_option=save_option)
        ctx.request_context.lifespan_context.worksheets.clear()
        return f"Successfully closed all worksheets with save_option='{save_option}'"
    except Exception as e:
        return f"Error closing all worksheets: {str(e)}"

@mcp.tool()
def quit_mathcad(ctx: Context, save_option: str = "Discard") -> str:
    """
    Quit the MathCAD application entirely.

    Args:
        save_option: One of 'Discard', 'Prompt', or 'Save'
    """
    app = ctx.request_context.lifespan_context.app
    if save_option not in ["Discard", "Prompt", "Save"]:
        return f"Error: save_option must be one of: Discard, Prompt, or Save."

    try:
        app.quit(save_option=save_option)
        # Clear from context just in case
        ctx.request_context.lifespan_context.worksheets.clear()
        return f"Successfully quit MathCAD with save_option='{save_option}'"
    except Exception as e:
        return f"Error quitting MathCAD: {str(e)}"

@mcp.tool()
def save_worksheet(ctx: Context, worksheet_name: str, path: str, format: str = "mcdx") -> str:
    """Save the worksheet to a specified path in a given format (mcdx, pdf, rtf, xps)."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"

    format = format.lower()
    if format not in ['mcdx', 'pdf', 'rtf', 'xps']:
        return f"Error: Unsupported format '{format}'. Must be one of: mcdx, pdf, rtf, xps."

    full_path = resolve_path(path)
    if not str(full_path).lower().endswith(f'.{format}'):
        full_path = Path(f"{full_path}.{format}")

    try:
        worksheets[worksheet_name].save_as(str(full_path))
        return f"Successfully saved worksheet '{worksheet_name}' to {full_path}"
    except Exception as e:
        return f"Error saving worksheet: {str(e)}"

@mcp.tool()
def save_worksheet_inplace(ctx: Context, worksheet_name: str) -> str:
    """
    Save the worksheet in place (using the current filename).
    """
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        worksheets[worksheet_name].save()
        return f"Successfully saved worksheet '{worksheet_name}' in place."
    except Exception as e:
        return f"Error saving worksheet in place: {str(e)}"

@mcp.tool()
def activate_worksheet(ctx: Context, worksheet_name: str) -> str:
    """Activate (bring to front) a specific worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"

    try:
        worksheets[worksheet_name].activate()
        return f"Successfully activated worksheet: {worksheet_name}"
    except Exception as e:
        return f"Error activating worksheet: {str(e)}"

@mcp.tool()
def set_real_input(ctx: Context, worksheet_name: str, input_name: str,
                   value: float, units: str = "") -> str:
    """Set a numerical input in the specified worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if input_name not in worksheet.inputs():
            return f"Error: Input '{input_name}' not found in worksheet '{worksheet_name}'"

        error = worksheet.set_real_input(input_name, value, units)
        if error > 0:
            return (f"Warning: COM returned error code {error} "
                    f"when setting input '{input_name}' to {value} {units}")
        return f"Successfully set input '{input_name}' to {value} {units}"
    except Exception as e:
        return f"Error setting input: {str(e)}"

@mcp.tool()
def set_string_input(ctx: Context, worksheet_name: str, input_name: str,
                     value: str) -> str:
    """Set a string input in the specified worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if input_name not in worksheet.inputs():
            return f"Error: Input '{input_name}' not found in worksheet '{worksheet_name}'"

        error = worksheet.set_string_input(input_name, value)
        if error > 0:
            return (f"Warning: COM returned error code {error} "
                    f"when setting input '{input_name}' to '{value}'")
        return f"Successfully set input '{input_name}' to '{value}'"
    except Exception as e:
        return f"Error setting input: {str(e)}"

@mcp.tool()
def set_matrix_input(ctx: Context, worksheet_name: str, input_name: str,
                     matrix: str, units: str = "") -> str:
    """Set a matrix input in the specified worksheet (matrix given as JSON)."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        matrix_data = json.loads(matrix)
        if (not isinstance(matrix_data, list)
            or not all(isinstance(row, list) for row in matrix_data)):
            return "Error: Matrix must be a 2D array in JSON format (e.g. [[1,2],[3,4]])"

        if input_name not in worksheet.inputs():
            return f"Error: Input '{input_name}' not found in worksheet '{worksheet_name}'"

        error = worksheet.set_matrix_input(input_name, matrix_data, units)
        if error > 0:
            return (f"Warning: COM returned error code {error} "
                    f"when setting matrix input '{input_name}'")
        return f"Successfully set matrix input '{input_name}'"
    except json.JSONDecodeError:
        return "Error: Invalid JSON format for matrix"
    except Exception as e:
        return f"Error setting matrix input: {str(e)}"

@mcp.tool()
def get_input(ctx: Context, worksheet_name: str, input_name: str) -> str:
    """
    Retrieve the current value of a designated input (it can be real/string/matrix).
    """
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if input_name not in worksheet.inputs():
            return f"Error: Input '{input_name}' not found in worksheet '{worksheet_name}'"

        val, units, error_code = worksheet.get_input(input_name)
        if error_code != 0:
            return f"Warning: Error code {error_code} retrieving input '{input_name}'"

        # If matrix, it will come back as a list of lists
        if isinstance(val, list) and val and isinstance(val[0], list):
            # It's a matrix
            fm = format_matrix(val)
            return f"Input '{input_name}' (matrix):\n{fm}\nUnits: {units}"
        elif isinstance(val, str):
            # It's a string
            return f"Input '{input_name}' (string): '{val}'"
        else:
            # It's a real value
            return f"Input '{input_name}' (real): {val} {units}"
    except Exception as e:
        return f"Error fetching input '{input_name}': {str(e)}"

@mcp.tool()
def get_matrix_input(ctx: Context, worksheet_name: str, input_name: str) -> str:
    """
    Retrieve the current value of a designated matrix input.
    """
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if input_name not in worksheet.inputs():
            return f"Error: Input '{input_name}' not found in worksheet '{worksheet_name}'"

        matrix, units, error_code = worksheet.get_matrix_input(input_name)
        if error_code != 0:
            return f"Warning: Error code {error_code} retrieving matrix input '{input_name}'"

        fm = format_matrix(matrix)
        return f"Matrix input '{input_name}':\n{fm}\nUnits: {units}"
    except Exception as e:
        return f"Error fetching matrix input '{input_name}': {str(e)}"

@mcp.tool()
def calculate_worksheet(ctx: Context, worksheet_name: str) -> str:
    """Synchronize (recalculate) the worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        worksheets[worksheet_name].calculate()
        return f"Successfully recalculated worksheet: {worksheet_name}"
    except Exception as e:
        return f"Error recalculating worksheet: {str(e)}"

@mcp.tool()
def sync_worksheet(ctx: Context, worksheet_name: str) -> str:
    """Alias for calculating the worksheet (syncronize)."""
    return calculate_worksheet(ctx, worksheet_name)

@mcp.tool()
def pause_calculation(ctx: Context, worksheet_name: str) -> str:
    """Pause the calculation of a worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        worksheets[worksheet_name].pause_calculation()
        return f"Successfully paused calculation for worksheet: {worksheet_name}"
    except Exception as e:
        return f"Error pausing calculation: {str(e)}"

@mcp.tool()
def resume_calculation(ctx: Context, worksheet_name: str) -> str:
    """Resume the calculation of a worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        worksheets[worksheet_name].resume_calculation()
        return f"Successfully resumed calculation for worksheet: {worksheet_name}"
    except Exception as e:
        return f"Error resuming calculation: {str(e)}"

@mcp.tool()
def get_output(ctx: Context, worksheet_name: str, output_name: str) -> str:
    """Get a designated output value (real, string, or matrix)."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        value, units, error_code = worksheet.get_output(output_name)
        if error_code > 0:
            return f"Warning: Error code {error_code} when getting output '{output_name}'"

        if isinstance(value, list) and value and isinstance(value[0], list):
            # It's a matrix
            fm = format_matrix(value)
            return f"Output '{output_name}' (matrix):\n{fm}\nUnits: {units}"
        elif isinstance(value, str):
            # It's a string
            return f"Output '{output_name}' (string): '{value}'"
        else:
            # It's numeric
            return f"Output '{output_name}' = {value} {units}"
    except Exception as e:
        return f"Error getting output: {str(e)}"

@mcp.tool()
def get_real_output(ctx: Context, worksheet_name: str, output_name: str) -> str:
    """Get a numerical output value from a worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        value, units, error_code = worksheet.get_real_output(output_name)
        if error_code > 0:
            return f"Warning: Error code {error_code} when getting output '{output_name}'"
        return f"Output '{output_name}' = {value} {units}"
    except Exception as e:
        return f"Error getting real output: {str(e)}"

@mcp.tool()
def get_real_output_with_units(ctx: Context, worksheet_name: str, output_name: str, units: str) -> str:
    """Get a numerical output value with conversion to specified units."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        value, result_units, error_code = worksheet.get_real_output(output_name, units)
        if error_code > 0:
            return (f"Warning: Error code {error_code} when getting output '{output_name}' "
                    f"in units '{units}'")
        return f"Output '{output_name}' = {value} {result_units}"
    except Exception as e:
        return f"Error getting output with units conversion: {str(e)}"

@mcp.tool()
def get_matrix_output(ctx: Context, worksheet_name: str, output_name: str) -> str:
    """Get a matrix output from a worksheet."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        matrix, units, error_code = worksheet.get_matrix_output(output_name)
        if error_code > 0:
            return f"Warning: Error code {error_code} when getting matrix output '{output_name}'"
        fm = format_matrix(matrix)
        return f"Matrix output '{output_name}':\n{fm}\nUnits: {units}"
    except Exception as e:
        return f"Error getting matrix output: {str(e)}"

@mcp.tool()
def get_matrix_output_with_units(ctx: Context, worksheet_name: str, output_name: str, units: str) -> str:
    """Get a matrix output from a worksheet with unit conversion."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        matrix, result_units, error_code = worksheet.get_matrix_output(output_name, units)
        if error_code > 0:
            return (f"Warning: Error code {error_code} when getting matrix output '{output_name}' "
                    f"in units '{units}'")
        fm = format_matrix(matrix)
        return f"Matrix output '{output_name}':\n{fm}\nUnits: {result_units}"
    except Exception as e:
        return f"Error getting matrix output with units: {str(e)}"

@mcp.tool()
def save_as_pdf(ctx: Context, worksheet_name: str, output_path: str) -> str:
    """Save the worksheet as a PDF file (Mathcad Prime 5+ only)."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    app = ctx.request_context.lifespan_context.app

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    full_path = resolve_path(output_path)
    if not str(full_path).lower().endswith('.pdf'):
        full_path = Path(f"{full_path}.pdf")

    try:
        if hasattr(app, 'version_major_int') and app.version_major_int < 5:
            return "Error: PDF export requires Mathcad Prime 5 or newer"
        worksheet.save_as(str(full_path))
        return f"Successfully saved worksheet '{worksheet_name}' as PDF to {full_path}"
    except Exception as e:
        return f"Error saving as PDF: {str(e)}"

@mcp.tool()
def save_as_rtf(ctx: Context, worksheet_name: str, output_path: str) -> str:
    """Save the worksheet as an RTF file."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    full_path = resolve_path(output_path)
    if not str(full_path).lower().endswith('.rtf'):
        full_path = Path(f"{full_path}.rtf")

    try:
        worksheet.save_as(str(full_path))
        return f"Successfully saved worksheet '{worksheet_name}' as RTF to {full_path}"
    except Exception as e:
        return f"Error saving as RTF: {str(e)}"

@mcp.tool()
def save_as_xps(ctx: Context, worksheet_name: str, output_path: str) -> str:
    """Save the worksheet as an XPS file."""
    worksheets = ctx.request_context.lifespan_context.worksheets

    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    full_path = resolve_path(output_path)
    if not str(full_path).lower().endswith('.xps'):
        full_path = Path(f"{full_path}.xps")

    try:
        worksheet.save_as(str(full_path))
        return f"Successfully saved worksheet '{worksheet_name}' as XPS to {full_path}"
    except Exception as e:
        return f"Error saving as XPS: {str(e)}"

@mcp.tool()
def is_worksheet_readonly(ctx: Context, worksheet_name: str) -> str:
    """Check if a worksheet is read-only."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        ro = worksheets[worksheet_name].is_readonly()
        return f"Worksheet '{worksheet_name}' read-only status: {ro}"
    except Exception as e:
        return f"Error checking read-only status: {str(e)}"

@mcp.tool()
def is_worksheet_modified(ctx: Context, worksheet_name: str) -> str:
    """Check if a worksheet has unsaved changes."""
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        mod = worksheets[worksheet_name].is_modified()
        return f"Worksheet '{worksheet_name}' modified status: {mod}"
    except Exception as e:
        return f"Error checking modified status: {str(e)}"

@mcp.tool()
def set_worksheet_modified(ctx: Context, worksheet_name: str, state: bool = True) -> str:
    """
    Force-set the 'modified' status of a worksheet (e.g., mark as changed).
    """
    worksheets = ctx.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        worksheets[worksheet_name].is_modified(setmodified=state)
        return f"Set worksheet '{worksheet_name}' modified status to {state}"
    except Exception as e:
        return f"Error setting modified status: {str(e)}"

# --------------------------------------------------------------------------
# Resources
# --------------------------------------------------------------------------

@mcp.resource("mathcad://worksheets")
def list_worksheets() -> str:
    """List all open worksheets in MathCAD."""
    app = mcp.request_context.lifespan_context.app
    try:
        names = app.worksheet_names()
        if not names:
            return "No worksheets are currently open in MathCAD."
        return "Open worksheets:\n" + "\n".join(f"- {n}" for n in names)
    except Exception as e:
        return f"Error listing worksheets: {str(e)}"

@mcp.resource("mathcad://worksheets/paths")
def list_worksheet_paths() -> str:
    """List all open worksheets in MathCAD with their full paths."""
    app = mcp.request_context.lifespan_context.app
    try:
        wpaths = app.worksheet_paths()
        wnames = app.worksheet_names()
        if not wpaths:
            return "No worksheets are currently open in MathCAD."
        result = "Open worksheets with paths:\n"
        for i, p in enumerate(wpaths):
            name = wnames[i] if i < len(wnames) else "Unknown"
            result += f"- {name}: {p}\n"
        return result
    except Exception as e:
        return f"Error listing worksheet paths: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/inputs")
def get_inputs(worksheet_name: str) -> str:
    """Get a list of all inputs in the specified worksheet."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        inp = worksheets[worksheet_name].inputs()
        if not inp:
            return f"No designated inputs found in worksheet '{worksheet_name}'."
        return (f"Inputs in worksheet '{worksheet_name}':\n" +
                "\n".join(f"- {x}" for x in inp))
    except Exception as e:
        return f"Error listing inputs: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/outputs")
def get_outputs(worksheet_name: str) -> str:
    """Get a list of all outputs in the specified worksheet."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    try:
        outs = worksheets[worksheet_name].outputs()
        if not outs:
            return f"No designated outputs found in worksheet '{worksheet_name}'."
        return (f"Outputs in worksheet '{worksheet_name}':\n" +
                "\n".join(f"- {o}" for o in outs))
    except Exception as e:
        return f"Error listing outputs: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/output/{output_name}")
def get_output_value(worksheet_name: str, output_name: str) -> str:
    """Get the value of a specific output from a worksheet."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        # Try get_output (handles real/string/matrix)
        value, outunits, error_code = worksheet.get_output(output_name)
        if error_code != 0:
            return f"Warning: Error code {error_code} when getting output '{output_name}'"
        if isinstance(value, list) and value and isinstance(value[0], list):
            fm = format_matrix(value)
            return f"Output '{output_name}' (matrix):\n{fm}\nUnits: {outunits}"
        elif isinstance(value, str):
            return f"Output '{output_name}' (string): '{value}'"
        else:
            return f"Output '{output_name}' = {value} {outunits}"
    except Exception as e:
        return f"Error getting output value: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/output/{output_name}/matrix")
def get_matrix_output_resource(worksheet_name: str, output_name: str) -> str:
    """Get a matrix output from a worksheet as formatted text."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        matrix, outunits, error_code = worksheet.get_matrix_output(output_name)
        if error_code != 0:
            return f"Warning: Error code {error_code} when getting matrix output '{output_name}'"
        fm = format_matrix(matrix)
        return f"Matrix output '{output_name}':\n{fm}\nUnits: {outunits}"
    except Exception as e:
        return f"Error getting matrix output: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/output/{output_name}/with_units/{units}")
def get_output_with_units(worksheet_name: str, output_name: str, units: str) -> str:
    """Get the value of a specific output with units conversion."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"
    worksheet = worksheets[worksheet_name]

    try:
        if output_name not in worksheet.outputs():
            return f"Error: Output '{output_name}' not found in worksheet '{worksheet_name}'"
        # Try numeric output with conversion
        val, outunits, error_code = worksheet.get_real_output(output_name, units)
        if error_code == 0:
            return f"Output '{output_name}' = {val} {outunits}"
        # else fallback
        val2, outunits2, err2 = worksheet.get_output(output_name)
        return (f"Output '{output_name}' = {val2} {outunits2} "
                f"(unit conversion not possible, code {error_code})")
    except Exception as e:
        return f"Error: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/info")
def get_worksheet_info(worksheet_name: str) -> str:
    """Get info about a specific worksheet."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"

    ws = worksheets[worksheet_name]
    try:
        ro = ws.is_readonly()
        mod = ws.is_modified()
        inps = len(ws.inputs())
        outs = len(ws.outputs())
        info = (
            f"Worksheet: {worksheet_name}\n"
            f"Read-only: {ro}\n"
            f"Modified: {mod}\n"
            f"Input count: {inps}\n"
            f"Output count: {outs}"
        )
        return info
    except Exception as e:
        return f"Error getting worksheet info: {str(e)}"

@mcp.resource("mathcad://worksheet/{worksheet_name}/status")
def get_worksheet_status(worksheet_name: str) -> str:
    """Get the status of a worksheet (read-only and modified)."""
    worksheets = mcp.request_context.lifespan_context.worksheets
    if worksheet_name not in worksheets:
        return f"Error: Worksheet '{worksheet_name}' not found or not open"

    ws = worksheets[worksheet_name]
    try:
        ro = ws.is_readonly()
        mod = ws.is_modified()
        return (f"Worksheet '{worksheet_name}' status:\n"
                f"Read-only: {ro}\n"
                f"Modified: {mod}")
    except Exception as e:
        return f"Error getting worksheet status: {str(e)}"

# --------------------------------------------------------------------------
# Prompts
# --------------------------------------------------------------------------

@mcp.prompt()
def solve_equation(equation: str, variables: str) -> str:
    return f"""Please solve this equation using MathCAD:
Equation: {equation}
Variables to solve for: {variables}

Use MathCAD's symbolic solver if possible.
"""

@mcp.prompt()
def plot_function(function: str, x_range: str) -> str:
    return f"""Please create a plot of this function in MathCAD:
Function: {function}
X-axis range: {x_range}"""

@mcp.prompt()
def analyze_engineering_problem(problem_description: str, relevant_equations: str) -> str:
    return f"""Solve this engineering problem in MathCAD:

Problem:
{problem_description}

Equations:
{relevant_equations}
"""

@mcp.prompt()
def unit_conversion(value: str, from_unit: str, to_unit: str) -> str:
    return f"""Convert value using MathCAD:
Value: {value} {from_unit}
Convert to: {to_unit}
Explain the steps."""

@mcp.prompt()
def create_worksheet_template(title: str, description: str, fields: str) -> str:
    return f"""Create a new MathCAD worksheet template:
Title: {title}
Description: {description}
Required Fields: {fields}
"""

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
if __name__ == "__main__":
    mcp.run(transport='stdio')
