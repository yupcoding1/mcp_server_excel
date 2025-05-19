import os
from dotenv import load_dotenv
from excel_fucntion import *
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp import Context
from pydantic import Field

load_dotenv()
EXCEL_FILES_DIR = os.getenv("EXCEL_FILES_DIR", "./excel_files")

mcp = FastMCP("Excel MCP Server", dependencies=["openpyxl", "python-dotenv"])

# Resource: List all Excel files
def list_excel_files() -> list[str]:
    files = [f for f in os.listdir(EXCEL_FILES_DIR) if f.endswith(".xlsx") or f.endswith(".xlsm")]
    return files

@mcp.resource("excel-files://list")
def resource_list_excel_files() -> list[str]:
    return list_excel_files()

@mcp.resource("excel-sheetnames://{filename}")
def resource_list_sheets(filename: str) -> list[str]:
    return list_sheets(os.path.join(EXCEL_FILES_DIR, filename))

@mcp.tool()
def tool_create_excel_file(
    filename: str = Field(description="The name of the Excel file to create"),
    sheet_name: str = Field(description="The name of the initial sheet", default="Sheet1")
) -> str:
    """Create a new Excel file with an initial sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return create_excel_file(path, sheet_name)

@mcp.tool()
def tool_add_sheet(
    filename: str = Field(description="The Excel file to add a sheet to"),
    sheet_name: str = Field(description="The name of the new sheet")
) -> str:
    """Add a new sheet to an existing Excel file."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return add_sheet(path, sheet_name)

@mcp.tool()
def tool_rename_sheet(
    filename: str = Field(description="The Excel file containing the sheet"),
    old_name: str = Field(description="The current name of the sheet"),
    new_name: str = Field(description="The new name for the sheet")
) -> str:
    """Rename a sheet in an Excel file."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return rename_sheet(path, old_name, new_name)

@mcp.tool()
def tool_delete_sheet(
    filename: str = Field(description="The Excel file to delete a sheet from"),
    sheet_name: str = Field(description="The name of the sheet to delete")
) -> str:
    """Delete a sheet from an Excel file."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return delete_sheet(path, sheet_name)

@mcp.tool()
def tool_write_cell(
    filename: str = Field(description="The Excel file to write to"),
    sheet: str = Field(description="The sheet to write to"),
    cell: str = Field(description="The cell address (e.g. A1)"),
    value: str = Field(description="The value to write")
) -> str:
    """Write a value to a specific cell in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return write_cell(path, sheet, cell, value)

@mcp.tool()
def tool_read_cell(
    filename: str = Field(description="The Excel file to read from"),
    sheet: str = Field(description="The sheet to read from"),
    cell: str = Field(description="The cell address (e.g. A1)")
) -> str:
    """Read the value from a specific cell in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return str(read_cell(path, sheet, cell))

@mcp.tool()
def tool_merge_cells(
    filename: str = Field(description="The Excel file to modify"),
    sheet: str = Field(description="The sheet to merge cells in"),
    cell_range: str = Field(description="The range of cells to merge (e.g. A1:B2)")
) -> str:
    """Merge a range of cells in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return merge_cells(path, sheet, cell_range)

@mcp.tool()
def tool_unmerge_cells(
    filename: str = Field(description="The Excel file to modify"),
    sheet: str = Field(description="The sheet to unmerge cells in"),
    cell_range: str = Field(description="The range of cells to unmerge (e.g. A1:B2)")
) -> str:
    """Unmerge a range of cells in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return unmerge_cells(path, sheet, cell_range)

@mcp.tool()
def tool_write_row(
    filename: str = Field(description="The Excel file to write to"),
    sheet: str = Field(description="The sheet to write to"),
    start_cell: str = Field(description="The starting cell address for the row (e.g. A1)"),
    data: list = Field(description="The list of values to write in the row")
) -> str:
    """Write a row of values starting at a specific cell in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return write_row(path, sheet, start_cell, data)

@mcp.tool()
def tool_write_column(
    filename: str = Field(description="The Excel file to write to"),
    sheet: str = Field(description="The sheet to write to"),
    start_cell: str = Field(description="The starting cell address for the column (e.g. A1)"),
    data: list = Field(description="The list of values to write in the column")
) -> str:
    """Write a column of values starting at a specific cell in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return write_column(path, sheet, start_cell, data)

@mcp.tool()
def tool_set_border(
    filename: str = Field(description="The Excel file to modify"),
    sheet: str = Field(description="The sheet to set borders in"),
    cell_range: str = Field(description="The range of cells to set borders for (e.g. A1:B2)")
) -> str:
    """Set borders for a range of cells in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return set_border(path, sheet, cell_range)

@mcp.tool()
def tool_auto_fit_columns(
    filename: str = Field(description="The Excel file to modify"),
    sheet: str = Field(description="The sheet to auto-fit columns in")
) -> str:
    """Auto-fit the width of all columns in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return auto_fit_columns(path, sheet)

@mcp.tool()
def tool_get_used_range(
    filename: str = Field(description="The Excel file to inspect"),
    sheet: str = Field(description="The sheet to get the used range from")
) -> dict:
    """Get the used range (min/max row/col) of an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return get_used_range(path, sheet)

@mcp.tool()
def tool_read_range(
    filename: str = Field(description="The Excel file to read from"),
    sheet: str = Field(description="The sheet to read from"),
    cell_range: str = Field(description="The range of cells to read (e.g. A1:B2)")
) -> list:
    """Read a range of cells from an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return read_range(path, sheet, cell_range)

@mcp.tool()
def tool_write_formula(
    filename: str = Field(description="The Excel file to modify"),
    sheet: str = Field(description="The sheet to write the formula in"),
    cell: str = Field(description="The cell address to write the formula to (e.g. A1)"),
    formula: str = Field(description="The formula to write (without the leading =)")
) -> str:
    """Write a formula to a specific cell in an Excel sheet."""
    path = os.path.join(EXCEL_FILES_DIR, filename)
    return write_formula(path, sheet, cell, formula)

@mcp.tool()
def tool_save_as_new_file(
    old_filename: str = Field(description="The original Excel file name"),
    new_filename: str = Field(description="The new Excel file name")
) -> str:
    """Save the Excel file as a new file with a different name."""
    old_path = os.path.join(EXCEL_FILES_DIR, old_filename)
    new_path = os.path.join(EXCEL_FILES_DIR, new_filename)
    return save_as_new_file(old_path, new_path)

@mcp.tool()
def greet_user(
    name: str = Field(description="The name of the person to greet"),
    title: str = Field(description="Optional title like Mr/Ms/Dr", default=""),
    times: int = Field(description="Number of times to repeat the greeting", default=1),
) -> str:
    """Greet a user with optional title and repetition"""
    greeting = f"Hello {title + ' ' if title else ''}{name}!"
    return "\n".join([greeting] * times)

if __name__ == "__main__":
    os.makedirs(EXCEL_FILES_DIR, exist_ok=True)
    mcp.run(transport="stdio")
