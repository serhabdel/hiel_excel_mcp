# Hiel Excel MCP

An optimized Excel MCP server that provides comprehensive Excel manipulation capabilities through a clean API. The server now includes 25+ powerful tools for advanced Excel operations including tables, pivot tables, advanced formatting, and data manipulation.

## Overview

The `hiel_excel_mcp` server provides AI agents with extensive Excel manipulation capabilities while maintaining performance and reliability. All operations are designed to be non-blocking and thread-safe with intelligent caching for optimal performance.

## Key Features

- **25+ Powerful Tools**: Comprehensive Excel manipulation capabilities
- **Advanced Features**: Tables, pivot tables, advanced formatting, data analysis
- **Performance Optimized**: Thread-safe operations with intelligent workbook caching
- **Backward Compatibility**: All existing functionality preserved with aliases
- **Multiple Transports**: Supports stdio, SSE, and streamable HTTP protocols
- **No Dependencies**: Full Excel manipulation without requiring Microsoft Excel

## Installation

### Prerequisites

- Python 3.8+ (recommended 3.10+)
- pip (package installer for Python)
- openpyxl (Excel manipulation library)

### Option 1: Install from PyPI (when published)

```bash
# Linux/macOS
pip3 install hiel-excel-mcp

# Windows
pip install hiel-excel-mcp
```

### Option 2: Install from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/hiel-excel-mcp.git
cd hiel-excel-mcp

# Linux/macOS
pip3 install -e .

# Windows
pip install -e .
```

## Building and Running the Server

### Option 1: Using Python Directly

#### Linux/macOS

```bash
# Run the server directly
python3 server.py

# Or with stdio transport
python3 -m hiel_excel_mcp stdio

# Or with HTTP transport
python3 -m hiel_excel_mcp streamable-http --host 0.0.0.0 --port 8017
```

#### Windows

```cmd
# Run the server directly
python server.py

# Or with stdio transport
python -m hiel_excel_mcp stdio

# Or with HTTP transport
python -m hiel_excel_mcp streamable-http --host 0.0.0.0 --port 8017
```

### Option 2: Using UVX (for improved performance)

UVX is a high-performance Python runtime that can significantly improve server performance.

#### Linux/macOS

```bash
# Install UVX
pip3 install uvx

# Run the server with UVX
uvx server.py

# Or with stdio transport
uvx -m hiel_excel_mcp stdio

# Or with HTTP transport
uvx -m hiel_excel_mcp streamable-http --host 0.0.0.0 --port 8017
```

#### Windows

Note: UVX may have limited support on Windows. Use Python directly if you encounter issues.

```cmd
# Install UVX
pip install uvx

# Run the server with UVX
uvx server.py
```

### Using with Claude Desktop

1. Configure the `claude_desktop_config.json` file:

```json
{
  "mcpServers": {
    "hiel-excel-mcp": {
      "command": "python3",  // Use "python" on Windows
      "args": [
        "server.py"
      ],
      "disabled": false
    }
  }
}
```

2. Place this file in your Claude Desktop configuration directory
3. Restart Claude Desktop to load the MCP server

## Available Tools

The Excel MCP server provides the following tools, organized by category:

### Workbook Operations
- **`workbook-create`** - Create a new Excel workbook
- **`workbook-metadata`** - Get workbook metadata

### Worksheet Operations
- **`worksheet-create`** - Create new worksheet
- **`worksheet-delete`** - Delete a worksheet from workbook

### Data Operations
- **`data-read`** - Read data from worksheet
- **`data-write`** - Write 2D array data to worksheet
- **`find-replace`** - Find and replace text in worksheet
- **`filter-apply`** - Apply filters to a data range
- **`sort-range`** - Sort data by one or multiple columns

### Cell Operations
- **`cell-write`** - Write value to a single cell
- **`formula-apply`** - Apply a formula to a cell
- **`range-merge`** - Merge cells in a range
- **`range-unmerge`** - Unmerge cells in a range

### Formatting
- **`format-range`** - Apply formatting to a cell range
- **`format-conditional`** - Apply conditional formatting to a range
- **`format-advanced`** - Apply advanced formatting (fonts, borders, fills, alignment, number formats)

### Data Structure
- **`table-create`** - Create an Excel table from a range with auto-filters and formatting
- **`pivot-create`** - Create a pivot table for data analysis
- **`chart-create`** - Create a chart in Excel
- **`named-range-create`** - Create a named range for easy reference

### Row and Column Operations
- **`rows-insert`** - Insert rows at specified position
- **`rows-delete`** - Delete rows at specified position
- **`columns-insert`** - Insert columns at specified position
- **`columns-delete`** - Delete columns at specified position

### Data Validation and Protection
- **`validation-add`** - Add data validation to a range
- **`protection-add`** - Add protection to worksheet or range

### Import/Export
- **`io-export-csv`** - Export Excel data to CSV
- **`io-import-csv`** - Import CSV data to Excel

### System
- **`server-status`** - Get MCP server status and information

## Usage Examples

### Creating and Populating a Workbook

```python
# Create a new workbook
result = await excel_mcp.call_tool("workbook-create", {"filepath": "sales_report.xlsx"})

# Write data to the workbook
data = [
    ["Product", "Q1", "Q2", "Q3", "Q4", "Total"],
    ["Product A", 100, 150, 120, 180, "=SUM(B2:E2)"],
    ["Product B", 200, 210, 190, 220, "=SUM(B3:E3)"],
    ["Product C", 150, 160, 140, 200, "=SUM(B4:E4)"]
]
result = await excel_mcp.call_tool("data-write", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "data": data,
    "start_cell": "A1"
})
```

### Creating Tables and Formatting

```python
# Create a table from the data range
result = await excel_mcp.call_tool("table-create", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "range": "A1:F4",
    "table_name": "SalesData",
    "style": "TableStyleMedium2"
})

# Apply advanced formatting to the header row
result = await excel_mcp.call_tool("format-advanced", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "range": "A1:F1",
    "formatting": {
        "font": {"bold": True, "color": "FFFFFF"},
        "fill": {"color": "4472C4", "type": "solid"},
        "alignment": {"horizontal": "center"}
    }
})
```

### Data Analysis with Pivot Tables

```python
# Create a pivot table for analysis
result = await excel_mcp.call_tool("pivot-create", {
    "filepath": "sales_report.xlsx",
    "source_sheet": "Sales",
    "source_range": "A1:F4",
    "target_sheet": "Analysis",
    "target_cell": "A1",
    "rows": ["Product"],
    "columns": [],
    "values": [{"field": 5, "function": "sum"}],
    "filters": []
})
```

### Data Manipulation

```python
# Sort data by Q4 sales (descending)
result = await excel_mcp.call_tool("sort-range", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "range": "A2:F4",
    "sort_by": [{"column": 4, "ascending": False}]
})

# Find and replace text
result = await excel_mcp.call_tool("find-replace", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "find_text": "Product",
    "replace_text": "Item",
    "match_case": True
})
```

### Protection and Named Ranges

```python
# Create a named range for the totals column
result = await excel_mcp.call_tool("named-range-create", {
    "filepath": "sales_report.xlsx",
    "name": "Totals",
    "sheet_name": "Sales",
    "range": "F2:F4"
})

# Add protection to the worksheet
result = await excel_mcp.call_tool("protection-add", {
    "filepath": "sales_report.xlsx",
    "sheet_name": "Sales",
    "password": "secure123",
    "allow_formatting": True
})
```

## Environment Variables

- `EXCEL_FILES_PATH`: Base directory for Excel files (default: current directory)
- `MAX_ROWS_PER_CALL`: Maximum number of rows allowed per operation (default: 10000)
- `MAX_COLS_PER_CALL`: Maximum number of columns allowed per operation (default: 1000)
- `MAX_FILE_SIZE`: Maximum file size in bytes (default: 50MB)
- `FASTMCP_HOST`: Server host (default: 0.0.0.0)
- `FASTMCP_PORT`: Server port (default: 8017)

## Development

```bash
# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Format code
black hiel_excel_mcp/
isort hiel_excel_mcp/

# Type checking
mypy hiel_excel_mcp/
```

## CI/CD Workflow

This project supports both GitHub Actions and GitLab CI/CD for continuous integration and deployment.

### GitHub Actions

The GitHub workflow includes:

- **Automated Testing**: Tests run on multiple Python versions (3.8-3.11) and operating systems (Ubuntu, Windows)
- **Code Quality Checks**: Linting with flake8, formatting with black, import sorting with isort, and type checking with mypy
- **Test Coverage**: Coverage reports generated and uploaded to Codecov
- **Package Building**: Python package built and verified with twine
- **Docker Image**: Docker image built from the Dockerfile in the deploy directory

To run the workflow manually, go to the Actions tab in the GitHub repository and select "Run workflow" on the "Build and Test Excel MCP" workflow.

### GitLab CI/CD

The GitLab pipeline includes:

- **Staged Pipeline**: Organized into lint, test, build, package, and docker stages
- **Multiple Python Versions**: Tests run on Python 3.8, 3.9, 3.10, and 3.11
- **Code Quality**: Separate jobs for flake8, black, isort, and mypy
- **Artifacts**: Test reports and built packages stored as artifacts
- **Docker Build**: Container image built from the Dockerfile in the deploy directory
- **Caching**: Dependency caching between jobs for faster builds

The pipeline automatically runs on all branches and can be viewed in the CI/CD section of your GitLab repository.

## License

MIT License - see LICENSE file for details.