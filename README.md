# Hiel Excel MCP

An optimized version of the Excel MCP server that consolidates 72 individual tools into 12 comprehensive, multi-functional tools for better AI agent interaction.

## Overview

The `hiel_excel_mcp` server reduces cognitive load on AI agents by grouping related Excel operations into logical domains while maintaining all existing functionality and backward compatibility.

## Key Features

- **12 Grouped Tools**: Consolidates 72+ individual operations into intuitive tool groups
- **Backward Compatibility**: All existing functionality preserved
- **Performance Optimized**: Intelligent workbook caching and batch operations
- **Multiple Transports**: Supports stdio, SSE, and streamable HTTP protocols
- **Comprehensive**: Full Excel manipulation without requiring Microsoft Excel

## Installation

```bash
# Install from PyPI (when published)
pip install hiel-excel-mcp

# Or install in development mode
pip install -e .
```

## Usage

### Local Development (Stdio)
```bash
hiel-excel-mcp stdio
```

### Remote Access (HTTP)
```bash
hiel-excel-mcp streamable-http --host 0.0.0.0 --port 8017
```

## Tool Groups

1. **workbook_manager** - Workbook lifecycle and metadata
2. **worksheet_manager** - Worksheet operations and structure  
3. **data_manager** - Data reading, writing, and basic operations
4. **cell_manager** - Cell-level operations and ranges
5. **formatting_manager** - Cell formatting and conditional formatting
6. **formula_manager** - Formula application and validation
7. **analysis_manager** - Charts, pivot tables, and Excel tables
8. **validation_manager** - Data validation and safety operations
9. **import_export_manager** - CSV import/export and format conversion
10. **advanced_manager** - Named ranges, hyperlinks, comments
11. **batch_manager** - Batch operations and templates
12. **system_manager** - Cache management, filtering, sorting

## Environment Variables

- `EXCEL_FILES_PATH`: Base directory for Excel files (default: current directory)
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

## License

MIT License - see LICENSE file for details.