"""
Test package structure and basic imports.
"""

import pytest
import sys
import os

# Add the package to the path for testing
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))


def test_package_import():
    """Test that the main package can be imported."""
    import hiel_excel_mcp
    assert hiel_excel_mcp.__version__ == "0.1.0"
    assert "Optimized Excel MCP Server" in hiel_excel_mcp.__description__


def test_server_import():
    """Test that the server module can be imported."""
    from hiel_excel_mcp.server import app
    assert app is not None


def test_cli_import():
    """Test that the CLI module can be imported."""
    from hiel_excel_mcp.__main__ import main
    assert callable(main)


def test_tools_package():
    """Test that the tools package exists."""
    import hiel_excel_mcp.tools
    # Package should exist and be importable


def test_core_package():
    """Test that the core package exists."""
    import hiel_excel_mcp.core
    # Package should exist and be importable