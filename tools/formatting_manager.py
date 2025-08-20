"""
Formatting Manager Tool for hiel_excel_mcp.

Provides comprehensive formatting operations including basic cell formatting,
conditional formatting rules, and advanced styling capabilities.
"""

import json
import logging
from typing import Dict, Any, Optional, Union

from ..core.base_tool import BaseTool, operation_route, OperationResponse, create_success_response, create_error_response
from ..core.workbook_context import workbook_context

# Import existing functionality
import sys
import os

# Add the src directory to the path to import existing modules
src_path = os.path.join(os.path.dirname(__file__), '..', '..', 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

    add_conditional_formatting as add_cf,
    remove_conditional_formatting as remove_cf,
    list_conditional_formatting as list_cf,
    create_highlight_cells_rule as create_highlight
)

logger = logging.getLogger(__name__)


class FormattingManager(BaseTool):
    """
    Comprehensive formatting management tool.
    
    Handles all formatting operations including basic cell formatting,
    conditional formatting rules, and advanced styling capabilities.
    """
    
    def get_tool_name(self) -> str:
        return "formatting_manager"
    
    def get_tool_description(self) -> str:
        return "Comprehensive cell formatting and conditional formatting management tool"
    
    @operation_route(
        name="apply_formatting",
        description="Apply basic formatting to a cell range",
        required_params=["filepath", "sheet_name", "start_cell"],
        optional_params=[
            "end_cell", "bold", "italic", "underline", "font_size", "font_color",
            "bg_color", "border_style", "border_color", "number_format", "alignment",
            "wrap_text", "merge_cells", "protection"
        ]
    )
    def apply_formatting(
        self,
        filepath: str,
        sheet_name: str,
        start_cell: str,
        end_cell: Optional[str] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: Optional[int] = None,
        font_color: Optional[str] = None,
        bg_color: Optional[str] = None,
        border_style: Optional[str] = None,
        border_color: Optional[str] = None,
        number_format: Optional[str] = None,
        alignment: Optional[str] = None,
        wrap_text: bool = False,
        merge_cells: bool = False,
        protection: Optional[Dict[str, Any]] = None,
        **kwargs
    ) -> OperationResponse:
        """
        Apply basic formatting to a cell range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            start_cell: Starting cell reference
            end_cell: Optional ending cell reference
            bold: Whether to make text bold
            italic: Whether to make text italic
            underline: Whether to underline text
            font_size: Font size in points
            font_color: Font color (hex code)
            bg_color: Background color (hex code)
            border_style: Border style (thin, medium, thick, double)
            border_color: Border color (hex code)
            number_format: Excel number format string
            alignment: Text alignment (left, center, right, justify)
            wrap_text: Whether to wrap text
            merge_cells: Whether to merge the range
            protection: Cell protection settings
            
        Returns:
            OperationResponse with formatting results
        """
        try:
            # Apply formatting using existing functionality
            result = format_range(
                filepath=filepath,
                sheet_name=sheet_name,
                start_cell=start_cell,
                end_cell=end_cell,
                bold=bold,
                italic=italic,
                underline=underline,
                font_size=font_size,
                font_color=font_color,
                bg_color=bg_color,
                border_style=border_style,
                border_color=border_color,
                number_format=number_format,
                alignment=alignment,
                wrap_text=wrap_text,
                merge_cells=merge_cells,
                protection=protection
            )
            
            return create_success_response(
                operation="apply_formatting",
                message=f"Applied formatting to range {result.get('range', start_cell)}",
                data={
                    "filepath": filepath,
                    "sheet_name": sheet_name,
                    "range": result.get("range", start_cell),
                    "formatting_applied": {
                        "bold": bold,
                        "italic": italic,
                        "underline": underline,
                        "font_size": font_size,
                        "font_color": font_color,
                        "bg_color": bg_color,
                        "border_style": border_style,
                        "border_color": border_color,
                        "number_format": number_format,
                        "alignment": alignment,
                        "wrap_text": wrap_text,
                        "merge_cells": merge_cells
                    }
                }
            )
            
        except Exception as e:
            logger.error(f"Failed to apply formatting: {e}")
            return create_error_response("apply_formatting", e)
    
    @operation_route(
        name="add_conditional_formatting",
        description="Add conditional formatting rule to a range",
        required_params=["filepath", "sheet_name", "range_ref", "rule_config"]
    )
    def add_conditional_formatting(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: str,
        rule_config: Dict[str, Any],
        **kwargs
    ) -> OperationResponse:
        """
        Add conditional formatting rule to a range.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            range_ref: Cell range for conditional formatting
            rule_config: Configuration for the formatting rule
            
        Returns:
            OperationResponse with conditional formatting results
        """
        try:
            # Add conditional formatting using existing functionality
            result_json = add_cf(filepath, sheet_name, range_ref, rule_config)
            result = json.loads(result_json)
            
            if result.get("success"):
                return create_success_response(
                    operation="add_conditional_formatting",
                    message=f"Added conditional formatting to range {range_ref}",
                    data={
                        "filepath": filepath,
                        "sheet_name": sheet_name,
                        "range": range_ref,
                        "rule_type": rule_config.get("type"),
                        "rule_config": rule_config
                    }
                )
            else:
                raise Exception(result.get("message", "Failed to add conditional formatting"))
            
        except Exception as e:
            logger.error(f"Failed to add conditional formatting: {e}")
            return create_error_response("add_conditional_formatting", e)
    
    @operation_route(
        name="remove_conditional_formatting",
        description="Remove conditional formatting from range or sheet",
        required_params=["filepath", "sheet_name"],
        optional_params=["range_ref"]
    )
    def remove_conditional_formatting(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: Optional[str] = None,
        **kwargs
    ) -> OperationResponse:
        """
        Remove conditional formatting from range or entire sheet.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            range_ref: Specific range (None for entire sheet)
            
        Returns:
            OperationResponse with removal results
        """
        try:
            # Remove conditional formatting using existing functionality
            result_json = remove_cf(filepath, sheet_name, range_ref)
            result = json.loads(result_json)
            
            if result.get("success"):
                return create_success_response(
                    operation="remove_conditional_formatting",
                    message=f"Removed conditional formatting from {range_ref or 'entire sheet'}",
                    data={
                        "filepath": filepath,
                        "sheet_name": sheet_name,
                        "range": range_ref or "entire sheet",
                        "removed_count": result.get("removed_count", 0)
                    }
                )
            else:
                raise Exception(result.get("message", "Failed to remove conditional formatting"))
            
        except Exception as e:
            logger.error(f"Failed to remove conditional formatting: {e}")
            return create_error_response("remove_conditional_formatting", e)
    
    @operation_route(
        name="list_conditional_formatting",
        description="List all conditional formatting rules in a sheet",
        required_params=["filepath", "sheet_name"]
    )
    def list_conditional_formatting(
        self,
        filepath: str,
        sheet_name: str,
        **kwargs
    ) -> OperationResponse:
        """
        List all conditional formatting rules in a sheet.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            
        Returns:
            OperationResponse with formatting rules information
        """
        try:
            # List conditional formatting using existing functionality
            result_json = list_cf(filepath, sheet_name)
            result = json.loads(result_json)
            
            if result.get("success"):
                return create_success_response(
                    operation="list_conditional_formatting",
                    message=f"Listed {result.get('total_rules', 0)} conditional formatting rules",
                    data={
                        "filepath": filepath,
                        "sheet_name": sheet_name,
                        "total_rules": result.get("total_rules", 0),
                        "conditional_formatting": result.get("conditional_formatting", [])
                    }
                )
            else:
                raise Exception(result.get("message", "Failed to list conditional formatting"))
            
        except Exception as e:
            logger.error(f"Failed to list conditional formatting: {e}")
            return create_error_response("list_conditional_formatting", e)
    
    @operation_route(
        name="create_highlight_rule",
        description="Create a simple cell highlighting rule",
        required_params=["filepath", "sheet_name", "range_ref", "operator", "value"],
        optional_params=["highlight_color"]
    )
    def create_highlight_rule(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: str,
        operator: str,
        value: Union[str, int, float],
        highlight_color: str = "FFFF00",
        **kwargs
    ) -> OperationResponse:
        """
        Create a simple cell highlighting rule.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            range_ref: Cell range for formatting
            operator: Comparison operator (equal, greater_than, etc.)
            value: Value to compare against
            highlight_color: Background color for highlighting
            
        Returns:
            OperationResponse with highlight rule results
        """
        try:
            # Create highlight rule using existing functionality
            result_json = create_highlight(
                filepath, sheet_name, range_ref, operator, value, highlight_color
            )
            result = json.loads(result_json)
            
            if result.get("success"):
                return create_success_response(
                    operation="create_highlight_rule",
                    message=f"Created highlight rule for range {range_ref}",
                    data={
                        "filepath": filepath,
                        "sheet_name": sheet_name,
                        "range": range_ref,
                        "operator": operator,
                        "value": value,
                        "highlight_color": highlight_color,
                        "rule_type": "cell_is"
                    }
                )
            else:
                raise Exception(result.get("message", "Failed to create highlight rule"))
            
        except Exception as e:
            logger.error(f"Failed to create highlight rule: {e}")
            return create_error_response("create_highlight_rule", e)


# Create global instance
formatting_manager = FormattingManager()


def formatting_manager_tool(operation: str, **kwargs) -> str:
    """
    MCP tool function for formatting management operations.
    
    Args:
        operation: The operation to perform
        **kwargs: Operation-specific parameters
        
    Returns:
        JSON string with operation results
    """
    try:
        response = formatting_manager.execute_operation(operation, **kwargs)
        return response.to_json()
    except Exception as e:
        logger.error(f"Unexpected error in formatting_manager_tool: {e}", exc_info=True)
        error_response = create_error_response(operation, e)
        return error_response.to_json()