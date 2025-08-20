"""
Validation Manager Tool for hiel_excel_mcp.

Provides comprehensive data validation management including dropdown creation,
number validation, date validation, and validation removal operations.
"""

import json
import logging
from typing import Dict, Any, Optional, List, Union
from datetime import datetime, date

from ..core.base_tool import BaseTool, operation_route, OperationResponse, create_success_response, create_error_response
from ..core.workbook_context import workbook_context

# Import existing functionality
import sys
import os

# Add the src directory to the path to import existing modules
src_path = os.path.join(os.path.dirname(__file__), '..', '..', 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)


logger = logging.getLogger(__name__)


class ValidationManager(BaseTool):
    """
    Comprehensive data validation management tool.
    
    Handles creation and management of Excel data validation rules including
    dropdowns, number validation, date validation, and validation removal.
    """
    
    def get_tool_name(self) -> str:
        return "validation_manager"
    
    def get_tool_description(self) -> str:
        return "Comprehensive data validation management tool for Excel files"
    
    @operation_route(
        name="create_dropdown",
        description="Create dropdown validation with predefined options",
        required_params=["filepath", "sheet_name", "cell_range", "options"],
        optional_params=["input_title", "input_message", "allow_blank"]
    )
    def create_dropdown(self, filepath: str, sheet_name: str, cell_range: str, 
                       options: List[str], input_title: Optional[str] = None,
                       input_message: Optional[str] = None, allow_blank: bool = True,
                       **kwargs) -> OperationResponse:
        """
        Create dropdown validation with predefined options.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            cell_range: Cell range to apply validation (e.g., "A1:A10")
            options: List of dropdown options
            input_title: Title for input prompt
            input_message: Message for input prompt
            allow_blank: Whether to allow blank cells
            
        Returns:
            OperationResponse with dropdown creation results
        """
        try:
            if not options:
                raise ValueError("Options list cannot be empty")
            
            if not isinstance(options, list):
                raise ValueError("Options must be provided as a list")
            
            # Use workbook context for optimization
            with workbook_context(filepath) as wb_ctx:
                result = AdvancedValidationManager.create_dropdown_validation(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    cell_range=cell_range,
                    options=options,
                    input_title=input_title,
                    input_message=input_message,
                    allow_blank=allow_blank
                )
            
            return create_success_response(
                operation="create_dropdown",
                message=f"Dropdown validation created for range {cell_range} with {len(options)} options",
                data={
                    "filepath": filepath,
                    "sheet_name": sheet_name,
                    "cell_range": cell_range,
                    "options": options,
                    "options_count": len(options),
                    "input_title": input_title,
                    "input_message": input_message,
                    "allow_blank": allow_blank
                }
            )
            
        except Exception as e:
            logger.error(f"Failed to create dropdown validation: {e}")
            return create_error_response("create_dropdown", e, {
                "filepath": filepath,
                "sheet_name": sheet_name,
                "cell_range": cell_range
            })
    
    @operation_route(
        name="create_number_validation",
        description="Create numeric validation rule with optional min/max values",
        required_params=["filepath", "sheet_name", "cell_range"],
        optional_params=["min_value", "max_value", "allow_decimals", "allow_blank", "error_message"]
    )
    def create_number_validation(self, filepath: str, sheet_name: str, cell_range: str,
                               min_value: Optional[Union[int, float]] = None,
                               max_value: Optional[Union[int, float]] = None,
                               allow_decimals: bool = True, allow_blank: bool = True,
                               error_message: Optional[str] = None, **kwargs) -> OperationResponse:
        """
        Create numeric validation rule with optional min/max values.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            cell_range: Cell range to apply validation
            min_value: Minimum allowed value
            max_value: Maximum allowed value
            allow_decimals: Whether to allow decimal numbers
            allow_blank: Whether to allow blank cells
            error_message: Custom error message
            
        Returns:
            OperationResponse with number validation results
        """
        try:
            # Validate numeric inputs
            if min_value is not None and max_value is not None:
                if min_value > max_value:
                    raise ValueError("min_value cannot be greater than max_value")
            
            # Use workbook context for optimization
            with workbook_context(filepath) as wb_ctx:
                result = AdvancedValidationManager.create_number_validation(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    cell_range=cell_range,
                    min_value=min_value,
                    max_value=max_value,
                    allow_decimals=allow_decimals,
                    allow_blank=allow_blank,
                    error_message=error_message
                )
            
            return create_success_response(
                operation="create_number_validation",
                message=f"Number validation created for range {cell_range}",
                data={
                    "filepath": filepath,
                    "sheet_name": sheet_name,
                    "cell_range": cell_range,
                    "min_value": min_value,
                    "max_value": max_value,
                    "allow_decimals": allow_decimals,
                    "allow_blank": allow_blank,
                    "error_message": error_message,
                    "validation_type": "decimal" if allow_decimals else "whole"
                }
            )
            
        except Exception as e:
            logger.error(f"Failed to create number validation: {e}")
            return create_error_response("create_number_validation", e, {
                "filepath": filepath,
                "sheet_name": sheet_name,
                "cell_range": cell_range
            })
    
    @operation_route(
        name="create_date_validation",
        description="Create date validation rule with optional start/end dates",
        required_params=["filepath", "sheet_name", "cell_range"],
        optional_params=["start_date", "end_date", "allow_blank", "error_message"]
    )
    def create_date_validation(self, filepath: str, sheet_name: str, cell_range: str,
                             start_date: Optional[str] = None, end_date: Optional[str] = None,
                             allow_blank: bool = True, error_message: Optional[str] = None,
                             **kwargs) -> OperationResponse:
        """
        Create date validation rule with optional start/end dates.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            cell_range: Cell range to apply validation
            start_date: Earliest allowed date (YYYY-MM-DD format)
            end_date: Latest allowed date (YYYY-MM-DD format)
            allow_blank: Whether to allow blank cells
            error_message: Custom error message
            
        Returns:
            OperationResponse with date validation results
        """
        try:
            # Validate date formats if provided
            if start_date:
                try:
                    datetime.strptime(start_date, "%Y-%m-%d")
                except ValueError:
                    raise ValueError(f"start_date must be in YYYY-MM-DD format, got: {start_date}")
            
            if end_date:
                try:
                    datetime.strptime(end_date, "%Y-%m-%d")
                except ValueError:
                    raise ValueError(f"end_date must be in YYYY-MM-DD format, got: {end_date}")
            
            # Validate date range
            if start_date and end_date:
                start_dt = datetime.strptime(start_date, "%Y-%m-%d")
                end_dt = datetime.strptime(end_date, "%Y-%m-%d")
                if start_dt > end_dt:
                    raise ValueError("start_date cannot be after end_date")
            
            # Use workbook context for optimization
            with workbook_context(filepath) as wb_ctx:
                result = AdvancedValidationManager.create_date_validation(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    cell_range=cell_range,
                    start_date=start_date,
                    end_date=end_date,
                    allow_blank=allow_blank,
                    error_message=error_message
                )
            
            return create_success_response(
                operation="create_date_validation",
                message=f"Date validation created for range {cell_range}",
                data={
                    "filepath": filepath,
                    "sheet_name": sheet_name,
                    "cell_range": cell_range,
                    "start_date": start_date,
                    "end_date": end_date,
                    "allow_blank": allow_blank,
                    "error_message": error_message
                }
            )
            
        except Exception as e:
            logger.error(f"Failed to create date validation: {e}")
            return create_error_response("create_date_validation", e, {
                "filepath": filepath,
                "sheet_name": sheet_name,
                "cell_range": cell_range
            })
    
    @operation_route(
        name="remove_validation",
        description="Remove data validation from cells or entire sheet",
        required_params=["filepath", "sheet_name"],
        optional_params=["cell_range"]
    )
    def remove_validation(self, filepath: str, sheet_name: str, 
                         cell_range: Optional[str] = None, **kwargs) -> OperationResponse:
        """
        Remove data validation from cells or entire sheet.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            cell_range: Specific range to remove validation from (None for all)
            
        Returns:
            OperationResponse with validation removal results
        """
        try:
            # Use workbook context for optimization
            with workbook_context(filepath) as wb_ctx:
                result = AdvancedValidationManager.remove_validation(
                    filepath=filepath,
                    sheet_name=sheet_name,
                    cell_range=cell_range
                )
            
            scope = cell_range if cell_range else "entire sheet"
            
            return create_success_response(
                operation="remove_validation",
                message=f"Validation removed from {scope} in sheet '{sheet_name}'",
                data={
                    "filepath": filepath,
                    "sheet_name": sheet_name,
                    "cell_range": cell_range,
                    "scope": scope,
                    "removed_count": result.get("removed_count", 0)
                }
            )
            
        except Exception as e:
            logger.error(f"Failed to remove validation: {e}")
            return create_error_response("remove_validation", e, {
                "filepath": filepath,
                "sheet_name": sheet_name,
                "cell_range": cell_range
            })


# Create global instance
validation_manager = ValidationManager()


def validation_manager_tool(operation: str, **kwargs) -> str:
    """
    MCP tool function for data validation management operations.
    
    Args:
        operation: The operation to perform (create_dropdown, create_number_validation, 
                  create_date_validation, remove_validation)
        **kwargs: Operation-specific parameters
        
    Returns:
        JSON string with operation results
    """
    try:
        response = validation_manager.execute_operation(operation, **kwargs)
        return response.to_json()
    except Exception as e:
        logger.error(f"Unexpected error in validation_manager_tool: {e}", exc_info=True)
        error_response = create_error_response(operation, e)
        return error_response.to_json()