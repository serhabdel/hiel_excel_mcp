"""
Worksheet Manager Tool for hiel_excel_mcp.

Provides comprehensive worksheet management including creation, copying, deletion,
renaming, cell merging operations, and validation information retrieval.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, List

from ..core.base_tool import BaseTool, operation_route, OperationResponse, create_success_response, create_error_response
from ..core.workbook_context import workbook_context

# Import existing functionality
import sys
import os

# Add the src directory to the path to import existing modules
src_path = os.path.join(os.path.dirname(__file__), '..', '..', 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

    copy_sheet, delete_sheet, rename_sheet, merge_range, unmerge_range, 
    get_merged_ranges
)

logger = logging.getLogger(__name__)


class WorksheetManager(BaseTool):
    """
    Comprehensive worksheet management tool.
    
    Handles worksheet operations including creation, copying, deletion, renaming,
    cell merging operations, and validation information retrieval.
    """
    
    def get_tool_name(self) -> str:
        return "worksheet_manager"
    
    def get_tool_description(self) -> str:
        return "Comprehensive worksheet management and operations tool"
    
    @operation_route(
        name="create",
        description="Create a new worksheet in the workbook",
        required_params=["filepath", "sheet_name"]
    )
    def create(self, filepath: str, sheet_name: str, **kwargs) -> OperationResponse:
        """
        Create a new worksheet in the workbook.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the new worksheet
            
        Returns:
            OperationResponse with creation results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Create worksheet using existing functionality
            result = create_sheet(validated_path, sheet_name)
            
            return create_success_response(
                operation="create",
                message=f"Worksheet '{sheet_name}' created successfully in {validated_path}",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to create worksheet {sheet_name} in {filepath}: {e}")
            return create_error_response("create", e)
    
    @operation_route(
        name="copy",
        description="Copy a worksheet within the same workbook",
        required_params=["filepath", "source_sheet", "target_sheet"]
    )
    def copy(self, filepath: str, source_sheet: str, target_sheet: str, **kwargs) -> OperationResponse:
        """
        Copy a worksheet within the same workbook.
        
        Args:
            filepath: Path to the Excel file
            source_sheet: Name of the source worksheet to copy
            target_sheet: Name of the new worksheet
            
        Returns:
            OperationResponse with copy results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Copy worksheet using existing functionality
            result = copy_sheet(validated_path, source_sheet, target_sheet)
            
            return create_success_response(
                operation="copy",
                message=f"Worksheet '{source_sheet}' copied to '{target_sheet}' in {validated_path}",
                data={
                    "filepath": validated_path,
                    "source_sheet": source_sheet,
                    "target_sheet": target_sheet,
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to copy worksheet {source_sheet} to {target_sheet} in {filepath}: {e}")
            return create_error_response("copy", e)
    
    @operation_route(
        name="delete",
        description="Delete a worksheet from the workbook",
        required_params=["filepath", "sheet_name"]
    )
    def delete(self, filepath: str, sheet_name: str, **kwargs) -> OperationResponse:
        """
        Delete a worksheet from the workbook.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet to delete
            
        Returns:
            OperationResponse with deletion results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Delete worksheet using existing functionality
            result = delete_sheet(validated_path, sheet_name)
            
            return create_success_response(
                operation="delete",
                message=f"Worksheet '{sheet_name}' deleted from {validated_path}",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to delete worksheet {sheet_name} from {filepath}: {e}")
            return create_error_response("delete", e)
    
    @operation_route(
        name="rename",
        description="Rename a worksheet",
        required_params=["filepath", "old_name", "new_name"]
    )
    def rename(self, filepath: str, old_name: str, new_name: str, **kwargs) -> OperationResponse:
        """
        Rename a worksheet.
        
        Args:
            filepath: Path to the Excel file
            old_name: Current name of the worksheet
            new_name: New name for the worksheet
            
        Returns:
            OperationResponse with rename results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Rename worksheet using existing functionality
            result = rename_sheet(validated_path, old_name, new_name)
            
            return create_success_response(
                operation="rename",
                message=f"Worksheet renamed from '{old_name}' to '{new_name}' in {validated_path}",
                data={
                    "filepath": validated_path,
                    "old_name": old_name,
                    "new_name": new_name,
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to rename worksheet {old_name} to {new_name} in {filepath}: {e}")
            return create_error_response("rename", e)
    
    @operation_route(
        name="get_merged_cells",
        description="Get all merged cell ranges in a worksheet",
        required_params=["filepath", "sheet_name"]
    )
    def get_merged_cells(self, filepath: str, sheet_name: str, **kwargs) -> OperationResponse:
        """
        Get all merged cell ranges in a worksheet.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            
        Returns:
            OperationResponse with merged cell ranges
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Get merged ranges using existing functionality
            merged_ranges = get_merged_ranges(validated_path, sheet_name)
            
            return create_success_response(
                operation="get_merged_cells",
                message=f"Retrieved {len(merged_ranges)} merged cell ranges from worksheet '{sheet_name}'",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "merged_ranges": merged_ranges,
                    "count": len(merged_ranges)
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to get merged cells from {sheet_name} in {filepath}: {e}")
            return create_error_response("get_merged_cells", e)
    
    @operation_route(
        name="merge_cells",
        description="Merge a range of cells",
        required_params=["filepath", "sheet_name", "start_cell", "end_cell"]
    )
    def merge_cells(self, filepath: str, sheet_name: str, start_cell: str, end_cell: str, **kwargs) -> OperationResponse:
        """
        Merge a range of cells.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            start_cell: Starting cell of the range (e.g., "A1")
            end_cell: Ending cell of the range (e.g., "C3")
            
        Returns:
            OperationResponse with merge results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Merge cells using existing functionality
            result = merge_range(validated_path, sheet_name, start_cell, end_cell)
            
            return create_success_response(
                operation="merge_cells",
                message=f"Cells merged from {start_cell} to {end_cell} in worksheet '{sheet_name}'",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "start_cell": start_cell,
                    "end_cell": end_cell,
                    "range": f"{start_cell}:{end_cell}",
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to merge cells {start_cell}:{end_cell} in {sheet_name} of {filepath}: {e}")
            return create_error_response("merge_cells", e)
    
    @operation_route(
        name="unmerge_cells",
        description="Unmerge a range of cells",
        required_params=["filepath", "sheet_name", "start_cell", "end_cell"]
    )
    def unmerge_cells(self, filepath: str, sheet_name: str, start_cell: str, end_cell: str, **kwargs) -> OperationResponse:
        """
        Unmerge a range of cells.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            start_cell: Starting cell of the range (e.g., "A1")
            end_cell: Ending cell of the range (e.g., "C3")
            
        Returns:
            OperationResponse with unmerge results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Unmerge cells using existing functionality
            result = unmerge_range(validated_path, sheet_name, start_cell, end_cell)
            
            return create_success_response(
                operation="unmerge_cells",
                message=f"Cells unmerged from {start_cell} to {end_cell} in worksheet '{sheet_name}'",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "start_cell": start_cell,
                    "end_cell": end_cell,
                    "range": f"{start_cell}:{end_cell}",
                    "message": result.get("message", "")
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to unmerge cells {start_cell}:{end_cell} in {sheet_name} of {filepath}: {e}")
            return create_error_response("unmerge_cells", e)
    
    @operation_route(
        name="get_validation_info",
        description="Get validation information for a range in the worksheet",
        required_params=["filepath", "sheet_name", "start_cell"],
        optional_params=["end_cell"]
    )
    def get_validation_info(self, filepath: str, sheet_name: str, start_cell: str, 
                          end_cell: Optional[str] = None, **kwargs) -> OperationResponse:
        """
        Get validation information for a range in the worksheet.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            start_cell: Starting cell of the range (e.g., "A1")
            end_cell: Ending cell of the range (optional)
            
        Returns:
            OperationResponse with validation information
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Get validation info using existing functionality
            validation_info = validate_range_in_sheet_operation(
                validated_path, sheet_name, start_cell, end_cell
            )
            
            return create_success_response(
                operation="get_validation_info",
                message=f"Retrieved validation information for range in worksheet '{sheet_name}'",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "start_cell": start_cell,
                    "end_cell": end_cell,
                    "validation_info": validation_info
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to get validation info for range in {sheet_name} of {filepath}: {e}")
            return create_error_response("get_validation_info", e)


# Create global instance
worksheet_manager = WorksheetManager()


def worksheet_manager_tool(operation: str, **kwargs) -> str:
    """
    MCP tool function for worksheet management operations.
    
    Args:
        operation: The operation to perform
        **kwargs: Operation-specific parameters
        
    Returns:
        JSON string with operation results
    """
    try:
        response = worksheet_manager.execute_operation(operation, **kwargs)
        return response.to_json()
    except Exception as e:
        logger.error(f"Unexpected error in worksheet_manager_tool: {e}", exc_info=True)
        error_response = create_error_response(operation, e)
        return error_response.to_json()