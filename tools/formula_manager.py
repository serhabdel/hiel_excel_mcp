"""
Formula Manager Tool for hiel_excel_mcp.

Provides comprehensive formula operations including applying formulas,
validating formula syntax, and batch formula operations.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, List

from ..core.base_tool import BaseTool, operation_route, OperationResponse, create_success_response, create_error_response
from ..core.workbook_context import workbook_context, _global_cache

# Import existing functionality
import sys
import os

# Add the src directory to the path to import existing modules
src_path = os.path.join(os.path.dirname(__file__), '..', '..', 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)


logger = logging.getLogger(__name__)


class FormulaManager(BaseTool):
    """
    Comprehensive formula management tool.
    
    Handles formula operations including applying formulas, validating syntax,
    and batch formula operations.
    """
    
    def get_tool_name(self) -> str:
        return "formula_manager"
    
    def get_tool_description(self) -> str:
        return "Comprehensive formula operations and validation management tool"
    
    @operation_route(
        name="apply_formula",
        description="Apply a formula to a specific cell",
        required_params=["filepath", "sheet_name", "cell", "formula"]
    )
    def apply_formula(self, filepath: str, sheet_name: str, cell: str, 
                     formula: str, **kwargs) -> OperationResponse:
        """
        Apply a formula to a specific cell.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            cell: Cell address (e.g., "A1")
            formula: Formula to apply (with or without leading =)
            
        Returns:
            OperationResponse with application results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=True)
            
            # Apply formula using existing functionality
            result = apply_formula(validated_path, sheet_name, cell, formula)
            
            return create_success_response(
                operation="apply_formula",
                message=result["message"],
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "cell": result["cell"],
                    "formula": result["formula"],
                    "applied_successfully": True
                },
                warnings=warnings if warnings else None
            )
            
        except (ValidationError, CalculationError) as e:
            logger.error(f"Failed to apply formula: {e}")
            return create_error_response("apply_formula", e)
        except Exception as e:
            logger.error(f"Unexpected error applying formula: {e}")
            return create_error_response("apply_formula", e)
    
    @operation_route(
        name="validate_formula",
        description="Validate formula syntax and safety",
        required_params=["formula"],
        optional_params=["filepath", "sheet_name", "cell"]
    )
    def validate_formula(self, formula: str, filepath: Optional[str] = None,
                        sheet_name: Optional[str] = None, cell: Optional[str] = None,
                        **kwargs) -> OperationResponse:
        """
        Validate formula syntax and safety.
        
        Args:
            formula: Formula to validate (with or without leading =)
            filepath: Optional path to Excel file for context validation
            sheet_name: Optional worksheet name for context validation
            cell: Optional cell address for context validation
            
        Returns:
            OperationResponse with validation results
        """
        try:
            # Basic formula validation
            is_valid, message = validate_formula(formula)
            
            validation_result = {
                "formula": formula,
                "is_valid": is_valid,
                "validation_message": message,
                "syntax_check": "passed" if is_valid else "failed"
            }
            
            # If file context provided, do additional validation
            if filepath and sheet_name and cell:
                try:
                    validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
                    
                    # Validate formula in context
                    context_result = validate_formula_in_cell_operation(
                        validated_path, sheet_name, cell, formula
                    )
                    
                    validation_result.update({
                        "context_validation": context_result,
                        "filepath": validated_path,
                        "sheet_name": sheet_name,
                        "cell": cell
                    })
                    
                except Exception as context_error:
                    validation_result["context_validation_error"] = str(context_error)
                    logger.warning(f"Context validation failed: {context_error}")
            
            status_message = (
                f"Formula validation {'passed' if is_valid else 'failed'}: {message}"
            )
            
            if is_valid:
                return create_success_response(
                    operation="validate_formula",
                    message=status_message,
                    data=validation_result
                )
            else:
                return OperationResponse(
                    success=False,
                    operation="validate_formula",
                    message=status_message,
                    data=validation_result,
                    errors=[message]
                )
            
        except Exception as e:
            logger.error(f"Failed to validate formula: {e}")
            return create_error_response("validate_formula", e)
    
    @operation_route(
        name="batch_apply_formulas",
        description="Apply multiple formulas to different cells in batch",
        required_params=["filepath", "sheet_name", "formulas"]
    )
    def batch_apply_formulas(self, filepath: str, sheet_name: str,
                           formulas: List[Dict[str, str]], **kwargs) -> OperationResponse:
        """
        Apply multiple formulas to different cells in batch.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            formulas: List of formula specifications, each containing 'cell' and 'formula' keys
            
        Returns:
            OperationResponse with batch application results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=True)
            
            if not formulas:
                raise ValueError("No formulas provided for batch operation")
            
            # Validate formula specifications
            for i, formula_spec in enumerate(formulas):
                if not isinstance(formula_spec, dict):
                    raise ValueError(f"Formula specification {i} must be a dictionary")
                if 'cell' not in formula_spec or 'formula' not in formula_spec:
                    raise ValueError(f"Formula specification {i} must contain 'cell' and 'formula' keys")
            
            results = []
            errors = []
            successful_applications = 0
            
            # Use workbook context for efficient batch operations
            context = _global_cache.get_context(validated_path)
            with context as wb:
                if sheet_name not in wb.sheetnames:
                    raise ValidationError(f"Sheet '{sheet_name}' not found")
                
                ws = wb[sheet_name]
                
                for i, formula_spec in enumerate(formulas):
                    cell_ref = formula_spec['cell']
                    formula = formula_spec['formula']
                    
                    try:
                        # Validate cell reference
                        if not validate_cell_reference(cell_ref):
                            raise ValidationError(f"Invalid cell reference: {cell_ref}")
                        
                        # Validate formula
                        is_valid, validation_message = validate_formula(formula)
                        if not is_valid:
                            raise ValidationError(f"Invalid formula syntax: {validation_message}")
                        
                        # Ensure formula starts with =
                        if not formula.startswith('='):
                            formula = f'={formula}'
                        
                        # Apply formula to cell
                        cell_obj = ws[cell_ref]
                        cell_obj.value = formula
                        
                        results.append({
                            "index": i,
                            "cell": cell_ref,
                            "formula": formula,
                            "status": "success",
                            "message": f"Formula applied to {cell_ref}"
                        })
                        successful_applications += 1
                        
                        # Mark context as dirty to ensure save
                        context.mark_dirty()
                        
                    except Exception as cell_error:
                        error_msg = f"Failed to apply formula to {cell_ref}: {str(cell_error)}"
                        results.append({
                            "index": i,
                            "cell": cell_ref,
                            "formula": formula,
                            "status": "error",
                            "error": error_msg
                        })
                        errors.append(error_msg)
                        logger.error(error_msg)
            
            # Determine overall success
            total_formulas = len(formulas)
            has_errors = len(errors) > 0
            
            message = (
                f"Batch formula application completed: {successful_applications}/{total_formulas} successful"
            )
            
            response_data = {
                "filepath": validated_path,
                "sheet_name": sheet_name,
                "total_formulas": total_formulas,
                "successful_applications": successful_applications,
                "failed_applications": len(errors),
                "results": results
            }
            
            if has_errors:
                return OperationResponse(
                    success=successful_applications > 0,  # Partial success if some worked
                    operation="batch_apply_formulas",
                    message=message,
                    data=response_data,
                    errors=errors,
                    warnings=warnings if warnings else None
                )
            else:
                return create_success_response(
                    operation="batch_apply_formulas",
                    message=message,
                    data=response_data,
                    warnings=warnings if warnings else None
                )
            
        except Exception as e:
            logger.error(f"Failed to apply formulas in batch: {e}")
            return create_error_response("batch_apply_formulas", e)


# Create global instance
formula_manager = FormulaManager()


def formula_manager_tool(operation: str, **kwargs) -> str:
    """
    MCP tool function for formula management operations.
    
    Args:
        operation: The operation to perform
        **kwargs: Operation-specific parameters
        
    Returns:
        JSON string with operation results
    """
    try:
        response = formula_manager.execute_operation(operation, **kwargs)
        return response.to_json()
    except Exception as e:
        logger.error(f"Unexpected error in formula_manager_tool: {e}", exc_info=True)
        error_response = create_error_response(operation, e)
        return error_response.to_json()