"""
Validation system for hiel_excel_mcp operations.

Provides parameter validation, type checking, and data validation utilities
for all grouped tools.
"""

import os
import re
from typing import Any, Dict, List, Optional, Union, Type
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class ValidationError(Exception):
    """Raised when validation fails."""
    pass


class ParameterValidator:
    """Validates parameters for tool operations."""
    
    @staticmethod
    def validate_filepath(filepath: str, must_exist: bool = True, 
                         allowed_extensions: Optional[List[str]] = None) -> str:
        """
        Validate file path parameter.
        
        Args:
            filepath: File path to validate
            must_exist: Whether file must exist
            allowed_extensions: List of allowed file extensions
            
        Returns:
            str: Validated file path
            
        Raises:
            ValidationError: If validation fails
        """
        if not filepath or not isinstance(filepath, str):
            raise ValidationError("filepath must be a non-empty string")
        
        # Convert to Path object for validation
        path = Path(filepath)
        
        # Check if file must exist
        if must_exist and not path.exists():
            raise ValidationError(f"File does not exist: {filepath}")
        
        # Check file extension if specified
        if allowed_extensions:
            if path.suffix.lower() not in [ext.lower() for ext in allowed_extensions]:
                raise ValidationError(
                    f"Invalid file extension. Allowed: {', '.join(allowed_extensions)}"
                )
        
        return str(path.resolve())
    
    @staticmethod
    def validate_sheet_name(sheet_name: str) -> str:
        """
        Validate worksheet name parameter.
        
        Args:
            sheet_name: Sheet name to validate
            
        Returns:
            str: Validated sheet name
            
        Raises:
            ValidationError: If validation fails
        """
        if not sheet_name or not isinstance(sheet_name, str):
            raise ValidationError("sheet_name must be a non-empty string")
        
        # Excel sheet name restrictions
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            if char in sheet_name:
                raise ValidationError(
                    f"Sheet name contains invalid character '{char}'. "
                    f"Invalid characters: {', '.join(invalid_chars)}"
                )
        
        if len(sheet_name) > 31:
            raise ValidationError("Sheet name cannot exceed 31 characters")
        
        return sheet_name
    
    @staticmethod
    def validate_cell_reference(cell_ref: str) -> str:
        """
        Validate Excel cell reference.
        
        Args:
            cell_ref: Cell reference to validate (e.g., "A1", "B5")
            
        Returns:
            str: Validated cell reference
            
        Raises:
            ValidationError: If validation fails
        """
        if not cell_ref or not isinstance(cell_ref, str):
            raise ValidationError("cell_ref must be a non-empty string")
        
        # Excel cell reference pattern
        pattern = r'^[A-Z]+[1-9]\d*$'
        if not re.match(pattern, cell_ref.upper()):
            raise ValidationError(
                f"Invalid cell reference format: {cell_ref}. "
                "Expected format: A1, B5, AA10, etc."
            )
        
        return cell_ref.upper()
    
    @staticmethod
    def validate_range_reference(range_ref: str) -> str:
        """
        Validate Excel range reference.
        
        Args:
            range_ref: Range reference to validate (e.g., "A1:B5")
            
        Returns:
            str: Validated range reference
            
        Raises:
            ValidationError: If validation fails
        """
        if not range_ref or not isinstance(range_ref, str):
            raise ValidationError("range_ref must be a non-empty string")
        
        # Handle single cell reference
        if ':' not in range_ref:
            return ParameterValidator.validate_cell_reference(range_ref)
        
        # Handle range reference
        parts = range_ref.split(':')
        if len(parts) != 2:
            raise ValidationError(
                f"Invalid range format: {range_ref}. Expected format: A1:B5"
            )
        
        start_cell = ParameterValidator.validate_cell_reference(parts[0])
        end_cell = ParameterValidator.validate_cell_reference(parts[1])
        
        return f"{start_cell}:{end_cell}"
    
    @staticmethod
    def validate_operation_name(operation: str, valid_operations: List[str]) -> str:
        """
        Validate operation name parameter.
        
        Args:
            operation: Operation name to validate
            valid_operations: List of valid operation names
            
        Returns:
            str: Validated operation name
            
        Raises:
            ValidationError: If validation fails
        """
        if not operation or not isinstance(operation, str):
            raise ValidationError("operation must be a non-empty string")
        
        if operation not in valid_operations:
            raise ValidationError(
                f"Invalid operation: {operation}. "
                f"Valid operations: {', '.join(valid_operations)}"
            )
        
        return operation
    
    @staticmethod
    def validate_type(value: Any, expected_type: Type, param_name: str) -> Any:
        """
        Validate parameter type.
        
        Args:
            value: Value to validate
            expected_type: Expected type
            param_name: Parameter name for error messages
            
        Returns:
            Any: Validated value
            
        Raises:
            ValidationError: If validation fails
        """
        if not isinstance(value, expected_type):
            raise ValidationError(
                f"Parameter '{param_name}' must be of type {expected_type.__name__}, "
                f"got {type(value).__name__}"
            )
        
        return value
    
    @staticmethod
    def validate_choice(value: Any, choices: List[Any], param_name: str) -> Any:
        """
        Validate parameter is one of allowed choices.
        
        Args:
            value: Value to validate
            choices: List of allowed choices
            param_name: Parameter name for error messages
            
        Returns:
            Any: Validated value
            
        Raises:
            ValidationError: If validation fails
        """
        if value not in choices:
            raise ValidationError(
                f"Parameter '{param_name}' must be one of: {', '.join(map(str, choices))}, "
                f"got: {value}"
            )
        
        return value
    
    @staticmethod
    def validate_range(value: Union[int, float], min_val: Optional[Union[int, float]] = None,
                      max_val: Optional[Union[int, float]] = None, param_name: str = "value") -> Union[int, float]:
        """
        Validate numeric parameter is within range.
        
        Args:
            value: Value to validate
            min_val: Minimum allowed value
            max_val: Maximum allowed value
            param_name: Parameter name for error messages
            
        Returns:
            Union[int, float]: Validated value
            
        Raises:
            ValidationError: If validation fails
        """
        if min_val is not None and value < min_val:
            raise ValidationError(
                f"Parameter '{param_name}' must be >= {min_val}, got: {value}"
            )
        
        if max_val is not None and value > max_val:
            raise ValidationError(
                f"Parameter '{param_name}' must be <= {max_val}, got: {value}"
            )
        
        return value


class DataValidator:
    """Validates data content and structure."""
    
    @staticmethod
    def validate_data_structure(data: List[List[Any]], min_rows: int = 1, 
                               min_cols: int = 1) -> List[List[Any]]:
        """
        Validate data structure for writing to Excel.
        
        Args:
            data: Data to validate
            min_rows: Minimum number of rows required
            min_cols: Minimum number of columns required
            
        Returns:
            List[List[Any]]: Validated data
            
        Raises:
            ValidationError: If validation fails
        """
        if not isinstance(data, list):
            raise ValidationError("Data must be a list of lists")
        
        if len(data) < min_rows:
            raise ValidationError(f"Data must have at least {min_rows} rows")
        
        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ValidationError(f"Row {i} must be a list")
            
            if len(row) < min_cols:
                raise ValidationError(f"Row {i} must have at least {min_cols} columns")
        
        return data
    
    @staticmethod
    def validate_formula(formula: str) -> str:
        """
        Validate Excel formula syntax.
        
        Args:
            formula: Formula to validate
            
        Returns:
            str: Validated formula
            
        Raises:
            ValidationError: If validation fails
        """
        if not formula or not isinstance(formula, str):
            raise ValidationError("Formula must be a non-empty string")
        
        # Ensure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Basic validation - check for balanced parentheses
        open_parens = formula.count('(')
        close_parens = formula.count(')')
        
        if open_parens != close_parens:
            raise ValidationError("Formula has unbalanced parentheses")
        
        return formula


def validate_common_parameters(**kwargs) -> Dict[str, Any]:
    """
    Validate common parameters used across multiple operations.
    
    Args:
        **kwargs: Parameters to validate
        
    Returns:
        Dict[str, Any]: Validated parameters
        
    Raises:
        ValidationError: If validation fails
    """
    validated = {}
    
    # Validate filepath if present
    if 'filepath' in kwargs:
        validated['filepath'] = ParameterValidator.validate_filepath(
            kwargs['filepath'], 
            allowed_extensions=['.xlsx', '.xlsm', '.xls']
        )
    
    # Validate sheet_name if present
    if 'sheet_name' in kwargs:
        validated['sheet_name'] = ParameterValidator.validate_sheet_name(
            kwargs['sheet_name']
        )
    
    # Validate cell references if present
    if 'cell' in kwargs:
        validated['cell'] = ParameterValidator.validate_cell_reference(
            kwargs['cell']
        )
    
    if 'start_cell' in kwargs:
        validated['start_cell'] = ParameterValidator.validate_cell_reference(
            kwargs['start_cell']
        )
    
    if 'end_cell' in kwargs:
        validated['end_cell'] = ParameterValidator.validate_cell_reference(
            kwargs['end_cell']
        )
    
    # Validate range references if present
    if 'range_ref' in kwargs:
        validated['range_ref'] = ParameterValidator.validate_range_reference(
            kwargs['range_ref']
        )
    
    # Copy other parameters as-is
    for key, value in kwargs.items():
        if key not in validated:
            validated[key] = value
    
    return validated