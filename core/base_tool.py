"""
Base tool interface and operation routing system for hiel_excel_mcp.

This module provides the abstract base class and infrastructure for all grouped tools,
including operation routing, validation, and standardized response handling.
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional, Callable
from functools import wraps
import json
import logging
import asyncio
from dataclasses import dataclass, asdict
from enum import Enum

from .utils import ExcelMCPUtils, ValidationError as UtilsValidationError
from .error_handler import ErrorHandler, handle_excel_errors
from .config import config

logger = logging.getLogger(__name__)


class OperationStatus(Enum):
    """Status codes for operation responses."""
    SUCCESS = "success"
    ERROR = "error"
    WARNING = "warning"


@dataclass
class OperationResponse:
    """Standardized response model for all tool operations."""
    success: bool
    operation: str
    message: str
    data: Optional[Dict[str, Any]] = None
    errors: Optional[List[str]] = None
    warnings: Optional[List[str]] = None
    status: OperationStatus = OperationStatus.SUCCESS
    
    def to_json(self) -> str:
        """Convert response to JSON string."""
        response_dict = asdict(self)
        response_dict['status'] = self.status.value
        return json.dumps(response_dict, default=str, indent=2)


@dataclass
class OperationMetadata:
    """Metadata for tool operations."""
    name: str
    description: str
    required_params: List[str]
    optional_params: List[str]
    examples: Optional[Dict[str, Any]] = None


class ValidationError(Exception):
    """Raised when operation validation fails."""
    pass


class OperationNotFoundError(Exception):
    """Raised when requested operation is not supported."""
    pass


def operation_route(name: str, description: str, required_params: List[str], 
                   optional_params: Optional[List[str]] = None):
    """
    Decorator for registering operations within grouped tools.
    
    Args:
        name: Operation name
        description: Operation description
        required_params: List of required parameter names
        optional_params: List of optional parameter names
    """
    if optional_params is None:
        optional_params = []
    
    def decorator(func: Callable) -> Callable:
        func._operation_metadata = OperationMetadata(
            name=name,
            description=description,
            required_params=required_params,
            optional_params=optional_params
        )
        
        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        
        return wrapper
    return decorator


class BaseTool(ABC):
    """
    Abstract base class for all grouped tools in hiel_excel_mcp.
    
    Provides common functionality for operation routing, validation,
    error handling, and response formatting.
    """
    
    def __init__(self):
        self._operations: Dict[str, Callable] = {}
        self._operation_metadata: Dict[str, OperationMetadata] = {}
        self._register_operations()
    
    def _register_operations(self):
        """Register all operations defined in the tool class."""
        for attr_name in dir(self):
            attr = getattr(self, attr_name)
            if hasattr(attr, '_operation_metadata'):
                metadata = attr._operation_metadata
                self._operations[metadata.name] = attr
                self._operation_metadata[metadata.name] = metadata
                logger.debug(f"Registered operation '{metadata.name}' for {self.__class__.__name__}")
    
    def get_available_operations(self) -> List[str]:
        """Get list of available operations for this tool."""
        return list(self._operations.keys())
    
    def get_operation_metadata(self, operation: str) -> Optional[OperationMetadata]:
        """Get metadata for a specific operation."""
        return self._operation_metadata.get(operation)
    
    def get_all_operations_metadata(self) -> Dict[str, OperationMetadata]:
        """Get metadata for all operations."""
        return self._operation_metadata.copy()
    
    def validate_operation(self, operation: str) -> None:
        """
        Validate that the requested operation is supported.
        
        Args:
            operation: Operation name to validate
            
        Raises:
            OperationNotFoundError: If operation is not supported
        """
        if operation not in self._operations:
            available = ", ".join(self.get_available_operations())
            raise OperationNotFoundError(
                f"Operation '{operation}' not supported. Available operations: {available}"
            )
    
    def validate_parameters(self, operation: str, **kwargs) -> None:
        """
        Validate parameters for a specific operation.
        
        Args:
            operation: Operation name
            **kwargs: Parameters to validate
            
        Raises:
            ValidationError: If required parameters are missing
        """
        metadata = self.get_operation_metadata(operation)
        if not metadata:
            return
        
        missing_params = []
        for param in metadata.required_params:
            if param not in kwargs or kwargs[param] is None:
                missing_params.append(param)
        
        if missing_params:
            raise ValidationError(
                f"Missing required parameters for operation '{operation}': {', '.join(missing_params)}"
            )
    
    @ExcelMCPUtils.performance_monitor
    def execute_operation(self, operation: str, **kwargs) -> OperationResponse:
        """
        Execute a specific operation with validation and error handling.
        
        Args:
            operation: Operation name to execute
            **kwargs: Operation parameters
            
        Returns:
            OperationResponse: Standardized response object
        """
        context = ExcelMCPUtils.create_operation_context(
            self.get_tool_name(), operation, **kwargs
        )
        
        try:
            # Validate operation exists
            self.validate_operation(operation)
            
            # Validate parameters
            self.validate_parameters(operation, **kwargs)
            
            # Validate file paths if present
            self._validate_file_paths(**kwargs)
            
            # Execute operation with error handling wrapper
            operation_func = self._operations[operation]
            wrapped_func = handle_excel_errors(operation, self.get_tool_name())(operation_func)
            result = wrapped_func(**kwargs)
            
            # Handle different return types
            if isinstance(result, OperationResponse):
                return result
            elif isinstance(result, dict):
                # Check if it's already an error response
                if result.get('success') is False:
                    return OperationResponse(
                        success=False,
                        operation=operation,
                        message=result.get('error', f"Operation '{operation}' failed"),
                        errors=[result.get('error', 'Unknown error')],
                        data=result,
                        status=OperationStatus.ERROR
                    )
                
                return OperationResponse(
                    success=True,
                    operation=operation,
                    message=f"Operation '{operation}' completed successfully",
                    data=result,
                    status=OperationStatus.SUCCESS
                )
            else:
                return OperationResponse(
                    success=True,
                    operation=operation,
                    message=f"Operation '{operation}' completed successfully",
                    data={"result": result} if result is not None else None,
                    status=OperationStatus.SUCCESS
                )
                
        except OperationNotFoundError as e:
            error_response = ErrorHandler.handle_error(e, context)
            return OperationResponse(
                success=False,
                operation=operation,
                message=error_response['error'],
                errors=[error_response['error']],
                data=error_response,
                status=OperationStatus.ERROR
            )
            
        except (ValidationError, UtilsValidationError) as e:
            error_response = ErrorHandler.handle_error(e, context)
            return OperationResponse(
                success=False,
                operation=operation,
                message=error_response['error'],
                errors=[error_response['error']],
                data=error_response,
                status=OperationStatus.ERROR
            )
            
        except Exception as e:
            error_response = ErrorHandler.handle_error(e, context)
            return OperationResponse(
                success=False,
                operation=operation,
                message=error_response['error'],
                errors=[error_response['error']],
                data=error_response,
                status=OperationStatus.ERROR
            )
    
    def _validate_file_paths(self, **kwargs):
        """Validate file paths in operation parameters."""
        # Common file path parameter names
        file_params = ['filepath', 'file_path', 'template_path', 'output_path', 
                      'csv_path', 'input_path', 'workbook_path']
        
        for param_name in file_params:
            if param_name in kwargs:
                file_path = kwargs[param_name]
                if file_path:
                    try:
                        # Use centralized validation
                        allow_create = param_name in ['output_path', 'csv_path']
                        validated_path, warnings = ExcelMCPUtils.validate_filepath(
                            file_path, allow_create=allow_create
                        )
                        # Update the parameter with validated path
                        kwargs[param_name] = validated_path
                    except Exception as e:
                        raise UtilsValidationError(f"Invalid {param_name}: {e}")
    
    async def execute_operation_async(self, operation: str, **kwargs) -> OperationResponse:
        """
        Execute operation asynchronously with timeout protection.
        
        Args:
            operation: Operation name to execute
            **kwargs: Operation parameters
            
        Returns:
            OperationResponse: Standardized response object
        """
        try:
            # Use async timeout from config
            async with ExcelMCPUtils.async_timeout(config.operation_timeout_seconds):
                # Run synchronous operation in executor
                loop = asyncio.get_event_loop()
                return await loop.run_in_executor(
                    None, 
                    self.execute_operation, 
                    operation, 
                    **kwargs
                )
        except Exception as e:
            context = ExcelMCPUtils.create_operation_context(
                self.get_tool_name(), operation, **kwargs
            )
            error_response = ErrorHandler.handle_error(e, context)
            return OperationResponse(
                success=False,
                operation=operation,
                message=error_response['error'],
                errors=[error_response['error']],
                data=error_response,
                status=OperationStatus.ERROR
            )
    
    @abstractmethod
    def get_tool_name(self) -> str:
        """Get the name of this tool."""
        pass
    
    @abstractmethod
    def get_tool_description(self) -> str:
        """Get the description of this tool."""
        pass
    
    def get_tool_info(self) -> Dict[str, Any]:
        """Get comprehensive information about this tool."""
        return {
            "name": self.get_tool_name(),
            "description": self.get_tool_description(),
            "operations": {
                name: {
                    "description": metadata.description,
                    "required_params": metadata.required_params,
                    "optional_params": metadata.optional_params,
                    "examples": metadata.examples
                }
                for name, metadata in self._operation_metadata.items()
            }
        }


def create_error_response(operation: str, error: Exception, 
                         context: Optional[Dict[str, Any]] = None) -> OperationResponse:
    """
    Create a standardized error response.
    
    Args:
        operation: Operation name that failed
        error: Exception that occurred
        context: Additional context information
        
    Returns:
        OperationResponse: Standardized error response
    """
    error_message = str(error)
    error_type = type(error).__name__
    
    return OperationResponse(
        success=False,
        operation=operation,
        message=f"Operation failed: {error_message}",
        errors=[f"{error_type}: {error_message}"],
        data={"context": context} if context else None,
        status=OperationStatus.ERROR
    )


def create_success_response(operation: str, message: str, 
                          data: Optional[Dict[str, Any]] = None,
                          warnings: Optional[List[str]] = None) -> OperationResponse:
    """
    Create a standardized success response.
    
    Args:
        operation: Operation name that succeeded
        message: Success message
        data: Response data
        warnings: Optional warnings
        
    Returns:
        OperationResponse: Standardized success response
    """
    return OperationResponse(
        success=True,
        operation=operation,
        message=message,
        data=data,
        warnings=warnings,
        status=OperationStatus.WARNING if warnings else OperationStatus.SUCCESS
    )