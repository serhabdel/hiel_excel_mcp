"""
Centralized error handling for Hiel Excel MCP.
Provides consistent error handling, logging, and response formatting.
"""

import logging
import time
import traceback
from typing import Any, Dict, Optional, Type, Union, Callable
from functools import wraps
import asyncio

from .utils import ExcelMCPException, ValidationError, SecurityError, PerformanceError
from .config import config

logger = logging.getLogger(__name__)


class ErrorHandler:
    """Centralized error handling for Excel MCP operations."""
    
    # Error code mappings for consistent error reporting
    ERROR_CODES = {
        'VALIDATION_ERROR': 4001,
        'SECURITY_ERROR': 4003,
        'FILE_NOT_FOUND': 4004,
        'PERMISSION_DENIED': 4005,
        'TIMEOUT_ERROR': 4008,
        'PERFORMANCE_ERROR': 4009,
        'INTERNAL_ERROR': 5000,
        'IMPORT_ERROR': 5001,
        'CONFIGURATION_ERROR': 5002
    }
    
    @staticmethod
    def wrap_operation(operation_name: str, tool_name: str):
        """
        Decorator to wrap operations with consistent error handling.
        
        Args:
            operation_name: Name of the operation being performed
            tool_name: Name of the tool performing the operation
        """
        def decorator(func: Callable):
            @wraps(func)
            def wrapper(*args, **kwargs):
                context = {
                    'tool': tool_name,
                    'operation': operation_name,
                    'timestamp': time.time(),
                    'function': func.__name__,
                    'args_count': len(args),
                    'kwargs_keys': list(kwargs.keys())
                }
                
                try:
                    # Add timeout protection for long operations
                    start_time = time.time()
                    result = func(*args, **kwargs)
                    duration = time.time() - start_time
                    
                    # Log performance warnings
                    if duration > 5.0:  # More than 5 seconds
                        logger.warning(f"Slow operation: {tool_name}.{operation_name} took {duration:.2f}s")
                    
                    # Ensure result is in correct format
                    if isinstance(result, dict):
                        if 'success' not in result:
                            result['success'] = True
                        if 'timestamp' not in result:
                            result['timestamp'] = time.time()
                    
                    return result
                    
                except Exception as e:
                    return ErrorHandler.handle_error(e, context)
            
            # Handle async functions
            if asyncio.iscoroutinefunction(func):
                @wraps(func)
                async def async_wrapper(*args, **kwargs):
                    context = {
                        'tool': tool_name,
                        'operation': operation_name,
                        'timestamp': time.time(),
                        'function': func.__name__,
                        'async': True
                    }
                    
                    try:
                        start_time = time.time()
                        
                        # Add timeout protection
                        try:
                            result = await asyncio.wait_for(
                                func(*args, **kwargs),
                                timeout=config.operation_timeout_seconds
                            )
                        except asyncio.TimeoutError:
                            raise PerformanceError(
                                f"Operation timed out after {config.operation_timeout_seconds} seconds",
                                context=context
                            )
                        
                        duration = time.time() - start_time
                        
                        if duration > 5.0:
                            logger.warning(f"Slow async operation: {tool_name}.{operation_name} took {duration:.2f}s")
                        
                        if isinstance(result, dict):
                            if 'success' not in result:
                                result['success'] = True
                            if 'timestamp' not in result:
                                result['timestamp'] = time.time()
                        
                        return result
                        
                    except Exception as e:
                        return ErrorHandler.handle_error(e, context)
                
                return async_wrapper
            
            return wrapper
        return decorator
    
    @staticmethod
    def handle_error(error: Exception, context: Dict[str, Any]) -> Dict[str, Any]:
        """
        Handle and format errors with consistent structure.
        
        Args:
            error: Exception that occurred
            context: Operation context information
            
        Returns:
            Formatted error response
        """
        error_type = type(error).__name__
        error_message = str(error)
        
        # Determine error code and category
        error_code, error_category = ErrorHandler._classify_error(error)
        
        # Build base error response
        error_response = {
            'success': False,
            'error': error_message,
            'error_type': error_type,
            'error_code': error_code,
            'error_category': error_category,
            'timestamp': time.time(),
            'context': {
                'tool': context.get('tool', 'unknown'),
                'operation': context.get('operation', 'unknown'),
                'function': context.get('function', 'unknown')
            }
        }
        
        # Add suggestions based on error type
        suggestions = ErrorHandler._get_error_suggestions(error, context)
        if suggestions:
            error_response['suggestions'] = suggestions
        
        # Add recovery actions if available
        recovery_actions = ErrorHandler._get_recovery_actions(error, context)
        if recovery_actions:
            error_response['recovery_actions'] = recovery_actions
        
        # Add debug information in development mode
        if config.log_level == 'DEBUG':
            error_response['debug'] = {
                'traceback': traceback.format_exc(),
                'full_context': context,
                'original_error': str(getattr(error, 'original_error', None)) if hasattr(error, 'original_error') else None
            }
        
        # Log the error with appropriate level
        ErrorHandler._log_error(error, error_response, context)
        
        return error_response
    
    @staticmethod
    def _classify_error(error: Exception) -> tuple[int, str]:
        """Classify error and return error code and category."""
        if isinstance(error, ValidationError):
            return ErrorHandler.ERROR_CODES['VALIDATION_ERROR'], 'validation'
        elif isinstance(error, SecurityError):
            return ErrorHandler.ERROR_CODES['SECURITY_ERROR'], 'security'
        elif isinstance(error, PerformanceError):
            return ErrorHandler.ERROR_CODES['PERFORMANCE_ERROR'], 'performance'
        elif isinstance(error, FileNotFoundError):
            return ErrorHandler.ERROR_CODES['FILE_NOT_FOUND'], 'file_system'
        elif isinstance(error, PermissionError):
            return ErrorHandler.ERROR_CODES['PERMISSION_DENIED'], 'file_system'
        elif isinstance(error, (ImportError, ModuleNotFoundError)):
            return ErrorHandler.ERROR_CODES['IMPORT_ERROR'], 'system'
        elif isinstance(error, asyncio.TimeoutError):
            return ErrorHandler.ERROR_CODES['TIMEOUT_ERROR'], 'performance'
        else:
            return ErrorHandler.ERROR_CODES['INTERNAL_ERROR'], 'internal'
    
    @staticmethod
    def _get_error_suggestions(error: Exception, context: Dict[str, Any]) -> list[str]:
        """Get helpful suggestions based on error type and context."""
        suggestions = []
        
        if isinstance(error, ValidationError):
            suggestions.extend([
                "Verify file path is correct and accessible",
                "Check file format and extension",
                "Ensure file is not corrupted or locked",
                "Verify input parameters are valid"
            ])
        
        elif isinstance(error, SecurityError):
            suggestions.extend([
                "Check if file path is within allowed directories",
                "Verify path does not contain '..' or other traversal patterns",
                "Contact administrator if access should be granted",
                "Review security configuration"
            ])
        
        elif isinstance(error, FileNotFoundError):
            suggestions.extend([
                "Verify the file path exists",
                "Check file permissions",
                "Ensure file has not been moved or deleted",
                "Use absolute path if relative path fails"
            ])
        
        elif isinstance(error, PermissionError):
            suggestions.extend([
                "Check file and directory permissions",
                "Ensure file is not locked by another application",
                "Verify user has sufficient privileges",
                "Close file in other applications if open"
            ])
        
        elif isinstance(error, (ImportError, ModuleNotFoundError)):
            suggestions.extend([
                "Verify Excel MCP dependencies are installed",
                "Check Python environment configuration",
                "Reinstall package if modules are missing"
            ])
        
        elif isinstance(error, PerformanceError) or isinstance(error, asyncio.TimeoutError):
            suggestions.extend([
                "Reduce file size or operation complexity",
                "Check available system memory",
                "Consider breaking operation into smaller chunks",
                "Increase timeout if operation is expected to be slow"
            ])
        
        # Add operation-specific suggestions
        operation = context.get('operation', '')
        if 'import' in operation.lower():
            suggestions.append("Verify source file format is supported")
        elif 'export' in operation.lower():
            suggestions.append("Check available disk space for output file")
        elif 'formula' in operation.lower():
            suggestions.append("Validate formula syntax")
        elif 'format' in operation.lower():
            suggestions.append("Check formatting parameters are valid")
        
        return suggestions[:5]  # Limit to 5 most relevant suggestions
    
    @staticmethod
    def _get_recovery_actions(error: Exception, context: Dict[str, Any]) -> list[str]:
        """Get potential recovery actions for the error."""
        actions = []
        
        if isinstance(error, (FileNotFoundError, PermissionError)):
            actions.extend([
                "retry_with_different_path",
                "create_missing_directories",
                "check_file_permissions"
            ])
        
        elif isinstance(error, ValidationError):
            actions.extend([
                "validate_input_parameters",
                "retry_with_corrected_input",
                "use_alternative_operation"
            ])
        
        elif isinstance(error, PerformanceError):
            actions.extend([
                "retry_with_smaller_dataset",
                "increase_timeout",
                "use_batch_processing"
            ])
        
        return actions[:3]  # Limit to 3 most relevant actions
    
    @staticmethod
    def _log_error(error: Exception, error_response: Dict[str, Any], context: Dict[str, Any]):
        """Log error with appropriate level and detail."""
        error_category = error_response.get('error_category', 'unknown')
        tool = context.get('tool', 'unknown')
        operation = context.get('operation', 'unknown')
        
        # Create log message
        log_message = f"Error in {tool}.{operation}: {error_response['error']}"
        
        # Determine log level based on error category
        if error_category in ['security', 'internal']:
            logger.error(log_message)
        elif error_category in ['performance', 'file_system']:
            logger.warning(log_message)
        else:
            logger.info(log_message)
        
        # Log stack trace for internal errors in debug mode
        if error_category == 'internal' and config.log_level == 'DEBUG':
            logger.debug(f"Stack trace for {tool}.{operation}:", exc_info=True)


class ExcelMCPError(ExcelMCPException):
    """Base exception class with enhanced error context."""
    
    def __init__(self, message: str, error_code: Optional[int] = None, 
                 context: Optional[Dict[str, Any]] = None, 
                 suggestions: Optional[list[str]] = None,
                 recovery_actions: Optional[list[str]] = None,
                 original_error: Optional[Exception] = None):
        super().__init__(message, context, original_error)
        self.error_code = error_code or ErrorHandler.ERROR_CODES['INTERNAL_ERROR']
        self.suggestions = suggestions or []
        self.recovery_actions = recovery_actions or []


def handle_excel_errors(operation_name: str, tool_name: str):
    """Convenience decorator for error handling."""
    return ErrorHandler.wrap_operation(operation_name, tool_name)