"""
Utility functions for Hiel Excel MCP.
Centralized utilities to eliminate code duplication and provide common functionality.
"""

import os
import time
import asyncio
import logging
import traceback
from typing import Any, Dict, List, Optional, Callable, Union, Tuple
from pathlib import Path
from contextlib import asynccontextmanager
from functools import wraps
import threading

from .config import config

logger = logging.getLogger(__name__)


class ExcelMCPException(Exception):
    """Base exception for Excel MCP operations."""
    
    def __init__(self, message: str, context: Optional[Dict[str, Any]] = None, original_error: Optional[Exception] = None):
        super().__init__(message)
        self.context = context or {}
        self.original_error = original_error
        self.timestamp = time.time()


class ValidationError(ExcelMCPException):
    """Exception for validation failures."""
    pass


class SecurityError(ExcelMCPException):
    """Exception for security violations."""
    pass


class PerformanceError(ExcelMCPException):
    """Exception for performance-related issues."""
    pass


class ExcelMCPUtils:
    """Centralized utilities for Excel MCP operations."""
    
    _performance_metrics = {}
    _metrics_lock = threading.Lock()
    
    @staticmethod
    def validate_filepath(filepath: str, allow_create: bool = False, check_size: bool = True) -> Tuple[str, List[str]]:
        """
        Comprehensive file path validation with security checks.
        
        Args:
            filepath: Path to validate
            allow_create: Whether to allow non-existent files
            check_size: Whether to validate file size
            
        Returns:
            Tuple of (validated_path, warnings)
            
        Raises:
            ValidationError: If path is invalid
            SecurityError: If path violates security constraints
        """
        warnings = []
        
        try:
            # Basic path validation
            if not filepath or not isinstance(filepath, str):
                raise ValidationError("Filepath must be a non-empty string")
            
            # Get absolute path
            abs_path = os.path.abspath(filepath)
            
            # Security checks
            if config.enable_path_validation:
                if not config.is_path_allowed(abs_path):
                    raise SecurityError(f"Path not allowed: {abs_path}")
                
                # Check for path traversal attempts
                if '..' in filepath or abs_path != os.path.normpath(abs_path):
                    raise SecurityError("Path traversal detected")
            
            # Extension validation
            if not config.is_extension_allowed(abs_path):
                ext = Path(abs_path).suffix
                allowed = ', '.join(config.allowed_extensions)
                raise ValidationError(f"File extension '{ext}' not allowed. Allowed: {allowed}")
            
            # File existence checks
            if os.path.exists(abs_path):
                if not os.path.isfile(abs_path):
                    raise ValidationError(f"Path exists but is not a file: {abs_path}")
                
                # File size validation
                if check_size:
                    file_size = os.path.getsize(abs_path)
                    if file_size > config.max_file_size:
                        size_mb = file_size / (1024 * 1024)
                        max_mb = config.max_file_size / (1024 * 1024)
                        raise ValidationError(f"File size ({size_mb:.1f}MB) exceeds limit ({max_mb:.1f}MB)")
                    
                    if file_size == 0:
                        warnings.append("File is empty")
                
                # Permission checks
                if not os.access(abs_path, os.R_OK):
                    raise ValidationError(f"No read permission for file: {abs_path}")
                
            elif not allow_create:
                raise ValidationError(f"File does not exist: {abs_path}")
            else:
                # Check parent directory for creation
                parent_dir = os.path.dirname(abs_path)
                if not os.path.exists(parent_dir):
                    raise ValidationError(f"Parent directory does not exist: {parent_dir}")
                
                if not os.access(parent_dir, os.W_OK):
                    raise ValidationError(f"No write permission for directory: {parent_dir}")
            
            return abs_path, warnings
            
        except (ValidationError, SecurityError):
            raise
        except Exception as e:
            raise ValidationError(f"Path validation failed: {str(e)}")
    
    @staticmethod
    def safe_import_excel_module(module_name: str) -> Any:
        """
        Safely import modules from the parent excel_mcp package.
        Eliminates the need for sys.path manipulation.
        
        Args:
            module_name: Module name to import (e.g., 'workbook', 'formatting')
            
        Returns:
            Imported module
            
        Raises:
            ImportError: If module cannot be imported
        """
        try:
            # Try importing from the src directory relative to this package
            full_module_name = f"src.excel_mcp.{module_name}"
            module = __import__(full_module_name, fromlist=[module_name])
            return module
        except ImportError as e:
            logger.error(f"Failed to import {full_module_name}: {e}")
            raise ImportError(f"Cannot import Excel MCP module '{module_name}': {e}")
    
    @staticmethod
    def create_operation_context(tool_name: str, operation: str, **kwargs) -> Dict[str, Any]:
        """Create standardized operation context for logging and error handling."""
        return {
            'tool': tool_name,
            'operation': operation,
            'timestamp': time.time(),
            'thread_id': threading.current_thread().ident,
            'parameters': {k: str(v)[:100] for k, v in kwargs.items()},  # Truncate long values
        }
    
    @staticmethod
    def format_error_response(error: Exception, context: Dict[str, Any]) -> Dict[str, Any]:
        """Format error response with consistent structure and helpful information."""
        error_type = type(error).__name__
        
        # Build helpful error message
        message = str(error)
        if hasattr(error, 'original_error') and error.original_error:
            message += f" (caused by: {error.original_error})"
        
        response = {
            'success': False,
            'error': message,
            'error_type': error_type,
            'context': {
                'tool': context.get('tool', 'unknown'),
                'operation': context.get('operation', 'unknown'),
                'timestamp': context.get('timestamp', time.time())
            }
        }
        
        # Add suggestions based on error type
        if isinstance(error, ValidationError):
            response['suggestions'] = [
                "Check file path and permissions",
                "Verify file format and extension",
                "Ensure file is not corrupted"
            ]
        elif isinstance(error, SecurityError):
            response['suggestions'] = [
                "Check allowed paths configuration",
                "Verify file path does not contain path traversal",
                "Contact administrator if path should be allowed"
            ]
        elif isinstance(error, PerformanceError):
            response['suggestions'] = [
                "Reduce file size or operation complexity",
                "Check system resources",
                "Consider using batch operations"
            ]
        
        # Add debug info in development
        if config.log_level == 'DEBUG':
            response['debug'] = {
                'traceback': traceback.format_exc(),
                'full_context': context
            }
        
        return response
    
    @staticmethod
    def performance_monitor(func: Optional[Callable] = None, *, threshold_seconds: float = 1.0):
        """
        Decorator to monitor function performance and log slow operations.
        
        Args:
            threshold_seconds: Log warning if operation takes longer than this
        """
        def decorator(f):
            @wraps(f)
            def wrapper(*args, **kwargs):
                start_time = time.time()
                func_name = f"{f.__module__}.{f.__name__}"
                
                try:
                    result = f(*args, **kwargs)
                    success = True
                except Exception as e:
                    success = False
                    raise
                finally:
                    duration = time.time() - start_time
                    
                    # Record metrics
                    with ExcelMCPUtils._metrics_lock:
                        if func_name not in ExcelMCPUtils._performance_metrics:
                            ExcelMCPUtils._performance_metrics[func_name] = {
                                'total_calls': 0,
                                'total_time': 0,
                                'max_time': 0,
                                'failures': 0
                            }
                        
                        metrics = ExcelMCPUtils._performance_metrics[func_name]
                        metrics['total_calls'] += 1
                        metrics['total_time'] += duration
                        metrics['max_time'] = max(metrics['max_time'], duration)
                        if not success:
                            metrics['failures'] += 1
                    
                    # Log slow operations
                    if config.enable_performance_logging and duration > threshold_seconds:
                        logger.warning(f"Slow operation: {func_name} took {duration:.2f}s")
                
                return result
            
            # Add async version if the function is async
            if asyncio.iscoroutinefunction(f):
                @wraps(f)
                async def async_wrapper(*args, **kwargs):
                    start_time = time.time()
                    func_name = f"{f.__module__}.{f.__name__}"
                    
                    try:
                        result = await f(*args, **kwargs)
                        success = True
                    except Exception as e:
                        success = False
                        raise
                    finally:
                        duration = time.time() - start_time
                        
                        # Record metrics (same as sync version)
                        with ExcelMCPUtils._metrics_lock:
                            if func_name not in ExcelMCPUtils._performance_metrics:
                                ExcelMCPUtils._performance_metrics[func_name] = {
                                    'total_calls': 0,
                                    'total_time': 0,
                                    'max_time': 0,
                                    'failures': 0
                                }
                            
                            metrics = ExcelMCPUtils._performance_metrics[func_name]
                            metrics['total_calls'] += 1
                            metrics['total_time'] += duration
                            metrics['max_time'] = max(metrics['max_time'], duration)
                            if not success:
                                metrics['failures'] += 1
                        
                        if config.enable_performance_logging and duration > threshold_seconds:
                            logger.warning(f"Slow async operation: {func_name} took {duration:.2f}s")
                    
                    return result
                
                return async_wrapper
            
            return wrapper
        
        if func is None:
            return decorator
        else:
            return decorator(func)
    
    @staticmethod
    def get_performance_metrics() -> Dict[str, Any]:
        """Get performance metrics for all monitored functions."""
        with ExcelMCPUtils._metrics_lock:
            metrics = {}
            for func_name, data in ExcelMCPUtils._performance_metrics.items():
                if data['total_calls'] > 0:
                    metrics[func_name] = {
                        'total_calls': data['total_calls'],
                        'total_time': round(data['total_time'], 3),
                        'average_time': round(data['total_time'] / data['total_calls'], 3),
                        'max_time': round(data['max_time'], 3),
                        'failure_rate': round(data['failures'] / data['total_calls'] * 100, 2) if data['total_calls'] > 0 else 0
                    }
            return metrics
    
    @staticmethod
    def clear_performance_metrics():
        """Clear all performance metrics."""
        with ExcelMCPUtils._metrics_lock:
            ExcelMCPUtils._performance_metrics.clear()
    
    @staticmethod
    @asynccontextmanager
    async def async_timeout(timeout_seconds: float):
        """Async context manager for operation timeouts."""
        try:
            async with asyncio.timeout(timeout_seconds):
                yield
        except asyncio.TimeoutError:
            raise PerformanceError(f"Operation timed out after {timeout_seconds} seconds")
    
    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """Sanitize filename to prevent security issues."""
        # Remove or replace dangerous characters
        dangerous_chars = '<>:"/\\|?*'
        for char in dangerous_chars:
            filename = filename.replace(char, '_')
        
        # Remove leading/trailing dots and spaces
        filename = filename.strip('. ')
        
        # Ensure filename is not empty
        if not filename:
            filename = 'unnamed_file'
        
        # Truncate if too long
        if len(filename) > 255:
            filename = filename[:255]
        
        return filename
    
    @staticmethod
    def check_system_health() -> Dict[str, Any]:
        """Check system health and resource usage."""
        import psutil
        import gc
        
        try:
            process = psutil.Process()
            
            health_info = {
                'memory_usage_mb': round(process.memory_info().rss / 1024 / 1024, 2),
                'cpu_percent': process.cpu_percent(),
                'open_files': len(process.open_files()),
                'thread_count': process.num_threads(),
                'gc_counts': gc.get_count(),
                'timestamp': time.time()
            }
            
            # Add warnings for concerning metrics
            warnings = []
            if health_info['memory_usage_mb'] > 500:  # More than 500MB
                warnings.append("High memory usage")
            
            if health_info['open_files'] > 50:
                warnings.append("Many open files")
            
            if health_info['thread_count'] > 20:
                warnings.append("High thread count")
            
            health_info['warnings'] = warnings
            health_info['status'] = 'warning' if warnings else 'healthy'
            
            return health_info
            
        except ImportError:
            return {
                'status': 'unknown',
                'message': 'psutil not available for health monitoring',
                'timestamp': time.time()
            }
        except Exception as e:
            return {
                'status': 'error',
                'message': f"Health check failed: {e}",
                'timestamp': time.time()
            }