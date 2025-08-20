"""
Configuration management for Hiel Excel MCP.
Centralized configuration to eliminate hard-coded values and sys.path manipulation.
"""

import os
from typing import List, Dict, Any, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class ExcelMCPConfig:
    """Centralized configuration management for Excel MCP operations."""
    
    def __init__(self):
        # File and path settings
        self.excel_files_path = os.getenv('EXCEL_MCP_FILES_PATH', '.')
        self.allowed_extensions = {'.xlsx', '.xls', '.csv', '.xlsm', '.xlsb'}
        self.max_file_size = int(os.getenv('EXCEL_MCP_MAX_FILE_SIZE', '104857600'))  # 100MB
        self.allowed_paths = self._parse_allowed_paths()
        
        # Cache settings
        self.cache_size = int(os.getenv('EXCEL_MCP_CACHE_SIZE', '10'))
        self.cache_age_seconds = int(os.getenv('EXCEL_MCP_CACHE_AGE', '300'))  # 5 minutes
        
        # Performance settings
        self.max_concurrent_operations = int(os.getenv('EXCEL_MCP_MAX_CONCURRENT', '5'))
        self.operation_timeout_seconds = int(os.getenv('EXCEL_MCP_TIMEOUT', '300'))  # 5 minutes
        
        # Security settings
        self.enable_path_validation = os.getenv('EXCEL_MCP_VALIDATE_PATHS', 'true').lower() == 'true'
        self.sandbox_mode = os.getenv('EXCEL_MCP_SANDBOX', 'false').lower() == 'true'
        
        # Logging settings
        self.log_level = os.getenv('EXCEL_MCP_LOG_LEVEL', 'INFO')
        self.enable_performance_logging = os.getenv('EXCEL_MCP_PERF_LOG', 'false').lower() == 'true'
        
        # Batch operation settings
        self.batch_max_workers = int(os.getenv('EXCEL_MCP_BATCH_WORKERS', '4'))
        self.batch_cleanup_age = int(os.getenv('EXCEL_MCP_BATCH_CLEANUP', '3600'))  # 1 hour
        
        # Template settings
        self.template_cache_size = int(os.getenv('EXCEL_MCP_TEMPLATE_CACHE', '5'))
        self.enable_template_validation = os.getenv('EXCEL_MCP_VALIDATE_TEMPLATES', 'true').lower() == 'true'
        
        self._validate_config()
    
    def _parse_allowed_paths(self) -> List[str]:
        """Parse allowed paths from environment variable."""
        paths_str = os.getenv('EXCEL_MCP_ALLOWED_PATHS', self.excel_files_path)
        paths = []
        
        for path_str in paths_str.split(os.pathsep):
            path_str = path_str.strip()
            if path_str:
                try:
                    abs_path = os.path.abspath(path_str)
                    paths.append(abs_path)
                except Exception as e:
                    logger.warning(f"Invalid path in EXCEL_MCP_ALLOWED_PATHS: {path_str} - {e}")
        
        return paths or [os.path.abspath('.')]
    
    def _validate_config(self):
        """Validate configuration values and log warnings for invalid settings."""
        warnings = []
        
        # Validate file size
        if self.max_file_size < 1024:  # Less than 1KB
            warnings.append(f"Very small max_file_size: {self.max_file_size} bytes")
        elif self.max_file_size > 1024 * 1024 * 1024:  # Greater than 1GB
            warnings.append(f"Very large max_file_size: {self.max_file_size} bytes")
        
        # Validate cache settings
        if self.cache_size < 1:
            warnings.append("cache_size must be at least 1")
            self.cache_size = 1
        
        if self.cache_age_seconds < 60:  # Less than 1 minute
            warnings.append("cache_age_seconds is very low, may cause frequent cache invalidation")
        
        # Validate concurrency settings
        if self.max_concurrent_operations < 1:
            warnings.append("max_concurrent_operations must be at least 1")
            self.max_concurrent_operations = 1
        elif self.max_concurrent_operations > 20:
            warnings.append("max_concurrent_operations is very high, may cause resource issues")
        
        # Validate timeout
        if self.operation_timeout_seconds < 30:  # Less than 30 seconds
            warnings.append("operation_timeout_seconds is very low")
        
        # Log all warnings
        for warning in warnings:
            logger.warning(f"Configuration warning: {warning}")
    
    def is_path_allowed(self, filepath: str) -> bool:
        """Check if a file path is within allowed directories."""
        if not self.enable_path_validation:
            return True
        
        try:
            abs_filepath = os.path.abspath(filepath)
            
            for allowed_path in self.allowed_paths:
                if abs_filepath.startswith(allowed_path):
                    return True
            
            return False
        except Exception:
            return False
    
    def is_extension_allowed(self, filepath: str) -> bool:
        """Check if file extension is allowed."""
        ext = Path(filepath).suffix.lower()
        return ext in self.allowed_extensions
    
    def get_safe_filepath(self, filepath: str) -> Optional[str]:
        """Get a safe, validated filepath or None if invalid."""
        try:
            # Basic path validation
            abs_path = os.path.abspath(filepath)
            
            # Check if path is allowed
            if not self.is_path_allowed(abs_path):
                return None
            
            # Check extension
            if not self.is_extension_allowed(abs_path):
                return None
            
            return abs_path
        except Exception:
            return None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary for serialization."""
        return {
            'excel_files_path': self.excel_files_path,
            'allowed_extensions': list(self.allowed_extensions),
            'max_file_size': self.max_file_size,
            'allowed_paths': self.allowed_paths,
            'cache_size': self.cache_size,
            'cache_age_seconds': self.cache_age_seconds,
            'max_concurrent_operations': self.max_concurrent_operations,
            'operation_timeout_seconds': self.operation_timeout_seconds,
            'enable_path_validation': self.enable_path_validation,
            'sandbox_mode': self.sandbox_mode,
            'log_level': self.log_level,
            'enable_performance_logging': self.enable_performance_logging,
            'batch_max_workers': self.batch_max_workers,
            'batch_cleanup_age': self.batch_cleanup_age,
            'template_cache_size': self.template_cache_size,
            'enable_template_validation': self.enable_template_validation
        }


# Global configuration instance
config = ExcelMCPConfig()