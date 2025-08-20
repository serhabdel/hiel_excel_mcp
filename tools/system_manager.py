"""
System manager tool for handling filtering, sorting, and cache operations.
This tool provides high-level system management operations for Excel files.
"""

from typing import Dict, Any, List, Optional, Union
from ..core.base_tool import BaseTool
from ..core.workbook_context import WorkbookContext


class SystemManager(BaseTool):
    """Tool for managing system operations like filtering, sorting, and caching."""
    
    def apply_filter(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: str,
        filter_type: str = 'auto',
        filter_config: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Apply various types of filters to Excel data.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            range_ref: Range to apply filter to
            filter_type: Type of filter ('auto', 'column', 'values')
            filter_config: Filter configuration (optional)
            
        Returns:
            Dict with operation results
        """
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.filters import FilterManager
            
            if filter_type == 'auto':
                return FilterManager.apply_auto_filter(filepath, sheet_name, range_ref)
            
            elif filter_type == 'column':
                if not filter_config or 'column_index' not in filter_config:
                    return {
                        "success": False,
                        "error": "column_index required for column filter"
                    }
                
                column_index = filter_config['column_index']
                filter_criteria = filter_config.get('criteria', {})
                
                return FilterManager.apply_column_filter(
                    filepath, sheet_name, column_index, filter_criteria
                )
            
            elif filter_type == 'values':
                if not filter_config or 'column_index' not in filter_config or 'values' not in filter_config:
                    return {
                        "success": False,
                        "error": "column_index and values required for values filter"
                    }
                
                column_index = filter_config['column_index']
                values = filter_config['values']
                
                filter_criteria = {
                    'type': 'values',
                    'values': values
                }
                
                return FilterManager.apply_column_filter(
                    filepath, sheet_name, column_index, filter_criteria
                )
            
            else:
                return {
                    "success": False,
                    "error": f"Unknown filter type: {filter_type}"
                }
    
    def clear_filters(
        self,
        filepath: str,
        sheet_name: str
    ) -> Dict[str, Any]:
        """Clear all filters from a worksheet."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.filters import FilterManager
            return FilterManager.clear_filters(filepath, sheet_name)
    
    def sort_range(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: str,
        sort_config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Sort a range with flexible configuration.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of worksheet
            range_ref: Range to sort
            sort_config: Sort configuration:
                - columns: List of sort column configurations
                - has_header: Whether range has header row
                - case_sensitive: Case-sensitive sorting (optional)
                
        Returns:
            Dict with operation results
        """
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.sorting import SortManager
            
            if 'columns' in sort_config:
                # Multi-column sort
                sort_columns = sort_config['columns']
                has_header = sort_config.get('has_header', True)
                
                return SortManager.sort_range(
                    filepath, sheet_name, range_ref, sort_columns, has_header
                )
            
            elif 'column_index' in sort_config:
                # Single column sort
                column_index = sort_config['column_index']
                ascending = sort_config.get('ascending', True)
                has_header = sort_config.get('has_header', True)
                
                return SortManager.sort_by_column(
                    filepath, sheet_name, range_ref, column_index, ascending, has_header
                )
            
            else:
                return {
                    "success": False,
                    "error": "Sort configuration must specify 'columns' or 'column_index'"
                }
    
    def sort_multi_column(
        self,
        filepath: str,
        sheet_name: str,
        range_ref: str,
        sort_columns: List[Dict[str, Any]],
        has_header: bool = True
    ) -> Dict[str, Any]:
        """Sort range by multiple columns."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.sorting import SortManager
            return SortManager.sort_range(
                filepath, sheet_name, range_ref, sort_columns, has_header
            )
    
    def invalidate_cache(
        self,
        filepath: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Invalidate workbook cache.
        
        Args:
            filepath: Specific file to invalidate (None for all cache)
            
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.cache_manager import workbook_cache
            
            if filepath:
                workbook_cache.invalidate(filepath)
                return {
                    "success": True,
                    "message": f"Cache invalidated for: {filepath}",
                    "filepath": filepath
                }
            else:
                workbook_cache.clear()
                return {
                    "success": True,
                    "message": "Entire cache cleared"
                }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """Get detailed cache statistics."""
        try:
            from ...src.excel_mcp.cache_manager import workbook_cache
            stats = workbook_cache.get_cache_stats()
            
            return {
                "success": True,
                **stats
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def clear_cache(self) -> Dict[str, Any]:
        """Clear entire workbook cache."""
        try:
            from ...src.excel_mcp.cache_manager import workbook_cache
            workbook_cache.clear()
            
            return {
                "success": True,
                "message": "Cache cleared successfully"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_batch_status(
        self,
        operation_id: str
    ) -> Dict[str, Any]:
        """Get status of a batch operation."""
        try:
            from ...src.excel_mcp.cache_manager import batch_manager
            status = batch_manager.get_operation_status(operation_id)
            
            if status is None:
                return {
                    "success": False,
                    "error": f"Operation not found: {operation_id}"
                }
            
            return {
                "success": True,
                **status
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def list_batch_operations(
        self,
        active_only: bool = True,
        cleanup_old: bool = True
    ) -> Dict[str, Any]:
        """
        List all batch operations.
        
        Args:
            active_only: Show only active operations
            cleanup_old: Clean up old completed operations
            
        Returns:
            Dict with operations list
        """
        try:
            from ...src.excel_mcp.cache_manager import batch_manager
            
            if cleanup_old:
                batch_manager.cleanup_completed()
            
            operations = []
            for op_id, op_data in batch_manager.active_operations.items():
                if active_only and op_data.get('status') != 'running':
                    continue
                
                op_info = {
                    "operation_id": op_id,
                    "description": op_data.get('description', 'Unknown'),
                    "status": op_data.get('status', 'unknown'),
                    "total_items": op_data.get('total_items', 0),
                    "completed_items": op_data.get('completed_items', 0),
                    "failed_items": op_data.get('failed_items', 0),
                    "started": op_data.get('started', 0)
                }
                
                # Calculate progress
                if op_info["total_items"] > 0:
                    completed = op_info["completed_items"] + op_info["failed_items"]
                    op_info["progress_percent"] = int(completed / op_info["total_items"] * 100)
                else:
                    op_info["progress_percent"] = 100
                
                operations.append(op_info)
            
            return {
                "success": True,
                "total_operations": len(operations),
                "active_only": active_only,
                "operations": operations
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def cancel_batch_operation(
        self,
        operation_id: str
    ) -> Dict[str, Any]:
        """
        Cancel a batch operation.
        
        Args:
            operation_id: ID of operation to cancel
            
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.cache_manager import batch_manager
            
            if operation_id in batch_manager.active_operations:
                op = batch_manager.active_operations[operation_id]
                op['status'] = 'cancelled'
                
                return {
                    "success": True,
                    "operation_id": operation_id,
                    "message": "Operation cancelled"
                }
            else:
                return {
                    "success": False,
                    "operation_id": operation_id,
                    "error": "Operation not found"
                }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_filtered_data(
        self,
        filepath: str,
        sheet_name: str,
        include_headers: bool = True
    ) -> Dict[str, Any]:
        """Get visible (filtered) data from worksheet."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.filters import FilterManager
            return FilterManager.get_filtered_data(filepath, sheet_name, include_headers)
    
    def optimize_performance(
        self,
        config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Optimize system performance based on configuration.
        
        Args:
            config: Performance optimization configuration:
                - cache_size: Maximum cache size
                - cache_age: Maximum cache age in seconds
                - cleanup_operations: Whether to cleanup old operations
                
        Returns:
            Dict with optimization results
        """
        try:
            from ...src.excel_mcp.cache_manager import workbook_cache, batch_manager
            
            results = []
            
            # Adjust cache settings
            if 'cache_size' in config:
                old_size = workbook_cache.max_size
                workbook_cache.max_size = config['cache_size']
                results.append(f"Cache size changed from {old_size} to {config['cache_size']}")
            
            if 'cache_age' in config:
                old_age = workbook_cache.max_age_seconds
                workbook_cache.max_age_seconds = config['cache_age']
                results.append(f"Cache age changed from {old_age} to {config['cache_age']} seconds")
            
            # Cleanup old operations
            if config.get('cleanup_operations', True):
                cleanup_age = config.get('operation_cleanup_age', 3600)
                batch_manager.cleanup_completed(cleanup_age)
                results.append(f"Cleaned up operations older than {cleanup_age} seconds")
            
            # Clear cache if requested
            if config.get('clear_cache', False):
                workbook_cache.clear()
                results.append("Cache cleared")
            
            return {
                "success": True,
                "message": "Performance optimization completed",
                "changes": results,
                "config": config
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def system_health_check(self) -> Dict[str, Any]:
        """
        Perform a comprehensive system health check.
        
        Returns:
            Dict with health check results
        """
        try:
            from ...src.excel_mcp.cache_manager import workbook_cache, batch_manager
            import threading
            import gc
            
            health_info = {
                "success": True,
                "timestamp": self._get_current_timestamp(),
                "cache": {},
                "operations": {},
                "system": {}
            }
            
            # Cache health
            cache_stats = workbook_cache.get_cache_stats()
            health_info["cache"] = {
                "size": cache_stats["size"],
                "max_size": cache_stats["max_size"],
                "utilization_percent": int(cache_stats["size"] / cache_stats["max_size"] * 100),
                "max_age_seconds": cache_stats["max_age_seconds"],
                "entries_count": len(cache_stats["entries"])
            }
            
            # Operations health
            active_ops = len([op for op in batch_manager.active_operations.values() 
                            if op.get('status') == 'running'])
            completed_ops = len([op for op in batch_manager.active_operations.values() 
                               if op.get('status') == 'completed'])
            
            health_info["operations"] = {
                "active_operations": active_ops,
                "completed_operations": completed_ops,
                "total_operations": len(batch_manager.active_operations)
            }
            
            # System health
            health_info["system"] = {
                "active_threads": threading.active_count(),
                "gc_counts": gc.get_count()
            }
            
            # Health assessment
            warnings = []
            if health_info["cache"]["utilization_percent"] > 90:
                warnings.append("Cache utilization is high (>90%)")
            
            if active_ops > 10:
                warnings.append(f"Many active operations ({active_ops})")
            
            health_info["warnings"] = warnings
            health_info["status"] = "healthy" if not warnings else "warning"
            
            return health_info
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "status": "error"
            }
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        from datetime import datetime
        return datetime.now().isoformat()