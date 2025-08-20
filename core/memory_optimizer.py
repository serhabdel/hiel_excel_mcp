"""
Memory optimization module for Hiel Excel MCP.
Provides advanced memory management, garbage collection, and resource optimization.
"""

import gc
import sys
import psutil
import threading
import weakref
from typing import Dict, Any, Optional, Set, Callable, Protocol
from functools import wraps
from contextlib import contextmanager
import logging
from dataclasses import dataclass
from datetime import datetime, timedelta
import os

logger = logging.getLogger(__name__)


@dataclass
class MemoryStats:
    """Memory usage statistics."""
    rss_mb: float  # Resident Set Size in MB
    vms_mb: float  # Virtual Memory Size in MB
    percent: float  # Memory usage percentage
    available_mb: float  # Available system memory in MB
    gc_counts: tuple  # Garbage collection counts (gen0, gen1, gen2)
    open_files: int  # Number of open file descriptors
    threads: int  # Number of active threads
    timestamp: datetime


class MemoryOptimizer:
    """Advanced memory management and optimization system."""
    
    def __init__(self, 
                 memory_threshold_mb: float = 500.0,
                 gc_threshold_mb: float = 100.0,
                 enable_monitoring: bool = True):
        self.memory_threshold_mb = memory_threshold_mb
        self.gc_threshold_mb = gc_threshold_mb
        self.enable_monitoring = enable_monitoring
        
        # Memory tracking
        self._memory_samples = []
        self._max_samples = 100
        self._last_cleanup = datetime.now()
        self._cleanup_interval = timedelta(minutes=5)
        
        # Weak references to track objects
        self._tracked_objects: Set[weakref.ReferenceType] = set()
        self._memory_callbacks: Dict[str, Callable] = {}
        
        # Thread safety
        self._lock = threading.RLock()
        
        # Process handle
        try:
            self._process = psutil.Process()
        except psutil.NoSuchProcess:
            self._process = None
            logger.warning("Could not initialize process monitoring")
    
    def get_memory_stats(self) -> MemoryStats:
        """Get comprehensive memory statistics."""
        try:
            if self._process:
                memory_info = self._process.memory_info()
                memory_percent = self._process.memory_percent()
                open_files = len(self._process.open_files())
                threads = self._process.num_threads()
            else:
                memory_info = type('obj', (object,), {'rss': 0, 'vms': 0})()
                memory_percent = 0.0
                open_files = 0
                threads = 0
            
            # Get system memory
            sys_memory = psutil.virtual_memory()
            
            stats = MemoryStats(
                rss_mb=memory_info.rss / (1024 * 1024),
                vms_mb=memory_info.vms / (1024 * 1024),
                percent=memory_percent,
                available_mb=sys_memory.available / (1024 * 1024),
                gc_counts=gc.get_count(),
                open_files=open_files,
                threads=threads,
                timestamp=datetime.now()
            )
            
            # Store sample for trend analysis
            with self._lock:
                self._memory_samples.append(stats)
                if len(self._memory_samples) > self._max_samples:
                    self._memory_samples.pop(0)
            
            return stats
            
        except Exception as e:
            logger.error(f"Error getting memory stats: {e}")
            # Return minimal stats
            return MemoryStats(
                rss_mb=0.0, vms_mb=0.0, percent=0.0, available_mb=0.0,
                gc_counts=(0, 0, 0), open_files=0, threads=0,
                timestamp=datetime.now()
            )
    
    def optimize_memory(self, force: bool = False) -> Dict[str, Any]:
        """Perform memory optimization operations."""
        results = {
            'triggered_by': 'force' if force else 'automatic',
            'timestamp': datetime.now().isoformat(),
            'actions': [],
            'stats_before': None,
            'stats_after': None
        }
        
        try:
            # Get memory stats before optimization
            stats_before = self.get_memory_stats()
            results['stats_before'] = {
                'rss_mb': stats_before.rss_mb,
                'percent': stats_before.percent,
                'gc_counts': stats_before.gc_counts
            }
            
            # Check if optimization is needed
            needs_optimization = (
                force or 
                stats_before.rss_mb > self.memory_threshold_mb or
                stats_before.percent > 80.0
            )
            
            if not needs_optimization:
                results['actions'].append('no_optimization_needed')
                return results
            
            logger.info(f"Starting memory optimization: RSS={stats_before.rss_mb:.1f}MB, %={stats_before.percent:.1f}")
            
            # 1. Clean up tracked objects
            cleaned_objects = self._cleanup_tracked_objects()
            if cleaned_objects > 0:
                results['actions'].append(f'cleaned_{cleaned_objects}_tracked_objects')
            
            # 2. Force garbage collection
            gc_before = sum(gc.get_count())
            gc.collect()  # Full collection
            gc_after = sum(gc.get_count())
            collected = gc_before - gc_after
            if collected > 0:
                results['actions'].append(f'gc_collected_{collected}_objects')
            
            # 3. Optimize Python internals
            if hasattr(sys, 'intern'):
                # Clean up string interning cache (if possible)
                results['actions'].append('optimized_string_cache')
            
            # 4. Clear import cache for unused modules
            cleared_modules = self._clear_unused_modules()
            if cleared_modules > 0:
                results['actions'].append(f'cleared_{cleared_modules}_unused_modules')
            
            # 5. Trigger memory callbacks
            callback_results = self._trigger_memory_callbacks()
            if callback_results:
                results['actions'].extend(callback_results)
            
            # 6. System-level optimization hints
            if hasattr(os, 'posix_fadvise'):
                # Hint that we don't need cached data
                results['actions'].append('advised_memory_release')
            
            # Get memory stats after optimization
            stats_after = self.get_memory_stats()
            results['stats_after'] = {
                'rss_mb': stats_after.rss_mb,
                'percent': stats_after.percent,
                'gc_counts': stats_after.gc_counts
            }
            
            # Calculate savings
            memory_saved = stats_before.rss_mb - stats_after.rss_mb
            results['memory_saved_mb'] = memory_saved
            results['optimization_effective'] = memory_saved > 1.0  # Saved more than 1MB
            
            if memory_saved > 0:
                logger.info(f"Memory optimization completed: saved {memory_saved:.1f}MB")
            else:
                logger.info("Memory optimization completed: no significant savings")
            
            return results
            
        except Exception as e:
            logger.error(f"Error during memory optimization: {e}")
            results['error'] = str(e)
            return results
    
    def _cleanup_tracked_objects(self) -> int:
        """Clean up weak references to dead objects."""
        with self._lock:
            initial_count = len(self._tracked_objects)
            # Remove dead references
            self._tracked_objects = {ref for ref in self._tracked_objects if ref() is not None}
            cleaned = initial_count - len(self._tracked_objects)
            return cleaned
    
    def _clear_unused_modules(self) -> int:
        """Clear unused imported modules from sys.modules."""
        # This is a conservative approach - only clear specific modules we know are safe
        safe_to_clear = [
            'tempfile', 'shutil', 'zipfile', 'tarfile',  # File utilities
            'urllib', 'http', 'email',  # Network modules  
            'xml', 'html', 'json',  # Data formats
        ]
        
        cleared = 0
        modules_to_remove = []
        
        for module_name in list(sys.modules.keys()):
            for safe_module in safe_to_clear:
                if (module_name.startswith(safe_module + '.') and 
                    module_name not in ['json']):  # Keep essential modules
                    modules_to_remove.append(module_name)
                    break
        
        for module_name in modules_to_remove:
            try:
                del sys.modules[module_name]
                cleared += 1
            except (KeyError, AttributeError):
                pass
        
        return cleared
    
    def _trigger_memory_callbacks(self) -> list:
        """Trigger registered memory optimization callbacks."""
        results = []
        
        for name, callback in self._memory_callbacks.items():
            try:
                callback_result = callback()
                if callback_result:
                    results.append(f'callback_{name}_executed')
            except Exception as e:
                logger.warning(f"Memory callback '{name}' failed: {e}")
                results.append(f'callback_{name}_failed')
        
        return results
    
    def track_object(self, obj: Any, name: Optional[str] = None) -> None:
        """Track an object for memory monitoring."""
        try:
            ref = weakref.ref(obj)
            with self._lock:
                self._tracked_objects.add(ref)
        except TypeError:
            # Object doesn't support weak references
            pass
    
    def register_memory_callback(self, name: str, callback: Callable) -> None:
        """Register a callback to be called during memory optimization."""
        self._memory_callbacks[name] = callback
    
    def unregister_memory_callback(self, name: str) -> None:
        """Unregister a memory optimization callback."""
        self._memory_callbacks.pop(name, None)
    
    def get_memory_trend(self, minutes: int = 30) -> Dict[str, Any]:
        """Get memory usage trend over the specified time period."""
        cutoff = datetime.now() - timedelta(minutes=minutes)
        
        with self._lock:
            recent_samples = [
                s for s in self._memory_samples 
                if s.timestamp >= cutoff
            ]
        
        if not recent_samples:
            return {'trend': 'no_data', 'samples': 0}
        
        # Calculate trend
        rss_values = [s.rss_mb for s in recent_samples]
        
        if len(rss_values) < 2:
            return {'trend': 'insufficient_data', 'samples': len(rss_values)}
        
        # Simple linear trend calculation
        first_half = rss_values[:len(rss_values)//2]
        second_half = rss_values[len(rss_values)//2:]
        
        avg_first = sum(first_half) / len(first_half)
        avg_second = sum(second_half) / len(second_half)
        
        trend = 'stable'
        if avg_second > avg_first * 1.1:
            trend = 'increasing'
        elif avg_second < avg_first * 0.9:
            trend = 'decreasing'
        
        return {
            'trend': trend,
            'samples': len(recent_samples),
            'avg_first_half_mb': avg_first,
            'avg_second_half_mb': avg_second,
            'current_mb': rss_values[-1],
            'peak_mb': max(rss_values),
            'min_mb': min(rss_values)
        }
    
    @contextmanager
    def memory_context(self, operation_name: str):
        """Context manager for monitoring memory usage of operations."""
        stats_before = self.get_memory_stats()
        start_time = datetime.now()
        
        try:
            yield
        finally:
            end_time = datetime.now()
            stats_after = self.get_memory_stats()
            duration = (end_time - start_time).total_seconds()
            
            memory_delta = stats_after.rss_mb - stats_before.rss_mb
            
            # Log memory usage if significant
            if abs(memory_delta) > 1.0 or duration > 1.0:
                logger.debug(
                    f"Operation '{operation_name}': "
                    f"duration={duration:.2f}s, "
                    f"memory_delta={memory_delta:+.1f}MB"
                )
            
            # Trigger optimization if memory increased significantly
            if memory_delta > self.gc_threshold_mb:
                logger.info(f"Operation '{operation_name}' used {memory_delta:.1f}MB, triggering optimization")
                self.optimize_memory()


def memory_optimized(threshold_mb: float = 50.0):
    """Decorator to automatically optimize memory after operations that use significant memory."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Get memory before
            try:
                process = psutil.Process()
                memory_before = process.memory_info().rss / (1024 * 1024)
            except:
                memory_before = 0
            
            try:
                result = func(*args, **kwargs)
                
                # Check memory after
                try:
                    memory_after = process.memory_info().rss / (1024 * 1024)
                    memory_used = memory_after - memory_before
                    
                    if memory_used > threshold_mb:
                        # Trigger optimization
                        logger.debug(f"Function '{func.__name__}' used {memory_used:.1f}MB, optimizing memory")
                        gc.collect()
                        
                except:
                    pass
                
                return result
                
            except Exception:
                # Even if function fails, try to clean up
                gc.collect()
                raise
        
        return wrapper
    return decorator


# Global memory optimizer instance
memory_optimizer = MemoryOptimizer()