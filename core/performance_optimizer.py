"""
Performance optimization module for Hiel Excel MCP.
Provides advanced performance optimizations, caching, and execution strategies.
"""

import asyncio
import time
import threading
from typing import Dict, Any, Optional, Callable, Union, List, Tuple
from functools import wraps, lru_cache
from contextlib import asynccontextmanager
import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta
import weakref
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
import multiprocessing
from queue import Queue, Empty
import pickle
import hashlib
import os

from .config import config
from .memory_optimizer import memory_optimizer

logger = logging.getLogger(__name__)


@dataclass
class PerformanceMetrics:
    """Performance metrics for operations."""
    operation_name: str
    execution_time: float
    memory_delta: float
    cpu_usage: float
    cache_hits: int
    cache_misses: int
    concurrent_ops: int
    timestamp: datetime = field(default_factory=datetime.now)


class SmartCache:
    """High-performance cache with TTL, LRU eviction, and memory management."""
    
    def __init__(self, max_size: int = 1000, ttl_seconds: float = 300):
        self.max_size = max_size
        self.ttl_seconds = ttl_seconds
        self._cache: Dict[str, Tuple[Any, float, int]] = {}  # key -> (value, timestamp, access_count)
        self._access_times: Dict[str, float] = {}
        self._lock = threading.RLock()
        self._hits = 0
        self._misses = 0
    
    def get(self, key: str) -> Optional[Any]:
        """Get value from cache with TTL check."""
        with self._lock:
            if key not in self._cache:
                self._misses += 1
                return None
            
            value, timestamp, access_count = self._cache[key]
            
            # Check TTL
            if time.time() - timestamp > self.ttl_seconds:
                del self._cache[key]
                self._access_times.pop(key, None)
                self._misses += 1
                return None
            
            # Update access statistics
            self._cache[key] = (value, timestamp, access_count + 1)
            self._access_times[key] = time.time()
            self._hits += 1
            return value
    
    def set(self, key: str, value: Any) -> None:
        """Set value in cache with LRU eviction."""
        with self._lock:
            current_time = time.time()
            
            # Evict if at capacity
            if len(self._cache) >= self.max_size and key not in self._cache:
                self._evict_lru()
            
            self._cache[key] = (value, current_time, 0)
            self._access_times[key] = current_time
    
    def _evict_lru(self) -> None:
        """Evict least recently used item."""
        if not self._access_times:
            return
        
        lru_key = min(self._access_times.keys(), key=lambda k: self._access_times[k])
        del self._cache[lru_key]
        del self._access_times[lru_key]
    
    def clear(self) -> None:
        """Clear entire cache."""
        with self._lock:
            self._cache.clear()
            self._access_times.clear()
            self._hits = 0
            self._misses = 0
    
    def stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        with self._lock:
            total_requests = self._hits + self._misses
            hit_rate = (self._hits / total_requests) if total_requests > 0 else 0.0
            
            return {
                'size': len(self._cache),
                'max_size': self.max_size,
                'hit_rate': hit_rate,
                'hits': self._hits,
                'misses': self._misses,
                'ttl_seconds': self.ttl_seconds
            }


class ExecutionPool:
    """Smart execution pool with adaptive scaling."""
    
    def __init__(self, max_workers: Optional[int] = None):
        self.max_workers = max_workers or min(32, (os.cpu_count() or 1) + 4)
        self._thread_pool: Optional[ThreadPoolExecutor] = None
        self._process_pool: Optional[ProcessPoolExecutor] = None
        self._active_tasks = 0
        self._lock = threading.Lock()
        
    def _get_thread_pool(self) -> ThreadPoolExecutor:
        """Get or create thread pool."""
        if self._thread_pool is None:
            self._thread_pool = ThreadPoolExecutor(
                max_workers=self.max_workers,
                thread_name_prefix="hiel-excel-worker"
            )
        return self._thread_pool
    
    def _get_process_pool(self) -> ProcessPoolExecutor:
        """Get or create process pool."""
        if self._process_pool is None:
            # Use fewer processes than threads
            max_processes = min(4, os.cpu_count() or 1)
            self._process_pool = ProcessPoolExecutor(max_workers=max_processes)
        return self._process_pool
    
    async def execute_async(self, func: Callable, *args, use_process: bool = False, **kwargs) -> Any:
        """Execute function asynchronously."""
        with self._lock:
            self._active_tasks += 1
        
        try:
            loop = asyncio.get_event_loop()
            if use_process:
                executor = self._get_process_pool()
            else:
                executor = self._get_thread_pool()
            
            return await loop.run_in_executor(executor, func, *args)
        finally:
            with self._lock:
                self._active_tasks -= 1
    
    def shutdown(self) -> None:
        """Shutdown execution pools."""
        if self._thread_pool:
            self._thread_pool.shutdown(wait=True)
            self._thread_pool = None
        
        if self._process_pool:
            self._process_pool.shutdown(wait=True)
            self._process_pool = None
    
    @property
    def active_tasks(self) -> int:
        """Get number of active tasks."""
        with self._lock:
            return self._active_tasks


class PerformanceOptimizer:
    """Advanced performance optimization system."""
    
    def __init__(self):
        self.cache = SmartCache(
            max_size=config.cache_size * 10,  # Larger performance cache
            ttl_seconds=config.cache_age_seconds
        )
        self.execution_pool = ExecutionPool(max_workers=config.max_concurrent_operations)
        self._metrics: List[PerformanceMetrics] = []
        self._max_metrics = 1000
        self._lock = threading.RLock()
        
        # Performance optimization flags
        self._optimizations_enabled = {
            'caching': True,
            'async_execution': True,
            'memory_management': True,
            'batch_processing': True,
            'compression': True
        }
    
    def cached_operation(self, 
                        cache_key_func: Optional[Callable] = None,
                        ttl_seconds: Optional[float] = None,
                        ignore_args: List[str] = None):
        """Decorator for caching operation results."""
        def decorator(func):
            @wraps(func)
            def wrapper(*args, **kwargs):
                if not self._optimizations_enabled['caching']:
                    return func(*args, **kwargs)
                
                # Generate cache key
                if cache_key_func:
                    cache_key = cache_key_func(*args, **kwargs)
                else:
                    # Default key generation
                    filtered_kwargs = kwargs.copy()
                    if ignore_args:
                        for arg in ignore_args:
                            filtered_kwargs.pop(arg, None)
                    
                    key_data = f"{func.__name__}:{str(args)}:{str(filtered_kwargs)}"
                    cache_key = hashlib.md5(key_data.encode()).hexdigest()
                
                # Try to get from cache
                cached_result = self.cache.get(cache_key)
                if cached_result is not None:
                    return cached_result
                
                # Execute function and cache result
                result = func(*args, **kwargs)
                self.cache.set(cache_key, result)
                
                return result
            
            return wrapper
        return decorator
    
    async def async_operation(self, 
                             func: Callable, 
                             *args, 
                             timeout: Optional[float] = None,
                             use_process: bool = False,
                             **kwargs) -> Any:
        """Execute operation asynchronously with optimization."""
        if not self._optimizations_enabled['async_execution']:
            return func(*args, **kwargs)
        
        timeout = timeout or config.operation_timeout_seconds
        
        try:
            return await asyncio.wait_for(
                self.execution_pool.execute_async(func, *args, use_process=use_process, **kwargs),
                timeout=timeout
            )
        except asyncio.TimeoutError:
            raise TimeoutError(f"Operation timed out after {timeout} seconds")
    
    def batch_optimize(self, operations: List[Callable], batch_size: int = 10) -> List[Any]:
        """Optimize batch operations for better performance."""
        if not self._optimizations_enabled['batch_processing']:
            return [op() for op in operations]
        
        results = []
        
        # Process in batches to manage memory
        for i in range(0, len(operations), batch_size):
            batch = operations[i:i + batch_size]
            
            # Execute batch
            batch_results = []
            for operation in batch:
                try:
                    result = operation()
                    batch_results.append(result)
                except Exception as e:
                    logger.error(f"Batch operation failed: {e}")
                    batch_results.append(None)
            
            results.extend(batch_results)
            
            # Optimize memory after each batch
            if self._optimizations_enabled['memory_management']:
                memory_optimizer.optimize_memory()
        
        return results
    
    def performance_monitor(self, operation_name: str):
        """Decorator to monitor operation performance."""
        def decorator(func):
            @wraps(func)
            async def async_wrapper(*args, **kwargs):
                start_time = time.time()
                memory_before = memory_optimizer.get_memory_stats()
                
                try:
                    if asyncio.iscoroutinefunction(func):
                        result = await func(*args, **kwargs)
                    else:
                        result = func(*args, **kwargs)
                    
                    return result
                finally:
                    end_time = time.time()
                    memory_after = memory_optimizer.get_memory_stats()
                    
                    # Record metrics
                    metrics = PerformanceMetrics(
                        operation_name=operation_name,
                        execution_time=end_time - start_time,
                        memory_delta=memory_after.rss_mb - memory_before.rss_mb,
                        cpu_usage=memory_after.percent,
                        cache_hits=self.cache.stats()['hits'],
                        cache_misses=self.cache.stats()['misses'],
                        concurrent_ops=self.execution_pool.active_tasks
                    )
                    
                    self._record_metrics(metrics)
            
            @wraps(func)
            def sync_wrapper(*args, **kwargs):
                start_time = time.time()
                memory_before = memory_optimizer.get_memory_stats()
                
                try:
                    result = func(*args, **kwargs)
                    return result
                finally:
                    end_time = time.time()
                    memory_after = memory_optimizer.get_memory_stats()
                    
                    metrics = PerformanceMetrics(
                        operation_name=operation_name,
                        execution_time=end_time - start_time,
                        memory_delta=memory_after.rss_mb - memory_before.rss_mb,
                        cpu_usage=memory_after.percent,
                        cache_hits=self.cache.stats()['hits'],
                        cache_misses=self.cache.stats()['misses'],
                        concurrent_ops=self.execution_pool.active_tasks
                    )
                    
                    self._record_metrics(metrics)
            
            return async_wrapper if asyncio.iscoroutinefunction(func) else sync_wrapper
        return decorator
    
    def _record_metrics(self, metrics: PerformanceMetrics) -> None:
        """Record performance metrics."""
        with self._lock:
            self._metrics.append(metrics)
            if len(self._metrics) > self._max_metrics:
                # Keep only recent metrics
                self._metrics = self._metrics[-self._max_metrics//2:]
    
    def get_performance_summary(self, operation_name: Optional[str] = None, 
                               minutes: int = 60) -> Dict[str, Any]:
        """Get performance summary for operations."""
        cutoff = datetime.now() - timedelta(minutes=minutes)
        
        with self._lock:
            filtered_metrics = [
                m for m in self._metrics
                if m.timestamp >= cutoff and (operation_name is None or m.operation_name == operation_name)
            ]
        
        if not filtered_metrics:
            return {'summary': 'no_data', 'operations': 0}
        
        # Calculate statistics
        execution_times = [m.execution_time for m in filtered_metrics]
        memory_deltas = [m.memory_delta for m in filtered_metrics]
        
        return {
            'operations': len(filtered_metrics),
            'time_period_minutes': minutes,
            'operation_filter': operation_name,
            'execution_time': {
                'avg': sum(execution_times) / len(execution_times),
                'min': min(execution_times),
                'max': max(execution_times),
                'total': sum(execution_times)
            },
            'memory_usage': {
                'avg_delta_mb': sum(memory_deltas) / len(memory_deltas),
                'max_delta_mb': max(memory_deltas),
                'min_delta_mb': min(memory_deltas),
                'total_delta_mb': sum(memory_deltas)
            },
            'cache_performance': self.cache.stats(),
            'concurrent_operations': {
                'avg': sum(m.concurrent_ops for m in filtered_metrics) / len(filtered_metrics),
                'max': max(m.concurrent_ops for m in filtered_metrics)
            }
        }
    
    def optimize_for_operation(self, operation_type: str) -> None:
        """Optimize system for specific operation types."""
        if operation_type == 'bulk_data':
            # Optimize for large data operations
            self._optimizations_enabled['batch_processing'] = True
            self._optimizations_enabled['memory_management'] = True
            self.cache.ttl_seconds = 60  # Shorter TTL for bulk data
            
        elif operation_type == 'frequent_access':
            # Optimize for frequent small operations
            self._optimizations_enabled['caching'] = True
            self.cache.max_size *= 2  # Larger cache
            self.cache.ttl_seconds = 600  # Longer TTL
            
        elif operation_type == 'memory_intensive':
            # Optimize for memory-intensive operations
            self._optimizations_enabled['memory_management'] = True
            self.cache.max_size //= 2  # Smaller cache
            
        elif operation_type == 'cpu_intensive':
            # Optimize for CPU-intensive operations
            self._optimizations_enabled['async_execution'] = True
            # Consider using process pool for CPU-bound tasks
    
    def clear_caches(self) -> Dict[str, Any]:
        """Clear all performance caches."""
        stats_before = self.cache.stats()
        self.cache.clear()
        
        with self._lock:
            metrics_cleared = len(self._metrics)
            self._metrics.clear()
        
        return {
            'cache_stats_before': stats_before,
            'metrics_cleared': metrics_cleared,
            'timestamp': datetime.now().isoformat()
        }
    
    def shutdown(self) -> None:
        """Shutdown performance optimizer."""
        self.execution_pool.shutdown()
        self.cache.clear()


# Global performance optimizer instance
performance_optimizer = PerformanceOptimizer()


# Convenience decorators
def cached(cache_key_func: Optional[Callable] = None, ttl_seconds: Optional[float] = None):
    """Convenience decorator for caching."""
    return performance_optimizer.cached_operation(cache_key_func, ttl_seconds)


def monitored(operation_name: str):
    """Convenience decorator for performance monitoring."""
    return performance_optimizer.performance_monitor(operation_name)


async def optimize_async(func: Callable, *args, timeout: Optional[float] = None, **kwargs):
    """Convenience function for async optimization."""
    return await performance_optimizer.async_operation(func, *args, timeout=timeout, **kwargs)