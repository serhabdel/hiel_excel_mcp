"""
Comprehensive monitoring and observability module for Hiel Excel MCP.
Provides metrics collection, health checks, alerting, and observability features.
"""

import json
import time
import asyncio
import threading
from typing import Dict, Any, List, Optional, Callable, Union
from datetime import datetime, timedelta
from dataclasses import dataclass, asdict, field
from enum import Enum
from functools import wraps
from collections import defaultdict, deque
import logging
import psutil
import os
from pathlib import Path

from .config import config
from .memory_optimizer import memory_optimizer
from .performance_optimizer import performance_optimizer

logger = logging.getLogger(__name__)


class MetricType(Enum):
    """Types of metrics we can collect."""
    COUNTER = "counter"
    GAUGE = "gauge"
    HISTOGRAM = "histogram"
    TIMER = "timer"


class HealthStatus(Enum):
    """Health check status levels."""
    HEALTHY = "healthy"
    WARNING = "warning"
    CRITICAL = "critical"
    UNKNOWN = "unknown"


@dataclass
class MetricPoint:
    """Individual metric data point."""
    name: str
    value: Union[int, float]
    type: MetricType
    labels: Dict[str, str] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)


@dataclass
class HealthCheck:
    """Health check result."""
    name: str
    status: HealthStatus
    message: str
    details: Dict[str, Any] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)
    check_duration_ms: float = 0.0


class MetricsCollector:
    """Advanced metrics collection and aggregation system."""
    
    def __init__(self, max_points: int = 10000, retention_hours: int = 24):
        self.max_points = max_points
        self.retention_hours = retention_hours
        self._metrics: Dict[str, deque] = defaultdict(lambda: deque(maxlen=max_points))
        self._counters: Dict[str, float] = defaultdict(float)
        self._gauges: Dict[str, float] = defaultdict(float)
        self._histograms: Dict[str, List[float]] = defaultdict(list)
        self._lock = threading.RLock()
        
        # Start cleanup thread
        self._cleanup_thread = threading.Thread(target=self._cleanup_old_metrics, daemon=True)
        self._cleanup_thread.start()
    
    def record_metric(
        self,
        name: str,
        value: Union[int, float],
        metric_type: MetricType,
        labels: Optional[Dict[str, str]] = None
    ) -> None:
        """Record a metric data point."""
        labels = labels or {}
        metric_key = f"{name}:{json.dumps(labels, sort_keys=True)}"
        point = MetricPoint(
            name=name,
            value=value,
            type=metric_type,
            labels=labels
        )
        
        with self._lock:
            if metric_type == MetricType.COUNTER:
                self._counters[metric_key] += value
                current_value = self._counters[metric_key]
            elif metric_type == MetricType.GAUGE:
                self._gauges[metric_key] = value
                current_value = value
            elif metric_type == MetricType.HISTOGRAM:
                self._histograms[metric_key].append(value)
                # Keep only last 1000 values for histograms
                if len(self._histograms[metric_key]) > 1000:
                    self._histograms[metric_key] = self._histograms[metric_key][-1000:]
                current_value = value
            else:  # TIMER
                current_value = value
            
            point = MetricPoint(
                name=name,
                value=current_value,
                metric_type=metric_type,
                labels=labels
            )
            
            self._metrics[metric_key].append(point)
    
    def get_metrics(self, name_pattern: Optional[str] = None,
                   since: Optional[datetime] = None) -> List[MetricPoint]:
        """Get metrics matching pattern and time range."""
        results = []
        
        with self._lock:
            for metric_key, points in self._metrics.items():
                if name_pattern and name_pattern not in metric_key:
                    continue
                
                for point in points:
                    if since and point.timestamp < since:
                        continue
                    results.append(point)
        
        return sorted(results, key=lambda p: p.timestamp)
    
    def get_metric_summary(self, name: str, 
                          labels: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Get statistical summary of a metric."""
        labels = labels or {}
        metric_key = f"{name}:{json.dumps(labels, sort_keys=True)}"
        
        with self._lock:
            points = list(self._metrics.get(metric_key, []))
        
        if not points:
            return {"name": name, "labels": labels, "summary": "no_data"}
        
        values = [p.value for p in points]
        
        return {
            "name": name,
            "labels": labels,
            "count": len(values),
            "min": min(values),
            "max": max(values),
            "avg": sum(values) / len(values),
            "latest": values[-1] if values else None,
            "first_timestamp": points[0].timestamp.isoformat(),
            "last_timestamp": points[-1].timestamp.isoformat()
        }
    
    def _cleanup_old_metrics(self) -> None:
        """Background thread to clean up old metric points."""
        while True:
            try:
                cutoff = datetime.now() - timedelta(hours=self.retention_hours)
                
                with self._lock:
                    for metric_key in list(self._metrics.keys()):
                        points = self._metrics[metric_key]
                        # Remove old points
                        while points and points[0].timestamp < cutoff:
                            points.popleft()
                        
                        # Remove empty metric keys
                        if not points:
                            del self._metrics[metric_key]
                
                # Sleep for 1 hour before next cleanup
                time.sleep(3600)
                
            except Exception as e:
                logger.error(f"Error in metrics cleanup: {e}")
                time.sleep(300)  # Sleep 5 minutes on error


class HealthChecker:
    """System health monitoring and checks."""
    
    def __init__(self):
        self._checks: Dict[str, Callable[[], HealthCheck]] = {}
        self._last_results: Dict[str, HealthCheck] = {}
        self._lock = threading.RLock()
        
        # Register default health checks
        self._register_default_checks()
    
    def _register_default_checks(self) -> None:
        """Register default system health checks."""
        self.register_check("memory_usage", self._check_memory_usage)
        self.register_check("disk_space", self._check_disk_space)
        self.register_check("file_descriptors", self._check_file_descriptors)
        self.register_check("cache_health", self._check_cache_health)
        self.register_check("performance_health", self._check_performance_health)
    
    def register_check(self, name: str, check_func: Callable[[], HealthCheck]) -> None:
        """Register a custom health check."""
        with self._lock:
            self._checks[name] = check_func
    
    def run_check(self, name: str) -> HealthCheck:
        """Run a specific health check."""
        with self._lock:
            check_func = self._checks.get(name)
        
        if not check_func:
            return HealthCheck(
                name=name,
                status=HealthStatus.UNKNOWN,
                message=f"Health check '{name}' not found"
            )
        
        start_time = time.time()
        try:
            result = check_func()
            result.check_duration_ms = (time.time() - start_time) * 1000
            
            with self._lock:
                self._last_results[name] = result
            
            return result
            
        except Exception as e:
            error_result = HealthCheck(
                name=name,
                status=HealthStatus.CRITICAL,
                message=f"Health check failed: {str(e)}",
                check_duration_ms=(time.time() - start_time) * 1000
            )
            
            with self._lock:
                self._last_results[name] = error_result
            
            return error_result
    
    def run_all_checks(self) -> Dict[str, HealthCheck]:
        """Run all registered health checks."""
        results = {}
        
        with self._lock:
            check_names = list(self._checks.keys())
        
        for name in check_names:
            results[name] = self.run_check(name)
        
        return results
    
    def get_overall_health(self) -> HealthStatus:
        """Get overall system health status."""
        results = self.run_all_checks()
        
        if not results:
            return HealthStatus.UNKNOWN
        
        statuses = [result.status for result in results.values()]
        
        if HealthStatus.CRITICAL in statuses:
            return HealthStatus.CRITICAL
        elif HealthStatus.WARNING in statuses:
            return HealthStatus.WARNING
        elif all(status == HealthStatus.HEALTHY for status in statuses):
            return HealthStatus.HEALTHY
        else:
            return HealthStatus.UNKNOWN
    
    def _check_memory_usage(self) -> HealthCheck:
        """Check system memory usage."""
        try:
            stats = memory_optimizer.get_memory_stats()
            
            if stats.percent > 90:
                status = HealthStatus.CRITICAL
                message = f"Critical memory usage: {stats.percent:.1f}%"
            elif stats.percent > 75:
                status = HealthStatus.WARNING
                message = f"High memory usage: {stats.percent:.1f}%"
            else:
                status = HealthStatus.HEALTHY
                message = f"Memory usage normal: {stats.percent:.1f}%"
            
            return HealthCheck(
                name="memory_usage",
                status=status,
                message=message,
                details={
                    "rss_mb": stats.rss_mb,
                    "vms_mb": stats.vms_mb,
                    "percent": stats.percent,
                    "available_mb": stats.available_mb
                }
            )
            
        except Exception as e:
            return HealthCheck(
                name="memory_usage",
                status=HealthStatus.CRITICAL,
                message=f"Failed to check memory usage: {e}"
            )
    
    def _check_disk_space(self) -> HealthCheck:
        """Check disk space availability."""
        try:
            excel_path = Path(config.excel_files_path)
            disk_usage = psutil.disk_usage(excel_path)
            
            free_percent = (disk_usage.free / disk_usage.total) * 100
            
            if free_percent < 5:
                status = HealthStatus.CRITICAL
                message = f"Critical disk space: {free_percent:.1f}% free"
            elif free_percent < 10:
                status = HealthStatus.WARNING
                message = f"Low disk space: {free_percent:.1f}% free"
            else:
                status = HealthStatus.HEALTHY
                message = f"Disk space normal: {free_percent:.1f}% free"
            
            return HealthCheck(
                name="disk_space",
                status=status,
                message=message,
                details={
                    "total_gb": disk_usage.total / (1024**3),
                    "free_gb": disk_usage.free / (1024**3),
                    "used_gb": disk_usage.used / (1024**3),
                    "free_percent": free_percent
                }
            )
            
        except Exception as e:
            return HealthCheck(
                name="disk_space",
                status=HealthStatus.CRITICAL,
                message=f"Failed to check disk space: {e}"
            )
    
    def _check_file_descriptors(self) -> HealthCheck:
        """Check file descriptor usage."""
        try:
            process = psutil.Process()
            open_files = len(process.open_files())
            
            # Most systems have a limit around 1024-4096
            if open_files > 900:
                status = HealthStatus.CRITICAL
                message = f"Critical file descriptor usage: {open_files}"
            elif open_files > 500:
                status = HealthStatus.WARNING
                message = f"High file descriptor usage: {open_files}"
            else:
                status = HealthStatus.HEALTHY
                message = f"File descriptor usage normal: {open_files}"
            
            return HealthCheck(
                name="file_descriptors",
                status=status,
                message=message,
                details={"open_files": open_files}
            )
            
        except Exception as e:
            return HealthCheck(
                name="file_descriptors",
                status=HealthStatus.WARNING,
                message=f"Could not check file descriptors: {e}"
            )
    
    def _check_cache_health(self) -> HealthCheck:
        """Check cache system health."""
        try:
            cache_stats = performance_optimizer.cache.stats()
            
            hit_rate = cache_stats.get('hit_rate', 0)
            cache_size = cache_stats.get('size', 0)
            max_size = cache_stats.get('max_size', 1)
            
            utilization = (cache_size / max_size) * 100
            
            if hit_rate < 0.1:  # Less than 10% hit rate
                status = HealthStatus.WARNING
                message = f"Low cache hit rate: {hit_rate:.1%}"
            elif utilization > 95:
                status = HealthStatus.WARNING
                message = f"Cache nearly full: {utilization:.1f}%"
            else:
                status = HealthStatus.HEALTHY
                message = f"Cache healthy: {hit_rate:.1%} hit rate, {utilization:.1f}% full"
            
            return HealthCheck(
                name="cache_health",
                status=status,
                message=message,
                details=cache_stats
            )
            
        except Exception as e:
            return HealthCheck(
                name="cache_health",
                status=HealthStatus.WARNING,
                message=f"Could not check cache health: {e}"
            )
    
    def _check_performance_health(self) -> HealthCheck:
        """Check performance metrics health."""
        try:
            perf_summary = performance_optimizer.get_performance_summary(minutes=15)
            
            if perf_summary.get('summary') == 'no_data':
                return HealthCheck(
                    name="performance_health",
                    status=HealthStatus.HEALTHY,
                    message="No recent performance data (system idle)"
                )
            
            operations = perf_summary.get('operations', 0)
            avg_time = perf_summary.get('execution_time', {}).get('avg', 0)
            max_time = perf_summary.get('execution_time', {}).get('max', 0)
            
            if max_time > 30:  # Operations taking more than 30 seconds
                status = HealthStatus.WARNING
                message = f"Slow operations detected: max {max_time:.1f}s"
            elif avg_time > 10:  # Average more than 10 seconds
                status = HealthStatus.WARNING
                message = f"High average response time: {avg_time:.1f}s"
            else:
                status = HealthStatus.HEALTHY
                message = f"Performance healthy: {operations} ops, avg {avg_time:.1f}s"
            
            return HealthCheck(
                name="performance_health",
                status=status,
                message=message,
                details={
                    "operations_15min": operations,
                    "avg_time_seconds": avg_time,
                    "max_time_seconds": max_time
                }
            )
            
        except Exception as e:
            return HealthCheck(
                name="performance_health",
                status=HealthStatus.WARNING,
                message=f"Could not check performance health: {e}"
            )


class MonitoringSystem:
    """Complete monitoring and observability system."""
    
    def __init__(self):
        self.metrics = MetricsCollector()
        self.health = HealthChecker()
        self._alerts: List[Dict[str, Any]] = []
        self._max_alerts = 100
        self._lock = threading.RLock()
        
        # Start monitoring loop
        self._monitoring_active = True
        self._monitoring_thread = threading.Thread(target=self._monitoring_loop, daemon=True)
        self._monitoring_thread.start()
    
    def _monitoring_loop(self) -> None:
        """Background monitoring loop."""
        while self._monitoring_active:
            try:
                # Collect system metrics every 30 seconds
                self._collect_system_metrics()
                
                # Run health checks every 2 minutes
                if int(time.time()) % 120 == 0:
                    self._run_health_monitoring()
                
                time.sleep(30)
                
            except Exception as e:
                logger.error(f"Error in monitoring loop: {e}")
                time.sleep(60)  # Sleep longer on error
    
    def _collect_system_metrics(self) -> None:
        """Collect system-level metrics."""
        try:
            # Memory metrics
            memory_stats = memory_optimizer.get_memory_stats()
            self.metrics.record_metric("system_memory_rss_mb", memory_stats.rss_mb, MetricType.GAUGE)
            self.metrics.record_metric("system_memory_percent", memory_stats.percent, MetricType.GAUGE)
            self.metrics.record_metric("system_memory_available_mb", memory_stats.available_mb, MetricType.GAUGE)
            
            # Cache metrics
            cache_stats = performance_optimizer.cache.stats()
            self.metrics.record_metric("cache_hit_rate", cache_stats.get('hit_rate', 0), MetricType.GAUGE)
            self.metrics.record_metric("cache_size", cache_stats.get('size', 0), MetricType.GAUGE)
            self.metrics.record_metric("cache_hits", cache_stats.get('hits', 0), MetricType.COUNTER)
            self.metrics.record_metric("cache_misses", cache_stats.get('misses', 0), MetricType.COUNTER)
            
            # Process metrics
            process = psutil.Process()
            self.metrics.record_metric("process_cpu_percent", process.cpu_percent(), MetricType.GAUGE)
            self.metrics.record_metric("process_threads", process.num_threads(), MetricType.GAUGE)
            self.metrics.record_metric("process_open_files", len(process.open_files()), MetricType.GAUGE)
            
        except Exception as e:
            logger.warning(f"Failed to collect system metrics: {e}")
    
    def _run_health_monitoring(self) -> None:
        """Run health checks and generate alerts."""
        try:
            health_results = self.health.run_all_checks()
            
            for name, result in health_results.items():
                # Record health status as metric
                status_value = {
                    HealthStatus.HEALTHY: 1,
                    HealthStatus.WARNING: 2,
                    HealthStatus.CRITICAL: 3,
                    HealthStatus.UNKNOWN: 0
                }[result.status]
                
                self.metrics.record_metric(
                    "health_check_status", 
                    status_value, 
                    MetricType.GAUGE,
                    labels={"check": name}
                )
                
                # Generate alerts for non-healthy status
                if result.status in [HealthStatus.WARNING, HealthStatus.CRITICAL]:
                    self._create_alert(
                        severity=result.status.value,
                        message=f"Health check '{name}': {result.message}",
                        details=result.details
                    )
                    
        except Exception as e:
            logger.error(f"Failed to run health monitoring: {e}")
    
    def _create_alert(self, severity: str, message: str, details: Optional[Dict[str, Any]] = None) -> None:
        """Create an alert."""
        alert = {
            "timestamp": datetime.now().isoformat(),
            "severity": severity,
            "message": message,
            "details": details or {},
            "id": f"alert_{int(time.time() * 1000)}"
        }
        
        with self._lock:
            self._alerts.append(alert)
            if len(self._alerts) > self._max_alerts:
                self._alerts = self._alerts[-self._max_alerts//2:]  # Keep half
        
        logger.warning(f"Alert [{severity}]: {message}")
    
    def record_operation(self, operation_name: str, duration_seconds: float, 
                        success: bool, labels: Optional[Dict[str, str]] = None) -> None:
        """Record operation metrics."""
        labels = labels or {}
        labels['operation'] = operation_name
        labels['success'] = str(success)
        
        self.metrics.record_metric("operation_duration", duration_seconds, MetricType.TIMER, labels)
        self.metrics.record_metric("operation_count", 1, MetricType.COUNTER, labels)
        
        if not success:
            self.metrics.record_metric("operation_errors", 1, MetricType.COUNTER, labels)
    
    def get_dashboard_data(self) -> Dict[str, Any]:
        """Get comprehensive dashboard data."""
        return {
            "timestamp": datetime.now().isoformat(),
            "system_health": {
                "overall_status": self.health.get_overall_health().value,
                "checks": {
                    name: asdict(result) 
                    for name, result in self.health.run_all_checks().items()
                }
            },
            "metrics_summary": {
                "memory_usage": self.metrics.get_metric_summary("system_memory_percent"),
                "cache_performance": self.metrics.get_metric_summary("cache_hit_rate"),
                "operation_count": self.metrics.get_metric_summary("operation_count"),
                "error_rate": self._calculate_error_rate()
            },
            "performance": performance_optimizer.get_performance_summary(minutes=60),
            "memory": {
                "current_stats": asdict(memory_optimizer.get_memory_stats()),
                "trend": memory_optimizer.get_memory_trend(minutes=60)
            },
            "alerts": {
                "recent": self._get_recent_alerts(hours=1),
                "total_count": len(self._alerts)
            }
        }
    
    def _calculate_error_rate(self) -> Dict[str, Any]:
        """Calculate current error rate."""
        since = datetime.now() - timedelta(minutes=15)
        
        total_ops = self.metrics.get_metrics("operation_count", since=since)
        error_ops = self.metrics.get_metrics("operation_errors", since=since)
        
        if not total_ops:
            return {"error_rate": 0, "total_operations": 0, "errors": 0}
        
        total_count = sum(p.value for p in total_ops)
        error_count = sum(p.value for p in error_ops)
        
        error_rate = (error_count / total_count) if total_count > 0 else 0
        
        return {
            "error_rate": error_rate,
            "total_operations": int(total_count),
            "errors": int(error_count)
        }
    
    def _get_recent_alerts(self, hours: int = 1) -> List[Dict[str, Any]]:
        """Get recent alerts."""
        cutoff = datetime.now() - timedelta(hours=hours)
        
        with self._lock:
            recent = []
            for alert in self._alerts:
                alert_time = datetime.fromisoformat(alert['timestamp'])
                if alert_time >= cutoff:
                    recent.append(alert)
            
            return sorted(recent, key=lambda a: a['timestamp'], reverse=True)
    
    def export_metrics(self, format: str = "json") -> str:
        """Export metrics in various formats."""
        all_metrics = self.metrics.get_metrics()
        
        if format == "json":
            # Convert metrics to JSON-serializable format
            metrics_data = []
            for m in all_metrics:
                metric_dict = asdict(m)
                # Convert MetricType enum to string
                if 'type' in metric_dict:
                    metric_dict['type'] = metric_dict['type'].value
                metrics_data.append(metric_dict)
            return json.dumps(metrics_data, indent=2)
        elif format == "prometheus":
            return self._format_prometheus_metrics(all_metrics)
        else:
            raise ValueError(f"Unsupported export format: {format}")
    
    def _format_prometheus_metrics(self, metrics: List[MetricPoint]) -> str:
        """Format metrics in Prometheus format."""
        lines = []
        grouped = defaultdict(list)
        
        # Group metrics by name
        for metric in metrics:
            grouped[metric.name].append(metric)
        
        for name, metric_list in grouped.items():
            lines.append(f"# HELP {name} {name}")
            lines.append(f"# TYPE {name} {metric_list[0].metric_type.value}")
            
            for metric in metric_list:
                labels_str = ""
                if metric.labels:
                    label_pairs = [f'{k}="{v}"' for k, v in metric.labels.items()]
                    labels_str = "{" + ",".join(label_pairs) + "}"
                
                timestamp = int(metric.timestamp.timestamp() * 1000)
                lines.append(f"{name}{labels_str} {metric.value} {timestamp}")
            
            lines.append("")
        
        return "\n".join(lines)
    
    def shutdown(self) -> None:
        """Shutdown monitoring system."""
        self._monitoring_active = False
        if self._monitoring_thread.is_alive():
            self._monitoring_thread.join(timeout=5)


def monitored(operation_name: str, labels: Optional[Dict[str, str]] = None):
    """Decorator to automatically monitor operations."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            start_time = time.time()
            success = True
            
            try:
                result = func(*args, **kwargs)
                return result
            except Exception as e:
                success = False
                raise
            finally:
                duration = time.time() - start_time
                monitoring_system.record_operation(
                    operation_name, duration, success, labels
                )
        
        return wrapper
    return decorator


# Global monitoring system instance
monitoring_system = MonitoringSystem()