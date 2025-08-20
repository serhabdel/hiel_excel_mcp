"""
Comprehensive tests for monitoring and observability system.
"""

import pytest
import asyncio
import time
from datetime import datetime, timedelta
from unittest.mock import Mock, patch, MagicMock

from core.monitoring import (
    MetricsCollector, HealthChecker, MonitoringSystem,
    MetricType, HealthStatus, MetricPoint, HealthCheck,
    monitored
)


class TestMetricsCollector:
    """Test MetricsCollector functionality."""
    
    @pytest.fixture
    def collector(self):
        """Create a fresh MetricsCollector instance."""
        return MetricsCollector(max_points=100, retention_hours=1)
    
    def test_record_counter_metric(self, collector):
        """Test recording counter metrics."""
        collector.record_metric("test_counter", 1, MetricType.COUNTER)
        collector.record_metric("test_counter", 2, MetricType.COUNTER)
        
        summary = collector.get_metric_summary("test_counter")
        assert summary["latest"] == 3  # Counter should accumulate
    
    def test_record_gauge_metric(self, collector):
        """Test recording gauge metrics."""
        collector.record_metric("test_gauge", 10, MetricType.GAUGE)
        collector.record_metric("test_gauge", 20, MetricType.GAUGE)
        
        summary = collector.get_metric_summary("test_gauge")
        assert summary["latest"] == 20  # Gauge should be replaced
    
    def test_record_histogram_metric(self, collector):
        """Test recording histogram metrics."""
        values = [10, 20, 30, 40, 50]
        for val in values:
            collector.record_metric("test_histogram", val, MetricType.HISTOGRAM)
        
        summary = collector.get_metric_summary("test_histogram")
        assert summary["count"] == len(values)
        assert summary["min"] == min(values)
        assert summary["max"] == max(values)
        assert summary["avg"] == sum(values) / len(values)
    
    def test_metric_with_labels(self, collector):
        """Test metrics with labels."""
        collector.record_metric("labeled_metric", 10, MetricType.GAUGE, 
                              labels={"service": "excel", "environment": "prod"})
        collector.record_metric("labeled_metric", 20, MetricType.GAUGE,
                              labels={"service": "excel", "environment": "dev"})
        
        prod_summary = collector.get_metric_summary("labeled_metric", 
                                                   labels={"service": "excel", "environment": "prod"})
        dev_summary = collector.get_metric_summary("labeled_metric",
                                                  labels={"service": "excel", "environment": "dev"})
        
        assert prod_summary["latest"] == 10
        assert dev_summary["latest"] == 20
    
    def test_get_metrics_with_pattern(self, collector):
        """Test getting metrics with pattern matching."""
        collector.record_metric("app_request_count", 1, MetricType.COUNTER)
        collector.record_metric("app_response_time", 100, MetricType.TIMER)
        collector.record_metric("system_memory", 512, MetricType.GAUGE)
        
        app_metrics = collector.get_metrics("app_")
        assert len(app_metrics) == 2
        
        all_metrics = collector.get_metrics()
        assert len(all_metrics) == 3
    
    def test_get_metrics_since_timestamp(self, collector):
        """Test getting metrics since a specific timestamp."""
        collector.record_metric("test_metric", 1, MetricType.COUNTER)
        time.sleep(0.1)
        since_time = datetime.now()
        time.sleep(0.1)
        collector.record_metric("test_metric", 2, MetricType.COUNTER)
        
        recent_metrics = collector.get_metrics("test_metric", since=since_time)
        assert len(recent_metrics) == 1
        assert recent_metrics[0].value == 3  # Accumulated counter value


class TestHealthChecker:
    """Test HealthChecker functionality."""
    
    @pytest.fixture
    def checker(self):
        """Create a fresh HealthChecker instance."""
        return HealthChecker()
    
    def test_register_custom_check(self, checker):
        """Test registering custom health check."""
        def custom_check():
            return HealthCheck(
                name="custom",
                status=HealthStatus.HEALTHY,
                message="Custom check passed"
            )
        
        checker.register_check("custom", custom_check)
        result = checker.run_check("custom")
        
        assert result.name == "custom"
        assert result.status == HealthStatus.HEALTHY
        assert result.message == "Custom check passed"
    
    def test_run_nonexistent_check(self, checker):
        """Test running a non-existent health check."""
        result = checker.run_check("nonexistent")
        
        assert result.status == HealthStatus.UNKNOWN
        assert "not found" in result.message
    
    def test_run_failing_check(self, checker):
        """Test handling of failing health check."""
        def failing_check():
            raise Exception("Check failed")
        
        checker.register_check("failing", failing_check)
        result = checker.run_check("failing")
        
        assert result.status == HealthStatus.CRITICAL
        assert "failed" in result.message.lower()
    
    def test_run_all_checks(self, checker):
        """Test running all registered checks."""
        def healthy_check():
            return HealthCheck("healthy", HealthStatus.HEALTHY, "OK")
        
        def warning_check():
            return HealthCheck("warning", HealthStatus.WARNING, "Warning")
        
        checker.register_check("healthy", healthy_check)
        checker.register_check("warning", warning_check)
        
        results = checker.run_all_checks()
        
        assert len(results) >= 2  # At least our custom checks
        assert results["healthy"].status == HealthStatus.HEALTHY
        assert results["warning"].status == HealthStatus.WARNING
    
    def test_overall_health_status(self, checker):
        """Test overall health status calculation."""
        def healthy_check():
            return HealthCheck("h1", HealthStatus.HEALTHY, "OK")
        
        def warning_check():
            return HealthCheck("w1", HealthStatus.WARNING, "Warning")
        
        def critical_check():
            return HealthCheck("c1", HealthStatus.CRITICAL, "Critical")
        
        # Test with only healthy checks
        checker._checks = {"h1": healthy_check}
        assert checker.get_overall_health() == HealthStatus.HEALTHY
        
        # Test with warning
        checker._checks = {"h1": healthy_check, "w1": warning_check}
        assert checker.get_overall_health() == HealthStatus.WARNING
        
        # Test with critical (takes precedence)
        checker._checks = {"h1": healthy_check, "w1": warning_check, "c1": critical_check}
        assert checker.get_overall_health() == HealthStatus.CRITICAL
    
    @patch('psutil.disk_usage')
    def test_disk_space_check(self, mock_disk_usage, checker):
        """Test disk space health check."""
        # Mock disk usage with low space
        mock_usage = Mock()
        mock_usage.free = 1024 * 1024 * 100  # 100MB free
        mock_usage.total = 1024 * 1024 * 1024 * 10  # 10GB total
        mock_usage.used = mock_usage.total - mock_usage.free
        mock_disk_usage.return_value = mock_usage
        
        result = checker._check_disk_space()
        
        assert result.status == HealthStatus.CRITICAL
        assert "Critical disk space" in result.message


class TestMonitoringSystem:
    """Test MonitoringSystem functionality."""
    
    @pytest.fixture
    def monitoring(self):
        """Create a fresh MonitoringSystem instance."""
        system = MonitoringSystem()
        system._monitoring_active = False  # Disable background thread
        return system
    
    def test_record_operation(self, monitoring):
        """Test recording operation metrics."""
        monitoring.record_operation("test_op", 1.5, success=True, 
                                   labels={"user": "test"})
        
        # Check that metrics were recorded
        duration_metrics = monitoring.metrics.get_metrics("operation_duration")
        count_metrics = monitoring.metrics.get_metrics("operation_count")
        
        assert len(duration_metrics) > 0
        assert len(count_metrics) > 0
        assert duration_metrics[0].value == 1.5
    
    def test_record_failed_operation(self, monitoring):
        """Test recording failed operation metrics."""
        monitoring.record_operation("test_op", 2.0, success=False)
        
        error_metrics = monitoring.metrics.get_metrics("operation_errors")
        assert len(error_metrics) > 0
        assert error_metrics[0].value == 1
    
    def test_create_alert(self, monitoring):
        """Test alert creation."""
        monitoring._create_alert("warning", "Test alert", {"detail": "test"})
        
        assert len(monitoring._alerts) == 1
        alert = monitoring._alerts[0]
        assert alert["severity"] == "warning"
        assert alert["message"] == "Test alert"
        assert alert["details"]["detail"] == "test"
    
    def test_alert_limit(self, monitoring):
        """Test alert limit enforcement."""
        monitoring._max_alerts = 10
        
        # Create more alerts than the limit
        for i in range(15):
            monitoring._create_alert("info", f"Alert {i}")
        
        # Should keep only half when limit exceeded
        assert len(monitoring._alerts) <= monitoring._max_alerts
    
    def test_get_recent_alerts(self, monitoring):
        """Test getting recent alerts."""
        # Create old alert
        old_alert = {
            "timestamp": (datetime.now() - timedelta(hours=2)).isoformat(),
            "severity": "info",
            "message": "Old alert",
            "details": {},
            "id": "old"
        }
        
        # Create recent alert
        monitoring._create_alert("warning", "Recent alert")
        
        monitoring._alerts.insert(0, old_alert)
        
        recent = monitoring._get_recent_alerts(hours=1)
        
        assert len(recent) == 1
        assert recent[0]["message"] == "Recent alert"
    
    def test_calculate_error_rate(self, monitoring):
        """Test error rate calculation."""
        # Record some operations
        for i in range(10):
            monitoring.record_operation("test", 1.0, success=(i < 8))  # 80% success rate
        
        error_rate_info = monitoring._calculate_error_rate()
        
        assert error_rate_info["total_operations"] == 10
        assert error_rate_info["errors"] == 2
        assert error_rate_info["error_rate"] == 0.2
    
    def test_get_dashboard_data(self, monitoring):
        """Test dashboard data generation."""
        # Record some metrics
        monitoring.record_operation("test", 1.0, success=True)
        
        dashboard = monitoring.get_dashboard_data()
        
        assert "timestamp" in dashboard
        assert "system_health" in dashboard
        assert "metrics_summary" in dashboard
        assert "performance" in dashboard
        assert "memory" in dashboard
        assert "alerts" in dashboard
    
    def test_export_metrics_json(self, monitoring):
        """Test exporting metrics in JSON format."""
        monitoring.metrics.record_metric("test", 42, MetricType.GAUGE)
        
        json_export = monitoring.export_metrics("json")
        data = eval(json_export)  # Safe since we control the output
        
        assert isinstance(data, list)
        assert len(data) > 0
        assert data[0]["name"] == "test"
        assert data[0]["value"] == 42
    
    def test_export_metrics_prometheus(self, monitoring):
        """Test exporting metrics in Prometheus format."""
        monitoring.metrics.record_metric("test_metric", 42, MetricType.GAUGE,
                                        labels={"env": "prod"})
        
        prometheus_export = monitoring.export_metrics("prometheus")
        
        assert "# HELP test_metric" in prometheus_export
        assert "# TYPE test_metric gauge" in prometheus_export
        assert 'test_metric{env="prod"} 42' in prometheus_export


def test_monitored_decorator():
    """Test the @monitored decorator."""
    mock_monitoring = Mock()
    
    with patch('core.monitoring.monitoring_system', mock_monitoring):
        @monitored("test_operation", labels={"test": "true"})
        def test_function():
            time.sleep(0.01)
            return "result"
        
        result = test_function()
        
        assert result == "result"
        mock_monitoring.record_operation.assert_called_once()
        
        call_args = mock_monitoring.record_operation.call_args
        assert call_args[0][0] == "test_operation"
        assert call_args[0][1] > 0  # Duration
        assert call_args[0][2] == True  # Success
        assert call_args[0][3] == {"test": "true"}  # Labels


def test_monitored_decorator_with_exception():
    """Test the @monitored decorator with exception."""
    mock_monitoring = Mock()
    
    with patch('core.monitoring.monitoring_system', mock_monitoring):
        @monitored("failing_operation")
        def failing_function():
            raise ValueError("Test error")
        
        with pytest.raises(ValueError):
            failing_function()
        
        mock_monitoring.record_operation.assert_called_once()
        
        call_args = mock_monitoring.record_operation.call_args
        assert call_args[0][0] == "failing_operation"
        assert call_args[0][2] == False  # Success = False
