"""
Analysis Manager Tool for hiel_excel_mcp.

Provides comprehensive analysis and visualization operations including chart creation,
pivot table generation, Excel table creation, and data analysis functionality.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, List, Union
import statistics
import math

from ..core.base_tool import BaseTool, operation_route, OperationResponse, create_success_response, create_error_response
from ..core.workbook_context import workbook_context

# Import existing functionality
import sys
import os

# Add the src directory to the path to import existing modules
src_path = os.path.join(os.path.dirname(__file__), '..', '..', 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)


logger = logging.getLogger(__name__)


class AnalysisManager(BaseTool):
    """
    Comprehensive analysis and visualization management tool.
    
    Handles chart creation, pivot table generation, Excel table creation,
    and data analysis operations.
    """
    
    def get_tool_name(self) -> str:
        return "analysis_manager"
    
    def get_tool_description(self) -> str:
        return "Comprehensive analysis and visualization management tool for charts, pivot tables, and data analysis"
    
    @operation_route(
        name="create_chart",
        description="Create various types of charts in Excel worksheets",
        required_params=["filepath", "sheet_name", "data_range", "chart_type", "target_cell"],
        optional_params=["title", "x_axis", "y_axis", "style"]
    )
    def create_chart(self, filepath: str, sheet_name: str, data_range: str, 
                    chart_type: str, target_cell: str, title: str = "",
                    x_axis: str = "", y_axis: str = "", 
                    style: Optional[Dict[str, Any]] = None, **kwargs) -> OperationResponse:
        """
        Create a chart in an Excel worksheet.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            data_range: Range of data for the chart (e.g., "A1:C10")
            chart_type: Type of chart (line, bar, pie, scatter, area)
            target_cell: Cell where chart should be placed
            title: Chart title (optional)
            x_axis: X-axis label (optional)
            y_axis: Y-axis label (optional)
            style: Chart styling options (optional)
            
        Returns:
            OperationResponse with chart creation results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Set default style if none provided
            if style is None:
                style = {
                    "show_legend": True,
                    "show_data_labels": True,
                    "legend_position": "r"
                }
            
            # Create chart using existing functionality
            result = create_chart_in_sheet(
                validated_path, sheet_name, data_range, chart_type, 
                target_cell, title, x_axis, y_axis, style
            )
            
            return create_success_response(
                operation="create_chart",
                message=result["message"],
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "chart_type": chart_type,
                    "data_range": data_range,
                    "target_cell": target_cell,
                    "title": title,
                    "details": result.get("details", {})
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to create chart: {e}")
            return create_error_response("create_chart", e)
    
    @operation_route(
        name="create_pivot_table",
        description="Create pivot tables for data analysis and summarization",
        required_params=["filepath", "sheet_name", "data_range", "rows", "values"],
        optional_params=["columns", "agg_func"]
    )
    def create_pivot_table(self, filepath: str, sheet_name: str, data_range: str,
                          rows: List[str], values: List[str], 
                          columns: Optional[List[str]] = None,
                          agg_func: str = "sum", **kwargs) -> OperationResponse:
        """
        Create a pivot table for data analysis.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet containing source data
            data_range: Source data range (e.g., "A1:D100")
            rows: List of fields for row labels
            values: List of fields for values/aggregation
            columns: Optional list of fields for column labels
            agg_func: Aggregation function (sum, count, average, max, min)
            
        Returns:
            OperationResponse with pivot table creation results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Create pivot table using existing functionality
            result = create_pivot_table(
                validated_path, sheet_name, data_range, rows, values, columns, agg_func
            )
            
            return create_success_response(
                operation="create_pivot_table",
                message=result["message"],
                data={
                    "filepath": validated_path,
                    "source_sheet": sheet_name,
                    "pivot_sheet": result["details"]["pivot_sheet"],
                    "data_range": data_range,
                    "rows": rows,
                    "values": values,
                    "columns": columns or [],
                    "aggregation": agg_func,
                    "details": result.get("details", {})
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to create pivot table: {e}")
            return create_error_response("create_pivot_table", e)
    
    @operation_route(
        name="create_table",
        description="Create native Excel tables with formatting and structure",
        required_params=["filepath", "sheet_name", "data_range"],
        optional_params=["table_name", "table_style"]
    )
    def create_table(self, filepath: str, sheet_name: str, data_range: str,
                    table_name: Optional[str] = None, 
                    table_style: str = "TableStyleMedium9", **kwargs) -> OperationResponse:
        """
        Create a native Excel table.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            data_range: Range for the table (e.g., "A1:D10")
            table_name: Optional name for the table
            table_style: Visual style for the table
            
        Returns:
            OperationResponse with table creation results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Create table using existing functionality
            result = create_excel_table(
                validated_path, sheet_name, data_range, table_name, table_style
            )
            
            return create_success_response(
                operation="create_table",
                message=result["message"],
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "table_name": result["table_name"],
                    "data_range": result["range"],
                    "table_style": table_style
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to create table: {e}")
            return create_error_response("create_table", e)
    
    @operation_route(
        name="analyze_data",
        description="Perform statistical analysis on data ranges",
        required_params=["filepath", "sheet_name", "data_range"],
        optional_params=["analysis_type", "include_charts"]
    )
    def analyze_data(self, filepath: str, sheet_name: str, data_range: str,
                    analysis_type: str = "descriptive", 
                    include_charts: bool = False, **kwargs) -> OperationResponse:
        """
        Perform statistical analysis on data ranges.
        
        Args:
            filepath: Path to the Excel file
            sheet_name: Name of the worksheet
            data_range: Range of data to analyze (e.g., "A1:C100")
            analysis_type: Type of analysis (descriptive, correlation, trend)
            include_charts: Whether to create visualization charts
            
        Returns:
            OperationResponse with analysis results
        """
        try:
            # Validate path
            validated_path, warnings = PathValidator.validate_path(filepath, allow_create=False)
            
            # Parse data range
            if ':' in data_range:
                start_cell, end_cell = data_range.split(':', 1)
            else:
                start_cell, end_cell = data_range, None
            
            # Read data for analysis
            data = read_excel_range(validated_path, sheet_name, start_cell, end_cell)
            
            if not data:
                return create_success_response(
                    operation="analyze_data",
                    message="No data found in specified range",
                    data={"analysis_type": analysis_type, "data_range": data_range},
                    warnings=["No data found in specified range"]
                )
            
            # Perform analysis based on type
            analysis_results = {}
            
            if analysis_type.lower() == "descriptive":
                analysis_results = self._perform_descriptive_analysis(data)
            elif analysis_type.lower() == "correlation":
                analysis_results = self._perform_correlation_analysis(data)
            elif analysis_type.lower() == "trend":
                analysis_results = self._perform_trend_analysis(data)
            else:
                # Default to descriptive
                analysis_results = self._perform_descriptive_analysis(data)
                analysis_results["note"] = f"Unknown analysis type '{analysis_type}', performed descriptive analysis"
            
            # Create charts if requested
            charts_created = []
            if include_charts and analysis_type.lower() == "descriptive":
                charts_created = self._create_analysis_charts(
                    validated_path, sheet_name, data_range, analysis_results
                )
            
            return create_success_response(
                operation="analyze_data",
                message=f"Data analysis completed using {analysis_type} method",
                data={
                    "filepath": validated_path,
                    "sheet_name": sheet_name,
                    "data_range": data_range,
                    "analysis_type": analysis_type,
                    "analysis_results": analysis_results,
                    "charts_created": charts_created,
                    "data_summary": {
                        "total_rows": len(data),
                        "total_columns": len(data[0]) if data else 0,
                        "has_headers": self._detect_headers(data)
                    }
                },
                warnings=warnings if warnings else None
            )
            
        except Exception as e:
            logger.error(f"Failed to analyze data: {e}")
            return create_error_response("analyze_data", e)
    
    def _perform_descriptive_analysis(self, data: List[List]) -> Dict[str, Any]:
        """Perform descriptive statistical analysis on data."""
        if not data or len(data) < 2:
            return {"error": "Insufficient data for analysis"}
        
        # Assume first row contains headers
        headers = data[0]
        numeric_data = data[1:]
        
        results = {
            "columns": len(headers),
            "rows": len(numeric_data),
            "column_analysis": {}
        }
        
        # Analyze each column
        for col_idx, header in enumerate(headers):
            column_values = []
            
            # Extract numeric values from column
            for row in numeric_data:
                if col_idx < len(row):
                    try:
                        value = float(row[col_idx]) if row[col_idx] is not None else None
                        if value is not None and not math.isnan(value):
                            column_values.append(value)
                    except (ValueError, TypeError):
                        # Skip non-numeric values
                        continue
            
            if column_values:
                col_stats = {
                    "count": len(column_values),
                    "mean": statistics.mean(column_values),
                    "median": statistics.median(column_values),
                    "min": min(column_values),
                    "max": max(column_values),
                    "range": max(column_values) - min(column_values)
                }
                
                # Add standard deviation if we have enough data points
                if len(column_values) > 1:
                    col_stats["std_dev"] = statistics.stdev(column_values)
                    col_stats["variance"] = statistics.variance(column_values)
                
                results["column_analysis"][str(header)] = col_stats
            else:
                results["column_analysis"][str(header)] = {
                    "count": 0,
                    "note": "No numeric data found"
                }
        
        return results
    
    def _perform_correlation_analysis(self, data: List[List]) -> Dict[str, Any]:
        """Perform correlation analysis between numeric columns."""
        if not data or len(data) < 3:  # Need at least header + 2 data rows
            return {"error": "Insufficient data for correlation analysis"}
        
        headers = data[0]
        numeric_data = data[1:]
        
        # Extract numeric columns
        numeric_columns = {}
        for col_idx, header in enumerate(headers):
            column_values = []
            for row in numeric_data:
                if col_idx < len(row):
                    try:
                        value = float(row[col_idx]) if row[col_idx] is not None else None
                        if value is not None and not math.isnan(value):
                            column_values.append(value)
                        else:
                            column_values.append(None)
                    except (ValueError, TypeError):
                        column_values.append(None)
            
            # Only include columns with sufficient numeric data
            valid_values = [v for v in column_values if v is not None]
            if len(valid_values) >= 2:
                numeric_columns[str(header)] = column_values
        
        results = {
            "numeric_columns": list(numeric_columns.keys()),
            "correlations": {}
        }
        
        # Calculate correlations between numeric columns
        column_names = list(numeric_columns.keys())
        for i, col1 in enumerate(column_names):
            for j, col2 in enumerate(column_names[i+1:], i+1):
                # Get paired values (both non-null)
                paired_values = [
                    (v1, v2) for v1, v2 in zip(numeric_columns[col1], numeric_columns[col2])
                    if v1 is not None and v2 is not None
                ]
                
                if len(paired_values) >= 2:
                    x_values = [pair[0] for pair in paired_values]
                    y_values = [pair[1] for pair in paired_values]
                    
                    # Calculate Pearson correlation coefficient
                    try:
                        correlation = statistics.correlation(x_values, y_values)
                        results["correlations"][f"{col1} vs {col2}"] = {
                            "correlation": correlation,
                            "sample_size": len(paired_values),
                            "strength": self._interpret_correlation(correlation)
                        }
                    except statistics.StatisticsError:
                        results["correlations"][f"{col1} vs {col2}"] = {
                            "correlation": None,
                            "note": "Unable to calculate correlation"
                        }
        
        return results
    
    def _perform_trend_analysis(self, data: List[List]) -> Dict[str, Any]:
        """Perform basic trend analysis on time series or sequential data."""
        if not data or len(data) < 3:
            return {"error": "Insufficient data for trend analysis"}
        
        headers = data[0]
        numeric_data = data[1:]
        
        results = {
            "columns": len(headers),
            "rows": len(numeric_data),
            "trends": {}
        }
        
        # Analyze trends in each numeric column
        for col_idx, header in enumerate(headers):
            column_values = []
            
            for row_idx, row in enumerate(numeric_data):
                if col_idx < len(row):
                    try:
                        value = float(row[col_idx]) if row[col_idx] is not None else None
                        if value is not None and not math.isnan(value):
                            column_values.append((row_idx, value))
                    except (ValueError, TypeError):
                        continue
            
            if len(column_values) >= 3:
                # Calculate simple trend metrics
                values = [v[1] for v in column_values]
                
                # Calculate differences between consecutive values
                differences = [values[i+1] - values[i] for i in range(len(values)-1)]
                
                trend_info = {
                    "data_points": len(values),
                    "start_value": values[0],
                    "end_value": values[-1],
                    "total_change": values[-1] - values[0],
                    "average_change": statistics.mean(differences) if differences else 0,
                    "direction": "increasing" if values[-1] > values[0] else "decreasing" if values[-1] < values[0] else "stable"
                }
                
                # Add volatility measure
                if len(differences) > 1:
                    trend_info["volatility"] = statistics.stdev(differences)
                
                results["trends"][str(header)] = trend_info
        
        return results
    
    def _interpret_correlation(self, correlation: float) -> str:
        """Interpret correlation coefficient strength."""
        abs_corr = abs(correlation)
        if abs_corr >= 0.8:
            return "very strong"
        elif abs_corr >= 0.6:
            return "strong"
        elif abs_corr >= 0.4:
            return "moderate"
        elif abs_corr >= 0.2:
            return "weak"
        else:
            return "very weak"
    
    def _detect_headers(self, data: List[List]) -> bool:
        """Detect if first row contains headers."""
        if not data or len(data) < 2:
            return False
        
        first_row = data[0]
        second_row = data[1] if len(data) > 1 else []
        
        # Simple heuristic: if first row has more strings than second row, likely headers
        first_row_strings = sum(1 for cell in first_row if isinstance(cell, str) and not str(cell).replace('.', '').replace('-', '').isdigit())
        second_row_strings = sum(1 for cell in second_row if isinstance(cell, str) and not str(cell).replace('.', '').replace('-', '').isdigit())
        
        return first_row_strings > second_row_strings
    
    def _create_analysis_charts(self, filepath: str, sheet_name: str, 
                               data_range: str, analysis_results: Dict[str, Any]) -> List[Dict[str, str]]:
        """Create charts based on analysis results."""
        charts_created = []
        
        try:
            # Create a summary chart if we have column analysis
            if "column_analysis" in analysis_results:
                numeric_columns = [
                    col for col, stats in analysis_results["column_analysis"].items()
                    if isinstance(stats, dict) and "mean" in stats
                ]
                
                if len(numeric_columns) >= 2:
                    # Try to create a simple bar chart of means
                    try:
                        chart_result = create_chart_in_sheet(
                            filepath, sheet_name, data_range, "bar", "F2",
                            title="Data Analysis Summary", 
                            x_axis="Categories", 
                            y_axis="Values",
                            style={"show_legend": True, "show_data_labels": True}
                        )
                        charts_created.append({
                            "type": "bar",
                            "location": "F2",
                            "title": "Data Analysis Summary"
                        })
                    except Exception as e:
                        logger.warning(f"Could not create analysis chart: {e}")
        
        except Exception as e:
            logger.warning(f"Error creating analysis charts: {e}")
        
        return charts_created


# Create global instance
analysis_manager = AnalysisManager()


def analysis_manager_tool(operation: str, **kwargs) -> str:
    """
    MCP tool function for analysis management operations.
    
    Args:
        operation: The operation to perform
        **kwargs: Operation-specific parameters
        
    Returns:
        JSON string with operation results
    """
    try:
        response = analysis_manager.execute_operation(operation, **kwargs)
        return response.to_json()
    except Exception as e:
        logger.error(f"Unexpected error in analysis_manager_tool: {e}", exc_info=True)
        error_response = create_error_response(operation, e)
        return error_response.to_json()