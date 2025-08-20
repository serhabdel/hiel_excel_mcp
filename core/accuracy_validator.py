"""
Accuracy validation module for Hiel Excel MCP.
Provides comprehensive data validation, integrity checks, and accuracy verification.
"""

import re
import math
from typing import Any, Dict, List, Optional, Union, Tuple, Callable, Pattern
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import logging
from dataclasses import dataclass
from enum import Enum
import json

from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


class ValidationSeverity(Enum):
    """Severity levels for validation issues."""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


@dataclass
class ValidationIssue:
    """Represents a validation issue."""
    severity: ValidationSeverity
    code: str
    message: str
    location: Optional[str] = None
    expected: Optional[Any] = None
    actual: Optional[Any] = None
    suggestions: List[str] = None


class DataTypeValidator:
    """Validates data types and formats."""
    
    # Common regex patterns for validation
    PATTERNS = {
        'email': re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'),
        'phone': re.compile(r'^\+?[\d\s\-\(\)\.]{10,}$'),
        'url': re.compile(r'^https?://(?:[-\w.])+(?:[:\d]+)?(?:/(?:[\w/_.])*(?:\?(?:[\w&=%.])*)?(?:#(?:\w*))?)?$'),
        'ip_address': re.compile(r'^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'),
        'uuid': re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE),
        'credit_card': re.compile(r'^\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}$'),
        'ssn': re.compile(r'^\d{3}-?\d{2}-?\d{4}$'),
        'postal_code': re.compile(r'^[A-Z0-9]{3,10}$', re.IGNORECASE),
    }
    
    @classmethod
    def validate_type(cls, value: Any, expected_type: type) -> ValidationIssue:
        """Validate value against expected type."""
        if value is None:
            return ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="NULL_VALUE",
                message=f"Expected {expected_type.__name__} but got None",
                expected=expected_type.__name__,
                actual="None"
            )
        
        if not isinstance(value, expected_type):
            # Try to convert
            try:
                if expected_type == int:
                    converted = int(float(str(value)))
                elif expected_type == float:
                    converted = float(str(value))
                elif expected_type == str:
                    converted = str(value)
                elif expected_type == bool:
                    converted = bool(value)
                else:
                    raise ValueError("Cannot convert")
                
                return ValidationIssue(
                    severity=ValidationSeverity.INFO,
                    code="TYPE_CONVERTED",
                    message=f"Value converted from {type(value).__name__} to {expected_type.__name__}",
                    expected=expected_type.__name__,
                    actual=type(value).__name__,
                    suggestions=[f"Consider storing as {expected_type.__name__} directly"]
                )
                
            except (ValueError, TypeError, OverflowError):
                return ValidationIssue(
                    severity=ValidationSeverity.ERROR,
                    code="TYPE_MISMATCH",
                    message=f"Expected {expected_type.__name__} but got {type(value).__name__}",
                    expected=expected_type.__name__,
                    actual=type(value).__name__,
                    suggestions=[f"Convert value to {expected_type.__name__}", "Check data source format"]
                )
        
        return ValidationIssue(
            severity=ValidationSeverity.INFO,
            code="TYPE_VALID",
            message="Type validation passed"
        )
    
    @classmethod
    def validate_numeric_range(cls, value: Union[int, float], min_val: Optional[float] = None, 
                              max_val: Optional[float] = None) -> List[ValidationIssue]:
        """Validate numeric value is within specified range."""
        issues = []
        
        try:
            num_val = float(value)
        except (TypeError, ValueError):
            return [ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="NOT_NUMERIC",
                message=f"Value '{value}' is not numeric",
                actual=str(value),
                suggestions=["Provide a numeric value"]
            )]
        
        # Check for special float values
        if math.isnan(num_val):
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="NAN_VALUE",
                message="Value is NaN (Not a Number)",
                actual="NaN",
                suggestions=["Replace with a valid number", "Check calculation source"]
            ))
            return issues
        
        if math.isinf(num_val):
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INFINITE_VALUE",
                message="Value is infinite",
                actual="Infinity",
                suggestions=["Check for division by zero", "Verify calculation logic"]
            ))
            return issues
        
        # Range validation
        if min_val is not None and num_val < min_val:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="BELOW_MINIMUM",
                message=f"Value {num_val} is below minimum {min_val}",
                expected=f">= {min_val}",
                actual=str(num_val),
                suggestions=[f"Increase value to at least {min_val}"]
            ))
        
        if max_val is not None and num_val > max_val:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="ABOVE_MAXIMUM",
                message=f"Value {num_val} is above maximum {max_val}",
                expected=f"<= {max_val}",
                actual=str(num_val),
                suggestions=[f"Decrease value to at most {max_val}"]
            ))
        
        return issues
    
    @classmethod
    def validate_string_format(cls, value: str, format_name: str) -> ValidationIssue:
        """Validate string against known format patterns."""
        if not isinstance(value, str):
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="NOT_STRING",
                message=f"Expected string for format validation but got {type(value).__name__}",
                actual=type(value).__name__,
                expected="string"
            )
        
        pattern = cls.PATTERNS.get(format_name)
        if not pattern:
            return ValidationIssue(
                severity=ValidationSeverity.WARNING,
                code="UNKNOWN_FORMAT",
                message=f"Unknown format pattern: {format_name}",
                suggestions=["Use standard format names like 'email', 'phone', 'url'"]
            )
        
        if not pattern.match(value):
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="FORMAT_MISMATCH",
                message=f"Value '{value}' does not match {format_name} format",
                actual=value,
                expected=f"Valid {format_name} format",
                suggestions=[f"Check {format_name} format requirements", "Verify data source"]
            )
        
        return ValidationIssue(
            severity=ValidationSeverity.INFO,
            code="FORMAT_VALID",
            message=f"Format validation passed for {format_name}"
        )
    
    @classmethod
    def validate_date(cls, value: Any, date_format: Optional[str] = None) -> ValidationIssue:
        """Validate date values and formats."""
        if isinstance(value, (date, datetime)):
            return ValidationIssue(
                severity=ValidationSeverity.INFO,
                code="DATE_VALID",
                message="Date validation passed"
            )
        
        if not isinstance(value, str):
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_DATE_TYPE",
                message=f"Expected date string but got {type(value).__name__}",
                actual=type(value).__name__,
                suggestions=["Provide date as string or datetime object"]
            )
        
        # Try common date formats if none specified
        formats_to_try = [date_format] if date_format else [
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d',
            '%d-%m-%Y', '%m-%d-%Y', '%Y-%m-%d %H:%M:%S',
            '%d/%m/%Y %H:%M:%S', '%Y-%m-%dT%H:%M:%S'
        ]
        
        for fmt in formats_to_try:
            if fmt is None:
                continue
            try:
                datetime.strptime(value, fmt)
                return ValidationIssue(
                    severity=ValidationSeverity.INFO,
                    code="DATE_VALID",
                    message=f"Date validation passed with format {fmt}"
                )
            except ValueError:
                continue
        
        return ValidationIssue(
            severity=ValidationSeverity.ERROR,
            code="INVALID_DATE_FORMAT",
            message=f"Unable to parse date '{value}' with any known format",
            actual=value,
            suggestions=["Use standard date format like YYYY-MM-DD", "Check date format consistency"]
        )


class ExcelDataValidator:
    """Validates Excel-specific data and structures."""
    
    @classmethod
    def validate_cell_reference(cls, cell_ref: str) -> ValidationIssue:
        """Validate Excel cell reference format."""
        if not isinstance(cell_ref, str):
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_CELL_REF_TYPE",
                message=f"Cell reference must be string, got {type(cell_ref).__name__}",
                suggestions=["Use format like 'A1', 'B2', etc."]
            )
        
        # Excel cell reference pattern (e.g., A1, AB123, $A$1)
        pattern = re.compile(r'^\$?[A-Z]{1,3}\$?[1-9]\d*$', re.IGNORECASE)
        
        if not pattern.match(cell_ref):
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_CELL_REFERENCE",
                message=f"Invalid cell reference format: '{cell_ref}'",
                actual=cell_ref,
                expected="Format like A1, B2, AB123",
                suggestions=["Use valid Excel cell reference format", "Check for typos"]
            )
        
        # Check column limits (Excel has 16384 columns max)
        try:
            col_part = re.sub(r'[\$\d]', '', cell_ref.upper())
            col_index = column_index_from_string(col_part)
            if col_index > 16384:
                return ValidationIssue(
                    severity=ValidationSeverity.ERROR,
                    code="COLUMN_OUT_OF_RANGE",
                    message=f"Column index {col_index} exceeds Excel limit (16384)",
                    suggestions=["Use column within Excel limits"]
                )
        except ValueError:
            return ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_COLUMN_REFERENCE",
                message=f"Cannot parse column from reference: '{cell_ref}'",
                suggestions=["Check cell reference format"]
            )
        
        return ValidationIssue(
            severity=ValidationSeverity.INFO,
            code="CELL_REF_VALID",
            message="Cell reference validation passed"
        )
    
    @classmethod
    def validate_range_reference(cls, range_ref: str) -> List[ValidationIssue]:
        """Validate Excel range reference format."""
        issues = []
        
        if not isinstance(range_ref, str):
            return [ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_RANGE_REF_TYPE",
                message=f"Range reference must be string, got {type(range_ref).__name__}",
                suggestions=["Use format like 'A1:B2', 'A:A', '1:1'"]
            )]
        
        # Check for colon separator
        if ':' not in range_ref:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="MISSING_RANGE_SEPARATOR",
                message=f"Range reference missing ':' separator: '{range_ref}'",
                suggestions=["Use format like 'A1:B2'"]
            ))
            return issues
        
        parts = range_ref.split(':')
        if len(parts) != 2:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_RANGE_FORMAT",
                message=f"Range reference must have exactly one ':' separator: '{range_ref}'",
                suggestions=["Use format like 'A1:B2'"]
            ))
            return issues
        
        start_ref, end_ref = parts
        
        # Validate each part
        for i, ref in enumerate([start_ref, end_ref]):
            ref_name = "start" if i == 0 else "end"
            
            # Handle full column/row references (A:A, 1:1)
            if ref.isdigit() or re.match(r'^[A-Z]+$', ref, re.IGNORECASE):
                continue  # Valid full column or row reference
            
            cell_issue = cls.validate_cell_reference(ref)
            if cell_issue.severity in [ValidationSeverity.ERROR, ValidationSeverity.CRITICAL]:
                cell_issue.message = f"Invalid {ref_name} reference: {cell_issue.message}"
                issues.append(cell_issue)
        
        return issues
    
    @classmethod
    def validate_formula(cls, formula: str) -> List[ValidationIssue]:
        """Validate Excel formula syntax and structure."""
        issues = []
        
        if not isinstance(formula, str):
            return [ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="INVALID_FORMULA_TYPE",
                message=f"Formula must be string, got {type(formula).__name__}"
            )]
        
        # Formula must start with =
        if not formula.startswith('='):
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="FORMULA_MISSING_EQUALS",
                message="Formula must start with '='",
                suggestions=["Add '=' at the beginning of the formula"]
            ))
            return issues
        
        # Check for balanced parentheses
        open_parens = formula.count('(')
        close_parens = formula.count(')')
        if open_parens != close_parens:
            issues.append(ValidationIssue(
                severity=ValidationSeverity.ERROR,
                code="UNBALANCED_PARENTHESES",
                message=f"Formula has unbalanced parentheses: {open_parens} open, {close_parens} close",
                suggestions=["Check parentheses balance", "Add missing parentheses"]
            ))
        
        # Check for common Excel functions
        common_functions = [
            'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP',
            'INDEX', 'MATCH', 'CONCATENATE', 'LEFT', 'RIGHT', 'MID', 'LEN',
            'UPPER', 'LOWER', 'PROPER', 'TRIM', 'TODAY', 'NOW', 'DATE', 'TIME'
        ]
        
        # Extract function names from formula
        function_pattern = re.compile(r'([A-Z][A-Z0-9_]*)\s*\(', re.IGNORECASE)
        found_functions = function_pattern.findall(formula)
        
        for func in found_functions:
            if func.upper() not in common_functions:
                issues.append(ValidationIssue(
                    severity=ValidationSeverity.WARNING,
                    code="UNKNOWN_FUNCTION",
                    message=f"Unknown or uncommon function: {func}",
                    suggestions=["Verify function name spelling", "Check if function is supported"]
                ))
        
        # Check for cell references in formula
        cell_refs = re.findall(r'\$?[A-Z]{1,3}\$?[1-9]\d*', formula, re.IGNORECASE)
        for cell_ref in cell_refs:
            ref_issue = cls.validate_cell_reference(cell_ref)
            if ref_issue.severity in [ValidationSeverity.ERROR, ValidationSeverity.CRITICAL]:
                ref_issue.message = f"Invalid cell reference in formula: {ref_issue.message}"
                issues.append(ref_issue)
        
        return issues


class AccuracyValidator:
    """Main accuracy validation orchestrator."""
    
    def __init__(self):
        self.data_validator = DataTypeValidator()
        self.excel_validator = ExcelDataValidator()
        self._validation_rules: Dict[str, Callable] = {}
        self._custom_validators: Dict[str, Callable] = {}
    
    def register_custom_validator(self, name: str, validator_func: Callable) -> None:
        """Register a custom validation function."""
        self._custom_validators[name] = validator_func
    
    def validate_cell_value(self, value: Any, expected_type: Optional[type] = None,
                           format_rules: Optional[Dict[str, Any]] = None) -> List[ValidationIssue]:
        """Validate a single cell value comprehensively."""
        issues = []
        
        # Type validation
        if expected_type:
            type_issue = self.data_validator.validate_type(value, expected_type)
            if type_issue.severity != ValidationSeverity.INFO or type_issue.code != "TYPE_VALID":
                issues.append(type_issue)
        
        # Format validation
        if format_rules and value is not None:
            for rule_type, rule_config in format_rules.items():
                if rule_type == 'numeric_range' and isinstance(value, (int, float)):
                    range_issues = self.data_validator.validate_numeric_range(
                        value, 
                        rule_config.get('min'), 
                        rule_config.get('max')
                    )
                    issues.extend(range_issues)
                
                elif rule_type == 'string_format' and isinstance(value, str):
                    format_issue = self.data_validator.validate_string_format(
                        value, rule_config.get('format_name', 'email')
                    )
                    if format_issue.severity != ValidationSeverity.INFO:
                        issues.append(format_issue)
                
                elif rule_type == 'date_format':
                    date_issue = self.data_validator.validate_date(
                        value, rule_config.get('format')
                    )
                    if date_issue.severity != ValidationSeverity.INFO:
                        issues.append(date_issue)
        
        return issues
    
    def validate_worksheet_data(self, worksheet: Worksheet, 
                               validation_config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Validate entire worksheet data for accuracy and integrity."""
        issues = []
        total_cells = 0
        validated_cells = 0
        
        # Default validation config
        if validation_config is None:
            validation_config = {
                'check_empty_cells': True,
                'check_data_types': True,
                'check_formulas': True,
                'max_errors': 100
            }
        
        # Iterate through used range
        for row in worksheet.iter_rows():
            for cell in row:
                total_cells += 1
                
                if cell.value is None and not validation_config.get('check_empty_cells', True):
                    continue
                
                validated_cells += 1
                cell_location = f"{cell.column_letter}{cell.row}"
                
                try:
                    # Validate cell based on its content
                    cell_issues = self._validate_cell_comprehensive(cell, validation_config)
                    
                    for issue in cell_issues:
                        issue.location = cell_location
                        issues.append(issue)
                        
                        # Stop if too many errors
                        if len(issues) >= validation_config.get('max_errors', 100):
                            break
                    
                except Exception as e:
                    issues.append(ValidationIssue(
                        severity=ValidationSeverity.ERROR,
                        code="VALIDATION_ERROR",
                        message=f"Error validating cell: {str(e)}",
                        location=cell_location
                    ))
                
                if len(issues) >= validation_config.get('max_errors', 100):
                    break
            
            if len(issues) >= validation_config.get('max_errors', 100):
                break
        
        # Categorize issues by severity
        critical_issues = [i for i in issues if i.severity == ValidationSeverity.CRITICAL]
        error_issues = [i for i in issues if i.severity == ValidationSeverity.ERROR]
        warning_issues = [i for i in issues if i.severity == ValidationSeverity.WARNING]
        info_issues = [i for i in issues if i.severity == ValidationSeverity.INFO]
        
        return {
            'total_cells': total_cells,
            'validated_cells': validated_cells,
            'total_issues': len(issues),
            'critical_issues': len(critical_issues),
            'error_issues': len(error_issues),
            'warning_issues': len(warning_issues),
            'info_issues': len(info_issues),
            'issues': [
                {
                    'severity': issue.severity.value,
                    'code': issue.code,
                    'message': issue.message,
                    'location': issue.location,
                    'expected': issue.expected,
                    'actual': issue.actual,
                    'suggestions': issue.suggestions or []
                }
                for issue in issues
            ],
            'validation_summary': {
                'accuracy_score': max(0, 100 - (len(error_issues) + len(critical_issues)) * 5),
                'data_quality': 'excellent' if len(error_issues) == 0 else 'good' if len(error_issues) < 5 else 'needs_improvement',
                'recommendations': self._generate_recommendations(issues)
            }
        }
    
    def _validate_cell_comprehensive(self, cell: Cell, config: Dict[str, Any]) -> List[ValidationIssue]:
        """Comprehensive validation of a single cell."""
        issues = []
        
        # Check if cell has formula
        if cell.data_type == 'f' and config.get('check_formulas', True):
            formula_issues = self.excel_validator.validate_formula(cell.value)
            issues.extend(formula_issues)
        
        # Data type and format validation
        if cell.value is not None and config.get('check_data_types', True):
            # Infer expected type from cell value
            if isinstance(cell.value, str) and cell.value.startswith('='):
                pass  # Formula, already validated above
            elif isinstance(cell.value, (int, float)):
                # Numeric validation
                numeric_issues = self.data_validator.validate_numeric_range(cell.value)
                issues.extend(numeric_issues)
            elif isinstance(cell.value, str):
                # String validation - check for common formats
                if '@' in cell.value:
                    email_issue = self.data_validator.validate_string_format(cell.value, 'email')
                    if email_issue.code == "FORMAT_VALID":
                        issues.append(ValidationIssue(
                            severity=ValidationSeverity.INFO,
                            code="DETECTED_EMAIL",
                            message="Detected email format"
                        ))
                elif cell.value.startswith(('http://', 'https://')):
                    url_issue = self.data_validator.validate_string_format(cell.value, 'url')
                    issues.append(url_issue)
        
        return issues
    
    def _generate_recommendations(self, issues: List[ValidationIssue]) -> List[str]:
        """Generate recommendations based on validation issues."""
        recommendations = []
        
        # Count issue types
        issue_counts = {}
        for issue in issues:
            issue_counts[issue.code] = issue_counts.get(issue.code, 0) + 1
        
        # Generate recommendations based on common issues
        if issue_counts.get('TYPE_MISMATCH', 0) > 0:
            recommendations.append("Review data types and ensure consistency across columns")
        
        if issue_counts.get('FORMULA_MISSING_EQUALS', 0) > 0:
            recommendations.append("Check formulas and ensure they start with '='")
        
        if issue_counts.get('INVALID_CELL_REFERENCE', 0) > 0:
            recommendations.append("Verify cell references follow Excel format (e.g., A1, B2)")
        
        if issue_counts.get('FORMAT_MISMATCH', 0) > 0:
            recommendations.append("Standardize data formats (dates, emails, phone numbers)")
        
        if issue_counts.get('BELOW_MINIMUM', 0) + issue_counts.get('ABOVE_MAXIMUM', 0) > 0:
            recommendations.append("Review numeric values for outliers and range violations")
        
        return recommendations


# Global accuracy validator instance
accuracy_validator = AccuracyValidator()