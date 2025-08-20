"""
Tests for ValidationManager tool.

Tests all validation operations including dropdown creation, number validation,
date validation, and validation removal.
"""

import pytest
import json
import tempfile
import os
from pathlib import Path
from datetime import datetime
from unittest.mock import patch, MagicMock

from hiel_excel_mcp.tools.validation_manager import ValidationManager, validation_manager_tool
from hiel_excel_mcp.core.base_tool import OperationResponse, OperationStatus


class TestValidationManager:
    """Test cases for ValidationManager class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
        self.test_filepath = "test_validation.xlsx"
        self.test_sheet = "Sheet1"
        self.test_range = "A1:A10"
    
    def test_tool_metadata(self):
        """Test tool metadata and operation registration."""
        assert self.manager.get_tool_name() == "validation_manager"
        assert "validation management" in self.manager.get_tool_description().lower()
        
        operations = self.manager.get_available_operations()
        expected_operations = [
            "create_dropdown",
            "create_number_validation", 
            "create_date_validation",
            "remove_validation"
        ]
        
        for op in expected_operations:
            assert op in operations
    
    def test_operation_metadata(self):
        """Test operation metadata is properly defined."""
        # Test create_dropdown metadata
        dropdown_meta = self.manager.get_operation_metadata("create_dropdown")
        assert dropdown_meta is not None
        assert "filepath" in dropdown_meta.required_params
        assert "sheet_name" in dropdown_meta.required_params
        assert "cell_range" in dropdown_meta.required_params
        assert "options" in dropdown_meta.required_params
        assert "input_title" in dropdown_meta.optional_params
        
        # Test create_number_validation metadata
        number_meta = self.manager.get_operation_metadata("create_number_validation")
        assert number_meta is not None
        assert "filepath" in number_meta.required_params
        assert "min_value" in number_meta.optional_params
        assert "max_value" in number_meta.optional_params
        
        # Test create_date_validation metadata
        date_meta = self.manager.get_operation_metadata("create_date_validation")
        assert date_meta is not None
        assert "start_date" in date_meta.optional_params
        assert "end_date" in date_meta.optional_params
        
        # Test remove_validation metadata
        remove_meta = self.manager.get_operation_metadata("remove_validation")
        assert remove_meta is not None
        assert "cell_range" in remove_meta.optional_params


class TestCreateDropdownOperation:
    """Test cases for create_dropdown operation."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.create_dropdown_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_create_dropdown_success(self, mock_context, mock_create_dropdown):
        """Test successful dropdown creation."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_create_dropdown.return_value = {
            "success": True,
            "message": "Dropdown created",
            "options": ["Option1", "Option2", "Option3"]
        }
        
        # Test operation
        response = self.manager.execute_operation(
            "create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options=["Option1", "Option2", "Option3"],
            input_title="Select Option",
            allow_blank=True
        )
        
        assert response.success is True
        assert response.operation == "create_dropdown"
        assert "dropdown validation created" in response.message.lower()
        assert response.data["options_count"] == 3
        assert response.data["input_title"] == "Select Option"
        
        # Verify the underlying function was called correctly
        mock_create_dropdown.assert_called_once_with(
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options=["Option1", "Option2", "Option3"],
            input_title="Select Option",
            input_message=None,
            allow_blank=True
        )
    
    def test_create_dropdown_empty_options(self):
        """Test dropdown creation with empty options list."""
        response = self.manager.execute_operation(
            "create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options=[]
        )
        
        assert response.success is False
        assert "options list cannot be empty" in response.message.lower()
    
    def test_create_dropdown_invalid_options_type(self):
        """Test dropdown creation with invalid options type."""
        response = self.manager.execute_operation(
            "create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options="not a list"
        )
        
        assert response.success is False
        assert "options must be provided as a list" in response.message.lower()
    
    def test_create_dropdown_missing_required_params(self):
        """Test dropdown creation with missing required parameters."""
        response = self.manager.execute_operation(
            "create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1"
            # Missing cell_range and options
        )
        
        assert response.success is False
        assert "missing required parameters" in response.message.lower()


class TestCreateNumberValidationOperation:
    """Test cases for create_number_validation operation."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.create_number_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_create_number_validation_success(self, mock_context, mock_create_number):
        """Test successful number validation creation."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_create_number.return_value = {
            "success": True,
            "message": "Number validation created"
        }
        
        # Test operation
        response = self.manager.execute_operation(
            "create_number_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="B1:B10",
            min_value=1,
            max_value=100,
            allow_decimals=False,
            error_message="Enter a number between 1 and 100"
        )
        
        assert response.success is True
        assert response.operation == "create_number_validation"
        assert "number validation created" in response.message.lower()
        assert response.data["min_value"] == 1
        assert response.data["max_value"] == 100
        assert response.data["allow_decimals"] is False
        assert response.data["validation_type"] == "whole"
        
        # Verify the underlying function was called correctly
        mock_create_number.assert_called_once_with(
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="B1:B10",
            min_value=1,
            max_value=100,
            allow_decimals=False,
            allow_blank=True,
            error_message="Enter a number between 1 and 100"
        )
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.create_number_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_create_number_validation_decimals(self, mock_context, mock_create_number):
        """Test number validation with decimals allowed."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_create_number.return_value = {"success": True}
        
        response = self.manager.execute_operation(
            "create_number_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="C1:C10",
            min_value=0.1,
            max_value=99.9,
            allow_decimals=True
        )
        
        assert response.success is True
        assert response.data["validation_type"] == "decimal"
    
    def test_create_number_validation_invalid_range(self):
        """Test number validation with invalid min/max range."""
        response = self.manager.execute_operation(
            "create_number_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="D1:D10",
            min_value=100,
            max_value=1  # Invalid: min > max
        )
        
        assert response.success is False
        assert "min_value cannot be greater than max_value" in response.message.lower()


class TestCreateDateValidationOperation:
    """Test cases for create_date_validation operation."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.create_date_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_create_date_validation_success(self, mock_context, mock_create_date):
        """Test successful date validation creation."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_create_date.return_value = {
            "success": True,
            "message": "Date validation created"
        }
        
        # Test operation
        response = self.manager.execute_operation(
            "create_date_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="E1:E10",
            start_date="2024-01-01",
            end_date="2024-12-31",
            error_message="Enter a date in 2024"
        )
        
        assert response.success is True
        assert response.operation == "create_date_validation"
        assert "date validation created" in response.message.lower()
        assert response.data["start_date"] == "2024-01-01"
        assert response.data["end_date"] == "2024-12-31"
        
        # Verify the underlying function was called correctly
        mock_create_date.assert_called_once_with(
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="E1:E10",
            start_date="2024-01-01",
            end_date="2024-12-31",
            allow_blank=True,
            error_message="Enter a date in 2024"
        )
    
    def test_create_date_validation_invalid_format(self):
        """Test date validation with invalid date format."""
        response = self.manager.execute_operation(
            "create_date_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="F1:F10",
            start_date="01/01/2024"  # Invalid format
        )
        
        assert response.success is False
        assert "yyyy-mm-dd format" in response.message.lower()
    
    def test_create_date_validation_invalid_range(self):
        """Test date validation with invalid date range."""
        response = self.manager.execute_operation(
            "create_date_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="G1:G10",
            start_date="2024-12-31",
            end_date="2024-01-01"  # Invalid: start > end
        )
        
        assert response.success is False
        assert "start_date cannot be after end_date" in response.message.lower()


class TestRemoveValidationOperation:
    """Test cases for remove_validation operation."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.remove_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_remove_validation_specific_range(self, mock_context, mock_remove):
        """Test removing validation from specific range."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_remove.return_value = {
            "success": True,
            "removed_count": 2
        }
        
        # Test operation
        response = self.manager.execute_operation(
            "remove_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10"
        )
        
        assert response.success is True
        assert response.operation == "remove_validation"
        assert "validation removed from a1:a10" in response.message.lower()
        assert response.data["scope"] == "A1:A10"
        assert response.data["removed_count"] == 2
        
        # Verify the underlying function was called correctly
        mock_remove.assert_called_once_with(
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10"
        )
    
    @patch('hiel_excel_mcp.tools.validation_manager.AdvancedValidationManager.remove_validation')
    @patch('hiel_excel_mcp.tools.validation_manager.workbook_context')
    def test_remove_validation_entire_sheet(self, mock_context, mock_remove):
        """Test removing validation from entire sheet."""
        # Setup mocks
        mock_context.return_value.__enter__.return_value = MagicMock()
        mock_context.return_value.__exit__.return_value = None
        mock_remove.return_value = {
            "success": True,
            "removed_count": 5
        }
        
        # Test operation (no cell_range specified)
        response = self.manager.execute_operation(
            "remove_validation",
            filepath="test.xlsx",
            sheet_name="Sheet1"
        )
        
        assert response.success is True
        assert "validation removed from entire sheet" in response.message.lower()
        assert response.data["scope"] == "entire sheet"
        assert response.data["cell_range"] is None


class TestValidationManagerTool:
    """Test cases for the MCP tool function."""
    
    @patch('hiel_excel_mcp.tools.validation_manager.validation_manager')
    def test_tool_function_success(self, mock_manager):
        """Test successful tool function execution."""
        # Setup mock response
        mock_response = OperationResponse(
            success=True,
            operation="create_dropdown",
            message="Success",
            data={"test": "data"}
        )
        mock_manager.execute_operation.return_value = mock_response
        
        # Test tool function
        result = validation_manager_tool(
            operation="create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options=["A", "B", "C"]
        )
        
        # Verify result is valid JSON
        parsed_result = json.loads(result)
        assert parsed_result["success"] is True
        assert parsed_result["operation"] == "create_dropdown"
        
        # Verify manager was called correctly
        mock_manager.execute_operation.assert_called_once_with(
            "create_dropdown",
            filepath="test.xlsx",
            sheet_name="Sheet1",
            cell_range="A1:A10",
            options=["A", "B", "C"]
        )
    
    @patch('hiel_excel_mcp.tools.validation_manager.validation_manager')
    def test_tool_function_error(self, mock_manager):
        """Test tool function error handling."""
        # Setup mock to raise exception
        mock_manager.execute_operation.side_effect = Exception("Test error")
        
        # Test tool function
        result = validation_manager_tool(
            operation="create_dropdown",
            filepath="test.xlsx"
        )
        
        # Verify error response
        parsed_result = json.loads(result)
        assert parsed_result["success"] is False
        assert "test error" in parsed_result["message"].lower()


class TestValidationManagerIntegration:
    """Integration tests for ValidationManager."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.manager = ValidationManager()
    
    def test_invalid_operation(self):
        """Test handling of invalid operation."""
        response = self.manager.execute_operation(
            "invalid_operation",
            filepath="test.xlsx"
        )
        
        assert response.success is False
        assert "operation 'invalid_operation' not supported" in response.message.lower()
        assert "available operations:" in response.message.lower()
    
    def test_missing_required_parameters(self):
        """Test handling of missing required parameters."""
        response = self.manager.execute_operation(
            "create_dropdown",
            filepath="test.xlsx"
            # Missing required parameters
        )
        
        assert response.success is False
        assert "missing required parameters" in response.message.lower()
    
    def test_get_tool_info(self):
        """Test tool information retrieval."""
        info = self.manager.get_tool_info()
        
        assert info["name"] == "validation_manager"
        assert "validation management" in info["description"].lower()
        assert "operations" in info
        assert len(info["operations"]) == 4
        
        # Check specific operations exist
        operations = info["operations"]
        assert "create_dropdown" in operations
        assert "create_number_validation" in operations
        assert "create_date_validation" in operations
        assert "remove_validation" in operations
        
        # Check operation details
        dropdown_op = operations["create_dropdown"]
        assert "filepath" in dropdown_op["required_params"]
        assert "options" in dropdown_op["required_params"]
        assert "input_title" in dropdown_op["optional_params"]


if __name__ == "__main__":
    pytest.main([__file__])