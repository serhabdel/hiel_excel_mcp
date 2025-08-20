"""
Comprehensive tests for WorkbookManager tool.

Tests all workbook operations including creation, metadata retrieval,
safety validation, backup operations, and path management.
"""

import json
import os
import tempfile
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock

from hiel_excel_mcp.tools.workbook_manager import WorkbookManager, workbook_manager_tool
from hiel_excel_mcp.core.base_tool import OperationResponse, OperationStatus


class TestWorkbookManager:
    """Test suite for WorkbookManager tool."""
    
    def setup_method(self):
        """Set up test environment."""
        self.manager = WorkbookManager()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_workbook.xlsx")
    
    def teardown_method(self):
        """Clean up test environment."""
        # Clean up temp files
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_get_tool_info(self):
        """Test tool information retrieval."""
        info = self.manager.get_tool_info()
        
        assert info["name"] == "workbook_manager"
        assert "description" in info
        assert "operations" in info
        
        # Check that all expected operations are present
        expected_operations = [
            "create", "get_metadata", "validate_safety", 
            "create_backup", "get_backup_info", "validate_path", "sanitize_filename"
        ]
        
        for operation in expected_operations:
            assert operation in info["operations"]
            assert "description" in info["operations"][operation]
            assert "required_params" in info["operations"][operation]
    
    def test_get_available_operations(self):
        """Test getting available operations."""
        operations = self.manager.get_available_operations()
        
        expected_operations = [
            "create", "get_metadata", "validate_safety", 
            "create_backup", "get_backup_info", "validate_path", "sanitize_filename"
        ]
        
        for operation in expected_operations:
            assert operation in operations
    
    @patch('hiel_excel_mcp.tools.workbook_manager.create_workbook')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_create_operation_success(self, mock_validate_path, mock_create_workbook):
        """Test successful workbook creation."""
        # Mock path validation
        mock_validate_path.return_value = (self.test_file, [])
        
        # Mock workbook creation
        mock_workbook = MagicMock()
        mock_create_workbook.return_value = {
            "message": f"Created workbook: {self.test_file}",
            "workbook": mock_workbook
        }
        
        # Execute operation
        response = self.manager.execute_operation("create", filepath=self.test_file)
        
        # Verify response
        assert isinstance(response, OperationResponse)
        assert response.success is True
        assert response.operation == "create"
        assert self.test_file in response.message
        assert response.data["filepath"] == self.test_file
        assert response.data["sheet_name"] == "Sheet1"
        
        # Verify mocks were called
        mock_validate_path.assert_called_once_with(self.test_file, allow_create=True)
        mock_create_workbook.assert_called_once_with(self.test_file, "Sheet1")
        mock_workbook.close.assert_called_once()
    
    @patch('hiel_excel_mcp.tools.workbook_manager.create_workbook')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_create_operation_with_custom_sheet(self, mock_validate_path, mock_create_workbook):
        """Test workbook creation with custom sheet name."""
        mock_validate_path.return_value = (self.test_file, [])
        mock_workbook = MagicMock()
        mock_create_workbook.return_value = {
            "message": f"Created workbook: {self.test_file}",
            "workbook": mock_workbook
        }
        
        response = self.manager.execute_operation(
            "create", 
            filepath=self.test_file, 
            sheet_name="CustomSheet"
        )
        
        assert response.success is True
        assert response.data["sheet_name"] == "CustomSheet"
        mock_create_workbook.assert_called_once_with(self.test_file, "CustomSheet")
    
    @patch('hiel_excel_mcp.tools.workbook_manager.get_workbook_info')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_get_metadata_operation_success(self, mock_validate_path, mock_get_info):
        """Test successful metadata retrieval."""
        mock_validate_path.return_value = (self.test_file, [])
        mock_metadata = {
            "filename": "test_workbook.xlsx",
            "sheets": ["Sheet1", "Sheet2"],
            "size": 12345,
            "success": True
        }
        mock_get_info.return_value = mock_metadata
        
        response = self.manager.execute_operation("get_metadata", filepath=self.test_file)
        
        assert response.success is True
        assert response.operation == "get_metadata"
        assert response.data == mock_metadata
        mock_validate_path.assert_called_once_with(self.test_file, allow_create=False)
        mock_get_info.assert_called_once_with(self.test_file, False)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.get_workbook_info')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_get_metadata_with_ranges(self, mock_validate_path, mock_get_info):
        """Test metadata retrieval with ranges."""
        mock_validate_path.return_value = (self.test_file, [])
        mock_metadata = {
            "filename": "test_workbook.xlsx",
            "sheets": ["Sheet1"],
            "used_ranges": {"Sheet1": "A1:C10"},
            "success": True
        }
        mock_get_info.return_value = mock_metadata
        
        response = self.manager.execute_operation(
            "get_metadata", 
            filepath=self.test_file, 
            include_ranges=True
        )
        
        assert response.success is True
        assert "used_ranges" in response.data
        mock_get_info.assert_called_once_with(self.test_file, True)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.FileSafetyManager.validate_file_safety')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_validate_safety_operation_success(self, mock_validate_path, mock_validate_safety):
        """Test successful safety validation."""
        mock_validate_path.return_value = (self.test_file, ["Path warning"])
        mock_safety_result = {
            "safe": True,
            "file_size_mb": 1.5,
            "warnings": ["Safety warning"],
            "recommendations": ["Recommendation"]
        }
        mock_validate_safety.return_value = mock_safety_result
        
        response = self.manager.execute_operation("validate_safety", filepath=self.test_file)
        
        assert response.success is True
        assert response.operation == "validate_safety"
        assert response.data["safe"] is True
        assert response.data["file_size_mb"] == 1.5
        assert len(response.warnings) == 2  # Path + safety warnings
        assert "Path warning" in response.warnings
        assert "Safety warning" in response.warnings
    
    @patch('hiel_excel_mcp.tools.workbook_manager.FileSafetyManager.create_backup')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_create_backup_operation_success(self, mock_validate_path, mock_create_backup):
        """Test successful backup creation."""
        mock_validate_path.return_value = (self.test_file, [])
        backup_path = "/path/to/backup.xlsx"
        mock_create_backup.return_value = backup_path
        
        response = self.manager.execute_operation("create_backup", filepath=self.test_file)
        
        assert response.success is True
        assert response.operation == "create_backup"
        assert response.data["original_filepath"] == self.test_file
        assert response.data["backup_filepath"] == backup_path
        assert response.data["operation_name"] == "manual"
        mock_create_backup.assert_called_once_with(self.test_file, "manual")
    
    @patch('hiel_excel_mcp.tools.workbook_manager.FileSafetyManager.create_backup')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_create_backup_with_custom_operation_name(self, mock_validate_path, mock_create_backup):
        """Test backup creation with custom operation name."""
        mock_validate_path.return_value = (self.test_file, [])
        backup_path = "/path/to/backup.xlsx"
        mock_create_backup.return_value = backup_path
        
        response = self.manager.execute_operation(
            "create_backup", 
            filepath=self.test_file, 
            operation_name="test_operation"
        )
        
        assert response.success is True
        assert response.data["operation_name"] == "test_operation"
        mock_create_backup.assert_called_once_with(self.test_file, "test_operation")
    
    @patch('hiel_excel_mcp.tools.workbook_manager.FileSafetyManager.get_backup_info')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_get_backup_info_operation_success(self, mock_validate_path, mock_get_backup_info):
        """Test successful backup info retrieval."""
        mock_validate_path.return_value = (self.test_file, [])
        mock_backup_info = {
            "backup_dir": "/path/to/backups",
            "backups": [{"filename": "backup1.xlsx"}],
            "total_backups": 1
        }
        mock_get_backup_info.return_value = mock_backup_info
        
        response = self.manager.execute_operation("get_backup_info", filepath=self.test_file)
        
        assert response.success is True
        assert response.operation == "get_backup_info"
        assert response.data == mock_backup_info
    
    @patch('hiel_excel_mcp.tools.workbook_manager.FileSafetyManager.get_backup_info')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_get_backup_info_nonexistent_file(self, mock_validate_path, mock_get_backup_info):
        """Test backup info retrieval for non-existent file."""
        # Simulate path validation failure for non-existent file
        mock_validate_path.side_effect = Exception("File not found")
        mock_backup_info = {"backup_dir": "/path/to/backups", "backups": [], "total_backups": 0}
        mock_get_backup_info.return_value = mock_backup_info
        
        response = self.manager.execute_operation("get_backup_info", filepath=self.test_file)
        
        assert response.success is True
        assert "Original file does not exist" in response.warnings
    
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_validate_path_operation_success(self, mock_validate_path):
        """Test successful path validation."""
        validated_path = "/validated/path/file.xlsx"
        mock_validate_path.return_value = (validated_path, ["Warning"])
        
        response = self.manager.execute_operation("validate_path", filepath=self.test_file)
        
        assert response.success is True
        assert response.operation == "validate_path"
        assert response.data["original_path"] == self.test_file
        assert response.data["validated_path"] == validated_path
        assert response.data["is_valid"] is True
        assert response.warnings == ["Warning"]
        mock_validate_path.assert_called_once_with(self.test_file, True)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_validate_path_with_allow_create_false(self, mock_validate_path):
        """Test path validation with allow_create=False."""
        validated_path = "/validated/path/file.xlsx"
        mock_validate_path.return_value = (validated_path, [])
        
        response = self.manager.execute_operation(
            "validate_path", 
            filepath=self.test_file, 
            allow_create=False
        )
        
        assert response.success is True
        assert response.data["allow_create"] is False
        mock_validate_path.assert_called_once_with(self.test_file, False)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.sanitize_filename')
    def test_sanitize_filename_operation_success(self, mock_sanitize):
        """Test successful filename sanitization."""
        original_filename = "bad<>filename.xlsx"
        sanitized_filename = "bad__filename.xlsx"
        mock_sanitize.return_value = sanitized_filename
        
        response = self.manager.execute_operation("sanitize_filename", filename=original_filename)
        
        assert response.success is True
        assert response.operation == "sanitize_filename"
        assert response.data["original_filename"] == original_filename
        assert response.data["sanitized_filename"] == sanitized_filename
        assert response.data["was_modified"] is True
        assert "Filename was modified during sanitization" in response.warnings
    
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.sanitize_filename')
    def test_sanitize_filename_no_changes(self, mock_sanitize):
        """Test filename sanitization when no changes needed."""
        filename = "good_filename.xlsx"
        mock_sanitize.return_value = filename
        
        response = self.manager.execute_operation("sanitize_filename", filename=filename)
        
        assert response.success is True
        assert response.data["was_modified"] is False
        assert response.warnings is None
    
    def test_invalid_operation(self):
        """Test handling of invalid operation."""
        response = self.manager.execute_operation("invalid_operation", filepath=self.test_file)
        
        assert response.success is False
        assert response.status == OperationStatus.ERROR
        assert "not supported" in response.message
        assert "invalid_operation" in response.errors[0]
    
    def test_missing_required_parameters(self):
        """Test handling of missing required parameters."""
        response = self.manager.execute_operation("create")  # Missing filepath
        
        assert response.success is False
        assert response.status == OperationStatus.ERROR
        assert "Missing required parameters" in response.message
        assert "filepath" in response.errors[0]
    
    @patch('hiel_excel_mcp.tools.workbook_manager.create_workbook')
    @patch('hiel_excel_mcp.tools.workbook_manager.PathValidator.validate_path')
    def test_operation_exception_handling(self, mock_validate_path, mock_create_workbook):
        """Test handling of exceptions during operation execution."""
        mock_validate_path.return_value = (self.test_file, [])
        mock_create_workbook.side_effect = Exception("Test exception")
        
        response = self.manager.execute_operation("create", filepath=self.test_file)
        
        assert response.success is False
        assert response.status == OperationStatus.ERROR
        assert "Test exception" in response.message
        assert "Test exception" in response.errors[0]


class TestWorkbookManagerTool:
    """Test suite for workbook_manager_tool function."""
    
    def setup_method(self):
        """Set up test environment."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_workbook.xlsx")
    
    def teardown_method(self):
        """Clean up test environment."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.workbook_manager')
    def test_tool_function_success(self, mock_manager):
        """Test successful tool function execution."""
        mock_response = MagicMock()
        mock_response.to_json.return_value = '{"success": true}'
        mock_manager.execute_operation.return_value = mock_response
        
        result = workbook_manager_tool("create", filepath=self.test_file)
        
        assert result == '{"success": true}'
        mock_manager.execute_operation.assert_called_once_with("create", filepath=self.test_file)
    
    @patch('hiel_excel_mcp.tools.workbook_manager.workbook_manager')
    def test_tool_function_exception_handling(self, mock_manager):
        """Test tool function exception handling."""
        mock_manager.execute_operation.side_effect = Exception("Test exception")
        
        result = workbook_manager_tool("create", filepath=self.test_file)
        
        # Should return error response as JSON
        response_data = json.loads(result)
        assert response_data["success"] is False
        assert "Test exception" in response_data["message"]
    
    def test_tool_function_with_various_operations(self):
        """Test tool function with various operations."""
        # This is an integration test that would require actual file operations
        # For now, we'll test that the function can be called without errors
        
        operations_to_test = [
            ("sanitize_filename", {"filename": "test.xlsx"}),
            ("validate_path", {"filepath": "/tmp/test.xlsx", "allow_create": True})
        ]
        
        for operation, kwargs in operations_to_test:
            result = workbook_manager_tool(operation, **kwargs)
            response_data = json.loads(result)
            
            # Should return a valid JSON response
            assert "success" in response_data
            assert "operation" in response_data
            assert response_data["operation"] == operation


if __name__ == "__main__":
    pytest.main([__file__])