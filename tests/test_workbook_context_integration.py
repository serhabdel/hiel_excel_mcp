"""
Integration tests for workbook context with base tool infrastructure.
"""

import tempfile
from pathlib import Path
from unittest.mock import patch

from hiel_excel_mcp.core.workbook_context import workbook_context, clear_cache
from hiel_excel_mcp.core.base_tool import BaseTool


class MockWorkbookTool(BaseTool):
    """Mock tool for testing workbook context integration."""
    
    def __init__(self):
        super().__init__()
        self.operations = {
            'create_and_write': self._create_and_write,
            'read_data': self._read_data,
        }
    
    def _create_and_write(self, filepath: str, data: str, **kwargs):
        """Create workbook and write data."""
        with workbook_context(filepath) as wb:
            ws = wb.active
            ws['A1'] = data
        
        return {
            'success': True,
            'message': f'Created workbook and wrote: {data}',
            'filepath': filepath
        }
    
    def _read_data(self, filepath: str, **kwargs):
        """Read data from workbook."""
        with workbook_context(filepath, read_only=True) as wb:
            ws = wb.active
            value = ws['A1'].value
        
        return {
            'success': True,
            'data': value,
            'filepath': filepath
        }


def test_workbook_context_with_base_tool():
    """Test workbook context integration with base tool."""
    tool = MockWorkbookTool()
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # Remove the empty file
        Path(tmp_path).unlink()
        
        # Test create and write operation
        result1 = tool._create_and_write(tmp_path, 'Integration Test')
        assert result1['success'] is True
        assert 'Integration Test' in result1['message']
        
        # Verify file was created
        assert Path(tmp_path).exists()
        
        # Test read operation
        result2 = tool._read_data(tmp_path)
        assert result2['success'] is True
        assert result2['data'] == 'Integration Test'
        
        print('✓ Workbook context integrates correctly with base tool')
        
    finally:
        if Path(tmp_path).exists():
            Path(tmp_path).unlink()
        clear_cache()


def test_multiple_operations_same_file():
    """Test multiple operations on same file use cached workbook."""
    tool = MockWorkbookTool()
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # Remove the empty file
        Path(tmp_path).unlink()
        
        # Perform multiple operations
        tool._create_and_write(tmp_path, 'First')
        
        # This should use cached workbook
        with workbook_context(tmp_path) as wb:
            wb.active['B1'] = 'Second'
        
        # Read back both values
        with workbook_context(tmp_path, read_only=True) as wb:
            value_a = wb.active['A1'].value
            value_b = wb.active['B1'].value
        
        assert value_a == 'First'
        assert value_b == 'Second'
        
        print('✓ Multiple operations correctly use cached workbook')
        
    finally:
        if Path(tmp_path).exists():
            Path(tmp_path).unlink()
        clear_cache()


if __name__ == '__main__':
    test_workbook_context_with_base_tool()
    test_multiple_operations_same_file()
    print('All integration tests passed!')