"""
Batch manager tool for handling batch operations and template processing.
This tool provides high-level batch operations for Excel files.
"""

from typing import Dict, Any, List, Optional, Union
from ..core.base_tool import BaseTool
from ..core.workbook_context import WorkbookContext


class BatchManager(BaseTool):
    """Tool for managing batch operations and template processing."""
    
    def batch_create_workbooks(
        self, 
        file_paths: List[str]
    ) -> Dict[str, Any]:
        """Create multiple workbooks in batch."""
        try:
            from ...src.excel_mcp.batch_processor import batch_processor
            operation_id = batch_processor.batch_create_workbooks(file_paths)
            
            return {
                "success": True,
                "operation_id": operation_id,
                "message": f"Started batch creation of {len(file_paths)} workbooks",
                "total_files": len(file_paths)
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def batch_export(
        self,
        excel_paths: List[str],
        export_format: str = 'csv',
        export_configs: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """
        Export multiple Excel files to various formats.
        
        Args:
            excel_paths: List of Excel file paths to export
            export_format: Target format ('csv', 'json', 'xml', 'html')
            export_configs: List of export configuration dicts (optional)
            
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.batch_processor import batch_processor
            
            if export_format == 'csv':
                # Prepare CSV export pairs
                export_pairs = []
                for i, excel_path in enumerate(excel_paths):
                    if export_configs and i < len(export_configs):
                        config = export_configs[i]
                        csv_path = config.get('output_path', excel_path.replace('.xlsx', '.csv'))
                        sheet_name = config.get('sheet_name', 'Sheet1')
                        start_cell = config.get('start_cell', 'A1')
                        include_header = config.get('include_header', True)
                    else:
                        csv_path = excel_path.replace('.xlsx', '.csv')
                        sheet_name = 'Sheet1'
                        start_cell = 'A1'
                        include_header = True
                    
                    export_pairs.append({
                        'excel_path': excel_path,
                        'csv_path': csv_path
                    })
                
                operation_id = batch_processor.batch_export_csv(
                    export_pairs, sheet_name, start_cell, include_header
                )
                
                return {
                    "success": True,
                    "operation_id": operation_id,
                    "message": f"Started batch CSV export of {len(excel_paths)} files",
                    "total_files": len(excel_paths),
                    "export_format": "csv"
                }
            
            else:
                # For other formats, use generic batch operation
                from ...src.excel_mcp.export_formats import export_to_format
                
                def export_operation(**kwargs):
                    return export_to_format(
                        filepath=kwargs['filepath'],
                        output_path=kwargs['output_path'],
                        format=export_format,
                        **kwargs.get('config', {})
                    )
                
                # Prepare operation arguments
                operation_args = {
                    'format': export_format,
                    'configs': export_configs or [{}] * len(excel_paths)
                }
                
                operation_id = batch_processor.batch_apply_operation(
                    excel_paths, export_operation, operation_args, 
                    f"Batch {export_format.upper()} Export"
                )
                
                return {
                    "success": True,
                    "operation_id": operation_id,
                    "message": f"Started batch {export_format.upper()} export of {len(excel_paths)} files",
                    "total_files": len(excel_paths),
                    "export_format": export_format
                }
                
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def fill_template(
        self,
        template_path: str,
        output_path: str,
        data: Dict[str, Any],
        preserve_formatting: bool = True
    ) -> Dict[str, Any]:
        """Fill a single Excel template with data."""
        try:
            from ...src.excel_mcp.template_engine import TemplateEngine
            return TemplateEngine.fill_template(
                template_path, output_path, data, preserve_formatting
            )
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def fill_table_template(
        self,
        template_path: str,
        output_path: str,
        table_data: List[Dict[str, Any]],
        start_cell: str = "A1",
        include_headers: bool = True
    ) -> Dict[str, Any]:
        """Fill template with tabular data."""
        try:
            from ...src.excel_mcp.template_engine import TemplateEngine
            return TemplateEngine.fill_table_template(
                template_path, output_path, table_data, start_cell, include_headers
            )
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def generate_report_template(
        self,
        output_path: str,
        report_config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Generate a report template with placeholders."""
        try:
            from ...src.excel_mcp.template_engine import TemplateEngine
            return TemplateEngine.generate_report_template(output_path, report_config)
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def batch_fill_templates(
        self,
        template_configs: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Fill multiple templates with data in batch.
        
        Args:
            template_configs: List of dicts with template configuration:
                - template_path: Path to template file
                - output_path: Path for output file
                - data: Data for template filling
                - preserve_formatting: Whether to preserve formatting (optional)
                
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.template_engine import TemplateEngine
            
            results = []
            successful = 0
            failed = 0
            
            for config in template_configs:
                try:
                    template_path = config['template_path']
                    output_path = config['output_path']
                    data = config['data']
                    preserve_formatting = config.get('preserve_formatting', True)
                    
                    result = TemplateEngine.fill_template(
                        template_path, output_path, data, preserve_formatting
                    )
                    
                    results.append({
                        'template_path': template_path,
                        'output_path': output_path,
                        'success': True,
                        'filled_cells': result.get('total_filled_cells', 0)
                    })
                    successful += 1
                    
                except Exception as e:
                    results.append({
                        'template_path': config.get('template_path', 'unknown'),
                        'output_path': config.get('output_path', 'unknown'),
                        'success': False,
                        'error': str(e)
                    })
                    failed += 1
            
            return {
                "success": True,
                "message": f"Batch template filling completed: {successful} successful, {failed} failed",
                "total_templates": len(template_configs),
                "successful": successful,
                "failed": failed,
                "results": results
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def batch_generate_reports(
        self,
        report_configs: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Generate multiple report templates in batch.
        
        Args:
            report_configs: List of dicts with report configuration:
                - output_path: Path for template file
                - config: Report structure configuration
                
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.template_engine import TemplateEngine
            
            results = []
            successful = 0
            failed = 0
            
            for config in report_configs:
                try:
                    output_path = config['output_path']
                    report_config = config['config']
                    
                    result = TemplateEngine.generate_report_template(
                        output_path, report_config
                    )
                    
                    results.append({
                        'output_path': output_path,
                        'success': True,
                        'sections': result.get('sections', [])
                    })
                    successful += 1
                    
                except Exception as e:
                    results.append({
                        'output_path': config.get('output_path', 'unknown'),
                        'success': False,
                        'error': str(e)
                    })
                    failed += 1
            
            return {
                "success": True,
                "message": f"Batch report generation completed: {successful} successful, {failed} failed",
                "total_reports": len(report_configs),
                "successful": successful,
                "failed": failed,
                "results": results
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
            return batch_manager.get_operation_status(operation_id)
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def list_batch_operations(
        self,
        active_only: bool = True
    ) -> Dict[str, Any]:
        """List all batch operations."""
        try:
            from ...src.excel_mcp.cache_manager import batch_manager
            operations = batch_manager.list_operations(active_only)
            
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
        """Cancel a running batch operation."""
        try:
            from ...src.excel_mcp.cache_manager import batch_manager
            success = batch_manager.cancel_operation(operation_id)
            
            return {
                "success": success,
                "operation_id": operation_id,
                "message": "Operation cancelled" if success else "Operation not found or already completed"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def batch_apply_formulas(
        self,
        formula_configs: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Apply formulas to multiple files in batch.
        
        Args:
            formula_configs: List of dicts with:
                - filepath: Excel file path
                - sheet_name: Worksheet name
                - cell: Target cell
                - formula: Formula to apply
                
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.batch_processor import batch_processor
            operation_id = batch_processor.batch_apply_formula(formula_configs)
            
            return {
                "success": True,
                "operation_id": operation_id,
                "message": f"Started batch formula application to {len(formula_configs)} operations",
                "total_operations": len(formula_configs)
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def batch_process_data(
        self,
        operation_type: str,
        file_configs: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Generic batch data processing operation.
        
        Args:
            operation_type: Type of operation ('import', 'export', 'transform')
            file_configs: List of configuration dicts for each file
            
        Returns:
            Dict with operation results
        """
        try:
            from ...src.excel_mcp.batch_processor import batch_processor
            
            if operation_type == 'import':
                # Batch CSV import
                csv_excel_pairs = []
                for config in file_configs:
                    csv_excel_pairs.append({
                        'csv_path': config['input_path'],
                        'excel_path': config['output_path']
                    })
                
                operation_id = batch_processor.batch_import_csv(
                    csv_excel_pairs,
                    config.get('sheet_name', 'Sheet1'),
                    config.get('start_cell', 'A1'),
                    config.get('has_header', True)
                )
                
                return {
                    "success": True,
                    "operation_id": operation_id,
                    "message": f"Started batch import of {len(file_configs)} files",
                    "operation_type": "import"
                }
                
            elif operation_type == 'export':
                return self.batch_export(
                    [config['input_path'] for config in file_configs],
                    file_configs[0].get('format', 'csv'),
                    file_configs
                )
                
            elif operation_type == 'transform':
                # Generic transformation operation
                def transform_operation(**kwargs):
                    # This would be implemented based on specific transformation needs
                    from ...src.excel_mcp.data_transform import DataTransformer
                    transformer = DataTransformer()
                    return transformer.transform_file(**kwargs)
                
                operation_id = batch_processor.batch_apply_operation(
                    [config['filepath'] for config in file_configs],
                    transform_operation,
                    {'configs': file_configs},
                    "Batch Data Transformation"
                )
                
                return {
                    "success": True,
                    "operation_id": operation_id,
                    "message": f"Started batch transformation of {len(file_configs)} files",
                    "operation_type": "transform"
                }
            
            else:
                return {
                    "success": False,
                    "error": f"Unknown operation type: {operation_type}"
                }
                
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }