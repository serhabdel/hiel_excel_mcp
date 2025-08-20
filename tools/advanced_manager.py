"""
Advanced manager tool for handling named ranges, hyperlinks, and comments operations.
This tool provides high-level operations for advanced Excel features.
"""

from typing import Dict, Any, List, Optional
from ..core.base_tool import BaseTool
from ..core.workbook_context import WorkbookContext


class AdvancedManager(BaseTool):
    """Tool for managing named ranges, hyperlinks, and comments operations."""
    
    def create_named_range(
        self, 
        filepath: str, 
        name: str,
        range_reference: str,
        sheet_name: Optional[str] = None,
        comment: Optional[str] = None,
        scope: str = 'workbook'
    ) -> Dict[str, Any]:
        """Create a named range in the workbook."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.named_ranges import NamedRangeManager
            return NamedRangeManager.create_named_range(
                filepath, name, range_reference, sheet_name, comment, scope
            )
    
    def delete_named_range(
        self, 
        filepath: str, 
        name: str
    ) -> Dict[str, Any]:
        """Delete a named range from the workbook."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.named_ranges import NamedRangeManager
            return NamedRangeManager.delete_named_range(filepath, name)
    
    def list_named_ranges(
        self, 
        filepath: str, 
        include_details: bool = True
    ) -> Dict[str, Any]:
        """List all named ranges in the workbook."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.named_ranges import NamedRangeManager
            return NamedRangeManager.list_named_ranges(filepath, include_details)
    
    def get_named_range_value(
        self, 
        filepath: str, 
        name: str
    ) -> Dict[str, Any]:
        """Get the value(s) from a named range."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.named_ranges import NamedRangeManager
            return NamedRangeManager.get_named_range_value(filepath, name)
    
    def add_hyperlink(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        target: str,
        display_text: Optional[str] = None,
        tooltip: Optional[str] = None,
        link_type: str = 'auto'
    ) -> Dict[str, Any]:
        """Add a hyperlink to a cell."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.hyperlinks import HyperlinkManager
            return HyperlinkManager.add_hyperlink(
                filepath, sheet_name, cell, target, display_text, tooltip, link_type
            )
    
    def remove_hyperlink(
        self,
        filepath: str,
        sheet_name: str,
        cell: str,
        keep_text: bool = True
    ) -> Dict[str, Any]:
        """Remove hyperlink from a cell."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.hyperlinks import HyperlinkManager
            return HyperlinkManager.remove_hyperlink(filepath, sheet_name, cell, keep_text)
    
    def list_hyperlinks(
        self,
        filepath: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """List all hyperlinks in workbook or specific sheet."""
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.hyperlinks import HyperlinkManager
            return HyperlinkManager.list_hyperlinks(filepath, sheet_name)
    
    def manage_comments(
        self,
        filepath: str,
        action: str,
        sheet_name: str,
        cell: str,
        text: Optional[str] = None,
        author: Optional[str] = None,
        width: int = 200,
        height: int = 100,
        append: bool = False
    ) -> Dict[str, Any]:
        """
        Manage cell comments with multiple actions.
        
        Args:
            filepath: Path to Excel file
            action: 'add', 'edit', 'delete', or 'get'
            sheet_name: Name of worksheet
            cell: Cell reference
            text: Comment text (for add/edit actions)
            author: Comment author (for add action)
            width: Comment box width (for add action)
            height: Comment box height (for add action)
            append: Whether to append text (for edit action)
            
        Returns:
            Dict with operation results
        """
        with WorkbookContext(filepath) as wb_context:
            from ...src.excel_mcp.comments import CommentManager
            
            if action == 'add':
                if not text:
                    raise ValueError("Text is required for adding comments")
                return CommentManager.add_comment(
                    filepath, sheet_name, cell, text, author, width, height
                )
            
            elif action == 'edit':
                if not text:
                    raise ValueError("Text is required for editing comments")
                return CommentManager.edit_comment(
                    filepath, sheet_name, cell, text, append
                )
            
            elif action == 'delete':
                return CommentManager.delete_comment(filepath, sheet_name, cell)
            
            elif action == 'get':
                return CommentManager.get_comment(filepath, sheet_name, cell)
            
            else:
                raise ValueError(f"Invalid action: {action}. Use 'add', 'edit', 'delete', or 'get'")
    
    def search_advanced_features(
        self,
        filepath: str,
        search_type: str,
        search_text: str,
        sheet_name: Optional[str] = None,
        case_sensitive: bool = False,
        **kwargs
    ) -> Dict[str, Any]:
        """
        Search across advanced features (comments, named ranges, hyperlinks).
        
        Args:
            filepath: Path to Excel file
            search_type: 'comments', 'named_ranges', or 'hyperlinks'
            search_text: Text to search for
            sheet_name: Name of specific sheet (None for all sheets)
            case_sensitive: Whether search is case-sensitive
            **kwargs: Additional search parameters
            
        Returns:
            Dict with search results
        """
        with WorkbookContext(filepath) as wb_context:
            if search_type == 'comments':
                from ...src.excel_mcp.comments import CommentManager
                search_author = kwargs.get('search_author', False)
                return CommentManager.search_comments(
                    filepath, search_text, sheet_name, case_sensitive, search_author
                )
            
            elif search_type == 'named_ranges':
                from ...src.excel_mcp.named_ranges import NamedRangeManager
                # Search in named range names and references
                ranges_result = NamedRangeManager.list_named_ranges(filepath, True)
                if not ranges_result["success"]:
                    return ranges_result
                
                search_lower = search_text.lower() if not case_sensitive else search_text
                matching_ranges = []
                
                for range_info in ranges_result["named_ranges"]:
                    name = range_info["name"]
                    reference = range_info["reference"]
                    
                    name_match = (search_text in name) if case_sensitive else (search_lower in name.lower())
                    ref_match = (search_text in reference) if case_sensitive else (search_lower in reference.lower())
                    
                    if name_match or ref_match:
                        match_info = range_info.copy()
                        match_info["matches"] = []
                        if name_match:
                            match_info["matches"].append("name")
                        if ref_match:
                            match_info["matches"].append("reference")
                        matching_ranges.append(match_info)
                
                return {
                    "success": True,
                    "search_text": search_text,
                    "case_sensitive": case_sensitive,
                    "total_matches": len(matching_ranges),
                    "matching_ranges": matching_ranges
                }
            
            elif search_type == 'hyperlinks':
                from ...src.excel_mcp.hyperlinks import HyperlinkManager
                # Get all hyperlinks and filter
                links_result = HyperlinkManager.list_hyperlinks(filepath, sheet_name)
                if not links_result["success"]:
                    return links_result
                
                search_lower = search_text.lower() if not case_sensitive else search_text
                matching_links = []
                
                for link_info in links_result["hyperlinks"]:
                    target = link_info["target"]
                    display_text = str(link_info["display_text"]) if link_info["display_text"] else ""
                    
                    target_match = (search_text in target) if case_sensitive else (search_lower in target.lower())
                    text_match = (search_text in display_text) if case_sensitive else (search_lower in display_text.lower())
                    
                    if target_match or text_match:
                        match_info = link_info.copy()
                        match_info["matches"] = []
                        if target_match:
                            match_info["matches"].append("target")
                        if text_match:
                            match_info["matches"].append("display_text")
                        matching_links.append(match_info)
                
                return {
                    "success": True,
                    "search_text": search_text,
                    "case_sensitive": case_sensitive,
                    "total_matches": len(matching_links),
                    "matching_hyperlinks": matching_links
                }
            
            else:
                raise ValueError(f"Invalid search_type: {search_type}. Use 'comments', 'named_ranges', or 'hyperlinks'")
    
    def get_advanced_summary(
        self,
        filepath: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Get a comprehensive summary of all advanced features in the workbook.
        
        Args:
            filepath: Path to Excel file
            sheet_name: Name of specific sheet (None for all sheets)
            
        Returns:
            Dict with comprehensive summary
        """
        with WorkbookContext(filepath) as wb_context:
            summary = {
                "success": True,
                "filepath": filepath,
                "sheet_scope": sheet_name or "all_sheets"
            }
            
            try:
                # Named ranges summary
                from ...src.excel_mcp.named_ranges import NamedRangeManager
                ranges_result = NamedRangeManager.list_named_ranges(filepath, False)
                summary["named_ranges"] = {
                    "count": ranges_result.get("total_ranges", 0),
                    "success": ranges_result.get("success", False)
                }
                
                # Hyperlinks summary
                from ...src.excel_mcp.hyperlinks import HyperlinkManager
                links_result = HyperlinkManager.list_hyperlinks(filepath, sheet_name)
                summary["hyperlinks"] = {
                    "count": links_result.get("total_hyperlinks", 0),
                    "success": links_result.get("success", False)
                }
                
                # Comments summary
                from ...src.excel_mcp.comments import CommentManager
                comments_result = CommentManager.list_comments(filepath, sheet_name, False)
                summary["comments"] = {
                    "count": comments_result.get("total_comments", 0),
                    "success": comments_result.get("success", False)
                }
                
                # Overall statistics
                summary["totals"] = {
                    "named_ranges": summary["named_ranges"]["count"],
                    "hyperlinks": summary["hyperlinks"]["count"],
                    "comments": summary["comments"]["count"],
                    "total_advanced_features": (
                        summary["named_ranges"]["count"] + 
                        summary["hyperlinks"]["count"] + 
                        summary["comments"]["count"]
                    )
                }
                
                return summary
                
            except Exception as e:
                summary["success"] = False
                summary["error"] = str(e)
                return summary