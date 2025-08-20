#!/usr/bin/env python3
"""
CLI entry point for Hiel Excel MCP Server.
Provides command-line interface with Typer for starting and managing the server.
Supports both pip and uvx installation methods.
"""

import os
import sys
import logging
from pathlib import Path
from typing import Optional, List
import asyncio
import shutil
import subprocess

# Check if running under uvx
RUNNING_IN_UVX = os.environ.get('UV_PROJECT_ENVIRONMENT') is not None

try:
    import typer
except ImportError:
    install_cmd = "uvx install typer" if RUNNING_IN_UVX else "pip install typer"
    print(f"Typer not installed. Install with: {install_cmd}")
    sys.exit(1)

try:
    import uvicorn
except ImportError:
    install_cmd = "uvx install 'uvicorn[standard]'" if RUNNING_IN_UVX else "pip install 'uvicorn[standard]'"
    print(f"Uvicorn not installed. Install with: {install_cmd}")
    sys.exit(1)

from .core.config import config
from .core.utils import ExcelMCPUtils
from .core.memory_optimizer import memory_optimizer
from .core.performance_optimizer import performance_optimizer
from .core.accuracy_validator import accuracy_validator

app = typer.Typer(
    name="hiel-excel-mcp",
    help="🚀 Hiel Excel MCP Server - Production-ready Excel operations via MCP protocol",
    add_completion=False,
    rich_markup_mode="rich"
)


def setup_logging(log_level: str = "INFO", log_file: Optional[str] = None):
    """Setup logging configuration."""
    level = getattr(logging, log_level.upper(), logging.INFO)
    
    handlers = []
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(
        logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    )
    handlers.append(console_handler)
    
    # File handler if specified
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        )
        handlers.append(file_handler)
    
    # Configure root logger
    logging.basicConfig(
        level=level,
        handlers=handlers,
        force=True
    )


@app.command()
def serve(
    host: str = typer.Option("0.0.0.0", "--host", "-h", help="Host to bind to"),
    port: int = typer.Option(8000, "--port", "-p", help="Port to bind to"),
    log_level: str = typer.Option("INFO", "--log-level", "-l", 
                                 help="Log level (DEBUG, INFO, WARNING, ERROR)"),
    log_file: Optional[str] = typer.Option(None, "--log-file", "-f", 
                                          help="Log file path"),
    workers: int = typer.Option(1, "--workers", "-w", help="Number of worker processes"),
    reload: bool = typer.Option(False, "--reload", "-r", help="Enable auto-reload"),
    excel_path: Optional[str] = typer.Option(None, "--excel-path", "-e", 
                                            help="Excel files base path"),
    max_file_size: Optional[str] = typer.Option(None, "--max-file-size", "-s",
                                               help="Maximum file size (e.g., 100MB)"),
    allowed_paths: Optional[str] = typer.Option(None, "--allowed-paths", "-a",
                                               help="Allowed paths (colon-separated)"),
    enable_security: bool = typer.Option(True, "--security/--no-security",
                                        help="Enable security validations")
):
    """Start the Hiel Excel MCP Server."""
    
    # Setup logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)
    
    # Override configuration with CLI options
    if excel_path:
        os.environ['EXCEL_MCP_FILES_PATH'] = excel_path
    if max_file_size:
        # Parse size string (e.g., "100MB" -> bytes)
        size_bytes = parse_size_string(max_file_size)
        os.environ['EXCEL_MCP_MAX_FILE_SIZE'] = str(size_bytes)
    if allowed_paths:
        os.environ['EXCEL_MCP_ALLOWED_PATHS'] = allowed_paths
    
    os.environ['EXCEL_MCP_VALIDATE_PATHS'] = str(enable_security).lower()
    os.environ['EXCEL_MCP_LOG_LEVEL'] = log_level.upper()
    
    # Validate configuration
    try:
        config_dict = config.to_dict()
        logger.info("Server configuration:")
        for key, value in config_dict.items():
            if 'password' not in key.lower():  # Don't log sensitive info
                logger.info(f"  {key}: {value}")
    except Exception as e:
        logger.error(f"Configuration validation failed: {e}")
        typer.echo(f"❌ Configuration error: {e}", err=True)
        raise typer.Exit(1)
    
    # Import and start server
    try:
        from .server import app as fastmcp_app
        
        logger.info(f"Starting Hiel Excel MCP Server on {host}:{port}")
        logger.info(f"Excel files path: {config.excel_files_path}")
        logger.info(f"Security enabled: {enable_security}")
        
        uvicorn.run(
            "hiel_excel_mcp.server:app",
            host=host,
            port=port,
            workers=workers,
            log_level=log_level.lower(),
            reload=reload,
            access_log=log_level.upper() == 'DEBUG'
        )
        
    except Exception as e:
        logger.error(f"Failed to start server: {e}")
        typer.echo(f"❌ Server startup failed: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def stdio():
    """Run the server using stdio transport (recommended for local use)."""
    import asyncio
    from mcp.server.stdio import stdio_server
    
    setup_logging(config.log_level)
    logger = logging.getLogger(__name__)
    
    logger.info("Starting Hiel Excel MCP Server with stdio transport")
    
    try:
        from .server import app as mcp_app
        
        async def main():
            async with stdio_server() as (read_stream, write_stream):
                await mcp_app.run(read_stream, write_stream, mcp_app.create_initialization_options())
        
        asyncio.run(main())
        
    except Exception as e:
        logger.error(f"Failed to start server: {e}")
        typer.echo(f"❌ Server startup failed: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def streamable_http(
    host: str = typer.Option("0.0.0.0", help="Host to bind to"),
    port: int = typer.Option(8017, help="Port to bind to"),
):
    """Run the server using streamable HTTP transport (recommended for remote connections)."""
    import asyncio
    from fastmcp.transports.http import create_http_transport
    
    setup_logging(config.log_level)
    logger = logging.getLogger(__name__)
    
    logger.info(f"Starting Hiel Excel MCP Server with HTTP transport on {host}:{port}")
    
    try:
        from .server import app as mcp_app
        
        async def main():
            transport = create_http_transport(host=host, port=port)
            await mcp_app.run_transport(transport)
        
        asyncio.run(main())
        
    except Exception as e:
        logger.error(f"Failed to start server: {e}")
        typer.echo(f"❌ Server startup failed: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def validate_config(
    config_file: Optional[str] = typer.Option(None, "--config", "-c", 
                                             help="Configuration file path")
):
    """Validate server configuration."""
    
    setup_logging("INFO")
    logger = logging.getLogger(__name__)
    
    typer.echo("🔍 Validating Hiel Excel MCP configuration...")
    
    try:
        # Load configuration
        config_dict = config.to_dict()
        
        # Validate paths
        validation_results = []
        
        # Check Excel files path
        excel_path = Path(config.excel_files_path)
        if excel_path.exists():
            if excel_path.is_dir():
                validation_results.append(("✅", f"Excel path exists: {excel_path}"))
            else:
                validation_results.append(("❌", f"Excel path is not a directory: {excel_path}"))
        else:
            validation_results.append(("⚠️", f"Excel path does not exist: {excel_path}"))
        
        # Check allowed paths
        for allowed_path in config.allowed_paths:
            path = Path(allowed_path)
            if path.exists():
                validation_results.append(("✅", f"Allowed path exists: {path}"))
            else:
                validation_results.append(("❌", f"Allowed path does not exist: {path}"))
        
        # Print validation results
        typer.echo("\n📋 Configuration validation results:")
        for status, message in validation_results:
            typer.echo(f"  {status} {message}")
        
        # Print configuration summary
        typer.echo(f"\n⚙️ Configuration summary:")
        typer.echo(f"  Cache size: {config.cache_size}")
        typer.echo(f"  Cache age: {config.cache_age_seconds}s")
        typer.echo(f"  Max file size: {config.max_file_size / (1024*1024):.1f}MB")
        typer.echo(f"  Max concurrent operations: {config.max_concurrent_operations}")
        typer.echo(f"  Security enabled: {config.enable_path_validation}")
        typer.echo(f"  Sandbox mode: {config.sandbox_mode}")
        
        # Check if any critical errors
        critical_errors = [r for r in validation_results if r[0] == "❌"]
        if critical_errors:
            typer.echo(f"\n❌ Found {len(critical_errors)} critical configuration errors")
            raise typer.Exit(1)
        else:
            typer.echo(f"\n✅ Configuration validation passed")
            
    except Exception as e:
        typer.echo(f"❌ Configuration validation failed: {e}", err=True)
        raise typer.Exit(1)


@app.command()
def metrics(
    clear: bool = typer.Option(False, "--clear", "-c", help="Clear metrics after display")
):
    """Get server performance metrics."""
    
    setup_logging("INFO")
    
    typer.echo("📊 Fetching server metrics...")
    
    try:
        metrics = ExcelMCPUtils.get_performance_metrics()
        
        if not metrics:
            typer.echo("ℹ️ No performance metrics available")
            return
        
        typer.echo("\n📈 Performance Metrics:")
        for func_name, data in metrics.items():
            typer.echo(f"  {func_name}:")
            typer.echo(f"    Total calls: {data['total_calls']}")
            typer.echo(f"    Average time: {data['average_time']}s")
            typer.echo(f"    Max time: {data['max_time']}s")
            typer.echo(f"    Failure rate: {data['failure_rate']}%")
        
        if clear:
            ExcelMCPUtils.clear_performance_metrics()
            typer.echo("\n🧹 Metrics cleared")
            
    except Exception as e:
        typer.echo(f"❌ Failed to get metrics: {e}")
        raise typer.Exit(1)


def parse_size_string(size_str: str) -> int:
    """Parse size string like '100MB' into bytes."""
    size_str = size_str.strip().upper()
    
    if size_str.endswith('KB'):
        return int(size_str[:-2]) * 1024
    elif size_str.endswith('MB'):
        return int(size_str[:-2]) * 1024 * 1024
    elif size_str.endswith('GB'):
        return int(size_str[:-2]) * 1024 * 1024 * 1024
    else:
        # Assume bytes
        return int(size_str)


@app.command()
def version():
    """Show version information."""
    typer.echo("Hiel Excel MCP Server v1.0.0")
    typer.echo("Advanced Excel operations via MCP protocol")
    typer.echo("Built with FastMCP and OpenPyXL")


def main():
    """Main entry point."""
    app()


if __name__ == "__main__":
    main()