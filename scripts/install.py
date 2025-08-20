#!/usr/bin/env python3
"""
Installation script for Hiel Excel MCP Server.
Supports both pip and uvx installation methods with optimal configuration.
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
from typing import Optional, List
import json

def detect_package_manager() -> str:
    """Detect available package managers."""
    if shutil.which("uvx"):
        return "uvx"
    elif shutil.which("uv"):
        return "uv"  
    elif shutil.which("pip"):
        return "pip"
    else:
        return "none"

def install_with_uvx() -> bool:
    """Install using uvx for isolated execution."""
    try:
        print("🚀 Installing Hiel Excel MCP Server with uvx...")
        
        # Install the package
        result = subprocess.run([
            "uvx", "install", "hiel-excel-mcp[dev]",
            "--force"  # Force reinstall if already exists
        ], check=True, capture_output=True, text=True)
        
        print("✅ Installation completed successfully!")
        print("\n📋 Usage examples:")
        print("   uvx hiel-excel-mcp serve --port 8000")
        print("   uvx hiel-excel-mcp validate-config")
        print("   uvx hiel-excel-mcp stdio")
        print("   uvx hiel-excel-mcp streamable-http --port 8017")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ uvx installation failed: {e}")
        print(f"Error output: {e.stderr}")
        return False
    except FileNotFoundError:
        print("❌ uvx not found. Please install uv first:")
        print("   curl -LsSf https://astral.sh/uv/install.sh | sh")
        return False

def install_with_pip() -> bool:
    """Install using pip."""
    try:
        print("📦 Installing Hiel Excel MCP Server with pip...")
        
        # Upgrade pip first
        subprocess.run([
            sys.executable, "-m", "pip", "install", "--upgrade", "pip"
        ], check=True)
        
        # Install the package with all dependencies
        result = subprocess.run([
            sys.executable, "-m", "pip", "install", 
            "hiel-excel-mcp[dev]",
            "--upgrade"
        ], check=True, capture_output=True, text=True)
        
        print("✅ Installation completed successfully!")
        print("\n📋 Usage examples:")
        print("   hiel-excel-mcp serve --port 8000")
        print("   hiel-excel-mcp validate-config")
        print("   hiel-excel-mcp stdio")
        print("   python -m hiel_excel_mcp streamable-http")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ pip installation failed: {e}")
        print(f"Error output: {e.stderr}")
        return False

def create_config_file() -> bool:
    """Create default configuration file."""
    try:
        config_dir = Path.home() / ".config" / "hiel-excel-mcp"
        config_dir.mkdir(parents=True, exist_ok=True)
        
        config_file = config_dir / "config.json"
        
        default_config = {
            "excel_files_path": str(Path.home() / "Documents"),
            "max_file_size_mb": 100,
            "cache_size": 10,
            "cache_age_seconds": 300,
            "log_level": "INFO",
            "enable_security": True,
            "allowed_extensions": [".xlsx", ".xls", ".csv", ".xlsm"],
            "performance": {
                "enable_monitoring": True,
                "memory_threshold_mb": 500,
                "max_concurrent_operations": 5
            }
        }
        
        with open(config_file, 'w') as f:
            json.dump(default_config, f, indent=2)
        
        print(f"📝 Created configuration file: {config_file}")
        print("   You can edit this file to customize settings")
        
        return True
        
    except Exception as e:
        print(f"⚠️ Could not create config file: {e}")
        return False

def setup_shell_completion() -> bool:
    """Set up shell completion."""
    try:
        # Try to set up bash completion
        bash_completion_dir = Path.home() / ".bash_completion.d"
        bash_completion_dir.mkdir(exist_ok=True)
        
        completion_script = bash_completion_dir / "hiel-excel-mcp"
        
        # Generate completion script
        try:
            result = subprocess.run([
                "hiel-excel-mcp", "--install-completion", "bash"
            ], capture_output=True, text=True, check=True)
            
            print("✅ Shell completion installed")
            print("   Restart your shell or run: source ~/.bashrc")
            return True
            
        except subprocess.CalledProcessError:
            print("⚠️ Could not install shell completion")
            return False
            
    except Exception as e:
        print(f"⚠️ Error setting up completion: {e}")
        return False

def verify_installation() -> bool:
    """Verify the installation works correctly."""
    try:
        print("\n🔍 Verifying installation...")
        
        # Try to run version command
        result = subprocess.run([
            "hiel-excel-mcp", "version"
        ], capture_output=True, text=True, check=True)
        
        print("✅ Installation verified successfully!")
        print(result.stdout)
        
        # Try config validation
        result = subprocess.run([
            "hiel-excel-mcp", "validate-config"
        ], capture_output=True, text=True, check=True)
        
        print("✅ Configuration validation passed!")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Installation verification failed: {e}")
        return False
    except FileNotFoundError:
        print("❌ hiel-excel-mcp command not found after installation")
        return False

def main():
    """Main installation function."""
    print("🚀 Hiel Excel MCP Server Installation")
    print("=" * 50)
    
    # Detect package manager
    pm = detect_package_manager()
    
    if pm == "none":
        print("❌ No suitable package manager found!")
        print("Please install one of the following:")
        print("  • uv: curl -LsSf https://astral.sh/uv/install.sh | sh")
        print("  • pip: python -m ensurepip --upgrade")
        sys.exit(1)
    
    print(f"📦 Detected package manager: {pm}")
    
    # Choose installation method
    if pm == "uvx":
        print("\n🎯 Recommended: uvx installation for isolated execution")
        use_uvx = input("Use uvx? [Y/n]: ").strip().lower()
        if use_uvx in ("", "y", "yes"):
            success = install_with_uvx()
        else:
            success = install_with_pip()
    else:
        success = install_with_pip()
    
    if not success:
        print("\n❌ Installation failed!")
        sys.exit(1)
    
    # Create configuration
    print("\n⚙️ Setting up configuration...")
    create_config_file()
    
    # Set up completion
    print("\n🔧 Setting up shell completion...")
    setup_shell_completion()
    
    # Verify installation
    verify_installation()
    
    print("\n🎉 Installation completed successfully!")
    print("\n🚀 Quick Start:")
    print("   # Start server with HTTP transport")
    if pm == "uvx":
        print("   uvx hiel-excel-mcp serve --port 8000")
        print("\n   # Or use stdio for local development")  
        print("   uvx hiel-excel-mcp stdio")
    else:
        print("   hiel-excel-mcp serve --port 8000")
        print("\n   # Or use stdio for local development")
        print("   hiel-excel-mcp stdio")
    
    print("\n📚 Documentation:")
    print("   hiel-excel-mcp --help")
    print("   hiel-excel-mcp validate-config")
    print("   hiel-excel-mcp metrics")
    
if __name__ == "__main__":
    main()