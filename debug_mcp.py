#!/usr/bin/env python3
"""
Debug script to test MCP server manually.
This simulates what Claude Desktop does when connecting to the MCP server.
"""

import sys
import os
sys.path.insert(0, '/home/serhabdel/Documents/repos/Agent/MCPs/hiel_excel_mcp')

def test_import():
    """Test importing the server."""
    try:
        from hiel_excel_mcp import __main__
        print("✅ Successfully imported hiel_excel_mcp.__main__")
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False

def test_server_info():
    """Test getting server info."""
    try:
        from hiel_excel_mcp.server import get_all_tools_info
        info = get_all_tools_info()
        print(f"✅ Server info: {info['server_name']} v{info['version']}")
        print(f"✅ Tools: {info['total_tools']}")
        return True
    except Exception as e:
        print(f"❌ Server info error: {e}")
        return False

def test_manual_run():
    """Test running the server manually."""
    print("🧪 Testing manual server run...")
    
    # Set up environment
    os.environ['EXCEL_FILES_PATH'] = '/home/serhabdel/Documents/repos/Agent/MCPs/hiel_excel_mcp'
    
    try:
        # Import and run
        from hiel_excel_mcp.server import app
        print("✅ Server app imported")
        
        # Test that tools are registered
        print(f"✅ Server name: {app.name}")
        
        return True
    except Exception as e:
        print(f"❌ Manual run error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Run all debug tests."""
    print("🔍 DEBUGGING HIEL EXCEL MCP SERVER")
    print("=" * 50)
    
    print(f"📁 Current working directory: {os.getcwd()}")
    print(f"🐍 Python executable: {sys.executable}")
    print(f"📦 Python path: {sys.path[:3]}...")
    print()
    
    tests = [
        ("Import Test", test_import),
        ("Server Info Test", test_server_info), 
        ("Manual Run Test", test_manual_run)
    ]
    
    passed = 0
    for name, test_func in tests:
        print(f"🧪 {name}...")
        if test_func():
            passed += 1
        print()
    
    print("=" * 50)
    print(f"📊 Results: {passed}/{len(tests)} tests passed")
    
    if passed == len(tests):
        print("🎉 All tests passed! The MCP server should work.")
        print("\n📋 Try this configuration:")
        print('{')
        print('  "mcpServers": {')
        print('    "hiel-excel-mcp": {')
        print('      "command": "/usr/bin/python3",')
        print('      "args": ["-m", "hiel_excel_mcp"],')
        print(f'      "cwd": "{os.path.abspath(".")}",')
        print('      "env": {')
        print(f'        "PYTHONPATH": "{os.path.abspath(".")}",')
        print(f'        "EXCEL_FILES_PATH": "{os.path.abspath(".")}"')
        print('      }')
        print('    }')
        print('  }')
        print('}')
    else:
        print("❌ Some tests failed. Check the errors above.")

if __name__ == "__main__":
    main()