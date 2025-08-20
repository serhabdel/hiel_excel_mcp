#!/bin/bash
# Hiel Excel MCP Server Runner Script

# Set environment variables
export EXCEL_FILES_PATH="${EXCEL_FILES_PATH:-$(pwd)/data}"
export LOG_LEVEL="${LOG_LEVEL:-INFO}"
export MAX_FILE_SIZE="${MAX_FILE_SIZE:-52428800}"
export CACHE_SIZE="${CACHE_SIZE:-20}"
export CACHE_TTL="${CACHE_TTL:-300}"

# Create data directory if it doesn't exist
mkdir -p "$EXCEL_FILES_PATH"

# Add current directory to Python path
export PYTHONPATH="${PYTHONPATH}:$(pwd)"

echo "🚀 Starting Hiel Excel MCP Server..."
echo "📂 Excel files path: $EXCEL_FILES_PATH"
echo "📊 Log level: $LOG_LEVEL"

# Run the server
python3 server.py "$@"
