#!/bin/bash
# gdrive-exporter Launcher for Unix/Linux/macOS
# This script launches the main.py application

set -e  # Exit on any error

echo "Starting gdrive-exporter..."

# Check if Python is installed
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "Error: Python is not installed or not in PATH"
    echo "Please install Python and add it to your PATH"
    exit 1
fi

# Determine which Python command to use
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
else
    PYTHON_CMD="python"
fi

echo "Using Python: $PYTHON_CMD"

# Check if main.py exists in gdrive-exporter folder
if [ ! -f "gdrive-exporter/main.py" ]; then
    echo "Error: main.py not found in gdrive-exporter directory"
    echo "Please ensure the gdrive-exporter folder contains main.py"
    exit 1
fi

# Make the script executable (if it isn't already)
chmod +x "$0"

# Run the main application with all passed arguments
exec "$PYTHON_CMD" gdrive-exporter/main.py "$@"
