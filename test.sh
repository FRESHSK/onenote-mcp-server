#!/bin/bash

# OneNote MCP Server Test Script

echo "OneNote MCP Server Test Script"
echo "=============================="
echo ""

# Check if server.js exists
if [ ! -f "server.js" ]; then
    echo "Error: server.js not found. Make sure you're in the correct directory."
    exit 1
fi

# Check if node_modules exists
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
fi

# Check if AZURE_CLIENT_ID is set
if [ -z "$AZURE_CLIENT_ID" ]; then
    echo "Warning: AZURE_CLIENT_ID is not set."
    echo "Please set it before running tests:"
    echo "  export AZURE_CLIENT_ID='your-client-id'"
    echo ""
fi

echo "Starting test sequence..."
echo ""

# Test 1: List notebooks
echo "Test 1: Listing notebooks"
echo '{"tool":"onenote-read","input":{"type":"list_notebooks"}}' | node server.js
echo ""

# Instructions for interactive testing
echo "Interactive Testing:"
echo "==================="
echo ""
echo "1. To list notebooks:"
echo '   echo '"'"'{"tool":"onenote-read","input":{"type":"list_notebooks"}}'"'"' | node server.js'
echo ""
echo "2. To list sections (replace notebook-id):"
echo '   echo '"'"'{"tool":"onenote-read","input":{"type":"list_sections","notebookId":"notebook-id"}}'"'"' | node server.js'
echo ""
echo "3. To create a test notebook:"
echo '   echo '"'"'{"tool":"onenote-create","input":{"type":"create_notebook","displayName":"Test Notebook"}}'"'"' | node server.js'
echo ""
echo "4. To run the server interactively:"
echo "   node server.js"
echo "   Then type JSON commands and press Enter"
echo ""
echo "Note: Check server.log for detailed logs"