#!/bin/bash

echo
echo "==================================="
echo " Excel Finance MCP Bridge Starter"
echo "==================================="
echo

# Check if we're in the bridge directory
if [ ! -f "package.json" ]; then
    echo "Error: Please run this script from the bridge directory"
    exit 1
fi

# Check if node_modules exists
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
    if [ $? -ne 0 ]; then
        echo "Failed to install dependencies"
        exit 1
    fi
fi

# Build the bridge
echo "Building bridge..."
npm run build
if [ $? -ne 0 ]; then
    echo "Build failed"
    exit 1
fi

# Check if parent MCP server is built
if [ ! -f "../dist/index.js" ]; then
    echo "Building parent MCP server..."
    cd ..
    npm run build
    if [ $? -ne 0 ]; then
        echo "Failed to build parent MCP server"
        exit 1
    fi
    cd bridge
fi

echo
echo "Starting Excel Finance MCP Bridge..."
echo "Web interface will be available at: http://localhost:3001"
echo "Press Ctrl+C to stop"
echo

# Start the bridge
npm start