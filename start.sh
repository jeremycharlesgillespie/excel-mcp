#!/bin/bash
echo "Building Excel Finance MCP..."
npm run build

if [ $? -ne 0 ]; then
    echo "Build failed!"
    exit 1
fi

echo "Build successful! Starting MCP server..."
echo ""
echo "Server is now running. Press Ctrl+C to stop."
echo "Configure Claude Desktop with this server path if not already done."
echo ""
node dist/index.js