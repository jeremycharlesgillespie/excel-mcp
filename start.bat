@echo off
echo Building Excel Finance MCP...
call npm run build
if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Build successful! Starting MCP server...
echo.
echo Server is now running. Press Ctrl+C to stop.
echo Configure Claude Desktop with this server path if not already done.
echo.
node dist/index.js