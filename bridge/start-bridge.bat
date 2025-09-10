@echo off
echo.
echo ===================================
echo  Excel Finance MCP Bridge Starter
echo ===================================
echo.

REM Check if we're in the bridge directory
if not exist "package.json" (
    echo Error: Please run this script from the bridge directory
    pause
    exit /b 1
)

REM Check if node_modules exists
if not exist "node_modules" (
    echo Installing dependencies...
    npm install
    if errorlevel 1 (
        echo Failed to install dependencies
        pause
        exit /b 1
    )
)

REM Build the bridge
echo Building bridge...
npm run build
if errorlevel 1 (
    echo Build failed
    pause
    exit /b 1
)

REM Check if parent MCP server is built
if not exist "..\dist\index.js" (
    echo Building parent MCP server...
    cd ..
    npm run build
    if errorlevel 1 (
        echo Failed to build parent MCP server
        pause
        exit /b 1
    )
    cd bridge
)

echo.
echo Starting Excel Finance MCP Bridge...
echo Web interface will be available at: http://localhost:3001
echo Press Ctrl+C to stop
echo.

REM Start the bridge
npm start