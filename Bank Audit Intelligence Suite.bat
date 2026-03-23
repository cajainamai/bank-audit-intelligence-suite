@echo off
TITLE Bank Audit Intelligence Suite - CA Jainam Shah
echo.
echo ============================================================
echo   Starting Bank Audit Intelligence Suite...
echo ============================================================
echo.

cd /d "%~dp0source-code"

:: Open the browser after a short delay to allow server startup
echo Opening browser at http://localhost:5180...
start "" "http://localhost:5180"

:: Start the development server on a specific port to avoid conflicts
npm run dev -- --host --port 5180

pause
