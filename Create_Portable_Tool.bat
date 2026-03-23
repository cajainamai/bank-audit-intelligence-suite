@echo off
TITLE Create Portable Bank Audit Tool
echo.
echo ============================================================
echo   Creating your Portable Offline Tool...
echo ============================================================
echo.

cd /d "%~dp0source-code"

echo [1/2] Installing build tools (this may take a minute first time)...
call npm install

echo.
echo [2/2] "Baking" into a single portable file...
call npm run build

echo.
echo ============================================================
echo   SUCCESS!
echo   Your portable tool is ready here:
echo   %~dp0dist\index.html
echo.
echo   You can rename index.html to "Bank_Audit_Tool.html" 
echo   and share it with anyone!
echo ============================================================
echo.
pause
