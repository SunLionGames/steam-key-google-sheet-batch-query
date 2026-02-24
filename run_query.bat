@echo off
TITLE Steam Key Batch Query
echo Checking dependencies...
IF NOT EXIST "node_modules\" (
    echo node_modules not found. Installing...
    npm install
)
echo.
echo Starting Script...
node index.js
echo.
echo Process finished.
pause