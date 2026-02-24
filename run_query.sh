#!/bin/bash
echo "Checking dependencies..."
if [ ! -d "node_modules" ]; then
    echo "node_modules not found. Installing..."
    npm install
fi
echo ""
echo "Starting Script..."
node index.js
echo ""
read -p "Process finished. Press enter to exit."