#!/bin/bash

# LGU Payroll System - Setup & Install Script for Linux/macOS

echo "=========================================="
echo "  LGU Payroll System - Setup & Install"
echo "=========================================="
echo

# Check for Node.js
echo "[1/3] Checking for Node.js..."
if ! command -v node &> /dev/null
then
    echo "[ERROR] Node.js is not installed!"
    echo "Please install Node.js from https://nodejs.org/"
    exit 1
fi
echo "[OK] Node.js is installed."

# Install NPM dependencies
echo
echo "[2/3] Installing system dependencies (node_modules)..."
echo "Please wait..."
echo

npm install

if [ $? -ne 0 ]; then
    echo
    echo "[ERROR] Failed to install dependencies."
    exit 1
fi
echo
echo "[OK] Dependencies installed successfully."

# Final Check
echo
echo "[3/3] Finalizing setup..."
mkdir -p data
echo "[OK] Data directory verified."

echo
echo "=========================================="
echo "  SETUP COMPLETE!"
echo "=========================================="
echo
echo "You can now run the system using: npm start"
echo
