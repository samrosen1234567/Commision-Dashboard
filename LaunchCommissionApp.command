#!/bin/zsh
# Commission Reconciliation App Launcher

# Change to the directory where this script resides
cd "$(dirname "$0")"

# 1) Install Python3 via Homebrew if missing
if ! command -v python3 >/dev/null; then
    echo "Python3 not found. Installing Homebrew and Python3..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    brew install python
fi

# 2) Ensure pip and install dependencies
python3 -m ensurepip --upgrade
python3 -m pip install --upgrade pip
python3 -m pip install PySimpleGUI pandas openpyxl

# 3) Run the application
python3 commission_reconciler_app.py
