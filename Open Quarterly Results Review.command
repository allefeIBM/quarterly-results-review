#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

/usr/bin/python3 ./start.py

echo
read -n 1 -s -r -p "Press any key to close..."
