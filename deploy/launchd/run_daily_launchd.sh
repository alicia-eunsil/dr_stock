#!/bin/bash

set -euo pipefail

if [ -f "$HOME/.bash_profile" ]; then
  # Load local environment variables such as KIS_APP_KEY.
  . "$HOME/.bash_profile"
fi

cd "/Users/alicia/Desktop/#python/knee_shoulder_stock"
/usr/bin/python3 run_daily.py
