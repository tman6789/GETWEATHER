#!/bin/bash
# Activate your virtual environment and run the weather tool

cd "$(dirname "$0")"  # go to the folder the script is in
source .venv/bin/activate
python weather_tool.py

