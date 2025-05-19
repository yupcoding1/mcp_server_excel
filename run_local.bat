@echo off
REM Run the MCP Excel server locally (Windows)
set EXCEL_FILES_DIR=./excel_files
if not exist .\excel_files mkdir .\excel_files
python main.py
