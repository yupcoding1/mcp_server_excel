# MCP Excel Server

A Python-based Model Context Protocol (MCP) server for advanced Excel file manipulation. This project exposes Excel file operations (create, read, write, format, etc.) as MCP resources and tools, making it easy to automate and integrate Excel workflows.

## Table of Contents
- [Project Structure](#project-structure)
- [Features](#features)
- [Requirements](#requirements)
- [Setup](#setup)
- [Running the Server](#running-the-server)
- [Usage](#usage)
- [Docker Usage](#docker-usage)
- [Environment Variables](#environment-variables)
- [Directory Details](#directory-details)
- [Development Notes](#development-notes)
- [References](#references)

---

## Project Structure

```
.
├── advanced_server.py        # Advanced MCP server with async support
├── main.py                   # Main MCP server (FastMCP)
├── excel_fucntion.py         # All Excel file manipulation functions
├── requirements.txt          # Python dependencies
├── Dockerfile                # Docker build for advanced_server.py
├── Docker_advanced.txt       # Alternate Dockerfile for advanced_server.py
├── run_local.bat             # Windows batch script to run the server
├── run_local.sh              # Bash script to run the server
├── .env                      # Environment variables (e.g., EXCEL_FILES_DIR)
├── excel_files/              # Directory for Excel files (auto-created)
│   └── trial.xlsx            # Example Excel file
└── __pycache__/              # Python bytecode cache
```

## Features
- List, create, rename, and delete Excel files and sheets
- Read/write cell values, rows, columns, and ranges
- Merge/unmerge cells, set borders, auto-fit columns
- Write formulas, save as new file
- All operations exposed as MCP tools/resources
- Async server (advanced_server.py) and FastMCP server (main.py)
- Docker support for easy deployment

## Requirements
- Python 3.11+
- [openpyxl](https://openpyxl.readthedocs.io/)
- [python-dotenv](https://pypi.org/project/python-dotenv/)
- [pydantic](https://pydantic-docs.helpmanual.io/)
- [mcp](https://github.com/modelcontext/model-context-protocol)

Install dependencies:
```sh
pip install -r requirements.txt
```

## Setup
1. Clone the repository and navigate to the project directory.
2. Ensure you have Python 3.11+ installed.
3. (Optional) Edit `.env` to set `EXCEL_FILES_DIR` (default: `./excel_files`).
4. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

## Running the Server

### On Windows (PowerShell or CMD):
```bat
run_local.bat
```

### On Linux/macOS:
```sh
./run_local.sh
```

This will start the MCP Excel server using `main.py` and create the `excel_files/` directory if it does not exist.

### To run the advanced async server:
```sh
python advanced_server.py
```

## Usage
- The server exposes Excel file operations as MCP resources and tools.
- Integrate with any MCP-compatible client or use the CLI for testing.
- See `main.py` and `advanced_server.py` for all available tools and their parameters.

## Docker Usage

### Build and run with the provided Dockerfile:
```sh
docker build -t mcp-excel-server .
docker run -it --rm -v $(pwd)/excel_files:/app/excel_files mcp-excel-server
```
- The default entrypoint runs `advanced_server.py`.
- To use `main.py`, edit the `CMD` in the Dockerfile.
- The Excel files directory is mounted for persistence.

## Environment Variables
- `EXCEL_FILES_DIR`: Directory for storing Excel files (default: `./excel_files`).
  - Set in `.env`, or via environment when running Docker or scripts.

## Directory Details
- `excel_files/`: All Excel files created/modified by the server are stored here.
- `__pycache__/`: Python bytecode cache (can be ignored).

## Development Notes
- All Excel logic is in `excel_fucntion.py`.
- Add new tools/resources by editing `main.py` or `advanced_server.py`.
- For custom environments, update `.env` or pass variables directly.
- For MCP protocol details, see [modelcontext/model-context-protocol](https://github.com/modelcontext/model-context-protocol).

## References
- [Model Context Protocol (MCP) Specification](https://github.com/modelcontextprotocol/servers.git)
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk.git)

---

**Author:** Mohammed Abbasi

**License:** Use as you want.
