# 🤖 SARA Power BI Server (Source Edition) - Installation Guide

Welcome to **SARA** (Strategic Analysis & Reporting Assistant). 
This package is the **Source Code Edition**, allowing you to run the server directly via Python.

Developed by **Mateus.brito**.

---

## 📦 What's inside?
1.  **sara_powerbi (src)**: The Python source code. Includes PBIR editing and TOM/Analysis Services management.
2.  **run_source.bat**: A helper script that automatically creates a virtual environment, installs dependencies, and runs the server.

---

## 🚀 Step 1: Pre-requisites
1.  **Python 3.10+**: You must have Python installed and added to your system PATH.
2.  **Power BI Desktop**: Required locally to provide the necessary DLLs (Analysis Services libraries).

---

## 🔌 Step 2: First Run (Setup)
1.  Double-click `run_source.bat`.
2.  Wait while it:
    - Creates a `venv` folder.
    - Installs libraries (`pythonnet`, etc).
    - Starts the server.
3.  Once you see "Starting SARA Power BI Server...", it is working! You can close this window now.

---

## ⚙️ Step 3: MCP Configuration
To use this with an AI Client (like Claude Desktop or any MCP Client), you need to configure the connection.

Use the `config.example.json` file as a template. You need to provide **Absolute Paths**.

**Example Configuration:**
```json
{
  "mcpServers": {
    "sara-powerbi": {
      "command": "C:\\FULL\\PATH\\TO\\THIS_FOLDER\\venv\\Scripts\\python.exe",
      "args": [
        "-u",
        "-m",
        "sara_powerbi.server"
      ],
      "env": {
        "PYTHONPATH": "C:\\FULL\\PATH\\TO\\THIS_FOLDER\\src",
        "PYTHONUTF8": "1"
      }
    }
  }
}
```

**Important:** 
- Replace `C:\\FULL\\PATH\\TO\\THIS_FOLDER` with the actual path where you're running this.
- The `env` section with `PYTHONPATH` is **CRITICAL** for the server to find its own modules.

---

## 🛠️ Usage
Once configured in your AI Client, the `sara-powerbi` toolset will be available for:
- Creating/Editing Power BI Reports (PBIR).
- Managing Semantic Models (TOM/XMLA).
- Running DAX Queries.

**Troubleshooting:**
- **DLL Errors:** If you see errors about missing DLLs, run Power BI Desktop once to ensure libraries are registered, then run `run_source.bat` again to let the auto-fix mechanism work.
