# 📊 SARA Power BI Server (MCP)

**The missing link between AI Agents and Power BI Desktop.**

This [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server enables LLMs (like Claude, Gemini, or ChatGPT) to programmatically control, analyze, and build Power BI reports. It leverages the new **PBIR (Power BI Enhanced Report Format)** for visual manipulation and **TOM (Tabular Object Model)** for semantic model management.

![SARA Briefing UI](https://placehold.co/600x400?text=SARA+Briefing+UI) 

---

## 🚀 Quick Start: The AI Briefing Workflow

We've included a specialized interface to help you generate structured prompts for your AI agent.

1. **Open the Briefing UI:**
   Navigate to the `ui/` folder and open `briefing.html` in your browser.

2. **Fill the Form:**
   Describe your project needs: KPIs, Dimensions, Data Sources, and visual preferences.

3. **Copy to Agent:**
   Click **"Salvar e Copiar para SARA"**. This generates a structured markdown prompt.
   Paste this prompt into your AI Chat (connected to this MCP server).

4. **Watch the Magic:**
   The AI will read the briefing and start building your report using the tools below!

---

## 🛠️ Installation

### Prerequisites
- Windows OS (Required for Power BI Desktop Automation).
- [Power BI Desktop](https://powerbi.microsoft.com/desktop/).
- Python 3.10+.
- **Recommended:** Enable "Power BI Project (.pbip)" save option in Power BI Preview Features.

### Setup
1. Clone this repository:
   ```bash
   git clone https://github.com/mateuscbrito/powerbi-server.git
   cd powerbi-server
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## ⚙️ Configuration

Add the server to your MCP Client configuration (e.g., `claude_desktop_config.json` or `mcp_config.json`).
**Note:** Ensure you use the absolute path to your Python executable and the `src/sara_powerbi/server.py` file.

```json
{
  "mcpServers": {
    "sara-powerbi": {
      "command": "C:\\path\\to\\python.exe",
      "args": [
        "-u",
        "C:\\path\\to\\powerbi-server\\src\\sara_powerbi\\server.py"
      ],
      "disabledTools": ["visualize_python"]
    }
  }
}
```

---

## 🧰 Tools Reference

### 1. Visual Report Management (PBIR)
*Targeting `Acompanhamento PET Fornecedor.Report/definition/...`*

| Tool | Description |
|------|-------------|
| `pbir_get_info` | Detects the active project and lists all report pages. |
| `pbir_create_page` | Creates a new blank page in the report. |
| `pbir_create_visual` | Creates basic visuals (Card, Textbox) on a specific page. |
| `pbir_create_bar_chart` | Creates a Clustered Bar Chart with Category/Value fields. |
| `pbir_format_visual` | Updates visual titles and renames axis labels/legend fields. |
| `pbir_update_visual_layout` | Moves and resizes visuals (X, Y, Width, Height). |
| `pbir_bind_measure` | Changes the measure displayed in a Card visual. |
| `pbir_refactor_field` | **Powerful:** Renames a measure/column in *all* visuals across the report. |
| `pbir_audit_usage` | **Audit:** Finds every visual where a specific measure/column is used. |
| `pbir_list_visuals` | Lists all visuals on a page with their IDs and positions. |
| `pbir_delete_object` | Deletes a Page or a Visual. |

### 2. Semantic Model Management (TOM)
*Interacting with the running Power BI Analysis Services instance.*

| Tool | Description |
|------|-------------|
| `manage_measure` | Create, Update, or Delete DAX measures. |
| `manage_column` | Rename, Hide, or Change Type of columns. |
| `manage_table` | Create Calculated Tables or M Tables. |
| `manage_relationship` | Create/Delete relationships between tables. |
| `manage_role` | Create Row-Level Security (RLS) roles with DAX filters. |
| `run_dax` | Execute any DAX query and get JSON results (Limit: 2000 rows). |
| `search_model` | Deep search for objects (Tables, Columns, Eras) by name. |
| `get_vertipaq_stats` | **Optimization:** List the top 20 heaviest columns (RAM usage). |
| `manage_model_connection` | Check connection status to Power BI Desktop. |

---

## 🏗️ Architecture

This project uses a modular Python architecture:
- **`src/sara_powerbi/server.py`**: Main entry point using `FastMCP`.
- **`src/sara_powerbi/tools/pbir.py`**: Logic for parsing and editing JSON report definitions.
- **`src/sara_powerbi/tools/tom.py`**: Logic for communicating with `msmdsrv.exe` via `pythonnet`.
- **`ui/`**: Contains the standalone Briefing Assistant.

---

## ⚠️ Disclaimer
This is an **experimental** tool. Always back up your Power BI projects (use Git!) before running automated refactoring tools. The PBIR format is subject to changes by Microsoft.

---

**License:** MIT
