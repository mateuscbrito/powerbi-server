# 📊 Agentic Power BI (SARA MCP)

**The missing link between AI Agents and Power BI Desktop.**

This [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server enables LLMs (like Claude, Gemini, or ChatGPT) to programmatically control, analyze, and build Power BI reports. It leverages the new **PBIR (Power BI Enhanced Report Format)** for report manipulation and **TOM (Tabular Object Model)** for semantic model management.

---

## 🚀 Key Features

### 1. Report Operations (PBIR)
Manipulate the visual layer of your report by directly editing JSON definitions.
- **✨ Create & Layout:** programmatic creation of Pages and Visuals (Bar Charts, Cards, Textboxes).
- **🎨 Formatting:** Bulk rename titles, axes, and field names across visuals.
- **🔄 Refactoring:** Swap measures/columns in *all* visuals at once (e.g., replace `[Sales]` with `[Total Sales]`).
- **audit 🔍 Audit:** Find exactly where a measure or column is used in the report (Page & Visual level).

### 2. Semantic Model Operations (TOM)
Interact with the live analysis services instance inside Power BI Desktop.
- **Create Measures/Columns:** Write DAX expressions and inject them into the model.
- **Manage Relationships & Roles:** Create relationships and RLS roles on the fly.
- **Schema Search:** Deep search for tables, columns, and expressions.

### 3. Analysis & Optimization
- **⚡ VertiPaq Analysis:** Identify the heavy columns consuming your RAM.
- **🧠 DAX Execution:** Run DAX queries directly against the model and get results as JSON.

---

## 🛠️ Installation

### Prerequisites
- Windows OS (Required for Power BI Desktop).
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

**Important:** Point `command` to your specific Python executable.

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

## 💡 Usage Examples

### Refactoring a Measure
> "Replace the measure 'Sum of Sales' with 'Total Sales' in all charts."

The agent will call `pbir_refactor_field(table="Sales", old_name="Sum of Sales", new_name="Total Sales")`.

### Creating a Dashboard
> "Create a new page 'Analysis' and add a Bar Chart showing Sales by Category."

The agent will call `pbir_create_page` followed by `pbir_create_bar_chart`.

### Audit
> "Where is the measure 'Profit Margin' being used?"

The agent will call `pbir_audit_usage` and list every visual containing that measure.

---

## ⚠️ Disclaimer
This is an **experimental** tool. Always back up your Power BI projects (use Git!) before running automated refactoring tools. The PBIR format is subject to changes by Microsoft.

---

**License:** MIT
