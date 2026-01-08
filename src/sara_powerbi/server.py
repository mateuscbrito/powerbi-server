import sys
import os

# === WINDOWS FIX for MCP (Binary Mode) ===
if sys.platform == "win32":
    import msvcrt
    try:
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
        msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
    except: pass
# =========================================

from mcp.server.fastmcp import FastMCP
from sara_powerbi.tools import pbir, tom

# Initialize Server
mcp = FastMCP("sara-powerbi-ultimate")

# --- REGISTER TOOLS ---

# PBIR Tools
mcp.add_tool(pbir.pbir_get_info)
mcp.add_tool(pbir.pbir_inspect_structure)
mcp.add_tool(pbir.pbir_create_page)
mcp.add_tool(pbir.pbir_create_visual)
mcp.add_tool(pbir.pbir_create_bar_chart)
mcp.add_tool(pbir.pbir_bind_measure)
mcp.add_tool(pbir.pbir_format_visual)
mcp.add_tool(pbir.pbir_refactor_field)
mcp.add_tool(pbir.pbir_audit_usage)
mcp.add_tool(pbir.pbir_list_visuals)
mcp.add_tool(pbir.pbir_delete_object)
mcp.add_tool(pbir.pbir_update_visual_layout)

# TOM Tools
mcp.add_tool(tom.manage_model_connection)
mcp.add_tool(tom.list_objects)
mcp.add_tool(tom.search_model)
mcp.add_tool(tom.run_dax)
mcp.add_tool(tom.manage_measure)
mcp.add_tool(tom.manage_column)
mcp.add_tool(tom.manage_table)
mcp.add_tool(tom.manage_relationship)
mcp.add_tool(tom.manage_role)
mcp.add_tool(tom.manage_calc_group)
mcp.add_tool(tom.get_model_info)
mcp.add_tool(tom.get_vertipaq_stats)

def main():
    """Entry point for the server."""
    print(f"Starting SARA Power BI Server...", file=sys.stderr)
    mcp.run()

if __name__ == "__main__":
    main()
