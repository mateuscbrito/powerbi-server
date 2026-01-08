"""
SARA MCP Server - Power BI Ultimate Edition v6 (Complete Toolset)
Implements FULL schema management including tables, partitions, calc groups, and roles.
"""
import sys
import os
import json
import uuid
from typing import Optional, List, Dict, Any
import subprocess
import tempfile

# === WINDOWS FIX ===
if sys.platform == "win32":
    import msvcrt
    try:
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
        msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
    except: pass
# ===================

from mcp.server.fastmcp import FastMCP

# Initialize Server
mcp = FastMCP("sara-powerbi-ultimate")

# --- GLOBAL STATE & DLL LOADING ---
GLOBAL_CONTEXT = {"port": None, "server": None}

def _load_libs():
    import psutil, clr
    search_paths = [r"C:\Program Files\Microsoft Power BI Desktop\bin", r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin"]
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        if proc.info['name'] and 'msmdsrv.exe' in proc.info['name'].lower():
            try:
                if proc.info['exe']: search_paths.insert(0, os.path.dirname(proc.info['exe']))
            except: pass
    
    loaded = False
    for path in search_paths:
        if not os.path.exists(path): continue
        try:
            sys.path.append(path)
            try: clr.AddReference("Microsoft.AnalysisServices.AdomdClient")
            except: clr.AddReference("Microsoft.PowerBI.AdomdClient")
            try: clr.AddReference("Microsoft.AnalysisServices.Tabular"); loaded = True
            except: 
                try: clr.AddReference("Microsoft.PowerBI.Tabular"); loaded = True
                except: pass
        except: pass
        if loaded: break
    return loaded

def _find_port():
    import psutil
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'msmdsrv.exe' in proc.info['name'].lower():
            try:
                for conn in proc.connections(kind='tcp'):
                    if conn.status == 'LISTEN' and conn.laddr.ip == '127.0.0.1':
                        return conn.laddr.port
            except: pass
    return None

def _get_server():
    if not _load_libs(): raise Exception("DLLs load failed")
    port = _find_port()
    if not port: raise Exception("Power BI not running")
    
    from Microsoft.AnalysisServices.Tabular import Server
    s = Server()
    s.Connect(f"Data Source=localhost:{port};")
    return s

def _run_dax_raw(query: str):
    _load_libs()
    port = _find_port()
    if not port: raise Exception("Power BI not running")
    
    try: from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
    except: from Microsoft.PowerBI.AdomdClient import AdomdConnection
    
    conn = AdomdConnection(f"Data Source=localhost:{port};")
    conn.Open()
    cmd = conn.CreateCommand(); cmd.CommandText = query
    reader = cmd.ExecuteReader()
    cols = [reader.GetName(i) for i in range(reader.FieldCount)]
    data = []
    c = 0
    while reader.Read() and c < 2000:
        row = {}
        for i in range(reader.FieldCount):
            val = reader.GetValue(i)
            row[cols[i]] = str(val) if val is not None else None
        data.append(row); c+=1
    reader.Close(); conn.Close()
    return data

# --- UNIFIED TOOLS ---

class PBIRManager:
    @staticmethod
    def detect_path() -> Optional[str]:
        """Auto-detects the .Report folder of the open Power BI Project"""
        import psutil
        try:
            # 1. Find Data Engine (msmdsrv)
            srv_pid = None
            for proc in psutil.process_iter(['pid', 'name']):
                if 'msmdsrv.exe' in (proc.info['name'] or '').lower():
                    # Check if it's the one we are connected to (optional, but good)
                    # For now, pick the first one owned by user
                    srv_pid = proc.info['pid']
                    break  # Assuming single instance for now
            
            if not srv_pid: return None

            # 2. Get Parent (PBIDesktop)
            try:
                parent = psutil.Process(srv_pid).parent()
                if not parent or 'PBIDesktop' not in parent.name(): return None
            except: return None

            # 3. Scan Command Line (If opened via double-click)
            cmdline = parent.cmdline()
            for arg in cmdline:
                if arg.endswith('.pbip'):
                    report_dir = arg.replace(".pbip", ".Report")
                    if os.path.exists(report_dir): return report_dir

            # 4. Scan Open Files (If opened via File > Open)
            # Note: This often requires Admin, might fail on some setups
            try:
                for f in parent.open_files():
                    if f.path.endswith('.pbip'):
                        report_dir = f.path.replace(".pbip", ".Report")
                        if os.path.exists(report_dir): return report_dir
            except: pass

        except Exception as e: 
            # print(f"Debug: {e}")
            pass
        return None

    @staticmethod
    def get_pages(report_path: str) -> List[Dict]:
        """Reads pages from the PBIR structure."""
        pages_dir = os.path.join(report_path, "definition", "pages")
        if not os.path.exists(pages_dir): return []
        
        # 1. Try to get order from pages.json
        pages_order = []
        pages_reg = os.path.join(pages_dir, "pages.json")
        if os.path.exists(pages_reg):
            try:
                with open(pages_reg, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    pages_order = data.get("pageOrder", [])
            except: pass
            
        # 2. Scan directories if order is empty or to supplement
        found_pages = []
        try:
            for entry in os.scandir(pages_dir):
                if entry.is_dir() and os.path.exists(os.path.join(entry.path, "page.json")):
                    found_pages.append(entry.name)
        except: pass
        
        # Merge (Ordered lines first, then others)
        final_list = []
        seen = set()
        
        for pid in pages_order:
            if pid in found_pages:
                final_list.append(pid)
                seen.add(pid)
        
        for pid in found_pages:
            if pid not in seen:
                final_list.append(pid)
                
        # 3. Read Metadata for each page
        res = []
        for pid in final_list:
            p_file = os.path.join(pages_dir, pid, "page.json")
            try:
                with open(p_file, 'r', encoding='utf-8') as f:
                    p_data = json.load(f)
                    pname = p_data.get("displayName", pid)
                    res.append({"id": pid, "name": pname})
            except:
                res.append({"id": pid, "name": pid})
                
        return res

@mcp.tool()
def pbir_inspect_structure() -> str:
    """Debugs the folder structure of the detected project."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    structure = []
    # Walk top 3 levels
    for root, dirs, files in os.walk(path):
        level = root.replace(path, '').count(os.sep)
        if level > 3: continue
        indent = '  ' * level
        structure.append(f"{indent}[DIR] {os.path.basename(root)}")
        for f in files:
            structure.append(f"{indent}  {f}")
    return "\n".join(structure[:50]) # Limit output

@mcp.tool()
def pbir_get_info() -> str:
    """Detects active PBIR project and lists pages. Returns detection status."""
    path = PBIRManager.detect_path()
    if not path:
        return json.dumps({"detected": False, "message": "Could not auto-detect .pbip path. Ensure project is open."}, indent=2)
    
    pages = PBIRManager.get_pages(path)
    return json.dumps({
        "detected": True, 
        "project_path": path,
        "pages": [p["id"] for p in pages],
        "total_pages": len(pages)
    }, indent=2)

@mcp.tool()
def manage_model_connection(operation: str = "get_current") -> str:
    """Manage connection (list/select/get_current)."""
    try:
        if operation == "get_current":
            s = _get_server()
            return json.dumps({"connected": True, "port": GLOBAL_CONTEXT["port"], "db": s.Databases[0].Name}, indent=2)
        elif operation == "list":
            import psutil
            return json.dumps([{"pid": p.info['pid'], "name": p.info['name']} for p in psutil.process_iter(['pid','name']) if 'msmdsrv' in (p.info['name'] or '').lower()], indent=2)
        return "Unknown op"
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def list_objects(object_type: str = "tables") -> str:
    """List objects: tables, measures, columns, relationships, roles, partitions, hierarchies."""
    try:
        m = _get_server().Databases[0].Model
        res = []
        if object_type == "tables":
            res = [{"Name": t.Name, "Description": t.Description or ""} for t in m.Tables]
        elif object_type == "measures":
            for t in m.Tables:
                for meas in t.Measures: res.append({"Name": meas.Name, "Table": t.Name, "Expression": meas.Expression})
        elif object_type == "relationships":
            for r in m.Relationships:
                res.append({"From": f"{r.FromTable.Name}[{r.FromColumn.Name}]", "To": f"{r.ToTable.Name}[{r.ToColumn.Name}]", "Active": r.IsActive})
        elif object_type == "roles":
            res = [{"Name": r.Name} for r in m.Roles]
        elif object_type == "partitions":
            for t in m.Tables:
                for p in t.Partitions: res.append({"Table": t.Name, "Partition": p.Name, "Mode": str(p.Mode), "SourceType": str(p.SourceType)})
        return json.dumps(res, indent=2)
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def search_model(query: str) -> str:
    """Deep search tables, columns, measures, expressions."""
    try:
        m = _get_server().Databases[0].Model
        res = []
        q = query.lower()
        for t in m.Tables:
            if q in t.Name.lower(): res.append({"Type": "Table", "Name": t.Name})
            for c in t.Columns:
                if q in c.Name.lower(): res.append({"Type": "Column", "Name": c.Name, "Table": t.Name})
            for meas in t.Measures:
                if q in meas.Name.lower() or (meas.Expression and q in meas.Expression.lower()):
                    res.append({"Type": "Measure", "Name": meas.Name, "Table": t.Name})
        return json.dumps(res[:50], indent=2)
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def pbir_create_page(name: str) -> str:
    """Create a new blank report page."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    import uuid
    page_guid = str(uuid.uuid4()).replace("-", "")[:20] # PBI uses short-ish IDs usually
    
    # 1. Update pages.json registry (pageOrder)
    pages_reg = os.path.join(path, "definition", "pages", "pages.json")
    try:
        with open(pages_reg, 'r+', encoding='utf-8') as f:
            data = json.load(f)
            # PBIR uses 'pageOrder' array
            if "pageOrder" not in data: data["pageOrder"] = []
            data["pageOrder"].append(page_guid)
            
            # Remove legacy 'pages' if exists (cleanup)
            if "pages" in data: del data["pages"]
            
            f.seek(0); json.dump(data, f, indent=2); f.truncate()
    except Exception as e: return f"Error updating registry: {e}"
    
    # 2. Create Page Folder and Metadata
    page_dir = os.path.join(path, "definition", "pages", page_guid)
    os.makedirs(page_dir, exist_ok=True)
    
    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json",
        "name": page_guid,
        "displayName": name,
        "width": 1280,
        "height": 720,
        "displayOption": "FitToPage"
    }
    
    with open(os.path.join(page_dir, "page.json"), 'w', encoding='utf-8') as f:
        json.dump(page_json, f, indent=2)
        
    return f"Page '{name}' created ({page_guid})"

@mcp.tool()
def pbir_create_visual(page_name: str, visual_type: str, title: str = "New Visual") -> str:
    """Create a visual on a page. Types: 'card', 'textbox'."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # Find Page GUID
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return "Page not found"
    
    import uuid
    vis_guid = str(uuid.uuid4()).replace("-", "")[:20]
    vis_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals", vis_guid)
    os.makedirs(vis_dir, exist_ok=True)
    
    visual_json = {}
    
    if visual_type == "textbox":
        visual_json = {
             "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
             "name": vis_guid,
             "position": {"x": 100, "y": 100, "width": 300, "height": 100},
             "visual": {
                "visualType": "shape",
                "objects": {
                    "general": [{"properties": {"title": {"expr": {"Literal": {"Value": f"'{title}'"}}}}} ]
                },
                "drillFilterOtherVisuals": True
             }
        }
    elif visual_type == "card":
        # Basic Card Shell
        visual_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
            "name": vis_guid,
            "position": {"x": 50, "y": 50, "z": 0, "width": 300, "height": 300, "tabOrder": 1000},
            "visual": {
                "visualType": "card",
                "objects": {
                    "general": [{"properties": {"title": {"expr": {"Literal": {"Value": f"'{title}'"}}}}}]
                },
                "drillFilterOtherVisuals": True
            }
        }

    with open(os.path.join(vis_dir, "visual.json"), 'w', encoding='utf-8') as f:
        json.dump(visual_json, f, indent=2)
        
    return f"Visual {visual_type} created on {page_name}"

@mcp.tool()
def pbir_bind_measure(page_name: str, visual_title: str, measure_table: str, measure_name: str) -> str:
    """Binds a measure to a visual (Card) identified by its Title."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # 1. Find Page
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    # 2. Find Visual by Title
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    if not os.path.exists(visuals_dir): return "No visuals found."
    
    vis_path = None
    vis_data = None
    
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        # Check Title
                        try:
                            # Path: visual -> objects -> general -> title -> expr
                            title_expr = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"]
                            if title_expr.replace("'", "") == visual_title:
                                vis_path = v_file
                                vis_data = data
                                break
                        except: pass
                except: pass
    
    if not vis_path: return f"Visual '{visual_title}' not found on '{page_name}'."
    
    # 3. Update Query
    query_structure = {
      "queryState": {
        "Values": {
          "projections": [
            {
              "field": {
                "Measure": {
                  "Expression": {
                    "SourceRef": {"Entity": measure_table}
                  },
                  "Property": measure_name
                }
              },
              "queryRef": f"{measure_table}.{measure_name}",
              "nativeQueryRef": measure_name,
              "displayName": measure_name
            }
          ]
        }
      }
    }
    
    if "visual" not in vis_data: vis_data["visual"] = {}
    vis_data["visual"]["query"] = query_structure
    
    with open(vis_path, 'w', encoding='utf-8') as f:
        json.dump(vis_data, f, indent=2)
        
    return f"Bound measure '{measure_table}[{measure_name}]' to visual '{visual_title}'."

@mcp.tool()
def pbir_create_bar_chart(page_name: str, visual_title: str, category_table: str, category_col: str, value_table: str, value_measure: str) -> str:
    """Create a Clustered Bar Chart (Ranking)."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # Find Page
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    import uuid
    vis_guid = str(uuid.uuid4()).replace("-", "")[:20]
    vis_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals", vis_guid)
    
    # Check if we should use existing (if create overwrites... standard create_visual creates new GUID)
    os.makedirs(vis_dir, exist_ok=True)

    # 2. Build JSON
    visual_json = {
      "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
      "name": vis_guid,
      "position": {
        "x": 50,
        "y": 200,
        "z": 0,
        "width": 400,
        "height": 300,
        "tabOrder": 2000
      },
      "visual": {
        "visualType": "clusteredBarChart",
        "query": {
          "queryState": {
            "Category": {
              "projections": [
                {
                  "field": {
                    "Column": {
                      "Expression": {
                        "SourceRef": {
                          "Entity": category_table
                        }
                      },
                      "Property": category_col
                    }
                  },
                  "queryRef": f"{category_table}.{category_col}",
                  "nativeQueryRef": category_col,
                  "displayName": category_col
                }
              ]
            },
            "Y": {
              "projections": [
                {
                  "field": {
                    "Measure": {
                      "Expression": {
                        "SourceRef": {
                          "Entity": value_table
                        }
                      },
                      "Property": value_measure
                    }
                  },
                  "queryRef": f"{value_table}.{value_measure}",
                  "nativeQueryRef": value_measure,
                  "displayName": value_measure
                }
              ]
            }
          },
          "sortDefinition": {
             "sort": [
                 {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": value_table}}, "Property": value_measure}}, "direction": "Descending"}
             ],
             "isDefaultSort": True
          }
        },
        "objects": {
          "general": [
            {
              "properties": {
                "title": {
                  "expr": {
                    "Literal": {
                      "Value": f"'{visual_title}'"
                    }
                  }
                }
              }
            }
          ]
        },
        "drillFilterOtherVisuals": True
      }
    }
    
    with open(os.path.join(vis_dir, "visual.json"), 'w', encoding='utf-8') as f:
        json.dump(visual_json, f, indent=2)
        
    return f"Created Bar Chart '{visual_title}' on '{page_name}'"

    with open(os.path.join(vis_dir, "visual.json"), 'w', encoding='utf-8') as f:
        json.dump(visual_json, f, indent=2)
        
    return f"Created Bar Chart '{visual_title}' on '{page_name}'"

@mcp.tool()
def pbir_format_visual(page_name: str, visual_title: str, new_title: str = None, rename_fields: str = None) -> str:
    """
    Format a visual: set title and rename fields (axis labels).
    rename_fields: JSON string mapping 'OriginalName' to 'NewName' (e.g. '{"Sum(Sales)": "Total Sales"}')
    """
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # 1. Find Page
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    # 2. Find Visual
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    if not os.path.exists(visuals_dir): return "No visuals found."
    
    vis_data = None
    vis_path = None
    
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        curr_title = ""
                        try:
                            curr_title = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                        except: pass
                        
                        if curr_title == visual_title:
                            vis_data = data
                            vis_path = v_file
                            break
                except: pass
    
    if not vis_data: return f"Visual '{visual_title}' not found."
    
    # 3. Apply Formats
    
    # Title
    if new_title:
        # Set Title Text
        if "visual" not in vis_data: vis_data["visual"] = {}
        if "objects" not in vis_data["visual"]: vis_data["visual"]["objects"] = {}
        if "general" not in vis_data["visual"]["objects"]: vis_data["visual"]["objects"]["general"] = [{"properties": {}}]
        
        vis_data["visual"]["objects"]["general"][0]["properties"]["title"] = {
            "expr": {"Literal": {"Value": f"'{new_title}'"}}
        }
        
        # Ensure Visible
        if "visible" not in vis_data["visual"]["objects"]: vis_data["visual"]["objects"]["visible"] = [{"properties": {}}]
        vis_data["visual"]["objects"]["visible"][0]["properties"]["title"] = {
             "expr": {"Literal": {"Value": "true"}}
        }
        
    # Field Renaming
    if rename_fields:
        try:
            mapping = json.loads(rename_fields)
            
            def recurse_rename(obj):
                if isinstance(obj, dict):
                    # Check if this is a Projection/Field definition
                    if "displayName" in obj and "nativeQueryRef" in obj:
                        if obj["nativeQueryRef"] in mapping:
                            obj["displayName"] = mapping[obj["nativeQueryRef"]]
                        elif obj["displayName"] in mapping: # Fallback
                             obj["displayName"] = mapping[obj["displayName"]]
                             
                    for k, v in obj.items(): recurse_rename(v)
                elif isinstance(obj, list):
                    for item in obj: recurse_rename(item)
            
            if "query" in vis_data.get("visual", {}):
                recurse_rename(vis_data["visual"]["query"])
                
        except Exception as e: return f"Error parsing mapping: {e}"
        
    with open(vis_path, 'w', encoding='utf-8') as f:
        json.dump(vis_data, f, indent=2)
        
    with open(vis_path, 'w', encoding='utf-8') as f:
        json.dump(vis_data, f, indent=2)
        
    return f"Formatted '{visual_title}'."

@mcp.tool()
def pbir_refactor_field(table_name: str, old_name: str, new_name: str) -> str:
    """
    Refactor (Rename) a field use in ALL visuals.
    Replaces references to 'table_name[old_name]' with 'table_name[new_name]' in JSONs.
    """
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    count = 0
    
    for page in pages:
        visuals_dir = os.path.join(path, "definition", "pages", page["id"], "visuals")
        if not os.path.exists(visuals_dir): continue
        
        for entry in os.scandir(visuals_dir):
            if entry.is_dir():
                v_file = os.path.join(entry.path, "visual.json")
                if os.path.exists(v_file):
                    try:
                        changed = False
                        with open(v_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        
                        # Recursive Search & Replace
                        def recurse_replace(obj):
                            nonlocal changed
                            if isinstance(obj, dict):
                                # Check for Property Ref pattern
                                # "Property": "Name", "Expression": {"SourceRef": {"Entity": "Table"}}
                                if "Property" in obj and obj["Property"] == old_name:
                                    # Verify Entity if possible (it's usually nested in Expression)
                                    # But structure varies (Measure vs Column). 
                                    # Usually: obj = {"Property": "Col", "Expression": {...}}
                                    expr = obj.get("Expression", {})
                                    source = expr.get("SourceRef", {})
                                    entity = source.get("Entity")
                                    
                                    if entity == table_name:
                                        obj["Property"] = new_name
                                        changed = True
                                        
                                # Also check simple references logic if needed, but QueryState is most important
                                for k, v in obj.items(): recurse_replace(v)
                            elif isinstance(obj, list):
                                for item in obj: recurse_replace(item)
                                
                        if "visual" in data:
                            recurse_replace(data["visual"])
                            
                        if changed:
                            with open(v_file, 'w', encoding='utf-8') as f:
                                json.dump(data, f, indent=2)
                            count += 1
                    except: pass
    
    return f"Refactored '{old_name}' to '{new_name}' in {count} visuals."

@mcp.tool()
def pbir_audit_usage(object_name: str) -> str:
    """Find which visuals use a specific measure or column (by name)."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    usage = []
    
    for page in pages:
        visuals_dir = os.path.join(path, "definition", "pages", page["id"], "visuals")
        if not os.path.exists(visuals_dir): continue
        
        for entry in os.scandir(visuals_dir):
            if entry.is_dir():
                v_file = os.path.join(entry.path, "visual.json")
                if os.path.exists(v_file):
                    try:
                        with open(v_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                        vis_title = "Untitled"
                        try:
                            vis_title = data["visual"]["objects"]["general"][0]["properties"]["title"]["expr"]["Literal"]["Value"]
                        except: pass
                        
                        # Search string representation of JSON is lazy but effective for audit
                        # precise search is better though
                        found = False
                        
                        def recurse_find(obj):
                            nonlocal found
                            if found: return
                            if isinstance(obj, dict):
                                if "Property" in obj and obj["Property"] == object_name:
                                    found = True
                                    return
                                for k, v in obj.items(): recurse_find(v)
                            elif isinstance(obj, list):
                                for item in obj: recurse_find(item)
                                
                        if "visual" in data:
                            recurse_find(data["visual"])
                            
                        if found:
                            usage.append(f"Page: {page['name']} | Visual: {vis_title}")
                    except: pass
                    
    if not usage: return f"Object '{object_name}' not found in report visuals."
    return json.dumps(usage, indent=2)
    """List all visuals on a page with their Title, Type, and Layout stats."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    if not os.path.exists(visuals_dir): return "No visuals found."
    
    res = []
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        vis_obj = data.get("visual", {})
                        
                        # Get Title
                        title = "Untitled"
                        try:
                            title = vis_obj.get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                        except: pass
                        
                        # Get Type
                        v_type = vis_obj.get("visualType", "unknown")
                        
                        # Get Pos
                        pos = data.get("position", {})
                        
                        res.append({
                            "id": entry.name,
                            "title": title,
                            "type": v_type,
                            "x": int(pos.get("x",0)),
                            "y": int(pos.get("y",0)),
                            "w": int(pos.get("width",0)),
                            "h": int(pos.get("height",0))
                        })
                except: pass
    
    return json.dumps(res, indent=2)

@mcp.tool()
def pbir_update_visual_layout(page_name: str, visual_title: str, x: int = None, y: int = None, width: int = None, height: int = None, z: int = None) -> str:
    """Update position and size of a visual identified by Title."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # 1. Find Page
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    # 2. Find Visual and Update
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    found = False
    
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r+', encoding='utf-8') as f:
                        data = json.load(f)
                        
                        # Check Title
                        curr_title = ""
                        try:
                            curr_title = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                        except: pass
                        
                        if curr_title == visual_title:
                            # Update
                            pos = data.get("position", {})
                            if x is not None: pos["x"] = x
                            if y is not None: pos["y"] = y
                            if width is not None: pos["width"] = width
                            if height is not None: pos["height"] = height
                            if z is not None: pos["z"] = z
                            
                            data["position"] = pos
                            
                            f.seek(0); json.dump(data, f, indent=2); f.truncate()
                            found = True
                            break
                except: pass
        if found: break
        
    if found: return f"Updated layout for '{visual_title}'."
    return f"Visual '{visual_title}' not found."

@mcp.tool()
def pbir_delete_object(page_name: str, visual_title: str = None, visual_id: str = None) -> str:
    """Delete a Page (if visual_title/id is None) or a Visual on that page (by Title or ID)."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    # 1. Find Page
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    # DELETE PAGE
    if not visual_title and not visual_id:
        # Remove from pages.json
        pages_reg = os.path.join(path, "definition", "pages", "pages.json")
        try:
            with open(pages_reg, 'r+', encoding='utf-8') as f:
                data = json.load(f)
                if tgt_page["id"] in data.get("pageOrder", []):
                    data["pageOrder"].remove(tgt_page["id"])
                    f.seek(0); json.dump(data, f, indent=2); f.truncate()
        except: pass
        
        # Remove Folder
        import shutil
        page_dir = os.path.join(path, "definition", "pages", tgt_page["id"])
        try: shutil.rmtree(page_dir)
        except Exception as e: return f"Error deleting folder: {e}"
        
        return f"Page '{page_name}' deleted."

    # DELETE VISUAL
    else:
        visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
        if not os.path.exists(visuals_dir): return "No visuals found."
        
        tgt_vis_id = None
        
        if visual_id:
            # Direct Delete by ID
            tgt_vis_id = visual_id
        else:
            # Find by Title
            for entry in os.scandir(visuals_dir):
                if entry.is_dir():
                    v_file = os.path.join(entry.path, "visual.json")
                    if os.path.exists(v_file):
                        try:
                            with open(v_file, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                curr_title = ""
                                try:
                                    curr_title = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                                except: pass
                                
                                if curr_title == visual_title:
                                    tgt_vis_id = entry.name
                                    break
                        except: pass
        
        if tgt_vis_id and os.path.exists(os.path.join(visuals_dir, tgt_vis_id)):
            import shutil
            shutil.rmtree(os.path.join(visuals_dir, tgt_vis_id))
            return f"Visual '{visual_title or visual_id}' deleted."
            
        return f"Visual not found."

# --- CORE OPS ---

@mcp.tool()
def run_dax(query: str) -> str:
    """Execute DAX query (limit 2000 rows)."""
    try: return json.dumps(_run_dax_raw(query), indent=2)
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def get_model_info() -> str:
    """Get basic model metadata (Name, Compatibility Level, Timestamps)."""
    try: return json.dumps(_run_dax_raw("SELECT [Name], [CompatibilityLevel], [CreatedTimestamp], [ModifiedTimestamp] FROM $SYSTEM.TMSCHEMA_MODEL"), indent=2)
    except Exception as e: return f"Error: {e}"

# --- MANAGEMENT ---

@mcp.tool()
def manage_measure(operation: str, table_name: str, measure_name: str, expression: str = None, description: str = None) -> str:
    """Create, Update, or Delete measures. Operation: create/update/delete."""
    try:
        t = _get_server().Databases[0].Model.Tables.Find(table_name)
        if not t: return "Table not found"
        m = t.Measures.Find(measure_name)
        
        if operation == "create":
            if m: return "Exists"
            from Microsoft.AnalysisServices.Tabular import Measure
            new_m = Measure(); new_m.Name = measure_name; new_m.Expression = expression
            if description: new_m.Description = description
            t.Measures.Add(new_m)
        elif operation == "update":
            if not m: return "Not found"
            if expression: m.Expression = expression
            if description: m.Description = description
        elif operation == "delete":
            if not m: return "Not found"
            t.Measures.Remove(m)
        else: return "Invalid Op"
        
        t.Model.SaveChanges()
        return "Success"
    except Exception as e: return f"Error: {e}"

    except Exception as e: return f"Error: {e}"

@mcp.tool()
def manage_column(operation: str, table_name: str, column_name: str, new_name: str = None, new_description: str = None, is_hidden: bool = None, data_type: str = None) -> str:
    """Manage Table Columns. Ops: update (rename, hide, type), delete."""
    try:
        t = _get_server().Databases[0].Model.Tables.Find(table_name)
        if not t: return "Table not found"
        
        # Find Column (try Name or SourceColumn)
        c = t.Columns.Find(column_name)
        if not c:
            # Try finding by source column if it's a renaming op
            for col in t.Columns:
                if col.SourceColumn == column_name: c = col; break
        
        if not c: return "Column not found"
        
        if operation == "update":
            if new_name: c.Name = new_name
            if new_description is not None: c.Description = new_description
            if is_hidden is not None: c.IsHidden = is_hidden
            if data_type:
                from Microsoft.AnalysisServices.Tabular import DataType
                # Map string to Enum
                dt_map = {
                    "string": DataType.String, "int": DataType.Int64, 
                    "double": DataType.Double, "datetime": DataType.DateTime, 
                    "boolean": DataType.Boolean, "decimal": DataType.Decimal
                }
                if data_type.lower() in dt_map: c.DataType = dt_map[data_type.lower()]
                
        elif operation == "delete":
            t.Columns.Remove(c)
            
        t.Model.SaveChanges()
        return "Success"
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def manage_table(operation: str, table_name: str, source_expression: str = None, type: str = "Global") -> str:
    """
    Create/Delete Tables. 
    Types: 'Global' (M Table), 'Calculated' (DAX Table).
    source_expression: M code or DAX expression.
    """
    try:
        model = _get_server().Databases[0].Model
        table = model.Tables.Find(table_name)

        if operation == "create":
            if table: return "Table exists"
            from Microsoft.AnalysisServices.Tabular import Table, Partition, ModeType, CalculatedPartitionSource, MPartitionSource
            
            new_t = Table(); new_t.Name = table_name
            part = Partition(); part.Name = "Partition"
            
            if type == "Calculated":
                # Calculated Table (DAX) uses CalculatedPartitionSource
                # SourceType is read-only in modern TOM; derived from Source
                source = CalculatedPartitionSource()
                source.Expression = source_expression
                part.Source = source
            else:
                # M Table (Import) uses MPartitionSource
                part.Mode = ModeType.Import
                source = MPartitionSource()
                source.Expression = source_expression
                part.Source = source
                
            new_t.Partitions.Add(part)
            model.Tables.Add(new_t)
            model.SaveChanges()
            return f"Table {table_name} created."

        elif operation == "delete":
            if not table: return "Table not found"
            model.Tables.Remove(table)
            model.SaveChanges()
            return "Deleted"
            
        return "Op not supported"
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def manage_relationship(operation: str, from_table: str, from_col: str, to_table: str, to_col: str, active: bool = True) -> str:
    """Manage relationships between tables. Operation: create/delete."""
    try:
        m = _get_server().Databases[0].Model
        if operation == "create":
            ft = m.Tables.Find(from_table); tt = m.Tables.Find(to_table)
            if not ft or not tt: return "Table not found"
            from Microsoft.AnalysisServices.Tabular import SingleColumnRelationship
            r = SingleColumnRelationship()
            r.FromColumn = ft.Columns.Find(from_col)
            r.ToColumn = tt.Columns.Find(to_col)
            r.IsActive = active
            m.Relationships.Add(r); m.SaveChanges()
            return "Created"
        elif operation == "delete":
            tgt = None
            for r in m.Relationships:
                if r.FromTable.Name == from_table and r.FromColumn.Name == from_col and r.ToTable.Name == to_table: tgt = r; break
            if tgt: m.Relationships.Remove(tgt); m.SaveChanges(); return "Deleted"
            return "Not found"
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def manage_role(operation: str, role_name: str, table_filters: List[Dict[str, str]] = None) -> str:
    """Manage RLS Roles. table_filters is [{"table": "Sales", "expression": "Sales[ID]=1"}]."""
    try:
        m = _get_server().Databases[0].Model
        r = m.Roles.Find(role_name)
        
        if operation == "create":
            if r: return "Role exists"
            from Microsoft.AnalysisServices.Tabular import ModelRole
            new_r = ModelRole(); new_r.Name = role_name
            m.Roles.Add(new_r)
            
            # Add filters
            if table_filters:
                for f in table_filters:
                    t = m.Tables.Find(f["table"])
                    if t: 
                        new_r.RowLevelSecurityPermissions.Add(
                            t, f.get("expression", "")
                        ) # Logic simplified, permissions usually handled via property setting on permission object
            m.SaveChanges()
            return "Role created"
        elif operation == "delete":
            if r: m.Roles.Remove(r); m.SaveChanges(); return "Deleted"
        return "Unknown"
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def manage_calc_group(operation: str, table_name: str, items: List[Dict[str, str]] = None) -> str:
    """Create Calculation Groups."""
    try:
        m = _get_server().Databases[0].Model
        if operation == "create":
            from Microsoft.AnalysisServices.Tabular import CalculationGroup, Table, Partition, PartitionSourceType, ModeType
            
            # Create Table first
            t = Table(); t.Name = table_name
            part = Partition(); part.Name = "Partition"; part.SourceType = PartitionSourceType.CalculationGroup
            part.Mode = ModeType.Import 
            t.Partitions.Add(part)
            
            # Add Calc Group
            cg = CalculationGroup()
            t.CalculationGroup = cg
            
            # Add Items
            if items:
                from Microsoft.AnalysisServices.Tabular import CalculationItem
                for i in items:
                    ci = CalculationItem(); ci.Name = i["name"]; ci.Expression = i["expression"]
                    cg.CalculationItems.Add(ci)
            
            m.Tables.Add(t)
            m.SaveChanges()
            return "Calc Group Created"
    except Exception as e: return f"Error: {e}"

# --- ANALYSIS ---

@mcp.tool()
def get_vertipaq_stats() -> str:
    """Analyze memory usage (Top 20 columns)."""
    try:
        # Robust discovery logic
        check = _run_dax_raw("SELECT TOP 1 * FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS")
        if not check: return "No stats"
        cols = check[0].keys()
        
        t_col = 'DIMENSION_NAME' if 'DIMENSION_NAME' in cols else 'TABLE_ID'
        c_col = 'ATTRIBUTE_NAME' if 'ATTRIBUTE_NAME' in cols else 'COLUMN_ID'
        d_col = 'DICTIONARY_SIZE' if 'DICTIONARY_SIZE' in cols else None
        u_col = 'USED_SIZE' if 'USED_SIZE' in cols else ('DATA_SIZE' if 'DATA_SIZE' in cols else None)
        
        if not d_col and not u_col: return "Columns not found"
        
        sel = f"[{t_col}] as T, [{c_col}] as C"
        ord_parts = []
        if d_col: sel += f", [{d_col}] as Dict"; ord_parts.append(f"[{d_col}]")
        if u_col: sel += f", [{u_col}] as Data"; ord_parts.append(f"[{u_col}]")
        
        q = f"SELECT {sel} FROM $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS WHERE [COLUMN_TYPE]='BASIC_DATA' ORDER BY {'+'.join(ord_parts)} DESC"
        data = _run_dax_raw(q)
        
        res = []
        for r in data[:20]:
            d = int(r.get('Dict') or 0); u = int(r.get('Data') or 0)
            res.append({"Table": r.get('T'), "Column": r.get('C'), "SizeKB": round((d+u)/1024, 2)})
        return json.dumps(res, indent=2)
    except Exception as e: return f"Error: {e}"

@mcp.tool()
def visualize_python(dax_query: str, python_code: str, title: str = "Chart") -> str:
    """Generate Chart via Subprocess."""
    try:
        raw = _run_dax_raw(dax_query)
        if not raw: return "No data"
        
        img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts")
        os.makedirs(img_dir, exist_ok=True)
        img_path = os.path.join(img_dir, f"chart_{uuid.uuid4().hex}.png")
        data_path = os.path.join(tempfile.gettempdir(), f"data_{uuid.uuid4().hex}.json")
        script_path = os.path.join(tempfile.gettempdir(), f"script_{uuid.uuid4().hex}.py")
        
        with open(data_path, 'w') as f: json.dump(raw, f)
        
        script = f"""
import sys, json, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

try:
    with open(r"{data_path}", 'r') as f: df = pd.DataFrame(json.load(f))
    for c in df.columns: 
        try: df[c] = pd.to_numeric(df[c])
        except: pass
    
    plt.figure(figsize=(10,6)); sns.set_theme(style="whitegrid")
    {python_code}
    plt.title("{title}"); plt.tight_layout(); plt.savefig(r"{img_path}"); plt.close()
except: pass
"""
        with open(script_path, 'w') as f: f.write(script)
        subprocess.run([sys.executable, script_path], capture_output=True, timeout=90)
        try: os.remove(data_path); os.remove(script_path)
        except: pass
        
        return json.dumps({"success": True, "path": img_path})
    except Exception as e: return f"Error: {e}"

if __name__ == "__main__":
    mcp.run(transport="stdio")
