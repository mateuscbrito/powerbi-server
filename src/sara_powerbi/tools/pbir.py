import os
import json
import uuid
import psutil
from typing import List, Dict, Optional

class PBIRManager:
    @staticmethod
    def detect_path() -> Optional[str]:
        """Auto-detects the .Report folder of the open Power BI Project"""
        try:
            # 1. Find Data Engine (msmdsrv)
            srv_pid = None
            for proc in psutil.process_iter(['pid', 'name']):
                if 'msmdsrv.exe' in (proc.info['name'] or '').lower():
                    srv_pid = proc.info['pid']
                    break 
            
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

            # 4. Scan Open Files 
            try:
                for f in parent.open_files():
                    if f.path.endswith('.pbip'):
                        report_dir = f.path.replace(".pbip", ".Report")
                        if os.path.exists(report_dir): return report_dir
            except: pass

        except Exception: 
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
            
        # 2. Scan directories
        found_pages = []
        try:
            for entry in os.scandir(pages_dir):
                if entry.is_dir() and os.path.exists(os.path.join(entry.path, "page.json")):
                    found_pages.append(entry.name)
        except: pass
        
        # Merge
        final_list = []
        seen = set()
        
        for pid in pages_order:
            if pid in found_pages:
                final_list.append(pid)
                seen.add(pid)
        
        for pid in found_pages:
            if pid not in seen:
                final_list.append(pid)
                
        # 3. Read Metadata
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

def pbir_inspect_structure() -> str:
    """Debugs the folder structure of the detected project."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    structure = []
    for root, dirs, files in os.walk(path):
        level = root.replace(path, '').count(os.sep)
        if level > 3: continue
        indent = '  ' * level
        structure.append(f"{indent}[DIR] {os.path.basename(root)}")
        for f in files:
            structure.append(f"{indent}  {f}")
    return "\n".join(structure[:50])

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

def pbir_create_page(name: str) -> str:
    """Create a new blank report page."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    page_guid = str(uuid.uuid4()).replace("-", "")[:20]
    
    # 1. Update pages.json (pageOrder)
    pages_reg = os.path.join(path, "definition", "pages", "pages.json")
    try:
        with open(pages_reg, 'r+', encoding='utf-8') as f:
            data = json.load(f)
            if "pageOrder" not in data: data["pageOrder"] = []
            data["pageOrder"].append(page_guid)
            if "pages" in data: del data["pages"]
            f.seek(0); json.dump(data, f, indent=2); f.truncate()
    except Exception as e: return f"Error updating registry: {e}"
    
    # 2. Create Page Folder
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

def pbir_create_visual(page_name: str, visual_type: str, title: str = "New Visual") -> str:
    """Create a visual on a page. Types: 'card', 'textbox'."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return "Page not found"
    
    vis_guid = str(uuid.uuid4()).replace("-", "")[:20]
    vis_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals", vis_guid)
    os.makedirs(vis_dir, exist_ok=True)
    
    visual_json = {}
    
    # Helper for Title Expression
    def create_title_expr(text):
        return {
            "title": {
                "expr": {
                    "Literal": {
                        "Value": f"'{text}'"
                    }
                }
            }
        }

    if visual_type == "textbox":
        visual_json = {
             "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
             "name": vis_guid,
             "position": {"x": 100, "y": 100, "width": 300, "height": 100},
             "visual": {
                "visualType": "shape",
                "objects": {
                    "general": [{"properties": create_title_expr(title)}]
                },
                "drillFilterOtherVisuals": True
             }
        }
    elif visual_type == "card":
        visual_json = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
            "name": vis_guid,
            "position": {"x": 50, "y": 50, "z": 0, "width": 300, "height": 300, "tabOrder": 1000},
            "visual": {
                "visualType": "card",
                "objects": {
                    "general": [{"properties": create_title_expr(title)}]
                },
                "drillFilterOtherVisuals": True
            }
        }

    with open(os.path.join(vis_dir, "visual.json"), 'w', encoding='utf-8') as f:
        json.dump(visual_json, f, indent=2)
        
    return f"Visual {visual_type} created on {page_name}"

def pbir_create_bar_chart(page_name: str, visual_title: str, category_table: str, category_col: str, value_table: str, value_measure: str) -> str:
    """Create a Clustered Bar Chart."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    vis_guid = str(uuid.uuid4()).replace("-", "")[:20]
    vis_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals", vis_guid)
    os.makedirs(vis_dir, exist_ok=True)

    # Title Property
    title_prop = {
        "title": {
            "expr": {
                "Literal": {
                    "Value": f"'{visual_title}'"
                }
            }
        }
    }

    visual_json = {
      "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json",
      "name": vis_guid,
      "position": {"x": 50, "y": 200, "z": 0, "width": 400, "height": 300, "tabOrder": 2000},
      "visual": {
        "visualType": "clusteredBarChart",
        "query": {
          "queryState": {
            "Category": {
              "projections": [{
                  "field": {"Column": {"Expression": {"SourceRef": {"Entity": category_table}}, "Property": category_col}},
                  "queryRef": f"{category_table}.{category_col}",
                  "nativeQueryRef": category_col,
                  "displayName": category_col
                }]
            },
            "Y": {
              "projections": [{
                  "field": {"Measure": {"Expression": {"SourceRef": {"Entity": value_table}}, "Property": value_measure}},
                  "queryRef": f"{value_table}.{value_measure}",
                  "nativeQueryRef": value_measure,
                  "displayName": value_measure
                }]
            }
          },
          "sortDefinition": {
             "sort": [{"field": {"Measure": {"Expression": {"SourceRef": {"Entity": value_table}}, "Property": value_measure}}, "direction": "Descending"}],
             "isDefaultSort": True
          }
        },
        "objects": {
          "general": [{"properties": title_prop}]
        },
        "drillFilterOtherVisuals": True
      }
    }
    
    with open(os.path.join(vis_dir, "visual.json"), 'w', encoding='utf-8') as f:
        json.dump(visual_json, f, indent=2)
        
    return f"Created Bar Chart '{visual_title}' on '{page_name}'"

def pbir_bind_measure(page_name: str, visual_title: str, measure_table: str, measure_name: str) -> str:
    """Binds a measure to a visual (Card) identified by its Title."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    if not os.path.exists(visuals_dir): return "No visuals found."
    
    vis_path = None; vis_data = None
    
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        try:
                            title_expr = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"]
                            if title_expr.replace("'", "") == visual_title:
                                vis_path = v_file; vis_data = data; break
                        except: pass
                except: pass
    
    if not vis_path: return f"Visual '{visual_title}' not found on '{page_name}'."
    
    query_structure = {
      "queryState": {
        "Values": {
          "projections": [{
              "field": {"Measure": {"Expression": {"SourceRef": {"Entity": measure_table}}, "Property": measure_name}},
              "queryRef": f"{measure_table}.{measure_name}",
              "nativeQueryRef": measure_name,
              "displayName": measure_name
            }]
        }
      }
    }
    
    if "visual" not in vis_data: vis_data["visual"] = {}
    vis_data["visual"]["query"] = query_structure
    
    with open(vis_path, 'w', encoding='utf-8') as f:
        json.dump(vis_data, f, indent=2)
        
    return f"Bound measure '{measure_table}[{measure_name}]' to visual '{visual_title}'."

def pbir_format_visual(page_name: str, visual_title: str, new_title: str = None, rename_fields: str = None) -> str:
    """Format a visual: set title and rename fields (axis labels)."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    if not os.path.exists(visuals_dir): return "No visuals found."
    
    vis_data = None; vis_path = None
    
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
                            vis_data = data; vis_path = v_file; break
                except: pass
    
    if not vis_data: return f"Visual '{visual_title}' not found."
    
    if new_title:
        if "visual" not in vis_data: vis_data["visual"] = {}
        if "objects" not in vis_data["visual"]: vis_data["visual"]["objects"] = {}
        if "general" not in vis_data["visual"]["objects"]: vis_data["visual"]["objects"]["general"] = [{"properties": {}}]
        vis_data["visual"]["objects"]["general"][0]["properties"]["title"] = {"expr": {"Literal": {"Value": f"'{new_title}'"}}}
        
    if rename_fields:
        try:
            mapping = json.loads(rename_fields)
            def recurse_rename(obj):
                if isinstance(obj, dict):
                    if "displayName" in obj and "nativeQueryRef" in obj:
                        if obj["nativeQueryRef"] in mapping: obj["displayName"] = mapping[obj["nativeQueryRef"]]
                        elif obj["displayName"] in mapping: obj["displayName"] = mapping[obj["displayName"]]
                    for k, v in obj.items(): recurse_rename(v)
                elif isinstance(obj, list):
                    for item in obj: recurse_rename(item)
            if "query" in vis_data.get("visual", {}): recurse_rename(vis_data["visual"]["query"])
        except Exception as e: return f"Error parsing mapping: {e}"
        
    with open(vis_path, 'w', encoding='utf-8') as f:
        json.dump(vis_data, f, indent=2)
    return f"Formatted '{visual_title}'."

def pbir_refactor_field(table_name: str, old_name: str, new_name: str) -> str:
    """Refactor (Rename) a field use in ALL visuals."""
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
                        with open(v_file, 'r', encoding='utf-8') as f: data = json.load(f)
                        
                        def recurse_replace(obj):
                            nonlocal changed
                            if isinstance(obj, dict):
                                if "Property" in obj and obj["Property"] == old_name:
                                    expr = obj.get("Expression", {})
                                    source = expr.get("SourceRef", {})
                                    if source.get("Entity") == table_name:
                                        obj["Property"] = new_name; changed = True
                                for k, v in obj.items(): recurse_replace(v)
                            elif isinstance(obj, list):
                                for item in obj: recurse_replace(item)
                                
                        if "visual" in data: recurse_replace(data["visual"])
                        if changed:
                            with open(v_file, 'w', encoding='utf-8') as f: json.dump(data, f, indent=2)
                            count += 1
                    except: pass
    return f"Refactored '{old_name}' to '{new_name}' in {count} visuals."

def pbir_audit_usage(object_name: str) -> str:
    """Find which visuals use a specific measure or column."""
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
                        with open(v_file, 'r', encoding='utf-8') as f: data = json.load(f)
                        found = False
                        def recurse_find(obj):
                            nonlocal found
                            if found: return
                            if isinstance(obj, dict):
                                if "Property" in obj and obj["Property"] == object_name: found = True; return
                                for k, v in obj.items(): recurse_find(v)
                            elif isinstance(obj, list):
                                for item in obj: recurse_find(item)
                        if "visual" in data: recurse_find(data["visual"])
                        if found:
                            title = "Untitled"
                            try: title = data["visual"]["objects"]["general"][0]["properties"]["title"]["expr"]["Literal"]["Value"]
                            except: pass
                            usage.append(f"Page: {page['name']} | Visual: {title}")
                    except: pass
    if not usage: return f"Object '{object_name}' not found."
    return json.dumps(usage, indent=2)

def pbir_list_visuals(page_name: str) -> str:
    """List all visuals on a page."""
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
                        title = "Untitled"
                        try: title = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                        except: pass
                        pos = data.get("position", {})
                        res.append({"id": entry.name, "title": title, "type": data.get("visual", {}).get("visualType"), "x": int(pos.get("x",0)), "y": int(pos.get("y",0))})
                except: pass
    return json.dumps(res, indent=2)

def pbir_delete_object(page_name: str, visual_title: str = None, visual_id: str = None) -> str:
    """Delete a Page or a Visual on that page."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return f"Page '{page_name}' not found."
    
    # DELETE PAGE
    if not visual_title and not visual_id:
        pages_reg = os.path.join(path, "definition", "pages", "pages.json")
        try:
            with open(pages_reg, 'r+', encoding='utf-8') as f:
                data = json.load(f)
                if tgt_page["id"] in data.get("pageOrder", []):
                    data["pageOrder"].remove(tgt_page["id"])
                    f.seek(0); json.dump(data, f, indent=2); f.truncate()
        except: pass
        import shutil
        shutil.rmtree(os.path.join(path, "definition", "pages", tgt_page["id"]), ignore_errors=True)
        return f"Page '{page_name}' deleted."
        
    # DELETE VISUAL
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    target_id = visual_id
    if not target_id and visual_title:
        for entry in os.scandir(visuals_dir):
            if entry.is_dir():
                v_file = os.path.join(entry.path, "visual.json")
                if os.path.exists(v_file):
                    try:
                        with open(v_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            t = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                            if t == visual_title: target_id = entry.name; break
                    except: pass
                    
    if target_id:
        import shutil
        shutil.rmtree(os.path.join(visuals_dir, target_id), ignore_errors=True)
        return f"Visual deleted (ID: {target_id})."
    return "Visual not found."

def pbir_update_visual_layout(page_name: str, visual_title: str, x: int = None, y: int = None, width: int = None, height: int = None, z: int = None) -> str:
    """Update position and size of a visual."""
    path = PBIRManager.detect_path()
    if not path: return "No project detected."
    pages = PBIRManager.get_pages(path)
    tgt_page = next((p for p in pages if p["name"] == page_name), None)
    if not tgt_page: return "Page not found."
    
    visuals_dir = os.path.join(path, "definition", "pages", tgt_page["id"], "visuals")
    vis_path = None; vis_data = None
    
    for entry in os.scandir(visuals_dir):
        if entry.is_dir():
            v_file = os.path.join(entry.path, "visual.json")
            if os.path.exists(v_file):
                try:
                    with open(v_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        t = data.get("visual", {}).get("objects", {}).get("general", [])[0]["properties"]["title"]["expr"]["Literal"]["Value"].replace("'", "")
                        if t == visual_title: vis_path = v_file; vis_data = data; break
                except: pass
                
    if not vis_data: return "Visual not found."
    
    if "position" not in vis_data: vis_data["position"] = {}
    if x is not None: vis_data["position"]["x"] = x
    if y is not None: vis_data["position"]["y"] = y
    if width is not None: vis_data["position"]["width"] = width
    if height is not None: vis_data["position"]["height"] = height
    if z is not None: vis_data["position"]["z"] = z
    
    with open(vis_path, 'w', encoding='utf-8') as f: json.dump(vis_data, f, indent=2)
    return f"Updated layout for '{visual_title}'."
