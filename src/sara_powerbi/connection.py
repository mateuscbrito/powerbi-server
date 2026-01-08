import sys
import os
import psutil
try:
    import clr
except ImportError:
    clr = None # Handle non-windows or missing pythonnet gracefully

# Global State for caching port/connection info if needed
GLOBAL_CONTEXT = {"port": None}

def load_libs() -> bool:
    """
    Loads the necessary Analysis Services (TOM/ADOMD) DLLs from Power BI Desktop installation.
    Returns True if loaded, False otherwise.
    """
    if not clr:
        return False
        
    search_paths = [
        r"C:\Program Files\Microsoft Power BI Desktop\bin", 
        r"C:\Program Files (x86)\Microsoft Power BI Desktop\bin"
    ]
    
    # Try dynamic detection from running process
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        if proc.info['name'] and 'msmdsrv.exe' in proc.info['name'].lower():
            try:
                exe_path = proc.info['exe']
                if exe_path: 
                    search_paths.insert(0, os.path.dirname(exe_path))
            except: pass
    
    loaded = False
    for path in search_paths:
        if not os.path.exists(path): continue
        try:
            sys.path.append(path)
            try: 
                clr.AddReference("Microsoft.AnalysisServices.AdomdClient")
            except: 
                clr.AddReference("Microsoft.PowerBI.AdomdClient")
                
            try: 
                clr.AddReference("Microsoft.AnalysisServices.Tabular")
                loaded = True
            except: 
                try: 
                    clr.AddReference("Microsoft.PowerBI.Tabular")
                    loaded = True
                except: pass
        except: pass
        
        if loaded: break
        
    return loaded

def find_port() -> int | None:
    """Finds the local TCP port of the running Power BI Analysis Services instance."""
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'msmdsrv.exe' in proc.info['name'].lower():
            try:
                for conn in proc.connections(kind='tcp'):
                    if conn.status == 'LISTEN' and conn.laddr.ip == '127.0.0.1':
                        return conn.laddr.port
            except: pass
    return None

def get_server():
    """Returns a connected TOM Server object."""
    if not load_libs(): 
        raise Exception("Failed to load Power BI DLLs. Is Power BI Desktop installed?")
        
    port = find_port()
    if not port: 
        raise Exception("Power BI Desktop is not running or no model is open.")
    
    GLOBAL_CONTEXT["port"] = port
    
    from Microsoft.AnalysisServices.Tabular import Server
    s = Server()
    s.Connect(f"Data Source=localhost:{port};")
    return s

def get_adomd_connection():
    """Returns an open AdomdConnection for querying."""
    if not load_libs():
        raise Exception("Failed to load Power BI DLLs.")
        
    port = find_port()
    if not port:
        raise Exception("Power BI Desktop is not running.")
        
    try: 
        from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
    except: 
        from Microsoft.PowerBI.AdomdClient import AdomdConnection
    
    conn = AdomdConnection(f"Data Source=localhost:{port};")
    conn.Open()
    return conn
