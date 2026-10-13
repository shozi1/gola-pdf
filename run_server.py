"""
Gola PDF Live Server Bootstrapper
Starts FastAPI web server using Uvicorn on localhost.
"""

import uvicorn
import os
import sys

# Ensure current directory is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    port = 8080
    print(f"=======================================================")
    print(f"   Gola PDF Complete Suite is LIVE!                   ")
    print(f"   URL: http://127.0.0.1:{port}                         ")
    print(f"   Open your browser at: http://127.0.0.1:{port}        ")
    print(f"=======================================================")
    uvicorn.run("app:app", host="127.0.0.1", port=port, reload=False, log_level="info")
