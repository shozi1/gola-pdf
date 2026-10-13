"""
Gola PDF Production Server Runner
Optimized for high concurrency and low memory on a $14 VPS.
"""

import os
import sys
import uvicorn

# Ensure current directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8080"))
    workers = int(os.getenv("WORKERS", "2"))
    log_level = os.getenv("LOG_LEVEL", "warning")

    print(f"Starting Gola PDF Production Engine on {host}:{port} with {workers} workers...")
    uvicorn.run(
        "app:app",
        host=host,
        port=port,
        workers=workers,
        reload=False,
        access_log=False,
        log_level=log_level
    )
