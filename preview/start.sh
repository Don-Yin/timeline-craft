#!/bin/bash
set -e

echo "Starting FastAPI preview server on port 8004..."
exec uvicorn server:app --host 0.0.0.0 --port 8004

