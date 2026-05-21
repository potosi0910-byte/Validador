#!/bin/bash
# Inicio del servidor en producción (GoDaddy VPS / Ubuntu)
cd "$(dirname "$0")"
source venv/bin/activate
uvicorn api:app --host 0.0.0.0 --port 8000 --workers 2
