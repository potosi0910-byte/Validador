#!/bin/bash

# Script para iniciar Cloudflare Tunnel y Flask en paralelo (Linux/Mac)

echo "======================================"
echo "Iniciando Extractor RIPS con Cloudflare Tunnel"
echo "======================================"

# Iniciar Flask en background
echo "Iniciando Flask en puerto 5000..."
python app_medicamentos_control.py &
FLASK_PID=$!

# Esperar 3 segundos para que Flask inicie
sleep 3

# Iniciar Cloudflare Tunnel
echo ""
echo "Iniciando Cloudflare Tunnel..."
cloudflared tunnel run extractor-rips

# Limpiar al salir
trap "kill $FLASK_PID" EXIT
