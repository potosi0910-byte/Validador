#!/usr/bin/env powershell
# Script rápido para configurar y ejecutar con Cloudflare Tunnel

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Extractor RIPS + Cloudflare Tunnel" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Verificar si cloudflared está instalado
$cloudflared = Get-Command cloudflared -ErrorAction SilentlyContinue
if (-not $cloudflared) {
    Write-Host "ERROR: cloudflared no está instalado o no está en PATH" -ForegroundColor Red
    Write-Host "Descarga desde: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ cloudflared encontrado" -ForegroundColor Green
Write-Host ""

# Mostrar pasos
Write-Host "PASOS PARA INICIAR:" -ForegroundColor Yellow
Write-Host "1. cloudflared login          (autentica tu cuenta Cloudflare - solo primera vez)" -ForegroundColor White
Write-Host "2. cloudflared tunnel create extractor-rips   (crea el tunnel - solo primera vez)" -ForegroundColor White
Write-Host "3. python app_medicamentos_control.py         (en terminal 1)" -ForegroundColor White
Write-Host "4. cloudflared tunnel run extractor-rips      (en terminal 2)" -ForegroundColor White
Write-Host ""

Write-Host "¿Ya completaste los pasos 1 y 2?" -ForegroundColor Cyan
Write-Host "Y" -ForegroundColor Green -NoNewline
Write-Host " = Iniciar Flask | " -ForegroundColor White -NoNewline
Write-Host "N" -ForegroundColor Green -NoNewline
Write-Host " = Ver instrucciones detalladas" -ForegroundColor White

$response = Read-Host "Tu respuesta"

if ($response -eq "Y" -or $response -eq "y") {
    Write-Host ""
    Write-Host "Iniciando Flask..." -ForegroundColor Green
    Write-Host "Después de iniciar, abre OTRA terminal y ejecuta:" -ForegroundColor Yellow
    Write-Host "cloudflared tunnel run extractor-rips" -ForegroundColor Cyan
    Write-Host ""
    python app_medicamentos_control.py
} else {
    Write-Host ""
    Write-Host "Lee el archivo: CLOUDFLARE_SETUP.md" -ForegroundColor Cyan
    Write-Host ""
}
