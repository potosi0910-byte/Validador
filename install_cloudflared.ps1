param(
    [switch]$SkipPolicyCheck
)

Write-Host "================================"
Write-Host "Instalador Cloudflared"
Write-Host "================================"
Write-Host ""

$installDir = "C:\Program Files\cloudflared"
$downloadUrl = "https://github.com/cloudflare/cloudflared/releases/download/2024.5.0/cloudflared-windows-amd64.exe"

Write-Host "[*] Descargando cloudflared..."
Write-Host "URL: $downloadUrl"
Write-Host ""

if (-not (Test-Path $installDir)) {
    Write-Host "[*] Creando carpeta: $installDir"
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    Write-Host "[OK] Carpeta creada"
}

$exePath = Join-Path $installDir "cloudflared.exe"

Write-Host "[*] Descargando archivo..."
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -UseBasicParsing
    Write-Host "[OK] Descargado: $exePath"
} catch {
    Write-Host "[ERROR] Error descargando: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Intenta descargar manualmente desde:"
    Write-Host $downloadUrl
    exit 1
}

$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($userPath -notlike "*$installDir*") {
    Write-Host ""
    Write-Host "[*] Actualizando PATH..."
    [Environment]::SetEnvironmentVariable("PATH", "$userPath;$installDir", "User")
    Write-Host "[OK] PATH actualizado"
}

Write-Host ""
Write-Host "[OK] Instalacion completada!"
Write-Host ""
Write-Host "PROXIMOS PASOS:"
Write-Host "1. Cierra PowerShell completamente"
Write-Host "2. Abre una NUEVA terminal de PowerShell"
Write-Host "3. Ejecuta: cloudflared login"
Write-Host ""
