# Configuración de Cloudflare Tunnel para Extractor RIPS

## ¿Qué es Cloudflare Tunnel?
Cloudflare Tunnel (cloudflared) permite exponer tu aplicación Flask local a través de una URL pública sin necesidad de abrir puertos en tu firewall.

---

## Pasos de Instalación y Configuración

### 1. Instalar cloudflared
Descarga desde: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/

O instala con Chocolatey (Windows):
```powershell
choco install cloudflare-warp
```

### 2. Autenticarse con Cloudflare
Ejecuta en PowerShell:
```powershell
cloudflared login
```

Esto abrirá tu navegador para autorizar cloudflared con tu cuenta de Cloudflare.

### 3. Crear un Tunnel
```powershell
cloudflared tunnel create extractor-rips
```

Guarda el ID del tunnel que se mostrará en pantalla.

### 4. Crear el archivo de configuración
El archivo `.cloudflared\config.yml` ya está creado, pero necesitas reemplazar `<TUNNEL_ID>` con tu ID real.

Edita:
```
.cloudflared\config.yml
```

Reemplaza la línea:
```
credentials-file: C:\Users\HAUDITORIA42\.cloudflared\<TUNNEL_ID>.json
```

Con tu ID real, por ejemplo:
```
credentials-file: C:\Users\HAUDITORIA42\.cloudflared\a1b2c3d4-e5f6-7890-abcd-ef1234567890.json
```

### 5. Crear la ruta DNS (opcional, para URL personalizada)
Si quieres usar un dominio personalizado:

```powershell
cloudflared tunnel route dns extractor-rips example.com
```

Reemplaza `example.com` con tu dominio Cloudflare.

---

## Ejecución

### Opción A: Script Automático
Ejecuta desde PowerShell:
```powershell
.\iniciar_cloudflared.bat
```

Esto inicia Flask y el tunnel automáticamente.

### Opción B: Manual (Recomendado para debug)
En terminal 1:
```powershell
python app_medicamentos_control.py
```

En terminal 2:
```powershell
cloudflared tunnel run extractor-rips
```

### Opción C: Ejecutar tunnel sin archivo config
```powershell
cloudflared tunnel run --url http://localhost:5000 extractor-rips
```

---

## Acceso a la aplicación

Una vez que cloudflared esté corriendo, verás una URL como:
```
https://extractor-rips-abc123.cloudflare.app
```

O si configuraste un dominio personalizado:
```
https://extractor-rips.example.com
```

Abre esa URL en tu navegador. ¡El aplicativo ya está expuesto públicamente!

---

## Cambiar el puerto de Flask (si es necesario)

Si quieres usar un puerto diferente, edita `app_medicamentos_control.py`:

```python
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)  # Cambia 5000 por tu puerto
```

Y actualiza la URL en `.cloudflared\config.yml`:
```yaml
ingress:
  - hostname: extractor-rips.cloudflare.app
    service: http://localhost:5000  # Cambia el puerto aquí
```

---

## Detener el Tunnel

Presiona `Ctrl+C` en la terminal donde corre cloudflared.

---

## Troubleshooting

| Problema | Solución |
|----------|----------|
| "tunnel not found" | Verifica que creaste el tunnel: `cloudflared tunnel list` |
| "invalid credentials" | Reautentícate: `cloudflared login` |
| "connection refused" | Asegúrate de que Flask está corriendo en puerto 5000 |
| CORS errors | El proxy ya maneja CORS correctamente |

---

## Seguridad

- El tunnel es privado y solo accesible con credenciales de Cloudflare
- No necesitas exponer puertos en tu router
- El tráfico está encriptado (HTTPS)
- Puedes agregar autenticación Cloudflare adicional si lo deseas

---

## Referencias
- https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/
- https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/install-and-setup/tunnel-guide/
