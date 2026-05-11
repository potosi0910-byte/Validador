"""
api.py — Backend FastAPI para la Malla Validadora RIPS (Resolución 2275/2023)
Expone la misma lógica de app_medicamentos_control.py como API REST JSON.
"""

from __future__ import annotations

import io
import json
from datetime import datetime
from typing import List

from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel

from auth import (
    autenticar_usuario,
    cambiar_password,
    crear_token,
    crear_usuario,
    eliminar_usuario,
    get_usuario_actual,
    listar_usuarios,
    solo_admin,
)
from app_medicamentos_control import (
    cargar_excel_autorizaciones,
    construir_excel,
    contar_registros_rips,
    extraer_autorizaciones_rips,
    extraer_medicamentos_invalidos,
    validar_autorizaciones,
    validar_consultas_malla_2275,
    validar_general_malla_2275,
    validar_hospitalizacion_malla_2275,
    validar_medicamentos_malla_2275,
    validar_otros_servicios_malla_2275,
    validar_procedimientos_malla_2275,
    validar_recien_nacidos_malla_2275,
    validar_urgencias_malla_2275,
    validar_usuarios_malla_2275,
)
from auditoria import validar_auditoria

# ══════════════════════════════════════════════════════════════
# APP
# ══════════════════════════════════════════════════════════════

app = FastAPI(
    title="Malla Validadora RIPS 2275",
    description="API REST para validación de archivos RIPS (Resolución 2275/2023)",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Caché en memoria del último resultado procesado (para exportar sin reenviar archivos)
_cache: dict = {}


# ══════════════════════════════════════════════════════════════
# WRAPPER — hace que bytes se comporten como un FileStorage de Flask
# (openpyxl y json lo usan como file-like)
# ══════════════════════════════════════════════════════════════

class _FileWrapper(io.BytesIO):
    """BytesIO con atributo filename para compatibilidad con Flask FileStorage."""

    def __init__(self, filename: str, content: bytes):
        super().__init__(content)
        self.filename = filename


# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def _count_sev(lista: list, sev: str) -> int:
    return sum(1 for v in lista if v.get("severidad") == sev)


def _json_safe(obj):
    """Convierte tipos no serializables para la respuesta JSON."""
    if isinstance(obj, set):
        return list(obj)
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not JSON serializable")


# ══════════════════════════════════════════════════════════════
# ENDPOINTS
# ══════════════════════════════════════════════════════════════

# ── Modelos Pydantic ──────────────────────────────────────────────────────────

class NuevoUsuario(BaseModel):
    username: str
    password: str
    role:     str = "auditor"
    nombre:   str = ""

class CambioPassword(BaseModel):
    username:     str
    nueva_password: str


# ══════════════════════════════════════════════════════════════
# ENDPOINTS AUTH
# ══════════════════════════════════════════════════════════════

@app.post("/api/auth/login", tags=["Autenticación"])
def login(form: OAuth2PasswordRequestForm = Depends()):
    usuario = autenticar_usuario(form.username, form.password)
    if not usuario:
        raise HTTPException(status_code=401, detail="Usuario o contraseña incorrectos.")
    token = crear_token(usuario["username"], usuario["role"], usuario.get("nombre", ""))
    return {
        "access_token": token,
        "token_type":   "bearer",
        "username":     usuario["username"],
        "role":         usuario["role"],
        "nombre":       usuario.get("nombre", ""),
    }


@app.get("/api/auth/me", tags=["Autenticación"])
def me(usuario: dict = Depends(get_usuario_actual)):
    return usuario


@app.get("/api/auth/usuarios", tags=["Autenticación"])
def get_usuarios(_: dict = Depends(solo_admin)):
    return listar_usuarios()


@app.post("/api/auth/usuarios", tags=["Autenticación"])
def post_usuario(body: NuevoUsuario, _: dict = Depends(solo_admin)):
    return crear_usuario(body.username, body.password, body.role, body.nombre)


@app.delete("/api/auth/usuarios/{username}", tags=["Autenticación"])
def delete_usuario(username: str, _: dict = Depends(solo_admin)):
    eliminar_usuario(username)
    return {"ok": True}


@app.put("/api/auth/password", tags=["Autenticación"])
def put_password(body: CambioPassword, _: dict = Depends(solo_admin)):
    cambiar_password(body.username, body.nueva_password)
    return {"ok": True}


# ══════════════════════════════════════════════════════════════
# SISTEMA
# ══════════════════════════════════════════════════════════════

@app.get("/api/health", tags=["Sistema"])
def health():
    return {"status": "ok", "version": "2.0.0"}


@app.post("/api/procesar", tags=["Validación"])
async def procesar(
    json_files: List[UploadFile] = File(..., description="Archivos RIPS JSON"),
    excel_files: List[UploadFile] = File(default=[], description="Excel de autorizaciones (opcional)"),
    usuario: dict = Depends(get_usuario_actual),
):
    """
    Procesa uno o más archivos RIPS JSON con autorizaciones Excel opcionales.
    Retorna todas las validaciones, alertas y estadísticas en formato JSON.
    """
    registros: list = []
    alertas: list = []
    validaciones_malla: list = []
    validaciones_general: list = []
    validaciones_auditoria: list = []
    errores_acum: list = []

    # ── Leer archivos JSON ────────────────────────────────────────────────
    json_wrappers: list[_FileWrapper] = []
    for f in json_files:
        if not f or not f.filename:
            continue
        content = await f.read()
        json_wrappers.append(_FileWrapper(f.filename, content))

    # ── Leer archivos Excel ───────────────────────────────────────────────
    excel_wrappers: list[_FileWrapper] = []
    for f in excel_files:
        if not f or not f.filename:
            continue
        content = await f.read()
        excel_wrappers.append(_FileWrapper(f.filename, content))

    hay_excel = bool(excel_wrappers)

    # ── Cargar autorizaciones desde Excel ─────────────────────────────────
    registros_excel: dict = {}
    set_aut_excel: set = set()
    if hay_excel:
        registros_excel, set_aut_excel, errores_excel = cargar_excel_autorizaciones(excel_wrappers)
        errores_acum.extend(errores_excel)

    # ── Procesar cada RIPS JSON ───────────────────────────────────────────
    archivos_procesados = 0
    total_rips = 0
    pacientes_rips_global: dict = {}

    for wrapper in json_wrappers:
        try:
            wrapper.seek(0)
            data = json.loads(wrapper.read())

            registros.extend(extraer_medicamentos_invalidos(data, wrapper.filename))

            pacientes = extraer_autorizaciones_rips(data, wrapper.filename)
            pacientes_rips_global.update(pacientes)

            total_rips += contar_registros_rips(data)

            validaciones_malla.extend(validar_medicamentos_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_general_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_usuarios_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_consultas_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_procedimientos_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_urgencias_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_hospitalizacion_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_recien_nacidos_malla_2275(data, wrapper.filename))
            validaciones_general.extend(validar_otros_servicios_malla_2275(data, wrapper.filename))
            validaciones_auditoria.extend(validar_auditoria(data, wrapper.filename))

            archivos_procesados += 1

        except Exception as exc:
            errores_acum.append(f"Error en {wrapper.filename}: {exc}")

    # ── Validar autorizaciones ────────────────────────────────────────────
    if hay_excel and (registros_excel or set_aut_excel):
        alertas = validar_autorizaciones(pacientes_rips_global, registros_excel, set_aut_excel)

    # ── Estadísticas resumen (mismo formato que Flask) ────────────────────
    t_proc = datetime.now()

    def _top_reglas(lista, n=5):
        cnt = {}
        for v in lista:
            r = v.get("id_regla", "")
            cnt[r] = cnt.get(r, 0) + 1
        return sorted(cnt.items(), key=lambda x: -x[1])[:n]

    total_auths_rips = sum(len(p.get("set_auths", set())) for p in pacientes_rips_global.values())

    def _alen(key):
        return len(alertas.get(key, [])) if alertas else 0

    stats = {
        "archivos_json":              archivos_procesados,
        "total_rips":                 total_rips,
        "alerta_volumen":             total_rips > 800,
        "med_invalidos":              len(registros),
        "auts_rips":                  total_auths_rips,
        "auts_excel":                 len(set_aut_excel),
        "tipo_mismatch":              _alen("tipo_doc_mismatch"),
        "amb_par_nc":                 _alen("amb_par_no_cruza"),
        "hosp_proc_nc":               _alen("hosp_proc_no_cruza"),
        "hosp_aut_no_rel":            _alen("hosp_aut_no_relacionada"),
        "estancia_sin_aut":           _alen("estancia_sin_aut"),
        "proc_qx_aut_hosp":           _alen("proc_qx_misma_aut_hosp"),
        "sin_aut_rel":                _alen("sin_num_aut_relacionado"),
        "amb_emision_post":           _alen("amb_aut_emision_posterior"),
        "hosp_cod_sin_aut":           _alen("hosp_cod_sin_aut"),
        "hosp_cups_duplicado":        _alen("hosp_cups_duplicado"),
        "proc_sin_aut_amb":           _alen("proc_sin_aut_amb"),
        "proc_aut_no_cruza":          _alen("proc_aut_no_cruza_amb"),
        "cups_noestandar_nc":         _alen("cups_noestandar_sin_aut"),
        "hosp_proc_cod_no_cruza":     _alen("hosp_proc_cod_no_cruza"),
        "malla_total":                len(validaciones_malla),
        "malla_criticas":             _count_sev(validaciones_malla, "critica"),
        "malla_notificaciones":       sum(1 for v in validaciones_malla if v.get("severidad") in {"media", "alta"}),
        "malla_top_reglas":           _top_reglas(validaciones_malla),
        "general_total":              len(validaciones_general),
        "general_criticas":           _count_sev(validaciones_general, "critica"),
        "general_notificaciones":     sum(1 for v in validaciones_general if v.get("severidad") in {"media", "alta"}),
        "general_top_reglas":         _top_reglas(validaciones_general),
        "auditoria_total":            len(validaciones_auditoria),
        "auditoria_criticas":         _count_sev(validaciones_auditoria, "critica"),
        "auditoria_notificaciones":   sum(1 for v in validaciones_auditoria if v.get("severidad") in {"media", "alta"}),
        "auditoria_top_reglas":       _top_reglas(validaciones_auditoria),
        "tiempo_procesamiento":       f"{(datetime.now() - t_proc).total_seconds():.1f}s",
        "errores_procesamiento":      errores_acum,
    }

    # ── Guardar en caché para exportar ────────────────────────────────────
    _cache["ultimo"] = {
        "registros": registros,
        "alertas": alertas,
        "validaciones_malla": validaciones_malla,
        "validaciones_general": validaciones_general,
        "validaciones_auditoria": validaciones_auditoria,
        "stats": stats,
    }

    return {
        "stats": stats,
        "registros": registros,
        "alertas": alertas,
        "validaciones_malla": validaciones_malla,
        "validaciones_general": validaciones_general,
        "validaciones_auditoria": validaciones_auditoria,
    }


@app.get("/api/exportar", tags=["Exportación"])
def exportar(usuario: dict = Depends(get_usuario_actual)):
    """Exporta los últimos resultados procesados a Excel (.xlsx)."""
    cached = _cache.get("ultimo", {})
    if not cached.get("stats"):
        return {"error": "No hay resultados. Procese los archivos RIPS primero."}

    output = construir_excel(
        cached["registros"],
        cached["alertas"],
        cached["validaciones_malla"],
        cached["validaciones_general"],
        cached.get("validaciones_auditoria"),
    )

    nombre = f"Alertas_Malla_Validadora_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={nombre}"},
    )


# ══════════════════════════════════════════════════════════════
# ARRANQUE DIRECTO
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="127.0.0.1", port=8000, reload=True)
