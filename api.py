"""
api.py — Backend FastAPI para la Malla Validadora RIPS (Resolución 2275/2023)
Expone la misma lógica de app_medicamentos_control.py como API REST JSON.
"""

from __future__ import annotations

import base64
import glob as _glob
import io
import json
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
from collections import defaultdict
from datetime import datetime
from time import time
from typing import List

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.security import OAuth2PasswordRequestForm
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from auth import (
    autenticar_usuario,
    cambiar_password,
    cambiar_password_verificado,
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

_ORIGINS = os.environ.get(
    "ALLOWED_ORIGINS",
    "http://localhost:5173,http://localhost:3000"
).split(",")

app.add_middleware(
    CORSMiddleware,
    allow_origins=_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["Authorization", "Content-Type"],
)

# Caché en memoria del último resultado procesado (para exportar sin reenviar archivos)
_cache: dict = {}

# ══════════════════════════════════════════════════════════════
# ESTADÍSTICAS — SQLite (solo conteos, sin datos de pacientes)
# ══════════════════════════════════════════════════════════════

_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "estadisticas.db")


def _init_db():
    con = sqlite3.connect(_DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS estadisticas (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha                 TEXT NOT NULL,
            usuario               TEXT NOT NULL,
            archivos_json         INTEGER,
            total_rips            INTEGER,
            med_invalidos         INTEGER,
            auts_rips             INTEGER,
            auts_excel            INTEGER,
            tipo_mismatch         INTEGER,
            amb_par_nc            INTEGER,
            hosp_proc_nc          INTEGER,
            hosp_aut_no_rel       INTEGER,
            estancia_sin_aut      INTEGER,
            proc_qx_aut_hosp      INTEGER,
            sin_aut_rel           INTEGER,
            amb_emision_post      INTEGER,
            hosp_cod_sin_aut      INTEGER,
            hosp_cups_duplicado   INTEGER,
            proc_sin_aut_amb      INTEGER,
            proc_aut_no_cruza     INTEGER,
            cups_noestandar_nc    INTEGER,
            hosp_proc_cod_no_cruza INTEGER,
            malla_total           INTEGER,
            malla_criticas        INTEGER,
            malla_notificaciones  INTEGER,
            general_total         INTEGER,
            general_criticas      INTEGER,
            general_notificaciones INTEGER,
            auditoria_total       INTEGER,
            auditoria_criticas    INTEGER,
            auditoria_notificaciones INTEGER,
            top_reglas_malla      TEXT,
            top_reglas_general    TEXT,
            top_reglas_auditoria  TEXT,
            tiempo_procesamiento  TEXT
        )
    """)
    con.commit()
    con.close()


def _guardar_estadistica(stats: dict, usuario: str):
    try:
        con = sqlite3.connect(_DB_PATH)
        con.execute("""
            INSERT INTO estadisticas (
                fecha, usuario,
                archivos_json, total_rips, med_invalidos, auts_rips, auts_excel,
                tipo_mismatch, amb_par_nc, hosp_proc_nc, hosp_aut_no_rel,
                estancia_sin_aut, proc_qx_aut_hosp, sin_aut_rel, amb_emision_post,
                hosp_cod_sin_aut, hosp_cups_duplicado, proc_sin_aut_amb,
                proc_aut_no_cruza, cups_noestandar_nc, hosp_proc_cod_no_cruza,
                malla_total, malla_criticas, malla_notificaciones,
                general_total, general_criticas, general_notificaciones,
                auditoria_total, auditoria_criticas, auditoria_notificaciones,
                top_reglas_malla, top_reglas_general, top_reglas_auditoria,
                tiempo_procesamiento
            ) VALUES (
                ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
            )
        """, (
            datetime.now().isoformat(),
            usuario,
            stats.get("archivos_json", 0),
            stats.get("total_rips", 0),
            stats.get("med_invalidos", 0),
            stats.get("auts_rips", 0),
            stats.get("auts_excel", 0),
            stats.get("tipo_mismatch", 0),
            stats.get("amb_par_nc", 0),
            stats.get("hosp_proc_nc", 0),
            stats.get("hosp_aut_no_rel", 0),
            stats.get("estancia_sin_aut", 0),
            stats.get("proc_qx_aut_hosp", 0),
            stats.get("sin_aut_rel", 0),
            stats.get("amb_emision_post", 0),
            stats.get("hosp_cod_sin_aut", 0),
            stats.get("hosp_cups_duplicado", 0),
            stats.get("proc_sin_aut_amb", 0),
            stats.get("proc_aut_no_cruza", 0),
            stats.get("cups_noestandar_nc", 0),
            stats.get("hosp_proc_cod_no_cruza", 0),
            stats.get("malla_total", 0),
            stats.get("malla_criticas", 0),
            stats.get("malla_notificaciones", 0),
            stats.get("general_total", 0),
            stats.get("general_criticas", 0),
            stats.get("general_notificaciones", 0),
            stats.get("auditoria_total", 0),
            stats.get("auditoria_criticas", 0),
            stats.get("auditoria_notificaciones", 0),
            json.dumps(stats.get("malla_top_reglas", [])),
            json.dumps(stats.get("general_top_reglas", [])),
            json.dumps(stats.get("auditoria_top_reglas", [])),
            stats.get("tiempo_procesamiento", ""),
        ))
        con.commit()
        con.close()
    except Exception:
        pass  # nunca bloquear el flujo principal por un error de estadísticas


_init_db()


# ══════════════════════════════════════════════════════════════
# RATE LIMITING — protección bruta-fuerza en login
# ══════════════════════════════════════════════════════════════

_login_attempts: dict = defaultdict(list)
_RATE_WINDOW  = 60   # segundos
_RATE_MAX     = 10   # intentos por ventana


def _check_login_rate(ip: str) -> None:
    now = time()
    _login_attempts[ip] = [t for t in _login_attempts[ip] if now - t < _RATE_WINDOW]
    if len(_login_attempts[ip]) >= _RATE_MAX:
        raise HTTPException(
            status_code=429,
            detail="Demasiados intentos de acceso. Espere 1 minuto.",
            headers={"Retry-After": "60"},
        )
    _login_attempts[ip].append(now)


def _safe_filename(name: str) -> str:
    """Elimina path traversal y caracteres peligrosos de un nombre de archivo."""
    name = os.path.basename(name)
    name = re.sub(r"[^\w\-_. ]", "_", name)
    return name[:200] or "upload"


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
    username:       str
    nueva_password: str

class CambioPasswordPropio(BaseModel):
    password_actual: str
    nueva_password:  str


# ══════════════════════════════════════════════════════════════
# ENDPOINTS AUTH
# ══════════════════════════════════════════════════════════════

@app.post("/api/auth/login", tags=["Autenticación"])
def login(request: Request, form: OAuth2PasswordRequestForm = Depends()):
    ip = request.client.host if request.client else "unknown"
    _check_login_rate(ip)
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


@app.put("/api/auth/me/password", tags=["Autenticación"])
def put_mi_password(body: CambioPasswordPropio, usuario: dict = Depends(get_usuario_actual)):
    cambiar_password_verificado(usuario["username"], body.password_actual, body.nueva_password)
    return {"ok": True}


# ══════════════════════════════════════════════════════════════
# SISTEMA
# ══════════════════════════════════════════════════════════════

@app.get("/api/health", tags=["Sistema"])
def health():
    return {"status": "ok", "version": "2.0.0"}


@app.post("/api/procesar", tags=["Validación"])
async def procesar(
    request: Request,
    usuario: dict = Depends(get_usuario_actual),
):
    """
    Procesa uno o más archivos RIPS JSON con autorizaciones Excel opcionales.
    Retorna todas las validaciones, alertas y estadísticas en formato JSON.
    """
    # python-multipart tiene límite de 1000 archivos por defecto; lo subimos
    form = await request.form(max_files=5000, max_fields=5000)
    json_files  = form.getlist("json_files")
    excel_files = form.getlist("excel_files")

    if not json_files:
        raise HTTPException(status_code=400, detail="Se requieren archivos RIPS JSON.")

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

    # ── Guardar estadísticas (solo conteos, sin datos de pacientes) ──────
    _guardar_estadistica(stats, usuario.get("username", "desconocido"))

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


@app.get("/api/estadisticas", tags=["Estadísticas"])
def get_estadisticas(
    limite: int = 100,
    _: dict = Depends(solo_admin),
):
    """Retorna el historial de procesamiento (solo conteos, sin datos de pacientes)."""
    con = sqlite3.connect(_DB_PATH)
    con.row_factory = sqlite3.Row
    filas = con.execute(
        "SELECT * FROM estadisticas ORDER BY id DESC LIMIT ?", (limite,)
    ).fetchall()
    con.close()
    resultado = []
    for f in filas:
        row = dict(f)
        row["top_reglas_malla"]     = json.loads(row.get("top_reglas_malla") or "[]")
        row["top_reglas_general"]   = json.loads(row.get("top_reglas_general") or "[]")
        row["top_reglas_auditoria"] = json.loads(row.get("top_reglas_auditoria") or "[]")
        resultado.append(row)
    return {"total": len(resultado), "registros": resultado}


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
# MACHINE LEARNING — Análisis Predictivo Estancia Hospitalaria
# ══════════════════════════════════════════════════════════════

_ML_SCRIPT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "analisis_hospitalario_v1.py",
)


@app.post("/api/ml/analizar", tags=["Machine Learning"])
async def ml_analizar(
    excel_file: UploadFile = File(..., description="Excel con hojas DX_TRIAGE y CUPS"),
    k: int = Form(default=4, ge=2, le=10, description="Número de clusters K-Means"),
    _usuario: dict = Depends(get_usuario_actual),
):
    """
    Ejecuta el análisis predictivo de estancia hospitalaria con K-Means +
    modelos supervisados. Retorna las gráficas generadas como imágenes base64.
    """
    if not os.path.isfile(_ML_SCRIPT):
        raise HTTPException(status_code=500, detail="Script de análisis no disponible.")

    content = await excel_file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Archivo Excel vacío.")

    safe_name = _safe_filename(excel_file.filename or "datos.xlsx")

    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, safe_name)
        with open(excel_path, "wb") as fh:
            fh.write(content)

        out_dir = os.path.join(tmpdir, "figuras")
        os.makedirs(out_dir, exist_ok=True)

        _ml_env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
        try:
            result = subprocess.run(
                [sys.executable, _ML_SCRIPT, "--excel", excel_path, "--k", str(k), "--out", out_dir],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=900,
                cwd=tmpdir,
                env=_ml_env,
            )
        except subprocess.TimeoutExpired:
            raise HTTPException(status_code=504, detail="El análisis superó el tiempo límite (15 min).")

        if result.returncode != 0:
            # No exponer stderr al cliente — solo loguear en servidor
            import sys as _sys
            print(result.stderr[-3000:], file=_sys.stderr)
            raise HTTPException(status_code=500, detail="Error al procesar el archivo. Verifique que el Excel tenga las hojas DX_TRIAGE y CUPS.")

        imagenes: dict = {}
        for img_path in sorted(_glob.glob(os.path.join(out_dir, "*.png"))):
            nombre = os.path.splitext(os.path.basename(img_path))[0]
            with open(img_path, "rb") as fh:
                imagenes[nombre] = base64.b64encode(fh.read()).decode("utf-8")

        return {
            "status": "ok",
            "k": k,
            "archivo": excel_file.filename,
            "n_graficas": len(imagenes),
            "log": result.stdout[-5000:] if result.stdout else "",
            "imagenes": imagenes,
        }


# ══════════════════════════════════════════════════════════════
# FRONTEND ESTÁTICO (producción)
# Sirve el build de React desde frontend/dist/
# ══════════════════════════════════════════════════════════════

_DIST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend", "dist")

if os.path.isdir(_DIST):
    _assets = os.path.join(_DIST, "assets")
    if os.path.isdir(_assets):
        app.mount("/assets", StaticFiles(directory=_assets), name="static-assets")

    @app.get("/styles.css", include_in_schema=False)
    async def _styles():
        return FileResponse(os.path.join(_DIST, "styles.css"))

    @app.get("/", include_in_schema=False)
    async def _index():
        return FileResponse(os.path.join(_DIST, "index.html"))

    @app.get("/{full_path:path}", include_in_schema=False)
    async def _spa(full_path: str):
        return FileResponse(os.path.join(_DIST, "index.html"))


# ══════════════════════════════════════════════════════════════
# ARRANQUE DIRECTO
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="127.0.0.1", port=8000, reload=True)
