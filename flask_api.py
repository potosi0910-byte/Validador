"""
flask_api.py — Backend Flask para la Malla Validadora RIPS
Versión WSGI nativa para GoDaddy cPanel + Passenger.
"""
import io
import json
import os
import re
import sqlite3
from datetime import datetime, timedelta
from functools import wraps

import bcrypt
from flask import Flask, jsonify, request, send_file
from jose import JWTError, jwt

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
from auditoria import validar_auditoria, validar_concepto_recaudo
from pertinencia import validar_pertinencia

# ══════════════════════════════════════════════════════════════
# APP
# ══════════════════════════════════════════════════════════════

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB
app.config['MAX_FORM_PARTS']     = 10000               # werkzeug 3.x: por defecto 1000, insuficiente para >1000 archivos
app.config['MAX_FORM_MEMORY_SIZE'] = 500 * 1024 * 1024 # 500 MB en memoria para el form

SECRET_KEY  = "drf-malla-validadora-2275-clave-secreta-cambiar-en-produccion"
ALGORITHM   = "HS256"
TOKEN_HORAS = 8

_BASE = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(_BASE, "users.json")
DB_PATH    = os.path.join(_BASE, "estadisticas.db")
_cache: dict = {}

_RIPS_FILENAME_RE = re.compile(r'^Rips_SL\d{6}\.json$')

# ══════════════════════════════════════════════════════════════
# CORS
# ══════════════════════════════════════════════════════════════

@app.after_request
def _cors(response):
    origin = os.environ.get("ALLOWED_ORIGINS", "*").split(",")[0]
    response.headers["Access-Control-Allow-Origin"]      = origin
    response.headers["Access-Control-Allow-Methods"]     = "GET,POST,PUT,DELETE,OPTIONS"
    response.headers["Access-Control-Allow-Headers"]     = "Content-Type,Authorization"
    response.headers["Access-Control-Allow-Credentials"] = "true"
    return response

@app.route("/api/<path:p>", methods=["OPTIONS"])
@app.route("/api/", methods=["OPTIONS"])
def _preflight(p=""):
    return "", 204

# ══════════════════════════════════════════════════════════════
# HELPERS — usuarios
# ══════════════════════════════════════════════════════════════

def _cargar_usuarios():
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _guardar_usuarios(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

def _verificar_password(plain, hashed):
    try:
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    except Exception:
        return False

def _hashear_password(plain):
    return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()

def _crear_token(username, role, nombre):
    expire = datetime.utcnow() + timedelta(hours=TOKEN_HORAS)
    return jwt.encode(
        {"sub": username, "role": role, "nombre": nombre, "exp": expire},
        SECRET_KEY, algorithm=ALGORITHM
    )

# ══════════════════════════════════════════════════════════════
# DECORADORES AUTH
# ══════════════════════════════════════════════════════════════

def _get_payload():
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return None
    try:
        return jwt.decode(auth[7:], SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None

def require_auth(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        p = _get_payload()
        if not p:
            return jsonify({"detail": "Token inválido o expirado."}), 401
        request.usuario = {"username": p["sub"], "role": p.get("role","auditor"), "nombre": p.get("nombre","")}
        return f(*args, **kwargs)
    return wrapper

def require_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        p = _get_payload()
        if not p:
            return jsonify({"detail": "Token inválido o expirado."}), 401
        if p.get("role") != "admin":
            return jsonify({"detail": "Se requiere rol de administrador."}), 403
        request.usuario = {"username": p["sub"], "role": "admin", "nombre": p.get("nombre","")}
        return f(*args, **kwargs)
    return wrapper

# ══════════════════════════════════════════════════════════════
# ESTADÍSTICAS SQLite
# ══════════════════════════════════════════════════════════════

def _init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS estadisticas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL, usuario TEXT NOT NULL,
            archivos_json INTEGER, total_rips INTEGER, med_invalidos INTEGER,
            auts_rips INTEGER, auts_excel INTEGER, tipo_mismatch INTEGER,
            amb_par_nc INTEGER, hosp_proc_nc INTEGER, hosp_aut_no_rel INTEGER,
            estancia_sin_aut INTEGER, proc_qx_aut_hosp INTEGER, sin_aut_rel INTEGER,
            amb_emision_post INTEGER, hosp_cod_sin_aut INTEGER, hosp_cups_duplicado INTEGER,
            proc_sin_aut_amb INTEGER, proc_aut_no_cruza INTEGER, cups_noestandar_nc INTEGER,
            hosp_proc_cod_no_cruza INTEGER, malla_total INTEGER, malla_criticas INTEGER,
            malla_notificaciones INTEGER, general_total INTEGER, general_criticas INTEGER,
            general_notificaciones INTEGER, auditoria_total INTEGER, auditoria_criticas INTEGER,
            auditoria_notificaciones INTEGER, top_reglas_malla TEXT, top_reglas_general TEXT,
            top_reglas_auditoria TEXT, tiempo_procesamiento TEXT
        )
    """)
    con.commit()
    con.close()

def _guardar_estadistica(stats, usuario):
    try:
        con = sqlite3.connect(DB_PATH)
        con.execute("""INSERT INTO estadisticas (
            fecha, usuario, archivos_json, total_rips, med_invalidos, auts_rips, auts_excel,
            tipo_mismatch, amb_par_nc, hosp_proc_nc, hosp_aut_no_rel, estancia_sin_aut,
            proc_qx_aut_hosp, sin_aut_rel, amb_emision_post, hosp_cod_sin_aut,
            hosp_cups_duplicado, proc_sin_aut_amb, proc_aut_no_cruza, cups_noestandar_nc,
            hosp_proc_cod_no_cruza, malla_total, malla_criticas, malla_notificaciones,
            general_total, general_criticas, general_notificaciones, auditoria_total,
            auditoria_criticas, auditoria_notificaciones, top_reglas_malla,
            top_reglas_general, top_reglas_auditoria, tiempo_procesamiento
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
            datetime.now().isoformat(), usuario,
            stats.get("archivos_json",0), stats.get("total_rips",0), stats.get("med_invalidos",0),
            stats.get("auts_rips",0), stats.get("auts_excel",0), stats.get("tipo_mismatch",0),
            stats.get("amb_par_nc",0), stats.get("hosp_proc_nc",0), stats.get("hosp_aut_no_rel",0),
            stats.get("estancia_sin_aut",0), stats.get("proc_qx_aut_hosp",0), stats.get("sin_aut_rel",0),
            stats.get("amb_emision_post",0), stats.get("hosp_cod_sin_aut",0), stats.get("hosp_cups_duplicado",0),
            stats.get("proc_sin_aut_amb",0), stats.get("proc_aut_no_cruza",0), stats.get("cups_noestandar_nc",0),
            stats.get("hosp_proc_cod_no_cruza",0), stats.get("malla_total",0), stats.get("malla_criticas",0),
            stats.get("malla_notificaciones",0), stats.get("general_total",0), stats.get("general_criticas",0),
            stats.get("general_notificaciones",0), stats.get("auditoria_total",0), stats.get("auditoria_criticas",0),
            stats.get("auditoria_notificaciones",0),
            json.dumps(stats.get("malla_top_reglas",[])), json.dumps(stats.get("general_top_reglas",[])),
            json.dumps(stats.get("auditoria_top_reglas",[])), stats.get("tiempo_procesamiento",""),
        ))
        con.commit()
        con.close()
    except Exception:
        pass

_init_db()

# ══════════════════════════════════════════════════════════════
# HELPERS — procesamiento
# ══════════════════════════════════════════════════════════════

class _FileWrapper(io.BytesIO):
    def __init__(self, filename, content):
        super().__init__(content)
        self.filename = filename

def _count_sev(lista, sev):
    return sum(1 for v in lista if v.get("severidad") == sev)

def _top_reglas(lista, n=5):
    cnt = {}
    for v in lista:
        r = v.get("id_regla", "")
        cnt[r] = cnt.get(r, 0) + 1
    return sorted(cnt.items(), key=lambda x: -x[1])[:n]

def _alen(alertas, key):
    return len(alertas.get(key, [])) if alertas else 0

def _json_safe(obj):
    if isinstance(obj, set):
        return list(obj)
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not JSON serializable")

# ══════════════════════════════════════════════════════════════
# ENDPOINTS
# ══════════════════════════════════════════════════════════════

@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "version": "2.0.0"})


# ── Auth ──────────────────────────────────────────────────────

@app.route("/api/auth/login", methods=["POST"])
def login():
    username = request.form.get("username", "")
    password = request.form.get("password", "")
    for u in _cargar_usuarios():
        if u.get("username") == username and _verificar_password(password, u.get("password", "")):
            token = _crear_token(u["username"], u["role"], u.get("nombre", ""))
            return jsonify({
                "access_token": token, "token_type": "bearer",
                "username": u["username"], "role": u["role"], "nombre": u.get("nombre", ""),
            })
    return jsonify({"detail": "Usuario o contraseña incorrectos."}), 401


@app.route("/api/auth/me")
@require_auth
def me():
    return jsonify(request.usuario)


@app.route("/api/auth/usuarios", methods=["GET"])
@require_admin
def get_usuarios():
    return jsonify([
        {"username": u["username"], "role": u["role"], "nombre": u.get("nombre", "")}
        for u in _cargar_usuarios()
    ])


@app.route("/api/auth/usuarios", methods=["POST"])
@require_admin
def post_usuario():
    body = request.get_json()
    username = body.get("username", "")
    users = _cargar_usuarios()
    if any(u["username"] == username for u in users):
        return jsonify({"detail": f"El usuario '{username}' ya existe."}), 400
    users.append({
        "username": username,
        "password": _hashear_password(body.get("password", "")),
        "role": body.get("role", "auditor"),
        "nombre": body.get("nombre", ""),
    })
    _guardar_usuarios(users)
    return jsonify({"username": username, "role": body.get("role","auditor"), "nombre": body.get("nombre","")})


@app.route("/api/auth/usuarios/<username>", methods=["DELETE"])
@require_admin
def delete_usuario(username):
    users = _cargar_usuarios()
    nuevos = [u for u in users if u["username"] != username]
    if len(nuevos) == len(users):
        return jsonify({"detail": f"Usuario '{username}' no encontrado."}), 404
    if not any(u["role"] == "admin" for u in nuevos):
        return jsonify({"detail": "No se puede eliminar el último administrador."}), 400
    _guardar_usuarios(nuevos)
    return jsonify({"ok": True})


@app.route("/api/auth/password", methods=["PUT"])
@require_admin
def put_password():
    body = request.get_json()
    username = body.get("username", "")
    users = _cargar_usuarios()
    for u in users:
        if u["username"] == username:
            u["password"] = _hashear_password(body.get("nueva_password", ""))
            _guardar_usuarios(users)
            return jsonify({"ok": True})
    return jsonify({"detail": f"Usuario '{username}' no encontrado."}), 404


# ── Validación ────────────────────────────────────────────────

@app.route("/api/procesar", methods=["POST"])
@require_auth
def procesar():
    import sys, traceback
    try:
        return _procesar_interno()
    except Exception as exc:
        tb = traceback.format_exc()
        print(f"[ERROR /api/procesar] {exc}\n{tb}", file=sys.stderr, flush=True)
        return jsonify({"detail": f"Error interno: {exc}"}), 500


def _procesar_interno():
    registros, alertas = [], []
    validaciones_malla, validaciones_general, validaciones_auditoria, validaciones_pertinencia = [], [], [], []
    errores_acum = []

    json_wrappers = [
        _FileWrapper(f.filename, f.read())
        for f in request.files.getlist("json_files")
        if f and f.filename and _RIPS_FILENAME_RE.match(os.path.basename(f.filename))
    ]
    excel_wrappers = [
        _FileWrapper(f.filename, f.read())
        for f in request.files.getlist("excel_files") if f and f.filename
    ]

    hay_excel = bool(excel_wrappers)
    registros_excel, set_aut_excel = {}, set()
    if hay_excel:
        try:
            registros_excel, set_aut_excel, errs = cargar_excel_autorizaciones(excel_wrappers)
            errores_acum.extend(errs)
        except Exception as exc:
            errores_acum.append(f"Error cargando Excel de autorizaciones: {exc}")

    archivos_procesados = 0
    total_rips = 0
    pacientes_rips_global = {}

    for wrapper in json_wrappers:
        try:
            wrapper.seek(0)
            data = json.loads(wrapper.read())
            registros.extend(extraer_medicamentos_invalidos(data, wrapper.filename))
            pacientes_rips_global.update(extraer_autorizaciones_rips(data, wrapper.filename))
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
            validaciones_auditoria.extend(validar_concepto_recaudo(data, wrapper.filename))
            validaciones_pertinencia.extend(validar_pertinencia(data, wrapper.filename))
            archivos_procesados += 1
        except Exception as exc:
            errores_acum.append(f"Error en {wrapper.filename}: {exc}")

    if hay_excel and (registros_excel or set_aut_excel):
        try:
            alertas = validar_autorizaciones(pacientes_rips_global, registros_excel, set_aut_excel)
        except Exception as exc:
            errores_acum.append(f"Error en cruce de autorizaciones: {exc}")

    t_proc = datetime.now()
    total_auths_rips = sum(len(p.get("set_auths", set())) for p in pacientes_rips_global.values())

    stats = {
        "archivos_json": archivos_procesados, "total_rips": total_rips,
        "alerta_volumen": total_rips > 800, "med_invalidos": len(registros),
        "auts_rips": total_auths_rips, "auts_excel": len(set_aut_excel),
        "tipo_mismatch":         _alen(alertas, "tipo_doc_mismatch"),
        "amb_par_nc":            _alen(alertas, "amb_par_no_cruza"),
        "hosp_proc_nc":          _alen(alertas, "hosp_proc_no_cruza"),
        "hosp_aut_no_rel":       _alen(alertas, "hosp_aut_no_relacionada"),
        "estancia_sin_aut":      _alen(alertas, "estancia_sin_aut"),
        "proc_qx_aut_hosp":      _alen(alertas, "proc_qx_misma_aut_hosp"),
        "sin_aut_rel":           _alen(alertas, "sin_num_aut_relacionado"),
        "amb_emision_post":      _alen(alertas, "amb_aut_emision_posterior"),
        "hosp_cod_sin_aut":      _alen(alertas, "hosp_cod_sin_aut"),
        "hosp_cups_duplicado":   _alen(alertas, "hosp_cups_duplicado"),
        "proc_sin_aut_amb":      _alen(alertas, "proc_sin_aut_amb"),
        "proc_aut_no_cruza":     _alen(alertas, "proc_aut_no_cruza_amb"),
        "cups_noestandar_nc":    _alen(alertas, "cups_noestandar_sin_aut"),
        "hosp_proc_cod_no_cruza":_alen(alertas, "hosp_proc_cod_no_cruza"),
        "proc_sin_aut_no_excel":        _alen(alertas, "proc_sin_aut_no_excel"),
        "internacion_sin_aut_no_excel": _alen(alertas, "internacion_sin_aut_no_excel"),
        "internacion_aut_es_cedula":    _alen(alertas, "internacion_aut_es_cedula"),
        "malla_total":        len(validaciones_malla),
        "malla_criticas":     _count_sev(validaciones_malla, "critica"),
        "malla_notificaciones": sum(1 for v in validaciones_malla if v.get("severidad") in {"media","alta"}),
        "malla_top_reglas":   _top_reglas(validaciones_malla),
        "general_total":      len(validaciones_general),
        "general_criticas":   _count_sev(validaciones_general, "critica"),
        "general_notificaciones": sum(1 for v in validaciones_general if v.get("severidad") in {"media","alta"}),
        "general_top_reglas": _top_reglas(validaciones_general),
        "auditoria_total":    len(validaciones_auditoria),
        "auditoria_criticas": _count_sev(validaciones_auditoria, "critica"),
        "auditoria_notificaciones": sum(1 for v in validaciones_auditoria if v.get("severidad") in {"media","alta"}),
        "auditoria_top_reglas": _top_reglas(validaciones_auditoria),
        "pertinencia_total":   len(validaciones_pertinencia),
        "pertinencia_top_reglas": _top_reglas(validaciones_pertinencia),
        "tiempo_procesamiento": f"{(datetime.now() - t_proc).total_seconds():.1f}s",
        "errores_procesamiento": errores_acum,
    }

    _guardar_estadistica(stats, request.usuario.get("username", "desconocido"))
    _cache["ultimo"] = {
        "registros": registros, "alertas": alertas,
        "validaciones_malla": validaciones_malla,
        "validaciones_general": validaciones_general,
        "validaciones_auditoria": validaciones_auditoria,
        "validaciones_pertinencia": validaciones_pertinencia,
        "stats": stats,
    }

    _MAX_TABLA = 5000
    return app.response_class(
        response=json.dumps({
            "stats": stats,
            "registros": registros,
            "alertas": alertas,
            # malla y general NO se renderizan en tablas del UI — solo sus stats.
            # Se omiten del response para evitar respuestas de 30-80 MB con >300 archivos.
            # El Excel de exportación usa el _cache y tiene los datos completos.
            "validaciones_malla": [],
            "validaciones_general": [],
            "validaciones_auditoria":   validaciones_auditoria[:_MAX_TABLA],
            "validaciones_pertinencia": validaciones_pertinencia[:_MAX_TABLA],
        }, default=_json_safe),
        status=200,
        mimetype="application/json",
    )


@app.route("/api/exportar")
@require_auth
def exportar():
    cached = _cache.get("ultimo", {})
    if not cached.get("stats"):
        return jsonify({"error": "No hay resultados. Procese los archivos RIPS primero."}), 400
    output = construir_excel(
        cached["registros"], cached["alertas"],
        cached["validaciones_malla"], cached["validaciones_general"],
        cached.get("validaciones_auditoria"),
        cached.get("validaciones_pertinencia"),
    )
    nombre = f"Alertas_Malla_Validadora_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    return send_file(output, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/estadisticas")
@require_admin
def get_estadisticas():
    limite = request.args.get("limite", 100, type=int)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    filas = con.execute("SELECT * FROM estadisticas ORDER BY id DESC LIMIT ?", (limite,)).fetchall()
    con.close()
    resultado = []
    for f in filas:
        row = dict(f)
        row["top_reglas_malla"]     = json.loads(row.get("top_reglas_malla") or "[]")
        row["top_reglas_general"]   = json.loads(row.get("top_reglas_general") or "[]")
        row["top_reglas_auditoria"] = json.loads(row.get("top_reglas_auditoria") or "[]")
        resultado.append(row)
    return jsonify({"total": len(resultado), "registros": resultado})
