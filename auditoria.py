"""
auditoria.py
Módulo de auditoría clínica para RIPS JSON (Resolución 2275/2023).

Categorías de reglas:
  - Estancia      : cálculo y validación de días de hospitalización
  - Transfusional : coherencia entre procesamientos y transfusiones
  - Urgencias     : facturación correcta de la consulta 890701
  - Laboratorio   : paquetes de exámenes complementarios requeridos
"""

from datetime import datetime, time as dtime


# ══════════════════════════════════════════════════════════════
# HELPERS INTERNOS
# ══════════════════════════════════════════════════════════════

def _nstr(v):
    """Convierte cualquier valor a string limpio."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _num_doc_usuario(usuario):
    return _nstr(
        usuario.get("numDocumentoIdentificacion") or
        usuario.get("numDocumentoldentificacion") or ""
    )


def _parse_dt(s):
    s = _nstr(s)
    if len(s) >= 16:
        try:
            return datetime.strptime(s[:16], "%Y-%m-%d %H:%M")
        except ValueError:
            pass
    return None


def _get_factura(data):
    for k in ("numFactura", "numeroFactura", "nroFactura", "noFactura", "factura"):
        v = data.get(k)
        if v and not isinstance(v, (dict, list)):
            return _nstr(v)
    return ""


def _calcular_dias_auditoria(f_ini_dt, f_egr_dt):
    """
    Días de estancia facturables según lógica de auditoría:
      - Cada día calendario completo (00:00–23:59) cuenta como un día.
      - El día de INGRESO cuenta siempre, SALVO que el paciente haya ingresado
        dentro de las 3 horas previas a medianoche (hora >= 21:00), en cuyo
        caso ese día NO se reconoce como día de estancia.
      - El día de EGRESO NUNCA se cuenta, independientemente de la hora de salida.

    Fórmula:
        dias = (fecha_egreso - fecha_ingreso).days          ← egreso excluido por defecto
        si hora_ingreso >= 21:00  →  dias -= 1             ← ingreso tardío no cuenta
        resultado = max(0, dias)

    Ejemplo: ingreso 01/01/2026 10:00, egreso 04/01/2026 15:00
        (04 - 01) = 3 días; 10:00 < 21:00 → no se resta → 3 días  ✓

    Ejemplo: ingreso 01/01/2026 22:00, egreso 04/01/2026 15:00
        (04 - 01) = 3 días; 22:00 >= 21:00 → restar 1 → 2 días  ✓
    """
    dias = (f_egr_dt.date() - f_ini_dt.date()).days    # egreso excluido automáticamente
    if f_ini_dt.time() >= dtime(21, 0):
        dias -= 1                                       # ingreso tardío (< 3h para medianoche)
    return max(0, dias)


def _horas(f_ini_dt, f_egr_dt):
    return (f_egr_dt - f_ini_dt).total_seconds() / 3600


def _n_proc(lista_codigos, codigo):
    """Cuenta cuántas veces aparece el código en la lista."""
    return lista_codigos.count(codigo)


# ══════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════

def validar_auditoria(data, nombre_archivo=""):
    """
    Aplica todas las reglas de auditoría clínica sobre el JSON RIPS.

    Retorna lista de dicts con la misma estructura que los validadores
    de la malla 2275:
      archivo, num_factura, num_doc, consecutivo,
      id_regla, severidad, campo, mensaje, valor_actual
    """
    errores = []
    if not isinstance(data, dict):
        return errores

    num_factura = _get_factura(data)

    for usuario in data.get("usuarios", []):
        if not isinstance(usuario, dict):
            continue

        nd        = _num_doc_usuario(usuario)
        servicios = usuario.get("servicios", {})
        if not isinstance(servicios, dict):
            continue

        procs = servicios.get("procedimientos", []) or []
        otros = servicios.get("otrosServicios",  []) or []
        hosps = servicios.get("hospitalizacion", []) or []
        urgs  = servicios.get("urgencias",       []) or []
        cons  = servicios.get("consultas",       []) or []

        ctx = {
            "archivo":     nombre_archivo,
            "num_factura": num_factura,
            "num_doc":     nd,
            "consecutivo": "",
        }

        def _e(id_regla, severidad, campo, mensaje, valor_actual=""):
            errores.append({
                **ctx,
                "id_regla":     id_regla,
                "severidad":    severidad,
                "campo":        campo,
                "mensaje":      mensaje,
                "valor_actual": _nstr(valor_actual),
            })

        # ── Fechas de la estancia (hospitalización o urgencias) ──────────────
        f_ini_dt = f_egr_dt = None
        for src in hosps + urgs:
            if not isinstance(src, dict):
                continue
            fi = _parse_dt(src.get("fechaInicioAtencion") or src.get("fechaInicio") or "")
            fe = _parse_dt(src.get("fechaEgreso") or "")
            if fi and not f_ini_dt:
                f_ini_dt = fi
            if fe and not f_egr_dt:
                f_egr_dt = fe

        horas_est       = None
        dias_calculados = None
        if f_ini_dt and f_egr_dt and f_egr_dt >= f_ini_dt:
            horas_est       = _horas(f_ini_dt, f_egr_dt)
            dias_calculados = _calcular_dias_auditoria(f_ini_dt, f_egr_dt)

        # ── Listas de códigos ────────────────────────────────────────────────
        cod_procs = [_nstr(r.get("codProcedimiento")) for r in procs if isinstance(r, dict)]
        cod_cons  = [_nstr(r.get("codConsulta"))      for r in cons  if isinstance(r, dict)]
        cod_otros = [_nstr(r.get("codTecnologiaSalud")) for r in otros if isinstance(r, dict)]

        def n_proc(cod):
            return _n_proc(cod_procs, cod)

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE: ESTANCIA
        # ════════════════════════════════════════════════════════════════════

        if horas_est is not None:

            # ── EST-24H-01: Estancia < 24h sin observación ni hospitalización ──
            if horas_est < 24:
                tiene_obs  = ("5DSM01" in cod_otros or "5DSM01" in cod_procs)
                tiene_hosp = bool(hosps)

                if not tiene_obs and not tiene_hosp:
                    _e("EST-24H-01", "alta",
                       "codTecnologiaSalud / codProcedimiento / hospitalizacion",
                       f"Estancia menor a 24 horas ({horas_est:.1f}h): no se encontró el código "
                       "de observación (5DSM01) ni ningún registro de hospitalización. "
                       "Debe facturarse al menos uno de los dos.")

            # ── EST-DIAS-01/02: Estancia >= 24h — validar días facturados ─────
            else:
                # Contar líneas de estancia en otrosServicios (tipoOS = "03")
                total_fac = 0
                for rec in otros:
                    if not isinstance(rec, dict):
                        continue
                    if _nstr(rec.get("tipoOS")) != "03":
                        continue
                    qty_raw = rec.get("cantidadOS") or rec.get("cantidad") or 1
                    try:
                        total_fac += int(str(qty_raw).strip())
                    except (ValueError, TypeError):
                        total_fac += 1

                if total_fac == 0:
                    _e("EST-DIAS-01", "alta", "otrosServicios (tipoOS=03)",
                       f"Estancia de {horas_est:.1f}h (≥24h): no se encontraron registros de "
                       f"estancia (tipoOS=03) en otrosServicios. "
                       f"Días calculados por auditoría: {dias_calculados}. "
                       f"Ingreso: {f_ini_dt.strftime('%Y-%m-%d %H:%M')}, "
                       f"Egreso: {f_egr_dt.strftime('%Y-%m-%d %H:%M')}.")
                elif total_fac != dias_calculados:
                    nota_ingreso = (
                        " (hora de ingreso >= 21:00: día de ingreso no reconocido)"
                        if f_ini_dt.time() >= dtime(21, 0) else ""
                    )
                    _e("EST-DIAS-02", "media", "otrosServicios (tipoOS=03)",
                       f"Días de estancia facturados ({total_fac}) no coinciden con los días "
                       f"calculados por auditoría ({dias_calculados}). "
                       f"Ingreso: {f_ini_dt.strftime('%Y-%m-%d %H:%M')}{nota_ingreso}, "
                       f"Egreso: {f_egr_dt.strftime('%Y-%m-%d %H:%M')} "
                       f"({horas_est:.1f}h). "
                       "Regla: cada día calendario cuenta; si el ingreso es >= 21:00 ese día "
                       "no se reconoce; el día de egreso nunca se cuenta.",
                       total_fac)

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE: TRANSFUSIONAL
        # ════════════════════════════════════════════════════════════════════

        hay_transfusion_gr    = n_proc("912002") > 0   # Transfusión GR
        hay_transfusion_pfc   = n_proc("912005") > 0   # Transfusión PFC
        hay_transfusion_crio  = n_proc("912001") > 0   # Transfusión crioprecipitado
        hay_transfusion_plaq  = n_proc("912003") > 0   # Transfusión plaquetas
        hay_cualquier_transf  = (hay_transfusion_gr or hay_transfusion_pfc or
                                 hay_transfusion_crio or hay_transfusion_plaq)

        hay_proc_pfc   = n_proc("911111") > 0   # Procesamiento PFC
        hay_proc_crio  = n_proc("911105") > 0   # Procesamiento crioprecipitado
        hay_proc_plaq  = n_proc("911102") > 0   # Procesamiento plaquetas estándar
        hay_proc_plaq2 = n_proc("911201") > 0   # Procesamiento plaquetas por aféresis
        hay_proc_gr    = n_proc("911116") > 0   # Procesamiento GR pobre en leucocitos
        hay_proc_ali   = n_proc("911118") > 0   # Alícuota pediátrica filtrada

        # Triada de hemoclasificación
        abo_d = n_proc("911017")   # ABO directa
        abo_i = n_proc("911019")   # ABO inversa
        rh    = n_proc("911015")   # RH

        # ── TRANS-01/02/03: Hemoclasificaciones máx. 1 vez; obligatorias si hay transfusión ─
        for cod, nombre, conteo in [
            ("911017", "Hemoclasificación ABO directa",  abo_d),
            ("911019", "Hemoclasificación ABO inversa",  abo_i),
            ("911015", "Hemoclasificación RH",           rh),
        ]:
            reg = {"911017": "TRANS-01", "911019": "TRANS-02", "911015": "TRANS-03"}[cod]
            if conteo > 1:
                _e(reg, "alta", "codProcedimiento",
                   f"{nombre} ({cod}) facturada {conteo} veces. "
                   "Solo se permite una vez por estancia.", conteo)
            if hay_cualquier_transf and conteo == 0:
                _e(reg, "alta", "codProcedimiento",
                   f"Se facturó una transfusión pero no se encontró {nombre} ({cod}). "
                   "Es obligatorio facturarla cuando se realiza transfusión.")

        # ── TRANS-04: Fenotipo ABO 911025 ─────────────────────────────────────
        # Siempre genera alerta informativa (solo aplica para grupo A).
        # Además: max 1 vez; si hay transfusión/procesamiento y es grupo A, alertar si falta.
        if n_proc("911025") > 0:
            if n_proc("911025") > 1:
                _e("TRANS-04", "alta", "codProcedimiento",
                   f"Fenotipo ABO 911025 facturado {n_proc('911025')} veces. "
                   "Solo se permite una vez por estancia.")
            _e("TRANS-04", "media", "codProcedimiento",
               "Fenotipo ABO subgrupos (911025) facturado: este código solo aplica para "
               "pacientes con grupo sanguíneo A. Verificar que el grupo del paciente sea A.")
        elif hay_cualquier_transf or hay_proc_gr or hay_proc_pfc or hay_proc_crio or hay_proc_plaq or hay_proc_plaq2:
            _e("TRANS-04", "media", "codProcedimiento",
               "Se registró transfusión o procesamiento: si el grupo sanguíneo del paciente "
               "es A, verifique si corresponde facturar Fenotipo ABO subgrupos (911025).")

        # ── TRANS-05: Hemoglobina (902213) y Hematocrito (902211) ─────────────
        # Solo aplica cuando existe al menos una transfusión.
        # Obligatorio mínimo 2 veces por estancia a menos que exista 902210 (hemograma completo).
        hay_hemograma = n_proc("902210") > 0
        if hay_cualquier_transf and not hay_hemograma:
            if n_proc("902213") < 2:
                _e("TRANS-05", "alta", "codProcedimiento",
                   f"Hemoglobina (902213) facturada {n_proc('902213')} vez/veces. "
                   "Ante la presencia de transfusión, debe facturarse al menos 2 veces por estancia "
                   "cuando no existe hemograma completo (902210).", n_proc("902213"))
            if n_proc("902211") < 2:
                _e("TRANS-05", "alta", "codProcedimiento",
                   f"Hematocrito (902211) facturado {n_proc('902211')} vez/veces. "
                   "Ante la presencia de transfusión, debe facturarse al menos 2 veces por estancia "
                   "cuando no existe hemograma completo (902210).", n_proc("902211"))

        # ── TRANS-06: Anticuerpos irregulares 911003 — solo si fue facturado y estancia > 72h ─
        if n_proc("911003") > 0 and horas_est is not None and horas_est > 72:
            _e("TRANS-06", "media", "codProcedimiento",
               f"Anticuerpos irregulares (911003) facturado con estancia de {horas_est:.1f}h "
               "(>72h). Revisar vigencia del examen y si corresponde refacturación.")

        # ── TRANS-07: Procesamiento GR pobre en leucocitos 911116 ─────────────
        # No puede ser menor al número de unidades transfundidas de GR (912002).
        # Es obligatorio cuando está ordenado el procesamiento.
        if hay_transfusion_gr:
            if not hay_proc_gr:
                _e("TRANS-07", "media", "codProcedimiento",
                   f"Se facturaron {n_proc('912002')} transfusión(es) de GR (912002) pero no se "
                   "encontró procesamiento de GR pobre en leucocitos (911116). Verificar en historia "
                   "clínica si fue ordenado pobre en leucocitos y facturar si corresponde.")
            elif n_proc("911116") < n_proc("912002"):
                _e("TRANS-07", "alta", "codProcedimiento",
                   f"Procesamiento GR (911116): {n_proc('911116')} es menor al número de "
                   f"transfusiones GR (912002): {n_proc('912002')}. "
                   "El procesamiento no puede ser menor a las unidades transfundidas.",
                   f"911116={n_proc('911116')}, 912002={n_proc('912002')}")

        # ── TRANS-08: Prueba cruzada 911021 — obligatoria cuando hay 912002 ────
        if hay_transfusion_gr and n_proc("911021") == 0:
            _e("TRANS-08", "alta", "codProcedimiento",
               f"Se facturaron {n_proc('912002')} transfusión(es) de GR (912002) sin prueba "
               "cruzada mayor (911021). Facturar prueba cruzada por cada unidad transfundida.")

        # ── TRANS-09: 912002 menor que procesamientos — solicitar triada ────────
        # Si las transfusiones son menos que los procesamientos, verificar triada hemoclasif.
        if hay_transfusion_gr and n_proc("912002") < n_proc("911116"):
            faltantes_triada = []
            if abo_d == 0: faltantes_triada.append("ABO directa (911017)")
            if abo_i == 0: faltantes_triada.append("ABO inversa (911019)")
            if rh    == 0: faltantes_triada.append("RH (911015)")
            if n_proc("911003") == 0: faltantes_triada.append("Anticuerpos irregulares (911003)")
            if faltantes_triada:
                _e("TRANS-09", "alta", "codProcedimiento",
                   f"Transfusiones GR (912002): {n_proc('912002')} < "
                   f"Procesamientos (911116): {n_proc('911116')}. "
                   f"Verificar y facturar las pruebas faltantes: {', '.join(faltantes_triada)}.",
                   f"912002={n_proc('912002')}, 911116={n_proc('911116')}")

        # ── TRANS-10: Alícuota pediátrica 911118 — no facturable > 5 años ──────
        if hay_proc_ali:
            fecha_nac_raw = _nstr(usuario.get("fechaNacimiento") or "")
            if len(fecha_nac_raw) == 10:
                try:
                    fn_dt = datetime.strptime(fecha_nac_raw, "%Y-%m-%d")
                    edad_años = (f_ini_dt or datetime.now() - fn_dt).days // 365 if f_ini_dt else (datetime.now() - fn_dt).days // 365
                    if f_ini_dt:
                        edad_años = (f_ini_dt.date() - fn_dt.date()).days // 365
                    else:
                        edad_años = (datetime.now().date() - fn_dt.date()).days // 365
                    if edad_años > 5:
                        _e("TRANS-10", "alta", "codProcedimiento",
                           f"Procesamiento alícuota pediátrica filtrada (911118) facturado para "
                           f"paciente de {edad_años} años. Este código no es facturable en "
                           "pacientes mayores de 5 años.", edad_años)
                except ValueError:
                    pass

        # ── TRANS-11/12/13/14: Triada obligatoria para procesamientos y transfusiones ──
        # Cuando exista cualquiera de estos códigos, deben estar los 3 de hemoclasificación.
        REQUIEREN_TRIADA = {
            "911111": "Procesamiento PFC (911111)",
            "912005": "Transfusión PFC (912005)",
            "911105": "Procesamiento crioprecipitado (911105)",
            "912001": "Transfusión crioprecipitado (912001)",
            "911102": "Procesamiento plaquetas estándar (911102)",
            "912003": "Transfusión plaquetas (912003)",
            "911201": "Procesamiento plaquetas aféresis (911201)",
        }
        TRANS_TRIADA_IDS = {
            "911111": "TRANS-11", "912005": "TRANS-11",
            "911105": "TRANS-12", "912001": "TRANS-12",
            "911102": "TRANS-13", "912003": "TRANS-13",
            "911201": "TRANS-14",
        }
        for cod, nombre in REQUIEREN_TRIADA.items():
            if n_proc(cod) == 0:
                continue
            faltantes_t = []
            if abo_d == 0: faltantes_t.append("ABO directa (911017)")
            if abo_i == 0: faltantes_t.append("ABO inversa (911019)")
            if rh    == 0: faltantes_t.append("RH (911015)")
            if faltantes_t:
                _e(TRANS_TRIADA_IDS.get(cod, "TRANS-11"), "alta", "codProcedimiento",
                   f"Se facturó {nombre} sin la triada completa de hemoclasificación. "
                   f"Faltan: {', '.join(faltantes_t)}.",
                   cod)

        # ── TRANS-15: Procesamientos no pueden ser menores a las unidades transfundidas ──
        PROC_VS_TRANSF = [
            ("911118", "912002", "Alícuota pediátrica (911118)",       "Transfusión GR (912002)"),
            ("911111", "912005", "Procesamiento PFC (911111)",         "Transfusión PFC (912005)"),
            ("911105", "912001", "Procesamiento crioprecipitado (911105)", "Transfusión crioprecipitado (912001)"),
            ("911102", "912003", "Procesamiento plaquetas estándar (911102)", "Transfusión plaquetas (912003)"),
            ("911201", "912003", "Procesamiento plaquetas aféresis (911201)", "Transfusión plaquetas (912003)"),
        ]
        for cod_proc_t, cod_transf, nom_proc, nom_transf in PROC_VS_TRANSF:
            np = n_proc(cod_proc_t)
            nt = n_proc(cod_transf)
            if np > 0 and np < nt:
                _e("TRANS-15", "alta", "codProcedimiento",
                   f"{nom_proc}: {np} es menor al número de {nom_transf}: {nt}. "
                   "El procesamiento no puede ser inferior a las unidades transfundidas.",
                   f"{cod_proc_t}={np}, {cod_transf}={nt}")

        # ── TRANS-16: Cantidad > 1 en una misma línea para códigos únicos ──────
        CODIGOS_UNICOS_TRANS = {
            "911017": "ABO directa",
            "911019": "ABO inversa",
            "911015": "RH",
        }
        for proc in procs:
            if not isinstance(proc, dict):
                continue
            cod = _nstr(proc.get("codProcedimiento"))
            if cod not in CODIGOS_UNICOS_TRANS:
                continue
            qty_raw = proc.get("cantidadOS") or proc.get("cantidad")
            if qty_raw is None:
                continue
            try:
                if int(str(qty_raw).strip()) > 1:
                    _e("TRANS-16", "alta", "codProcedimiento",
                       f"Código {cod} ({CODIGOS_UNICOS_TRANS[cod]}) reportado con "
                       f"cantidad {qty_raw} en una misma línea. "
                       "Solo se permite cantidad 1 por registro.", cod)
            except (ValueError, TypeError):
                pass

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE: URGENCIAS
        # ════════════════════════════════════════════════════════════════════

        total_890701 = n_proc("890701") + cod_cons.count("890701")

        # ── URG-01: 890701 facturado más de una vez ───────────────────────────
        if total_890701 > 1:
            _e("URG-01", "alta", "codProcedimiento / codConsulta",
               f"Consulta de urgencias (890701) facturada {total_890701} veces. "
               "Solo se permite una vez por estancia.", total_890701)

        # ── URG-02: Urgencias sin consulta 890701 ─────────────────────────────
        if urgs and total_890701 == 0:
            _e("URG-02", "media", "codProcedimiento / codConsulta",
               "Existen registros de urgencias pero no se encontró la consulta de urgencias "
               "(890701). Verificar si aplica su facturación.")

        # ════════════════════════════════════════════════════════════════════
        # BLOQUE: LABORATORIO
        # ════════════════════════════════════════════════════════════════════

        # ── LAB-CALCIO-01: 903839 requiere electrolitos y gases ───────────────
        # Si aparece calcio (903839) deben estar: 903864, 903859, 903813, 903810
        if n_proc("903839") > 0:
            COMPLEMENTARIOS = {
                "903864": "Gases arteriales (903864)",
                "903859": "Magnesio (903859)",
                "903813": "Fósforo (903813)",
                "903810": "Calcio ionizado (903810)",
            }
            faltantes = [nombre for cod, nombre in COMPLEMENTARIOS.items() if n_proc(cod) == 0]
            if faltantes:
                _e("LAB-CALCIO-01", "alta", "codProcedimiento",
                   f"Se facturó calcio (903839) sin los complementarios requeridos: "
                   f"{', '.join(faltantes)}. Revisar y facturar los códigos faltantes.",
                   ", ".join(
                       cod for cod in COMPLEMENTARIOS if n_proc(cod) == 0
                   ))

    return errores
