"""
pertinencia.py
Módulo de observaciones de pertinencia clínica para RIPS JSON (Res. 2275/2023).
Valida la coherencia entre diagnósticos CIE-10 y procedimientos CUPS reportados.
Severidad: "observacion" (color amarillo en la interfaz).
"""


# ── Helpers ──────────────────────────────────────────────────────────────────

def _nstr(v):
    return str(v).strip() if v is not None else ""


def _get_factura(data):
    for k in ("numFactura", "numeroFactura", "num_factura"):
        v = data.get(k)
        if v and not isinstance(v, (dict, list)):
            return _nstr(v)
    return ""


def _num_doc_usuario(u):
    for k in ("numDocumentoIdentificacion", "numDocumento", "num_doc"):
        v = u.get(k)
        if v and not isinstance(v, (dict, list)):
            return _nstr(v)
    return ""


def _cie_starts(cod, *prefixes):
    """True si el código CIE-10 empieza con alguno de los prefijos (sin punto)."""
    c = cod.replace(".", "").upper()
    return any(c.startswith(p.upper()) for p in prefixes)


def _tiene_cups(cod_procs, *prefixes):
    """True si algún CUPS en la lista empieza con alguno de los prefijos."""
    for cp in cod_procs:
        for p in prefixes:
            if cp.startswith(p):
                return True
    return False


def _tiene_cups_exacto(cod_procs, *codigos):
    return any(c in cod_procs for c in codigos)


def _cie_en_rango(cod, letra, num_ini, num_fin):
    """True si el código CIE-10 está en el rango letraNNN–letraNNN (sin punto)."""
    c = cod.replace(".", "").upper()
    if not c.startswith(letra.upper()):
        return False
    try:
        num = int(c[1:3])
        return num_ini <= num <= num_fin
    except (ValueError, IndexError):
        return False


# ── Recolector de diagnósticos de un usuario ─────────────────────────────────

def _diagnosticos_usuario(servicios):
    """Recoge todos los CIE-10 de consultas, urgencias, hospitalización y procedimientos."""
    diags = set()
    for seccion in ("consultas", "urgencias", "hospitalizacion"):
        for item in servicios.get(seccion, []) or []:
            if not isinstance(item, dict):
                continue
            for campo in (
                "codDiagnosticoPrincipal",
                "codDiagnosticoRelacionado1",
                "codDiagnosticoRelacionado2",
                "codDiagnosticoRelacionado3",
                "codDiagnosticoEgreso1",
                "codDiagnosticoEgreso2",
                "codDiagnosticoEgreso3",
            ):
                v = _nstr(item.get(campo))
                if v:
                    diags.add(v.replace(".", "").upper())
    return diags


# ══════════════════════════════════════════════════════════════════════════════
# REGLAS DE PERTINENCIA
# Cada regla devuelve (True, cups_encontrados_str) si hay alerta,
# o (False, "") si la pertinencia es adecuada.
# Parámetros: diags (set CIE-10 sin punto), cod_procs (list CUPS str)
# ══════════════════════════════════════════════════════════════════════════════

REGLAS = []


def _reg(id_regla, grupo, cie_desc, cups_esperados_desc, fn_aplica, fn_ok):
    """Registra una regla de pertinencia."""
    REGLAS.append({
        "id": id_regla,
        "grupo": grupo,
        "cie_desc": cie_desc,
        "cups_esperados": cups_esperados_desc,
        "aplica": fn_aplica,
        "ok": fn_ok,
    })


# ── TRAUMA CRANEOENCEFÁLICO ───────────────────────────────────────────────────
_reg(
    "PERT-TRA-01", "Trauma craneoencefálico",
    "S06.x (lesiones intracraneales)",
    "TAC/RMN cerebral (893101/893200) y hemograma (903801)",
    lambda d, p: any(_cie_starts(c, "S06") for c in d),
    lambda d, p: (
        _tiene_cups(p, "893101", "893100", "893200") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── TRAUMA FACIAL ─────────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-02", "Trauma facial",
    "S02.x (fractura huesos cráneo/cara)",
    "Radiografía cráneo/facial (895200/895201) o TAC facial (893101)",
    lambda d, p: any(_cie_starts(c, "S02") for c in d),
    lambda d, p: _tiene_cups(p, "895200", "895201", "895202", "893101", "893100"),
)

# ── TRAUMA CERVICAL ───────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-03", "Trauma cervical",
    "S12.x / S13.x (fractura/luxación columna cervical)",
    "Radiografía columna cervical (895210) o TAC columna (893107)",
    lambda d, p: any(_cie_starts(c, "S12", "S13") for c in d),
    lambda d, p: _tiene_cups(p, "895210", "895211", "893107", "893106", "893100"),
)

# ── TRAUMA TORÁCICO ───────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-04", "Trauma torácico",
    "S22.x / S27.x (fractura costillas/lesión órganos torácicos)",
    "Radiografía tórax (895201/895203) o TAC tórax (893103)",
    lambda d, p: any(_cie_starts(c, "S22", "S27") for c in d),
    lambda d, p: _tiene_cups(p, "895201", "895203", "893103", "893100"),
)

# ── TRAUMA ABDOMINAL ─────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-05", "Trauma abdominal",
    "S36.x / S37.x (lesión órganos abdominales/pélvicos)",
    "Ecografía FAST (881601) o TAC abdominal (893104) y hemograma (903801)",
    lambda d, p: any(_cie_starts(c, "S36", "S37") for c in d),
    lambda d, p: (
        _tiene_cups(p, "881601", "881602", "893104", "893100") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── TRAUMA HOMBRO ─────────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-06", "Trauma hombro",
    "S42.x / S43.x (fractura/luxación hombro/clavícula)",
    "Radiografía hombro/clavícula (895221/895222)",
    lambda d, p: any(_cie_starts(c, "S42", "S43") for c in d),
    lambda d, p: _tiene_cups(p, "895221", "895222", "895220", "895200"),
)

# ── TRAUMA CODO/ANTEBRAZO ────────────────────────────────────────────────────
_reg(
    "PERT-TRA-07", "Trauma codo/antebrazo",
    "S52.x / S53.x (fractura/luxación codo y antebrazo)",
    "Radiografía codo/antebrazo (895224/895225)",
    lambda d, p: any(_cie_starts(c, "S52", "S53") for c in d),
    lambda d, p: _tiene_cups(p, "895224", "895225", "895223", "895200"),
)

# ── TRAUMA MUÑECA/MANO ───────────────────────────────────────────────────────
_reg(
    "PERT-TRA-08", "Trauma muñeca/mano",
    "S62.x / S63.x / S64.x (fractura/luxación muñeca y mano)",
    "Radiografía muñeca/mano (895226/895227)",
    lambda d, p: any(_cie_starts(c, "S62", "S63", "S64") for c in d),
    lambda d, p: _tiene_cups(p, "895226", "895227", "895228", "895200"),
)

# ── TRAUMA CADERA/FÉMUR ──────────────────────────────────────────────────────
_reg(
    "PERT-TRA-09", "Trauma cadera/fémur",
    "S72.x (fractura fémur/cadera)",
    "Radiografía cadera/pelvis (895231/895232) y hemograma (903801)",
    lambda d, p: any(_cie_starts(c, "S72") for c in d),
    lambda d, p: (
        _tiene_cups(p, "895231", "895232", "895230", "895200") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── TRAUMA RODILLA/PIERNA ────────────────────────────────────────────────────
_reg(
    "PERT-TRA-10", "Trauma rodilla/pierna",
    "S82.x / S83.x (fractura/luxación rodilla y pierna)",
    "Radiografía rodilla/pierna (895234/895235)",
    lambda d, p: any(_cie_starts(c, "S82", "S83") for c in d),
    lambda d, p: _tiene_cups(p, "895234", "895235", "895233", "895200"),
)

# ── TRAUMA TOBILLO/PIE ───────────────────────────────────────────────────────
_reg(
    "PERT-TRA-11", "Trauma tobillo/pie",
    "S92.x / S93.x / S99.x (fractura/luxación tobillo y pie)",
    "Radiografía tobillo/pie (895237/895238)",
    lambda d, p: any(_cie_starts(c, "S92", "S93", "S99") for c in d),
    lambda d, p: _tiene_cups(p, "895237", "895238", "895236", "895200"),
)

# ── POLITRAUMATISMO ──────────────────────────────────────────────────────────
_reg(
    "PERT-TRA-12", "Politraumatismo",
    "T00–T07 (traumatismos múltiples)",
    "TAC cerebral + tórax + abdomen (893100/893103/893104) y hemograma (903801) y tipificación (903901)",
    lambda d, p: any(_cie_en_rango(c, "T", 0, 7) for c in d),
    lambda d, p: (
        _tiene_cups(p, "893100", "893101", "893103", "893104") and
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903901", "903900")
    ),
)

# ── QUEMADURAS ───────────────────────────────────────────────────────────────
_reg(
    "PERT-QUE-01", "Quemaduras",
    "T20–T32 (quemaduras y corrosiones)",
    "Curación quemados + hemograma (903801) + proteínas (903301/903302)",
    lambda d, p: any(_cie_en_rango(c, "T", 20, 32) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903301", "903302", "903300")
    ),
)

# ── CUERPO EXTRAÑO ───────────────────────────────────────────────────────────
_reg(
    "PERT-CEX-01", "Cuerpo extraño",
    "T15–T19 (efectos de cuerpos extraños)",
    "Radiografía o endoscopia según localización (895200/870100)",
    lambda d, p: any(_cie_en_rango(c, "T", 15, 19) for c in d),
    lambda d, p: _tiene_cups(p, "895200", "895201", "870100", "870101", "870200"),
)

# ── HEPATOPATÍA ──────────────────────────────────────────────────────────────
_reg(
    "PERT-HEP-01", "Hepatopatía",
    "K70–K77 (enfermedades hepáticas)",
    "Transaminasas AST/ALT (903701/903702), bilirrubinas (903305) y ecografía hepática (881611/881601)",
    lambda d, p: any(_cie_en_rango(c, "K", 70, 77) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903701", "903702", "903700") and
        _tiene_cups(p, "903305", "903300") and
        _tiene_cups(p, "881601", "881611", "881610")
    ),
)

# ── PATOLOGÍA BILIAR ─────────────────────────────────────────────────────────
_reg(
    "PERT-BIL-01", "Patología biliar",
    "K80–K87 (colelitiasis, colecistitis, colangitis)",
    "Ecografía abdominal (881601) y bilirrubinas (903305) y amilasa/lipasa (904301/904302)",
    lambda d, p: any(_cie_en_rango(c, "K", 80, 87) for c in d),
    lambda d, p: (
        _tiene_cups(p, "881601", "881611") and
        _tiene_cups(p, "903305", "904301", "904302")
    ),
)

# ── PATOLOGÍA ABDOMINAL QUIRÚRGICA ───────────────────────────────────────────
_reg(
    "PERT-ABD-01", "Patología abdominal aguda",
    "K35–K40 (apendicitis, hernia, peritonitis)",
    "Ecografía o TAC abdominal (881601/893104) y hemograma (903801) y PCR (904501)",
    lambda d, p: any(_cie_en_rango(c, "K", 35, 40) for c in d),
    lambda d, p: (
        _tiene_cups(p, "881601", "893104", "893100") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── PATOLOGÍA RENAL ──────────────────────────────────────────────────────────
_reg(
    "PERT-REN-01", "Patología renal aguda/crónica",
    "N10–N23 (pielonefritis, cólico nefrítico, IRA, IRC)",
    "Creatinina (903601) + uroanálisis (903200) + ecografía renal (881601/881605)",
    lambda d, p: any(_cie_en_rango(c, "N", 10, 23) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903601", "903602", "903600") and
        _tiene_cups(p, "903200", "903201")
    ),
)

# ── CARDIOPATÍA ISQUÉMICA ────────────────────────────────────────────────────
_reg(
    "PERT-CAR-01", "Cardiopatía isquémica / SCA",
    "I20–I25 (angina, IAM, cardiopatía isquémica)",
    "ECG (884001) + troponina (904601) + CK-MB (904602) y hemograma (903801)",
    lambda d, p: any(_cie_en_rango(c, "I", 20, 25) for c in d),
    lambda d, p: (
        _tiene_cups(p, "884001") and
        _tiene_cups(p, "904601", "904602", "904600")
    ),
)

# ── OTRAS CARDIOPATÍAS ───────────────────────────────────────────────────────
_reg(
    "PERT-CAR-02", "Insuficiencia cardíaca / arritmias",
    "I26–I52 (ICC, arritmias, pericarditis, endocarditis, embolia pulmonar)",
    "ECG (884001) y Rx tórax (895201) y BNP/pro-BNP o ecocardiograma",
    lambda d, p: any(_cie_en_rango(c, "I", 26, 52) for c in d),
    lambda d, p: (
        _tiene_cups(p, "884001") and
        _tiene_cups(p, "895201", "895203", "895200")
    ),
)

# ── NEUMONÍA / INFECCIÓN RESPIRATORIA ────────────────────────────────────────
_reg(
    "PERT-RES-01", "Infección respiratoria aguda",
    "J12–J18 (neumonía vírica/bacteriana) / J20–J22 (bronquitis/IVAS)",
    "Rx tórax (895201/895203) y hemograma (903801) y PCR (904501)",
    lambda d, p: any(
        _cie_en_rango(c, "J", 12, 22) for c in d
    ),
    lambda d, p: (
        _tiene_cups(p, "895201", "895203", "895200") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── EPOC / ASMA ──────────────────────────────────────────────────────────────
_reg(
    "PERT-RES-02", "EPOC / Asma",
    "J44–J45 (EPOC, asma)",
    "Rx tórax (895201) y espirometría (940200) o gases arteriales (903100)",
    lambda d, p: any(_cie_en_rango(c, "J", 44, 46) for c in d),
    lambda d, p: (
        _tiene_cups(p, "895201", "895203", "895200") and
        _tiene_cups(p, "940200", "903100", "903101")
    ),
)

# ── EPILEPSIA / CONVULSIÓN ───────────────────────────────────────────────────
_reg(
    "PERT-NEU-01", "Epilepsia / Convulsión",
    "G40–G41 (epilepsia, estado epiléptico)",
    "EEG (940100) y TAC/RMN cerebral (893100/893200) y anticonvulsivantes",
    lambda d, p: any(_cie_en_rango(c, "G", 40, 41) for c in d),
    lambda d, p: (
        _tiene_cups(p, "940100", "893100", "893101", "893200")
    ),
)

# ── ECV / ACV ────────────────────────────────────────────────────────────────
_reg(
    "PERT-NEU-02", "ECV / ACV / Hemorragia cerebral",
    "I60–I64 (hemorragia subaracnoidea, intracerebral, ACV isquémico)",
    "TAC cerebral urgente (893100/893101) y hemograma (903801) y coagulación (903401)",
    lambda d, p: any(_cie_en_rango(c, "I", 60, 64) for c in d),
    lambda d, p: (
        _tiene_cups(p, "893100", "893101", "893200") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── CEFALEA / MIGRAÑA ────────────────────────────────────────────────────────
_reg(
    "PERT-NEU-03", "Cefalea / Migraña",
    "G43–G44 / R51 (migraña, cefalea en racimos, cefalea NE)",
    "TAC cerebral (893100) si primera consulta o cefalea severa; EEG (940100) si convulsión asociada",
    lambda d, p: any(_cie_starts(c, "G43", "G44", "R51") for c in d),
    lambda d, p: (
        _tiene_cups(p, "893100", "893101", "893200", "940100") or
        not any(_cie_starts(c, "G43", "G44", "R51") for c in d)
    ),
)

# ── DIARREA INFECCIOSA ───────────────────────────────────────────────────────
_reg(
    "PERT-INF-01", "Diarrea / Gastroenteritis infecciosa",
    "A09 (otras gastroenteritis y colitis infecciosas)",
    "Coprocultivo o cultivo de materia fecal (9051xx) y electrolitos (903401)",
    lambda d, p: any(_cie_starts(c, "A09") for c in d),
    lambda d, p: (
        _tiene_cups(p, "903401", "903400") or
        _tiene_cups(p, "905100", "905101", "905102")
    ),
)

# ── TUBERCULOSIS ─────────────────────────────────────────────────────────────
_reg(
    "PERT-INF-02", "Tuberculosis",
    "A15–A19 (tuberculosis respiratoria y miliar)",
    "Baciloscopia (903600/903614) y Rx tórax (895201) y cultivo BK",
    lambda d, p: any(_cie_en_rango(c, "A", 15, 19) for c in d),
    lambda d, p: (
        _tiene_cups(p, "895201", "895203") and
        _tiene_cups(p, "903614", "903615", "903600")
    ),
)

# ── VIH/SIDA ─────────────────────────────────────────────────────────────────
_reg(
    "PERT-INF-03", "VIH / SIDA",
    "B20 (enfermedad por VIH)",
    "Recuento CD4 (903810) y carga viral VIH (903820) y hemograma (903801)",
    lambda d, p: any(_cie_starts(c, "B20") for c in d),
    lambda d, p: (
        _tiene_cups(p, "903810", "903820", "903800", "903801")
    ),
)

# ── DIABETES DESCOMPENSADA ───────────────────────────────────────────────────
_reg(
    "PERT-MET-01", "Diabetes / Hiperglucemia descompensada",
    "E10–E14 (diabetes mellitus tipo 1 y 2 y complicaciones)",
    "Glucemia (903401) y HbA1c (903403) y creatinina (903601)",
    lambda d, p: any(_cie_en_rango(c, "E", 10, 14) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903401", "903400") and
        _tiene_cups(p, "903601", "903600")
    ),
)

# ── TRASTORNO HIDROELECTROLÍTICO ─────────────────────────────────────────────
_reg(
    "PERT-MET-02", "Trastorno hidroelectrolítico / Deshidratación",
    "E86–E87 (depleción de volumen, trastornos del sodio/potasio)",
    "Ionograma (electrolitos séricos: Na/K/Cl) (903401/903402) y creatinina (903601)",
    lambda d, p: any(_cie_en_rango(c, "E", 86, 87) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903401", "903402", "903400") and
        _tiene_cups(p, "903601", "903600")
    ),
)

# ── ANEMIA ───────────────────────────────────────────────────────────────────
_reg(
    "PERT-HEM-01", "Anemia",
    "D50–D64 (anemias ferropénica, hemolítica, aplásica y otras)",
    "Hemograma (903801) y reticulocitos (903802) y ferritina/hierro sérico (903803)",
    lambda d, p: any(_cie_en_rango(c, "D", 50, 64) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903802", "903803", "903804")
    ),
)

# ── INFECCIÓN GINECOLÓGICA ───────────────────────────────────────────────────
_reg(
    "PERT-GIN-01", "Infección ginecológica / EPI",
    "N70–N77 (salpingitis, vaginitis, endometritis, EPI)",
    "Ecografía pélvica (881602) y cultivo cervical/vaginal y hemograma (903801)",
    lambda d, p: any(_cie_en_rango(c, "N", 70, 77) for c in d),
    lambda d, p: (
        _tiene_cups(p, "881602", "881601") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── COMPLICACIÓN OBSTÉTRICA ──────────────────────────────────────────────────
_reg(
    "PERT-OBS-01", "Complicación obstétrica",
    "O20–O48 (amenaza aborto, hemorragias, preeclampsia, trabajo de parto)",
    "Ecografía obstétrica (881604) y hemograma (903801) y coagulación (903401)",
    lambda d, p: any(_cie_en_rango(c, "O", 20, 48) for c in d),
    lambda d, p: (
        _tiene_cups(p, "881604", "881603", "881602") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# ── LUMBALGIA ────────────────────────────────────────────────────────────────
_reg(
    "PERT-LUM-01", "Lumbalgia / Dorsalgia",
    "M54.x (lumbago, ciática, dorsalgia)",
    "Rx columna lumbosacra (895214) o RMN columna (893202) si crónica/severa",
    lambda d, p: any(_cie_starts(c, "M54") for c in d),
    lambda d, p: _tiene_cups(p, "895214", "895215", "895210", "893202", "893200", "895200"),
)

# ── DOLOR PÉLVICO / CÓLICO ───────────────────────────────────────────────────
_reg(
    "PERT-PEL-01", "Dolor pélvico / Cólico",
    "N94 / R10 (dolor pélvico, cólico menstrual, dolor abdominal)",
    "Ecografía pélvica (881602) y uroanálisis (903200)",
    lambda d, p: any(_cie_starts(c, "N94", "R10") for c in d),
    lambda d, p: (
        _tiene_cups(p, "881602", "881601") and
        _tiene_cups(p, "903200", "903201")
    ),
)

# ── FIEBRE PEDIÁTRICA ────────────────────────────────────────────────────────
_reg(
    "PERT-PED-01", "Fiebre sin foco / Síndrome febril",
    "R50 (fiebre de otro origen y de origen desconocido)",
    "Hemograma (903801) + hemocultivo (905001) + uroanálisis (903200)",
    lambda d, p: any(_cie_starts(c, "R50") for c in d),
    lambda d, p: (
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "905001", "905000", "903200", "903201")
    ),
)

# ── INTOXICACIÓN / SOBREDOSIS ────────────────────────────────────────────────
_reg(
    "PERT-INT-01", "Intoxicación / Sobredosis",
    "T36–T65 (intoxicaciones por fármacos, plaguicidas, sustancias)",
    "Hemograma (903801) + función renal (903601) + función hepática (903701) + niveles tóxicos",
    lambda d, p: any(_cie_en_rango(c, "T", 36, 65) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903601", "903600") and
        _tiene_cups(p, "903701", "903702", "903700")
    ),
)

# ── FRACTURA VERTEBRAL ───────────────────────────────────────────────────────
_reg(
    "PERT-FRX-01", "Fractura vertebral",
    "S32.x / S22.x (fractura columna lumbar/torácica)",
    "Rx columna (895214/895211) y TAC columna (893107) y hemograma (903801)",
    lambda d, p: any(_cie_starts(c, "S32", "S22") for c in d),
    lambda d, p: (
        _tiene_cups(p, "895214", "895211", "895210", "893107", "893106")
    ),
)

# ── AMPUTACIÓN / LACERACIÓN GRAVE ────────────────────────────────────────────
_reg(
    "PERT-AMP-01", "Amputación traumática / Herida compleja",
    "S48.x / S58.x / S68.x / S78.x / S88.x / S98.x (amputaciones traumáticas)",
    "Rx del segmento afectado (895200) + hemograma (903801) + tipificación (903901)",
    lambda d, p: any(_cie_starts(c, "S48", "S58", "S68", "S78", "S88", "S98") for c in d),
    lambda d, p: (
        _tiene_cups(p, "895200") and
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903901", "903900")
    ),
)

# ── PATOLOGÍA MUSCULOESQUELÉTICA ─────────────────────────────────────────────
_reg(
    "PERT-MSK-01", "Patología musculoesquelética inflamatoria",
    "M00–M25 (artritis, artropatías) y M40–M54 (dorsopatías)",
    "Rx articulación afectada (895200) y reactantes de fase aguda PCR/VSG (904501/904601)",
    lambda d, p: any(
        _cie_en_rango(c, "M", 0, 25) or _cie_en_rango(c, "M", 40, 54) for c in d
    ),
    lambda d, p: _tiene_cups(p, "895200", "895210", "895220", "895230", "881601"),
)

# ══════════════════════════════════════════════════════════════════════════════
# REGLAS TRANSVERSALES
# ══════════════════════════════════════════════════════════════════════════════

# TRV-01: Diagnóstico de infección sin hemocultivo ni cultivo
_reg(
    "PERT-TRV-01", "Infección grave sin cultivo",
    "A40–A41 / A49 / B99 (sepsis, bacteriemia, infección NE)",
    "Hemocultivo (905001) y/o cultivo del foco + hemograma (903801)",
    lambda d, p: any(_cie_en_rango(c, "A", 40, 49) or _cie_starts(c, "B99") for c in d),
    lambda d, p: (
        _tiene_cups(p, "905001", "905000", "905100", "905101") and
        _tiene_cups(p, "903801", "903800")
    ),
)

# TRV-02: Diagnóstico oncológico sin imagen confirmatoria
_reg(
    "PERT-TRV-02", "Diagnóstico oncológico sin imagen",
    "C00–C97 (tumores malignos de cualquier localización)",
    "TAC (893103/893104) o RMN (893200) o ecografía (881601) confirmatoria",
    lambda d, p: any(_cie_en_rango(c, "C", 0, 97) for c in d),
    lambda d, p: _tiene_cups(p, "893100", "893103", "893104", "893200", "881601", "881602"),
)

# TRV-03: Alta complejidad sin soporte diagnóstico (múltiples CIE-10 graves sin labs)
_reg(
    "PERT-TRV-03", "Múltiples diagnósticos graves sin soporte laboratorial",
    "Dos o más diagnósticos críticos (trauma + infección + falla orgánica)",
    "Hemograma (903801) + creatinina (903601) + transaminasas (903701) + coagulación",
    lambda d, p: (
        sum(1 for c in d if any([
            _cie_en_rango(c, "S", 0, 99), _cie_en_rango(c, "T", 0, 65),
            _cie_en_rango(c, "A", 40, 49), _cie_en_rango(c, "I", 20, 64),
            _cie_en_rango(c, "K", 70, 87), _cie_en_rango(c, "N", 10, 23),
        ])) >= 2
    ),
    lambda d, p: (
        _tiene_cups(p, "903801", "903800") and
        _tiene_cups(p, "903601", "903600")
    ),
)

# TRV-04: Insuficiencia renal aguda sin diálisis documentada
_reg(
    "PERT-TRV-04", "IRA sin seguimiento analítico",
    "N17.x (insuficiencia renal aguda)",
    "BUN (903602) + creatinina (903601) + uroanálisis (903200) seriados",
    lambda d, p: any(_cie_starts(c, "N17") for c in d),
    lambda d, p: (
        _tiene_cups(p, "903601", "903600") and
        _tiene_cups(p, "903602") and
        _tiene_cups(p, "903200", "903201")
    ),
)

# TRV-05: Diagnóstico de coagulopatía sin pruebas de coagulación
_reg(
    "PERT-TRV-05", "Coagulopatía sin perfil de coagulación",
    "D65–D69 (trastornos de la coagulación, CID, trombocitopenia)",
    "TP/INR (903401), TPT (903402), fibrinógeno (903403) y hemograma (903801)",
    lambda d, p: any(_cie_en_rango(c, "D", 65, 69) for c in d),
    lambda d, p: (
        _tiene_cups(p, "903401", "903402", "903403") and
        _tiene_cups(p, "903801", "903800")
    ),
)


# ══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def validar_pertinencia(data, nombre_archivo=""):
    """
    Aplica todas las reglas de pertinencia clínica sobre el JSON RIPS.

    Retorna lista de dicts con el mismo esquema que validar_auditoria,
    con severidad="observacion" (color amarillo).
    """
    observaciones = []
    if not isinstance(data, dict):
        return observaciones

    num_factura = _get_factura(data)

    for usuario in data.get("usuarios", []):
        if not isinstance(usuario, dict):
            continue

        nd = _num_doc_usuario(usuario)
        servicios = usuario.get("servicios", {})
        if not isinstance(servicios, dict):
            continue

        procs = servicios.get("procedimientos", []) or []
        cons  = servicios.get("consultas",      []) or []
        urgs  = servicios.get("urgencias",      []) or []
        hosps = servicios.get("hospitalizacion",[]) or []
        otros = servicios.get("otrosServicios", []) or []

        # Recolectar CIE-10 (sin punto, mayúscula)
        diags = _diagnosticos_usuario(servicios)
        if not diags:
            continue

        # Recolectar CUPS de procedimientos + consultas + otros servicios
        cod_procs = []
        for r in procs:
            if isinstance(r, dict):
                v = _nstr(r.get("codProcedimiento"))
                if v:
                    cod_procs.append(v)
        for r in cons:
            if isinstance(r, dict):
                v = _nstr(r.get("codConsulta"))
                if v:
                    cod_procs.append(v)
        for r in otros:
            if isinstance(r, dict):
                v = _nstr(r.get("codTecnologiaSalud"))
                if v:
                    cod_procs.append(v)

        for regla in REGLAS:
            try:
                if not regla["aplica"](diags, cod_procs):
                    continue
                if regla["ok"](diags, cod_procs):
                    continue

                diags_aplican = [
                    c for c in diags
                    if regla["aplica"]({c}, cod_procs)
                ]
                cie_str = ", ".join(sorted(diags_aplican)[:5])

                observaciones.append({
                    "archivo":     nombre_archivo,
                    "num_factura": num_factura,
                    "num_doc":     nd,
                    "consecutivo": "",
                    "id_regla":    regla["id"],
                    "severidad":   "observacion",
                    "campo":       "codDiagnosticoPrincipal / codProcedimiento",
                    "mensaje": (
                        f"{regla['grupo']}: diagnóstico(s) {regla['cie_desc']} "
                        f"sin los procedimientos de soporte esperados. "
                        f"Se esperan: {regla['cups_esperados']}."
                    ),
                    "valor_actual": f"CIE-10: {cie_str}",
                })
            except Exception:
                pass

    return observaciones
