"""
pertinencia.py
Módulo de observaciones de pertinencia clínica para RIPS JSON (Res. 2275/2023).
Valida la coherencia entre diagnósticos CIE-10 y procedimientos CUPS reportados.
Códigos CUPS actualizados con el catálogo 2026 (actualización 2026-04-29).
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
    for k in ("numDocumentoIdentificacion", "numDocumentoldentificacion", "num_doc"):
        v = u.get(k)
        if v and not isinstance(v, (dict, list)):
            return _nstr(v)
    return ""


def _cie_starts(cod, *prefixes):
    c = cod.replace(".", "").upper()
    return any(c.startswith(p.upper()) for p in prefixes)


def _cie_en_rango(cod, letra, num_ini, num_fin):
    c = cod.replace(".", "").upper()
    if not c.startswith(letra.upper()):
        return False
    try:
        num = int(c[1:3])
        return num_ini <= num <= num_fin
    except (ValueError, IndexError):
        return False


def _tiene(cod_procs, *codigos):
    """True si algún CUPS de la lista coincide exactamente con alguno de los dados."""
    s = set(cod_procs)
    return any(c in s for c in codigos)


# ── Catálogo CUPS 2026 — grupos de códigos por tipo de examen ────────────────

# Imagenología
TC_CRANEO   = ("879111","879112","879113")
TC_TORAX    = ("879301","879303","879304","879391")
TC_ABDOMEN  = ("879410","879420","879391")
TC_COLUMNA  = ("879201","879205")
TC_CUELLO   = ("879161",)
RMN_CEREBRO = ("883101","883104","883106")
RMN_COLUMNA = ("883210","883220","883230","883231","883234","883235")
RX_TORAX    = ("871121",)
RX_CRANEO   = ("870001","870003")
RX_COL_CERV = ("871010","871019")
RX_COL_TOR  = ("871020","871030")
RX_COL_LUMB = ("871040","871060")
RX_HOMBRO   = ("873204","873202","873112")
RX_CODO     = ("873205",)
RX_MUNECA   = ("873206","873210")
RX_CADERA   = ("873411","873412")
RX_RODILLA  = ("873420","873422")
RX_PIERNA   = ("873313",)
RX_TOBILLO  = ("873431",)
RX_PIE      = ("873333","873303")
RX_FEMUR    = ("873312",)
ECO_ABDOM   = ("881302","881305","881306","881390")
ECO_PELV    = ("881401","881360","881390")
ECO_OBSTET  = ("881436","881437","881432","882298")
ECO_VBILIAR = ("881306",)
ECG         = ("895101","895001","895401")
EEG         = ("891401","891402","891410","891901")
ESPIRO      = ("893703","893805","893808")

# Laboratorio
HEMOGRAMA   = ("902207","902208","902209","902210","902227")
GLUCOSA     = ("903841",)
CREATININA  = ("903895",)
BUN         = ("903856",)
ALT_AST     = ("903866","903867")
BILIRRUBINAS= ("903809",)
AMILASA     = ("903805",)
LIPASA      = ("903847",)
TROPONINA   = ("903436","903437","903438","903439")
GASES_ART   = ("903839","903062")
HBA1C       = ("903426","903427")
IONOGRAMA   = ("903605",)
RETICULOCIT = ("902223","902224")
FERRITINA   = ("903016",)
TP_INR      = ("902045",)
TPT         = ("902049",)
FIBRINOGENO = ("902024",)
PCR         = ("906913","906914")
BACILOSCOPIA= ("901101","901111")
CD4         = ("906714","906712","906713")
VIH_CONFIRM = ("906249","906250","906302","908802")
UROANAL     = ("907106",)
HEMOCULTIVO = ("901221","901222","901223","901224")
UROCULTIVO  = ("901235","901236","901237")
COPROCULTIVO= ("901206",)
GRUPO_SANG  = ("911011","911012")   # fenotipo eritrocitario (más cercano disponible)


# ── Recolector de diagnósticos de un usuario ─────────────────────────────────

def _diagnosticos_usuario(servicios):
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
# TABLA DE REGLAS DE PERTINENCIA
# Formato: (id, grupo, cie_desc, cups_desc, fn_aplica, fn_ok)
# ══════════════════════════════════════════════════════════════════════════════

REGLAS = [

    # ── TRAUMA CRANEOENCEFÁLICO ───────────────────────────────────────────────
    ("PERT-TRA-01", "Trauma craneoencefálico",
     "S06.x — lesiones intracraneales",
     "TC de cráneo (879111-879113) y hemograma (902207-902210)",
     lambda d, p: any(_cie_starts(c, "S06") for c in d),
     lambda d, p: _tiene(p, *TC_CRANEO, *RMN_CEREBRO) and _tiene(p, *HEMOGRAMA)),

    # ── TRAUMA FACIAL ─────────────────────────────────────────────────────────
    ("PERT-TRA-02", "Trauma facial",
     "S02.x — fractura huesos cráneo y cara",
     "RX cráneo/cara (870001) o TC cráneo (879111-879113)",
     lambda d, p: any(_cie_starts(c, "S02") for c in d),
     lambda d, p: _tiene(p, *RX_CRANEO, *TC_CRANEO)),

    # ── TRAUMA CERVICAL ───────────────────────────────────────────────────────
    ("PERT-TRA-03", "Trauma cervical",
     "S12.x / S13.x — fractura y luxación columna cervical",
     "RX columna cervical (871010/871019) o TC columna (879201)",
     lambda d, p: any(_cie_starts(c, "S12", "S13") for c in d),
     lambda d, p: _tiene(p, *RX_COL_CERV, *TC_COLUMNA, *TC_CUELLO)),

    # ── TRAUMA TORÁCICO ───────────────────────────────────────────────────────
    ("PERT-TRA-04", "Trauma torácico",
     "S22.x / S27.x — fractura costillas y lesión órganos torácicos",
     "RX tórax (871121) o TC tórax (879301/879303/879304)",
     lambda d, p: any(_cie_starts(c, "S22", "S27") for c in d),
     lambda d, p: _tiene(p, *RX_TORAX, *TC_TORAX)),

    # ── TRAUMA ABDOMINAL ─────────────────────────────────────────────────────
    ("PERT-TRA-05", "Trauma abdominal",
     "S36.x / S37.x — lesión órganos abdominales y pélvicos",
     "Ecografía abdominal (881302/881305) o TC abdomen (879410/879420) y hemograma",
     lambda d, p: any(_cie_starts(c, "S36", "S37") for c in d),
     lambda d, p: _tiene(p, *ECO_ABDOM, *TC_ABDOMEN) and _tiene(p, *HEMOGRAMA)),

    # ── TRAUMA HOMBRO ─────────────────────────────────────────────────────────
    ("PERT-TRA-06", "Trauma hombro / clavícula",
     "S42.x / S43.x — fractura y luxación hombro y clavícula",
     "RX hombro (873204) o RX clavícula (873112)",
     lambda d, p: any(_cie_starts(c, "S42", "S43") for c in d),
     lambda d, p: _tiene(p, *RX_HOMBRO)),

    # ── TRAUMA CODO / ANTEBRAZO ──────────────────────────────────────────────
    ("PERT-TRA-07", "Trauma codo / antebrazo",
     "S52.x / S53.x — fractura y luxación codo y antebrazo",
     "RX codo (873205)",
     lambda d, p: any(_cie_starts(c, "S52", "S53") for c in d),
     lambda d, p: _tiene(p, *RX_CODO)),

    # ── TRAUMA MUÑECA / MANO ─────────────────────────────────────────────────
    ("PERT-TRA-08", "Trauma muñeca / mano",
     "S62.x / S63.x / S64.x — fractura y luxación muñeca y mano",
     "RX puño o muñeca (873206) o RX mano (873210)",
     lambda d, p: any(_cie_starts(c, "S62", "S63", "S64") for c in d),
     lambda d, p: _tiene(p, *RX_MUNECA)),

    # ── TRAUMA CADERA / FÉMUR ────────────────────────────────────────────────
    ("PERT-TRA-09", "Trauma cadera / fémur",
     "S72.x — fractura de fémur y cadera",
     "RX cadera (873411/873412) o RX fémur (873312) y hemograma",
     lambda d, p: any(_cie_starts(c, "S72") for c in d),
     lambda d, p: _tiene(p, *RX_CADERA, *RX_FEMUR) and _tiene(p, *HEMOGRAMA)),

    # ── TRAUMA RODILLA / PIERNA ──────────────────────────────────────────────
    ("PERT-TRA-10", "Trauma rodilla / pierna",
     "S82.x / S83.x — fractura y luxación rodilla y pierna",
     "RX rodilla (873420/873422) o RX pierna (873313)",
     lambda d, p: any(_cie_starts(c, "S82", "S83") for c in d),
     lambda d, p: _tiene(p, *RX_RODILLA, *RX_PIERNA)),

    # ── TRAUMA TOBILLO / PIE ─────────────────────────────────────────────────
    ("PERT-TRA-11", "Trauma tobillo / pie",
     "S92.x / S93.x / S99.x — fractura y luxación tobillo y pie",
     "RX tobillo (873431) o RX pie (873333)",
     lambda d, p: any(_cie_starts(c, "S92", "S93", "S99") for c in d),
     lambda d, p: _tiene(p, *RX_TOBILLO, *RX_PIE)),

    # ── POLITRAUMATISMO ──────────────────────────────────────────────────────
    ("PERT-TRA-12", "Politraumatismo",
     "T00–T07 — traumatismos múltiples de varias regiones",
     "TC cráneo (879111) + TC tórax (879301) + TC abdomen (879410) y hemograma",
     lambda d, p: any(_cie_en_rango(c, "T", 0, 7) for c in d),
     lambda d, p: (
         _tiene(p, *TC_CRANEO, *RMN_CEREBRO) and
         _tiene(p, *TC_TORAX) and
         _tiene(p, *TC_ABDOMEN) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── QUEMADURAS ───────────────────────────────────────────────────────────
    ("PERT-QUE-01", "Quemaduras",
     "T20–T32 — quemaduras y corrosiones",
     "Hemograma (902207-902210) y proteínas/albúmina",
     lambda d, p: any(_cie_en_rango(c, "T", 20, 32) for c in d),
     lambda d, p: _tiene(p, *HEMOGRAMA)),

    # ── HEPATOPATÍA ──────────────────────────────────────────────────────────
    ("PERT-HEP-01", "Hepatopatía",
     "K70–K77 — enfermedades hepáticas",
     "Transaminasas ALT/AST (903866/903867) + bilirrubinas (903809) + ecografía hepática (881306)",
     lambda d, p: any(_cie_en_rango(c, "K", 70, 77) for c in d),
     lambda d, p: (
         _tiene(p, *ALT_AST) and
         _tiene(p, *BILIRRUBINAS) and
         _tiene(p, *ECO_VBILIAR, *ECO_ABDOM)
     )),

    # ── PATOLOGÍA BILIAR ─────────────────────────────────────────────────────
    ("PERT-BIL-01", "Patología biliar",
     "K80–K87 — colelitiasis, colecistitis, colangitis",
     "Ecografía hígado/vías biliares (881306) + bilirrubinas (903809) + amilasa (903805)",
     lambda d, p: any(_cie_en_rango(c, "K", 80, 87) for c in d),
     lambda d, p: (
         _tiene(p, *ECO_VBILIAR, *ECO_ABDOM) and
         _tiene(p, *BILIRRUBINAS, *AMILASA, *LIPASA)
     )),

    # ── PATOLOGÍA ABDOMINAL AGUDA ────────────────────────────────────────────
    ("PERT-ABD-01", "Patología abdominal aguda",
     "K35–K40 — apendicitis, hernia, peritonitis",
     "Ecografía abdominal (881302/881305) o TC abdomen (879410) + hemograma + PCR",
     lambda d, p: any(_cie_en_rango(c, "K", 35, 40) for c in d),
     lambda d, p: (
         _tiene(p, *ECO_ABDOM, *TC_ABDOMEN) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── PATOLOGÍA RENAL ──────────────────────────────────────────────────────
    ("PERT-REN-01", "Patología renal aguda / crónica",
     "N10–N23 — pielonefritis, cólico nefrítico, IRA, IRC",
     "Creatinina (903895) + uroanálisis (907106) + BUN (903856)",
     lambda d, p: any(_cie_en_rango(c, "N", 10, 23) for c in d),
     lambda d, p: (
         _tiene(p, *CREATININA) and
         _tiene(p, *UROANAL)
     )),

    # ── CARDIOPATÍA ISQUÉMICA / SCA ──────────────────────────────────────────
    ("PERT-CAR-01", "Cardiopatía isquémica / SCA",
     "I20–I25 — angina, IAM, cardiopatía isquémica",
     "ECG (895101) + troponina (903436-903439) + hemograma",
     lambda d, p: any(_cie_en_rango(c, "I", 20, 25) for c in d),
     lambda d, p: (
         _tiene(p, *ECG) and
         _tiene(p, *TROPONINA)
     )),

    # ── ICC / ARRITMIAS / EMBOLIA PULMONAR ───────────────────────────────────
    ("PERT-CAR-02", "Insuficiencia cardíaca / arritmias / embolia",
     "I26–I52 — ICC, arritmias, pericarditis, embolia pulmonar",
     "ECG (895101) + RX tórax (871121)",
     lambda d, p: any(_cie_en_rango(c, "I", 26, 52) for c in d),
     lambda d, p: (
         _tiene(p, *ECG) and
         _tiene(p, *RX_TORAX, *TC_TORAX)
     )),

    # ── INFECCIÓN RESPIRATORIA AGUDA ─────────────────────────────────────────
    ("PERT-RES-01", "Infección respiratoria aguda / neumonía",
     "J12–J22 — neumonías víricas y bacterianas, bronquitis aguda",
     "RX tórax (871121) + hemograma (902207-902210)",
     lambda d, p: any(_cie_en_rango(c, "J", 12, 22) for c in d),
     lambda d, p: (
         _tiene(p, *RX_TORAX, *TC_TORAX) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── EPOC / ASMA ──────────────────────────────────────────────────────────
    ("PERT-RES-02", "EPOC / Asma",
     "J44–J46 — EPOC, estado asmático",
     "RX tórax (871121) + espirometría (893703/893805/893808) o gases arteriales (903839)",
     lambda d, p: any(_cie_en_rango(c, "J", 44, 46) for c in d),
     lambda d, p: (
         _tiene(p, *RX_TORAX, *TC_TORAX) and
         _tiene(p, *ESPIRO, *GASES_ART)
     )),

    # ── EPILEPSIA / CONVULSIÓN ───────────────────────────────────────────────
    ("PERT-NEU-01", "Epilepsia / Convulsión",
     "G40–G41 — epilepsia y estado epiléptico",
     "EEG (891401/891402) o TC/RMN cerebral (879111/883101)",
     lambda d, p: any(_cie_en_rango(c, "G", 40, 41) for c in d),
     lambda d, p: _tiene(p, *EEG, *TC_CRANEO, *RMN_CEREBRO)),

    # ── ECV / ACV / HEMORRAGIA CEREBRAL ─────────────────────────────────────
    ("PERT-NEU-02", "ECV / ACV / Hemorragia cerebral",
     "I60–I64 — hemorragia subaracnoidea, intracerebral, ACV isquémico",
     "TC cráneo urgente (879111-879113) + hemograma + TP/TPT",
     lambda d, p: any(_cie_en_rango(c, "I", 60, 64) for c in d),
     lambda d, p: (
         _tiene(p, *TC_CRANEO, *RMN_CEREBRO) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── CEFALEA / MIGRAÑA ────────────────────────────────────────────────────
    ("PERT-NEU-03", "Cefalea / Migraña",
     "G43–G44 / R51 — migraña, cefalea en racimos, cefalea NE",
     "TC cráneo (879111) o RMN cerebral (883101) si primera vez o cefalea severa",
     lambda d, p: any(_cie_starts(c, "G43", "G44", "R51") for c in d),
     lambda d, p: _tiene(p, *TC_CRANEO, *RMN_CEREBRO, *EEG)),

    # ── DIARREA / GASTROENTERITIS INFECCIOSA ─────────────────────────────────
    ("PERT-INF-01", "Diarrea / Gastroenteritis infecciosa",
     "A09 — otras gastroenteritis y colitis infecciosas",
     "Coprocultivo (901206) o ionograma (903605) + hemograma",
     lambda d, p: any(_cie_starts(c, "A09") for c in d),
     lambda d, p: (
         _tiene(p, *COPROCULTIVO, *IONOGRAMA) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── TUBERCULOSIS ─────────────────────────────────────────────────────────
    ("PERT-INF-02", "Tuberculosis",
     "A15–A19 — tuberculosis respiratoria y miliar",
     "Baciloscopia (901101/901111) + RX tórax (871121)",
     lambda d, p: any(_cie_en_rango(c, "A", 15, 19) for c in d),
     lambda d, p: (
         _tiene(p, *BACILOSCOPIA) and
         _tiene(p, *RX_TORAX, *TC_TORAX)
     )),

    # ── VIH / SIDA ───────────────────────────────────────────────────────────
    ("PERT-INF-03", "VIH / SIDA",
     "B20 — enfermedad por VIH",
     "Prueba confirmatoria VIH (906249/906250) + linfocitos T CD4 (906714) + hemograma",
     lambda d, p: any(_cie_starts(c, "B20") for c in d),
     lambda d, p: (
         _tiene(p, *VIH_CONFIRM) and
         _tiene(p, *CD4) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── DIABETES DESCOMPENSADA ───────────────────────────────────────────────
    ("PERT-MET-01", "Diabetes / Hiperglucemia descompensada",
     "E10–E14 — diabetes mellitus tipo 1, 2 y complicaciones",
     "Glucosa (903841) + HbA1c (903426/903427) + creatinina (903895)",
     lambda d, p: any(_cie_en_rango(c, "E", 10, 14) for c in d),
     lambda d, p: (
         _tiene(p, *GLUCOSA) and
         _tiene(p, *CREATININA)
     )),

    # ── TRASTORNO HIDROELECTROLÍTICO ─────────────────────────────────────────
    ("PERT-MET-02", "Trastorno hidroelectrolítico / Deshidratación",
     "E86–E87 — depleción de volumen, hiponatremia, hipopotasemia",
     "Ionograma / electrolitos séricos (903605) + creatinina (903895)",
     lambda d, p: any(_cie_en_rango(c, "E", 86, 87) for c in d),
     lambda d, p: (
         _tiene(p, *IONOGRAMA) and
         _tiene(p, *CREATININA)
     )),

    # ── ANEMIA ───────────────────────────────────────────────────────────────
    ("PERT-HEM-01", "Anemia",
     "D50–D64 — anemias ferropénica, hemolítica, aplásica y otras",
     "Hemograma (902207-902210) + reticulocitos (902223/902224) + ferritina (903016)",
     lambda d, p: any(_cie_en_rango(c, "D", 50, 64) for c in d),
     lambda d, p: (
         _tiene(p, *HEMOGRAMA) and
         _tiene(p, *RETICULOCIT, *FERRITINA)
     )),

    # ── INFECCIÓN GINECOLÓGICA / EPI ─────────────────────────────────────────
    ("PERT-GIN-01", "Infección ginecológica / EPI",
     "N70–N77 — salpingitis, endometritis, vaginitis, EPI",
     "Ecografía pélvica transvaginal (881401) + hemograma + urocultivo (901235)",
     lambda d, p: any(_cie_en_rango(c, "N", 70, 77) for c in d),
     lambda d, p: (
         _tiene(p, *ECO_PELV) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── COMPLICACIÓN OBSTÉTRICA ──────────────────────────────────────────────
    ("PERT-OBS-01", "Complicación obstétrica",
     "O20–O48 — amenaza aborto, hemorragia, preeclampsia, trabajo de parto",
     "Ecografía obstétrica (881436/881437/881432) + hemograma + TP/TPT",
     lambda d, p: any(_cie_en_rango(c, "O", 20, 48) for c in d),
     lambda d, p: (
         _tiene(p, *ECO_OBSTET) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── LUMBALGIA / DORSALGIA ────────────────────────────────────────────────
    ("PERT-LUM-01", "Lumbalgia / Dorsalgia",
     "M54.x — lumbago, ciática, dorsalgia",
     "RX columna lumbosacra (871040) o RMN columna (883230/883231)",
     lambda d, p: any(_cie_starts(c, "M54") for c in d),
     lambda d, p: _tiene(p, *RX_COL_LUMB, *RX_COL_TOR, *RMN_COLUMNA, *TC_COLUMNA)),

    # ── DOLOR PÉLVICO / CÓLICO ───────────────────────────────────────────────
    ("PERT-PEL-01", "Dolor pélvico / Cólico",
     "N94 / R10 — dolor pélvico, dismenorrea, dolor abdominal",
     "Ecografía pélvica (881401) + uroanálisis (907106)",
     lambda d, p: any(_cie_starts(c, "N94", "R10") for c in d),
     lambda d, p: (
         _tiene(p, *ECO_PELV, *ECO_ABDOM) and
         _tiene(p, *UROANAL)
     )),

    # ── FIEBRE SIN FOCO / SÍNDROME FEBRIL ────────────────────────────────────
    ("PERT-PED-01", "Fiebre sin foco / Síndrome febril",
     "R50 — fiebre de origen desconocido o sin foco identificado",
     "Hemograma (902207-902210) + hemocultivo (901221) + uroanálisis (907106)",
     lambda d, p: any(_cie_starts(c, "R50") for c in d),
     lambda d, p: (
         _tiene(p, *HEMOGRAMA) and
         _tiene(p, *HEMOCULTIVO, *UROCULTIVO, *UROANAL)
     )),

    # ── INTOXICACIÓN / SOBREDOSIS ────────────────────────────────────────────
    ("PERT-INT-01", "Intoxicación / Sobredosis",
     "T36–T65 — intoxicaciones por fármacos, plaguicidas y sustancias",
     "Hemograma (902207-902210) + creatinina (903895) + transaminasas ALT/AST (903866/903867)",
     lambda d, p: any(_cie_en_rango(c, "T", 36, 65) for c in d),
     lambda d, p: (
         _tiene(p, *HEMOGRAMA) and
         _tiene(p, *CREATININA) and
         _tiene(p, *ALT_AST)
     )),

    # ── FRACTURA VERTEBRAL ───────────────────────────────────────────────────
    ("PERT-FRX-01", "Fractura vertebral",
     "S32.x — fractura de columna lumbar",
     "RX columna lumbosacra (871040) o TC columna (879201)",
     lambda d, p: any(_cie_starts(c, "S32") for c in d),
     lambda d, p: _tiene(p, *RX_COL_LUMB, *TC_COLUMNA, *RMN_COLUMNA)),

    # ── AMPUTACIÓN TRAUMÁTICA ────────────────────────────────────────────────
    ("PERT-AMP-01", "Amputación traumática",
     "S48.x / S58.x / S68.x / S78.x / S88.x / S98.x — amputaciones traumáticas",
     "RX del segmento (870001/873xxx) + hemograma + fenotipo eritrocitario (911011)",
     lambda d, p: any(_cie_starts(c, "S48","S58","S68","S78","S88","S98") for c in d),
     lambda d, p: (
         _tiene(p, *HEMOGRAMA) and
         _tiene(p, *GRUPO_SANG, *TP_INR)
     )),

    # ── PATOLOGÍA MUSCULOESQUELÉTICA INFLAMATORIA ────────────────────────────
    ("PERT-MSK-01", "Patología musculoesquelética inflamatoria",
     "M00–M25 — artritis, artropatías y M40–M54 — dorsopatías",
     "RX de la articulación afectada + PCR (906913/906914)",
     lambda d, p: any(
         _cie_en_rango(c, "M", 0, 25) or _cie_en_rango(c, "M", 40, 54) for c in d
     ),
     lambda d, p: _tiene(p,
         *RX_HOMBRO, *RX_CODO, *RX_MUNECA, *RX_CADERA,
         *RX_RODILLA, *RX_COL_LUMB, *RX_COL_CERV, *RX_TORAX
     )),

    # ── SEPSIS / BACTERIEMIA ─────────────────────────────────────────────────
    ("PERT-TRV-01", "Sepsis / Bacteriemia",
     "A40–A41 / A49 — sepsis estreptocócica, estafilocócica y bacteriemia NE",
     "Hemocultivo (901221-901224) + hemograma + PCR (906913)",
     lambda d, p: any(_cie_en_rango(c, "A", 40, 49) for c in d),
     lambda d, p: (
         _tiene(p, *HEMOCULTIVO) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── NEOPLASIA SIN IMAGEN ─────────────────────────────────────────────────
    ("PERT-TRV-02", "Diagnóstico oncológico sin imagen confirmatoria",
     "C00–C97 — tumores malignos de cualquier localización",
     "TC (879301/879410) o RMN (883101) o ecografía (881302) confirmatoria",
     lambda d, p: any(_cie_en_rango(c, "C", 0, 97) for c in d),
     lambda d, p: _tiene(p, *TC_CRANEO, *TC_TORAX, *TC_ABDOMEN,
                         *RMN_CEREBRO, *ECO_ABDOM, *ECO_PELV)),

    # ── IRA SIN SEGUIMIENTO ANALÍTICO ────────────────────────────────────────
    ("PERT-TRV-03", "IRA sin seguimiento analítico",
     "N17.x — insuficiencia renal aguda",
     "BUN (903856) + creatinina (903895) + uroanálisis (907106) seriados",
     lambda d, p: any(_cie_starts(c, "N17") for c in d),
     lambda d, p: (
         _tiene(p, *CREATININA) and
         _tiene(p, *BUN) and
         _tiene(p, *UROANAL)
     )),

    # ── COAGULOPATÍA SIN PERFIL ──────────────────────────────────────────────
    ("PERT-TRV-04", "Coagulopatía sin perfil de coagulación",
     "D65–D69 — trastornos de coagulación, CID, trombocitopenia",
     "TP/INR (902045) + TPT (902049) + fibrinógeno (902024) + hemograma",
     lambda d, p: any(_cie_en_rango(c, "D", 65, 69) for c in d),
     lambda d, p: (
         _tiene(p, *TP_INR) and
         _tiene(p, *TPT) and
         _tiene(p, *HEMOGRAMA)
     )),

    # ── MULTI-DIAGNÓSTICO GRAVE SIN SOPORTE LAB ──────────────────────────────
    ("PERT-TRV-05", "Múltiples diagnósticos graves sin soporte laboratorial",
     "Dos o más diagnósticos de alta complejidad (trauma + infección + falla orgánica)",
     "Hemograma (902207-902210) + creatinina (903895) imprescindibles",
     lambda d, p: sum(1 for c in d if any([
         _cie_en_rango(c,"S",0,99), _cie_en_rango(c,"T",0,65),
         _cie_en_rango(c,"A",40,49), _cie_en_rango(c,"I",20,64),
         _cie_en_rango(c,"K",70,87), _cie_en_rango(c,"N",10,23),
     ])) >= 2,
     lambda d, p: _tiene(p, *HEMOGRAMA) and _tiene(p, *CREATININA)),
]


# ══════════════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def validar_pertinencia(data, nombre_archivo=""):
    """
    Aplica todas las reglas de pertinencia clínica sobre el JSON RIPS.
    Solo evalúa usuarios que tienen urgencias u hospitalización.
    Retorna lista de dicts con severidad='observacion' (color amarillo).
    """
    observaciones = []
    if not isinstance(data, dict):
        return observaciones

    num_factura = _get_factura(data)

    for usuario in data.get("usuarios", []):
        if not isinstance(usuario, dict):
            continue

        nd        = _num_doc_usuario(usuario)
        servicios = usuario.get("servicios", {})
        if not isinstance(servicios, dict):
            continue

        urgs  = servicios.get("urgencias",       []) or []
        hosps = servicios.get("hospitalizacion",  []) or []

        # Solo aplica a pacientes con urgencias u hospitalización
        if not urgs and not hosps:
            continue

        diags = _diagnosticos_usuario(servicios)
        if not diags:
            continue

        # Recolectar CUPS de todos los servicios
        cod_procs = []
        for seccion in ("procedimientos", "consultas", "urgencias",
                        "hospitalizacion", "otrosServicios", "medicamentos"):
            for r in servicios.get(seccion, []) or []:
                if not isinstance(r, dict):
                    continue
                for campo in ("codProcedimiento", "codConsulta",
                              "codTecnologiaSalud", "codServicio", "codigo"):
                    v = _nstr(r.get(campo))
                    if v:
                        cod_procs.append(v)

        for id_regla, grupo, cie_desc, cups_desc, fn_aplica, fn_ok in REGLAS:
            try:
                if not fn_aplica(diags, cod_procs):
                    continue
                if fn_ok(diags, cod_procs):
                    continue

                diags_aplican = sorted(
                    c for c in diags if fn_aplica({c}, cod_procs)
                )[:5]

                observaciones.append({
                    "archivo":     nombre_archivo,
                    "num_factura": num_factura,
                    "num_doc":     nd,
                    "consecutivo": "",
                    "id_regla":    id_regla,
                    "severidad":   "observacion",
                    "campo":       "codDiagnosticoPrincipal / codProcedimiento",
                    "mensaje": (
                        f"{grupo}: diagnóstico(s) [{cie_desc}] sin los "
                        f"procedimientos de soporte esperados. "
                        f"Se esperan: {cups_desc}."
                    ),
                    "valor_actual": f"CIE-10: {', '.join(diags_aplican)}",
                })
            except Exception:
                pass

    return observaciones
