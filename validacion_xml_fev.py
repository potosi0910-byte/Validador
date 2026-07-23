"""
validacion_xml_fev.py
Validación del XML de factura electrónica de venta (archivo AttachedDocument,
nombre "Ad##############.xml") contra el Anexo Técnico 2 de la Resolución 948
de 2026 ("Campos de datos del sector salud, adicionales a la generación de la
FEV"), y cruce de esa información con el RIPS JSON de la misma factura.

Estructura del contenedor:
  AttachedDocument (Ad*.xml)
    └─ cac:Attachment[1] → CDATA con el Invoice UBL 2.1 (factura)
    └─ cac:Attachment[2] → CDATA con el ApplicationResponse (respuesta DIAN)

Este módulo solo necesita el Invoice embebido en el primer Attachment.

Cada validador devuelve una lista de dicts con la misma forma que el resto
de módulos de validación del proyecto:
    archivo, num_factura, num_doc, consecutivo,
    id_regla, severidad, campo, mensaje, valor_actual
"""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

_NS = {
    "ad":  "urn:oasis:names:specification:ubl:schema:xsd:AttachedDocument-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "ext": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
}

RE_NOMBRE_AD_XML = re.compile(r'^[Aa][Dd].+\.xml$')

_COBERTURA_VALIDAS = {str(i) for i in range(1, 18)}  # 1..17, Tabla "coberturaPlan"
_MODALIDAD_VALIDAS = {"01", "02", "03", "04", "05"}

# ── Tabla "CoberturaPlanBeneficios" (Anexo Técnico 2, Res. 948/2026) ─────────
# 3  Prima EPS, no asegurados SOAT | 4  Cobertura Póliza SOAT
_COBERTURA_SOAT             = {"3", "4"}
# SOAT (4) + planes voluntarios de salud: medicina prepagada(11), pólizas(12),
# plan complementario(10) — según numeral 3.5 NUMERO_POLIZA del Anexo Técnico 2.
_COBERTURA_REQUIERE_POLIZA  = {"4", "10", "11", "12"}

# ── Tabla "RIPSTipoUsuarioVersion2" — combinaciones documentadas cobertura↔tipoUsuario ──
# Solo se incluyen los pares que el Anexo Técnico 2 / la definición de cada
# código dejan inequívocos; el resto de coberturas no se restringen para
# evitar falsos positivos por combinaciones no documentadas explícitamente.
_MATRIZ_COBERTURA_TIPOUSUARIO = {
    "16": {"01", "02", "03"},   # UPC Régimen Contributivo → cotizante/beneficiario/adicional
    "17": {"04"},               # UPC Régimen Subsidiado → subsidiado
    "5":  {"09"},                # Cobertura ARL → Tomador/Amparado ARL
    "4":  {"10"},                # Cobertura Póliza SOAT → Tomador/Amparado SOAT
    "10": {"11"},                # Plan complementario en salud → Amparado planes voluntarios
    "11": {"11"},                # Plan medicina prepagada → Amparado planes voluntarios
    "12": {"11"},                # Pólizas en salud → Amparado planes voluntarios
    "13": {"06", "07", "13"},    # Régimen Especial o Excepción
    "14": {"08"},                 # Fondo Nacional de Salud PPL
    "15": {"12"},                 # Particular
}


def _nstr(v):
    if v is None:
        return ""
    return str(v).strip()


def _sin_ceros(v):
    """Normaliza '02' -> '2' para comparar contra tablas de referencia numéricas."""
    v = _nstr(v)
    try:
        return str(int(v))
    except (ValueError, TypeError):
        return v


def _num(v):
    try:
        return float(_nstr(v).replace(",", ""))
    except (ValueError, TypeError):
        return None


def es_archivo_ad_xml(nombre_archivo: str) -> bool:
    """True si el nombre de archivo corresponde al AttachedDocument (Ad####.xml)."""
    return bool(RE_NOMBRE_AD_XML.match((nombre_archivo or "").strip()))


# ══════════════════════════════════════════════════════════════
# EXTRACCIÓN: AttachedDocument → Invoice embebido
# ══════════════════════════════════════════════════════════════

def extraer_invoice_de_ad_xml(contenido: bytes):
    """
    Recibe los bytes del archivo Ad####.xml y retorna el Element raíz del
    Invoice UBL embebido en el primer cac:Attachment/.../cbc:Description
    (CDATA). Retorna None si no se pudo extraer/parsear.
    """
    try:
        root = ET.fromstring(contenido)
    except ET.ParseError:
        return None

    for attachment in root.findall(".//cac:Attachment", _NS):
        desc = attachment.find(".//cbc:Description", _NS)
        if desc is None or not desc.text:
            continue
        texto = desc.text.strip()
        if "<Invoice" not in texto:
            continue
        try:
            return ET.fromstring(texto)
        except ET.ParseError:
            continue
    return None


def _txt(elem, path):
    if elem is None:
        return ""
    found = elem.find(path, _NS)
    return _nstr(found.text) if found is not None else ""


def _attr(elem, path, attrib):
    if elem is None:
        return ""
    found = elem.find(path, _NS)
    return _nstr(found.get(attrib)) if found is not None else ""


def _local(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _find_by_local_path(root, path_parts):
    """Busca elementos por nombre local, ignorando el namespace heredado
    (CustomTagGeneral y sus hijos no declaran namespace propio y heredan el
    default del Invoice raíz, por lo que no se pueden ubicar por prefijo)."""
    current = [root]
    for part in path_parts:
        siguiente = []
        for elem in current:
            siguiente.extend(child for child in elem if _local(child.tag) == part)
        current = siguiente
        if not current:
            return []
    return current


def _additional_info_value(invoice, name, attrib=None):
    """
    Busca dentro de CustomTagGeneral/Interoperabilidad/Group/Collection
    el AdditionalInformation cuyo Name == name, y retorna su Value
    (o el atributo `attrib` del Value si se indica).
    """
    for ai in _find_by_local_path(
        invoice,
        ["UBLExtensions", "UBLExtension", "ExtensionContent", "CustomTagGeneral",
         "Interoperabilidad", "Group", "Collection", "AdditionalInformation"],
    ):
        nombres = [c for c in ai if _local(c.tag) == "Name"]
        if nombres and _nstr(nombres[0].text) == name:
            valores = [c for c in ai if _local(c.tag) == "Value"]
            if not valores:
                return None
            if attrib:
                return _nstr(valores[0].get(attrib))
            return _nstr(valores[0].text)
    return None


def extraer_datos_factura(invoice) -> dict:
    """Extrae del Invoice UBL los campos relevantes al Anexo Técnico 2."""
    if invoice is None:
        return {}

    lineas = []
    for line in invoice.findall("cac:InvoiceLine", _NS):
        item = line.find("cac:Item", _NS)
        lineas.append({
            "id":              _txt(line, "cbc:ID"),
            "uuid_paciente":   _txt(line, "cbc:UUID"),
            "valor":           _num(_txt(line, "cbc:LineExtensionAmount")),
            "descripcion":     _txt(item, "cbc:Description") if item is not None else "",
            "num_autorizacion": _txt(item, "cac:BuyersItemIdentification/cbc:ID") if item is not None else "",
            "cod_prestador":    _attr(item, "cac:BuyersItemIdentification/cbc:ID", "schemeAgencyID") if item is not None else "",
            "cod_estandar":     _txt(item, "cac:StandardItemIdentification/cbc:ID") if item is not None else "",
        })

    persona_doc = ""
    person = invoice.find(".//cac:AccountingCustomerParty/cac:Party/cac:Person", _NS)
    if person is not None:
        persona_doc = _txt(person, "cac:IdentityDocumentReference/cbc:ID") or _txt(person, "cbc:ID")

    return {
        "invoice_id":        _txt(invoice, "cbc:ID"),
        "cufe":              _txt(invoice, "cbc:UUID"),
        "issue_date":        _txt(invoice, "cbc:IssueDate"),
        "tipo_operacion":    _txt(invoice, "cbc:InvoiceTypeCode"),
        "customization_id":  _txt(invoice, "cbc:CustomizationID"),
        "periodo_inicio":    _txt(invoice, "cac:InvoicePeriod/cbc:StartDate"),
        "periodo_fin":       _txt(invoice, "cac:InvoicePeriod/cbc:EndDate"),
        "nit_prestador":     _txt(invoice, ".//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID"),
        "nit_prestador_dv":  _attr(invoice, ".//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID", "schemeID"),
        "nit_prestador_scheme_name": _attr(invoice, ".//cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID", "schemeName"),
        "nit_pagador":       _txt(invoice, ".//cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cbc:CompanyID"),
        "documento_paciente_factura": persona_doc,
        "valor_total":       _num(_txt(invoice, "cac:LegalMonetaryTotal/cbc:PayableAmount")),
        "codigo_prestador":  _additional_info_value(invoice, "CODIGO_PRESTADOR"),
        "modalidad_pago":         _additional_info_value(invoice, "MODALIDAD_PAGO"),
        "modalidad_pago_scheme":  _additional_info_value(invoice, "MODALIDAD_PAGO", "schemeID"),
        "cobertura_plan":         _additional_info_value(invoice, "COBERTURA_PLAN_BENEFICIOS"),
        "cobertura_plan_scheme":  _additional_info_value(invoice, "COBERTURA_PLAN_BENEFICIOS", "schemeID"),
        "numero_contrato":   _additional_info_value(invoice, "NUMERO_CONTRATO"),
        "numero_poliza":     _additional_info_value(invoice, "NUMERO_POLIZA"),
        "copago":            _num(_additional_info_value(invoice, "COPAGO")),
        "cuota_moderadora":  _num(_additional_info_value(invoice, "CUOTA_MODERADORA")),
        "pagos_compartidos": _num(_additional_info_value(invoice, "PAGOS_COMPARTIDOS")),
        "lineas":            lineas,
    }


# ══════════════════════════════════════════════════════════════
# VALIDACIÓN ESTRUCTURAL — Anexo Técnico 2, numerales 10.2 y 10.3
# ══════════════════════════════════════════════════════════════

def validar_estructura_xml_fev(datos: dict, nombre_archivo: str = "") -> list:
    """
    Valida presencia y formato de los campos obligatorios del Invoice según
    el Anexo Técnico 2 de la Resolución 948/2026, sin necesidad del RIPS.
    """
    errores = []
    if not datos:
        return [{
            "archivo": nombre_archivo, "num_factura": "", "num_doc": "", "consecutivo": "",
            "id_regla": "PFE002", "severidad": "critica", "campo": "Invoice",
            "mensaje": "[Invoice] El apartado no existe o no se pudo extraer del XML del documento electrónico.",
            "valor_actual": "",
        }]

    ctx = {"archivo": nombre_archivo, "num_factura": datos.get("invoice_id", ""), "num_doc": "", "consecutivo": ""}

    def _e(id_regla, severidad, campo, mensaje, valor_actual=""):
        errores.append({**ctx, "id_regla": id_regla, "severidad": severidad,
                         "campo": campo, "mensaje": mensaje, "valor_actual": _nstr(valor_actual)})

    if not datos.get("invoice_id"):
        _e("VFE005", "critica", "Invoice.ID", "El número de factura (Invoice.ID) está vacío.")

    if not datos.get("cufe") or len(datos.get("cufe", "")) < 90:
        _e("VFE006", "critica", "Invoice.UUID", "El CUFE está vacío o no cumple la longitud esperada (SHA-384).",
           datos.get("cufe", ""))

    if not datos.get("issue_date"):
        _e("FED005", "critica", "Invoice.IssueDate", "La fecha de expedición de la factura está vacía.")

    if datos.get("valor_total") is None:
        _e("FED017", "critica", "LegalMonetaryTotal.PayableAmount",
           "El valor total de la factura (PayableAmount) no existe o no tiene valor.")

    if not datos.get("nit_prestador"):
        _e("FED021", "critica", "AccountingSupplierParty.PartyTaxScheme.CompanyID",
           "El NIT del prestador (facturador) no existe o no tiene valor.")
    if not datos.get("nit_prestador_dv"):
        _e("FED022", "critica", "AccountingSupplierParty.PartyTaxScheme.CompanyID.schemeID",
           "Falta el dígito de verificación del NIT del prestador.")
    if datos.get("nit_prestador_scheme_name") != "31":
        _e("VFE011", "critica", "AccountingSupplierParty.PartyTaxScheme.CompanyID.schemeName",
           "El schemeName del NIT del prestador debe ser el literal '31'.", datos.get("nit_prestador_scheme_name", ""))

    if not datos.get("nit_pagador"):
        _e("FED034", "critica", "AccountingCustomerParty.PartyTaxScheme.CompanyID",
           "El NIT del pagador (EPS/ERP) no existe o no tiene valor.")

    if not datos.get("periodo_inicio"):
        _e("VFE020", "critica", "InvoicePeriod.StartDate",
           "Debe reportar la fecha de inicio del periodo de facturación.")
    if not datos.get("periodo_fin"):
        _e("VFE021", "critica", "InvoicePeriod.EndDate",
           "Debe reportar la fecha final del periodo de facturación.")

    if not datos.get("codigo_prestador"):
        _e("DND124", "critica", "Interoperabilidad.CODIGO_PRESTADOR",
           "No existe el grupo CODIGO_PRESTADOR en la extensión de Salud del XML.")

    modalidad = datos.get("modalidad_pago")
    modalidad_scheme = datos.get("modalidad_pago_scheme")
    if not modalidad:
        _e("DND126", "critica", "Interoperabilidad.MODALIDAD_PAGO",
           "No existe el grupo MODALIDAD_PAGO en la extensión de Salud del XML.")
    elif not modalidad_scheme:
        _e("DND125", "critica", "Interoperabilidad.MODALIDAD_PAGO.schemeID",
           "Falta el atributo schemeID en MODALIDAD_PAGO.")
    elif modalidad_scheme not in _MODALIDAD_VALIDAS and _sin_ceros(modalidad_scheme).zfill(2) not in _MODALIDAD_VALIDAS:
        _e("DND133", "alta", "Interoperabilidad.MODALIDAD_PAGO",
           "El código schemeID de MODALIDAD_PAGO no está en la tabla de referencia 'modalidadPago'.",
           modalidad_scheme)

    cobertura = datos.get("cobertura_plan")
    cobertura_scheme = datos.get("cobertura_plan_scheme")
    if not cobertura:
        _e("DND128", "critica", "Interoperabilidad.COBERTURA_PLAN_BENEFICIOS",
           "No existe el grupo COBERTURA_PLAN_BENEFICIOS en la extensión de Salud del XML.")
    elif not cobertura_scheme:
        _e("DND127", "critica", "Interoperabilidad.COBERTURA_PLAN_BENEFICIOS.schemeID",
           "Falta el atributo schemeID en COBERTURA_PLAN_BENEFICIOS.")
    elif _sin_ceros(cobertura_scheme) not in _COBERTURA_VALIDAS:
        _e("DND134", "alta", "Interoperabilidad.COBERTURA_PLAN_BENEFICIOS",
           "El código schemeID de COBERTURA_PLAN_BENEFICIOS no está en la tabla de referencia "
           "'coberturaPlan' (1-17).", cobertura_scheme)

    for campo, valor in (("COPAGO", datos.get("copago")),
                         ("CUOTA_MODERADORA", datos.get("cuota_moderadora")),
                         ("PAGOS_COMPARTIDOS", datos.get("pagos_compartidos"))):
        if valor is not None and valor < 0:
            _e("VFE-NEG", "alta", f"Interoperabilidad.{campo}",
               f"El valor de {campo} no puede ser negativo.", valor)

    if not datos.get("lineas"):
        _e("FED-LINE", "critica", "Invoice.InvoiceLine",
           "La factura no contiene líneas de detalle (InvoiceLine).")
    else:
        for linea in datos["lineas"]:
            if not linea.get("num_autorizacion"):
                _e("FED-AUT", "alta", f"InvoiceLine[{linea.get('id')}].BuyersItemIdentification",
                   "La línea de factura no tiene número de autorización asociado.", linea.get("id"))

    return errores


# ══════════════════════════════════════════════════════════════
# VALIDACIÓN DE CRUCE — XML (FEV) vs RIPS JSON
# ══════════════════════════════════════════════════════════════

def _todos_servicios(usuario):
    servicios = usuario.get("servicios", {})
    if not isinstance(servicios, dict):
        return []
    todos = []
    for grupo in servicios.values():
        if isinstance(grupo, list):
            todos.extend(s for s in grupo if isinstance(s, dict))
    return todos


def validar_cruce_xml_rips(datos: dict, rips_data: dict, nombre_archivo: str = "") -> list:
    """
    Compara la información extraída del XML de la FEV con el RIPS JSON de la
    misma factura, aplicando reglas equivalentes a PFP001 y a la validación
    de consistencia FEV-RIPS del numeral 10.3 del Anexo Técnico 2.
    """
    errores = []
    if not datos or not isinstance(rips_data, dict):
        return errores

    num_factura_rips = _nstr(rips_data.get("numFactura", ""))
    ctx = {"archivo": nombre_archivo, "num_factura": datos.get("invoice_id") or num_factura_rips,
           "num_doc": "", "consecutivo": ""}

    def _e(id_regla, severidad, campo, mensaje, valor_actual=""):
        errores.append({**ctx, "id_regla": id_regla, "severidad": severidad,
                         "campo": campo, "mensaje": mensaje, "valor_actual": _nstr(valor_actual)})

    # ── PFP001: número de factura XML vs RIPS ────────────────────────────
    if datos.get("invoice_id") and num_factura_rips and datos["invoice_id"] != num_factura_rips:
        _e("PFP001", "critica", "Invoice.ID vs RIPS.numFactura",
           f"El número de factura del XML ('{datos['invoice_id']}') no corresponde al reportado "
           f"por el RIPS ('{num_factura_rips}').", datos["invoice_id"])

    # ── NIT del prestador (facturador) vs numDocumentoIdObligado ─────────
    nit_rips = _nstr(rips_data.get("numDocumentoIdObligado", ""))
    if datos.get("nit_prestador") and nit_rips and datos["nit_prestador"] != nit_rips:
        _e("XML-RIPS-002", "critica", "AccountingSupplierParty.CompanyID vs RIPS.numDocumentoIdObligado",
           f"El NIT del prestador en el XML ('{datos['nit_prestador']}') no coincide con "
           f"numDocumentoIdObligado del RIPS ('{nit_rips}').", datos["nit_prestador"])

    # ── Recolectar usuarios/servicios del RIPS ───────────────────────────
    usuarios = rips_data.get("usuarios", [])
    if not isinstance(usuarios, list):
        usuarios = []

    docs_pacientes_rips = set()
    autorizaciones_rips = set()
    prestadores_rips = set()
    tipos_usuario_rips = set()
    faltan_siras = []
    suma_vrservicio = 0.0
    suma_copago = 0.0
    suma_cuota_moderadora = 0.0
    suma_pagos_compartidos = 0.0

    cobertura_scheme = _sin_ceros(datos.get("cobertura_plan_scheme") or "")
    es_soat = cobertura_scheme in _COBERTURA_SOAT

    for usuario in usuarios:
        if not isinstance(usuario, dict):
            continue
        doc_pac = _nstr(usuario.get("numDocumentoIdentificacion") or usuario.get("numDocumentoldentificacion") or "")
        if doc_pac:
            docs_pacientes_rips.add(doc_pac)
        tipo_usr = _nstr(usuario.get("tipoUsuario"))
        if tipo_usr:
            tipos_usuario_rips.add(tipo_usr)
        for serv in _todos_servicios(usuario):
            num_aut = _nstr(serv.get("numAutorizacion"))
            if num_aut:
                autorizaciones_rips.add(num_aut)
            if serv.get("codPrestador"):
                prestadores_rips.add(_nstr(serv["codPrestador"]))
            vr = _num(serv.get("vrServicio"))
            if vr is not None:
                suma_vrservicio += vr
            valor_mod = _num(serv.get("valorPagoModerador"))
            if valor_mod:
                concepto = _nstr(serv.get("conceptoRecaudo"))
                if concepto == "01":
                    suma_copago += valor_mod
                elif concepto == "02":
                    suma_cuota_moderadora += valor_mod
                elif concepto == "03":
                    suma_pagos_compartidos += valor_mod
            # ── SOAT-REQUIERE-SIRAS: el radicado SIRAS reemplaza a la autorización
            # tradicional para las coberturas 3 (Prima EPS no aseg. SOAT) y 4
            # (Cobertura Póliza SOAT); debe existir y ser estrictamente numérico.
            if es_soat and (not num_aut or not num_aut.isdigit()):
                faltan_siras.append((doc_pac, num_aut))

    # ── Valor total factura (PayableAmount) vs sumatoria vrServicio RIPS ──
    # El valor facturado a la entidad responsable de pago NO incluye lo que ya
    # fue recaudado directamente del usuario (copago, cuota moderadora, pagos
    # compartidos): PayableAmount = Σ vrServicio − copago − cuota_moderadora −
    # pagos_compartidos (numerales 3.6/3.7/3.8 del Anexo Técnico 2).
    valor_total = datos.get("valor_total")
    if valor_total is not None and usuarios:
        valor_esperado = suma_vrservicio - suma_copago - suma_cuota_moderadora - suma_pagos_compartidos
        if round(valor_total, 2) != round(valor_esperado, 2):
            _e("XML-RIPS-004", "critica", "LegalMonetaryTotal.PayableAmount vs suma(RIPS.vrServicio) − recaudos",
               f"El valor total de la factura en el XML ({valor_total:,.2f}) no coincide con la "
               f"sumatoria de vrServicio del RIPS ({suma_vrservicio:,.2f}) menos copago "
               f"({suma_copago:,.2f}), cuota moderadora ({suma_cuota_moderadora:,.2f}) y pagos "
               f"compartidos ({suma_pagos_compartidos:,.2f}) = {valor_esperado:,.2f}.", valor_total)

    # ── COPAGO / CUOTA_MODERADORA / PAGOS_COMPARTIDOS vs RIPS ────────────
    for campo_xml, valor_xml, suma_rips in (
        ("COPAGO", datos.get("copago"), suma_copago),
        ("CUOTA_MODERADORA", datos.get("cuota_moderadora"), suma_cuota_moderadora),
        ("PAGOS_COMPARTIDOS", datos.get("pagos_compartidos"), suma_pagos_compartidos),
    ):
        if valor_xml is not None and round(valor_xml, 2) != round(suma_rips, 2):
            _e("XML-RIPS-005", "alta", f"Interoperabilidad.{campo_xml} vs RIPS",
               f"El valor de {campo_xml} en el XML ({valor_xml:,.2f}) no corresponde al valor total "
               f"registrado en RIPS ({suma_rips:,.2f}).", valor_xml)

    # ── Autorizaciones y pacientes de cada línea de factura ──────────────
    for linea in datos.get("lineas", []):
        num_aut = linea.get("num_autorizacion")
        if num_aut and autorizaciones_rips and num_aut not in autorizaciones_rips:
            _e("XML-RIPS-006", "alta", f"InvoiceLine[{linea.get('id')}].BuyersItemIdentification",
               f"El número de autorización '{num_aut}' de la línea de factura no se encuentra "
               f"en ningún servicio del RIPS.", num_aut)

        cod_prestador_linea = linea.get("cod_prestador")
        if cod_prestador_linea and prestadores_rips and cod_prestador_linea not in prestadores_rips:
            _e("XML-RIPS-003", "alta", f"InvoiceLine[{linea.get('id')}].BuyersItemIdentification.schemeAgencyID",
               f"El código de prestador de la línea de factura ('{cod_prestador_linea}') no coincide "
               f"con ningún codPrestador reportado en los servicios del RIPS.", cod_prestador_linea)

        doc_pac_linea = linea.get("uuid_paciente")
        if doc_pac_linea and docs_pacientes_rips and doc_pac_linea not in docs_pacientes_rips:
            _e("XML-RIPS-007", "alta", f"InvoiceLine[{linea.get('id')}].UUID (documento paciente)",
               f"El documento de identificación del paciente en la línea de factura ('{doc_pac_linea}') "
               f"no corresponde a ningún usuario reportado en el RIPS.", doc_pac_linea)

    # ── Documento del paciente en AccountingCustomerParty (factura unipersonal) ──
    doc_factura = datos.get("documento_paciente_factura")
    if doc_factura and docs_pacientes_rips and doc_factura not in docs_pacientes_rips:
        _e("XML-RIPS-008", "alta", "AccountingCustomerParty.Party.Person vs RIPS.numDocumentoIdentificacion",
           f"El documento del paciente en el XML ('{doc_factura}') no corresponde a ningún usuario "
           f"reportado en el RIPS.", doc_factura)

    # ── SOAT no debe llevar número de contrato (CUCON) ───────────────────
    # Las atenciones SOAT (cobertura 3 y 4) se pagan contra la póliza, no
    # contra un contrato/acuerdo de voluntades (numeral 3.4 NUMERO_CONTRATO).
    numero_contrato = _nstr(datos.get("numero_contrato"))
    if es_soat and numero_contrato:
        _e("XML-SOAT-CUCON", "critica", "Interoperabilidad.NUMERO_CONTRATO",
           "La factura es de cobertura SOAT (código {}) y no debe reportar NUMERO_CONTRATO/CUCON."
           .format(cobertura_scheme), numero_contrato)

    # ── SOAT y planes voluntarios deben reportar NUMERO_POLIZA ───────────
    numero_poliza = _nstr(datos.get("numero_poliza"))
    if cobertura_scheme in _COBERTURA_REQUIERE_POLIZA and not numero_poliza:
        _e("XML-POLIZA-001", "critica", "Interoperabilidad.NUMERO_POLIZA",
           "La cobertura informada (código {}) exige reportar NUMERO_POLIZA (SOAT o planes "
           "voluntarios de salud) y el campo está vacío.".format(cobertura_scheme))
    elif cobertura_scheme and cobertura_scheme not in _COBERTURA_REQUIERE_POLIZA and numero_poliza:
        _e("XML-POLIZA-002", "alta", "Interoperabilidad.NUMERO_POLIZA",
           "Se reportó NUMERO_POLIZA pero la cobertura informada (código {}) no corresponde a "
           "SOAT ni a planes voluntarios de salud.".format(cobertura_scheme), numero_poliza)

    # ── Códigos SOAT deben tener radicado SIRAS (numérico) ───────────────
    if es_soat and faltan_siras:
        for doc_pac, num_aut in faltan_siras:
            _e("XML-SOAT-SIRAS", "critica", "servicios.numAutorizacion (radicado SIRAS)",
               "Las atenciones con cobertura SOAT deben reportar el radicado SIRAS en "
               "numAutorizacion; el valor está vacío o no es numérico.", num_aut)

    # ── El plan de beneficios (XML) debe cuadrar con el tipoUsuario (RIPS) ──
    combinaciones_esperadas = _MATRIZ_COBERTURA_TIPOUSUARIO.get(cobertura_scheme)
    if combinaciones_esperadas and tipos_usuario_rips:
        no_cuadran = tipos_usuario_rips - combinaciones_esperadas
        if no_cuadran:
            _e("XML-PLAN-CUADRE", "alta", "COBERTURA_PLAN_BENEFICIOS vs RIPS.tipoUsuario",
               "El plan de beneficios informado (código {}) no cuadra con el/los tipoUsuario "
               "reportado(s) en el RIPS: {}. Se esperaba tipoUsuario en {}."
               .format(cobertura_scheme, sorted(no_cuadran), sorted(combinaciones_esperadas)),
               ", ".join(sorted(no_cuadran)))

    return errores


# ══════════════════════════════════════════════════════════════
# API DE ALTO NIVEL
# ══════════════════════════════════════════════════════════════

def validar_xml_fev(contenido_xml: bytes, rips_data: dict | None, nombre_archivo: str = "") -> list:
    """
    Punto de entrada único: extrae el Invoice del AttachedDocument, valida su
    estructura contra el Anexo Técnico 2 y, si se dispone del RIPS JSON de la
    misma factura, valida el cruce entre ambos documentos.
    """
    invoice = extraer_invoice_de_ad_xml(contenido_xml)
    datos = extraer_datos_factura(invoice)

    errores = validar_estructura_xml_fev(datos, nombre_archivo)
    if rips_data is not None and datos:
        errores.extend(validar_cruce_xml_rips(datos, rips_data, nombre_archivo))
    return errores


# ══════════════════════════════════════════════════════════════
# CORRESPONDENCIA DE ARCHIVOS POR CARPETA (Ad####.xml ↔ Rips_SL######.json)
# ══════════════════════════════════════════════════════════════
# Solo se valida un Ad####.xml cuando existe su RIPS JSON en la misma carpeta
# (identificados ambos por el mismo número de factura). Este helper detecta
# los casos en que la carpeta de un lote quedó incompleta — le falta el RIPS,
# le falta el XML, o el XML no trae un número de factura identificable — para
# reportarlos en una hoja aparte del Excel en vez de mezclarlos con las
# alertas de validación normales.

def detectar_archivos_faltantes(rips_por_factura: dict, facturas_xml: dict, xml_sin_factura: list | None = None) -> list:
    """
    rips_por_factura: { numFactura -> nombre_archivo_json }
    facturas_xml:      { numFactura -> nombre_archivo_xml } (solo XML con
                        invoice_id identificado)
    xml_sin_factura:   lista de nombres de archivo XML de los que no se pudo
                        extraer un número de factura (Invoice.ID vacío o XML
                        no parseable)
    Retorna una lista de dicts: factura, tipo_faltante, archivo, mensaje.
    """
    faltantes = []

    for factura, archivo_json in rips_por_factura.items():
        if factura not in facturas_xml:
            faltantes.append({
                "factura": factura,
                "tipo_faltante": "Falta XML",
                "archivo": archivo_json,
                "mensaje": f"La carpeta de la factura '{factura}' tiene RIPS JSON ('{archivo_json}') "
                           f"pero no se encontró su archivo Ad####.xml correspondiente.",
            })

    for factura, archivo_xml in facturas_xml.items():
        if factura not in rips_por_factura:
            faltantes.append({
                "factura": factura,
                "tipo_faltante": "Falta RIPS",
                "archivo": archivo_xml,
                "mensaje": f"La carpeta de la factura '{factura}' tiene el archivo XML "
                           f"('{archivo_xml}') pero no se encontró su RIPS JSON correspondiente.",
            })

    for nombre in (xml_sin_factura or []):
        faltantes.append({
            "factura": "",
            "tipo_faltante": "XML sin factura identificable",
            "archivo": nombre,
            "mensaje": f"No se pudo extraer un número de factura (Invoice.ID) del archivo "
                       f"'{nombre}'; no se pudo emparejar con ningún RIPS ni validar.",
        })

    return faltantes
