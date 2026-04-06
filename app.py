from flask import Flask, render_template, request
import json

app = Flask(__name__)


def extraer_registros(data, nombre_archivo=""):
    """
    Recorre todo el JSON y, en cada diccionario donde exista 'vrServicio',
    guarda:
      - idrips (1.{tipo}.{consecutivo})
      - codConsulta/codTecnologiaSalud/codProcedimiento (según aplique)
      - nomTecnologiaSalud
      - vrServicio
      - numeroFactura (detectado en el JSON y propagado)
      - archivo (opcional)
    """

    tipo_servicio_index = {
        "consultas": 1,
        "medicamentos": 2,
        "procedimientos": 3,
        "urgencias": 4,
        "hospitalizacion": 5,
        "recienNacidos": 6,
        "otrosServicios": 7
    }

    # Posibles llaves donde suele venir el número de factura
    FACTURA_KEYS = {"numFactura", "numeroFactura", "nroFactura", "noFactura", "factura"}

    resultados = []

    def normalizar_factura(valor):
        if valor is None:
            return ""
        # Si viene como dict (ej: {"numFactura": "..."}), lo resolvemos arriba, aquí solo string/número
        if isinstance(valor, (int, float)):
            # Evitar .0 en floats que son enteros
            if isinstance(valor, float) and valor.is_integer():
                return str(int(valor))
            return str(valor)
        return str(valor).strip()

    def buscar_factura_en_dict(d):
        """
        Intenta encontrar una factura en el dict actual.
        Retorna "" si no encuentra.
        """
        for k in FACTURA_KEYS:
            if k in d:
                v = d.get(k)
                # Si "factura" fuera un objeto, intenta sacar numFactura
                if isinstance(v, dict):
                    for kk in FACTURA_KEYS:
                        if kk in v:
                            return normalizar_factura(v.get(kk))
                    return ""
                return normalizar_factura(v)
        return ""

    def recorrer(nodo, tipo_idx=None, factura_actual=""):
        if isinstance(nodo, dict):
            # Si en este nivel aparece factura, la actualizamos y se hereda hacia abajo
            f_local = buscar_factura_en_dict(nodo)
            if f_local:
                factura_actual = f_local

            # Si el diccionario tiene vrServicio, extraemos el registro
            if "vrServicio" in nodo:
                consecutivo = nodo.get("consecutivo")
                if tipo_idx is not None and consecutivo is not None:
                    idrips = f"1.{tipo_idx}.{consecutivo}"
                else:
                    idrips = "no tiene consecutivo, validar estructura rips"

                resultados.append({
                    "archivo": nombre_archivo,
                    "numeroFactura": factura_actual,
                    "idrips": idrips,
                    "codConsulta": (
                        nodo.get("codConsulta")
                        or nodo.get("codTecnologiaSalud")
                        or nodo.get("codProcedimiento")
                        or ""
                    ),
                    "nomTecnologiaSalud": nodo.get("nomTecnologiaSalud", ""),
                    "vrServicio": nodo.get("vrServicio", "")
                })

            # Recorrer hijos
            for clave, valor in nodo.items():
                nuevo_tipo = tipo_idx
                if clave in tipo_servicio_index:
                    nuevo_tipo = tipo_servicio_index[clave]
                recorrer(valor, nuevo_tipo, factura_actual)

        elif isinstance(nodo, list):
            for item in nodo:
                recorrer(item, tipo_idx, factura_actual)

    recorrer(data)
    return resultados


@app.route('/', methods=['GET', 'POST'])
def index():
    registros = []
    errores = []

    if request.method == 'POST':
        archivos = request.files.getlist('json_files')

        if not archivos or all(a.filename == "" for a in archivos):
            errores.append("Por favor seleccione uno o más archivos JSON.")
        else:
            for archivo in archivos:
                if not archivo or archivo.filename == "":
                    continue
                try:
                    data = json.load(archivo)
                    regs = extraer_registros(data, nombre_archivo=archivo.filename)
                    if not regs:
                        errores.append(f"[{archivo.filename}] No se encontraron nodos con 'vrServicio'.")
                    registros.extend(regs)
                except Exception as e:
                    errores.append(f"[{archivo.filename}] Error al leer el JSON (validar si es RIPS): {e}")

    # Mantengo 'error' por compatibilidad con tu template actual,
    # pero también mando 'errores' (lista) por si quieres mostrarlos mejor.
    error = "\n".join(errores) if errores else None

    return render_template('index.html', registros=registros if registros else None, error=error, errores=errores)


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
