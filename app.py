from flask import (
    Flask,
    render_template,
    request,
    send_file,
    session,
    make_response
)
from datetime import datetime, date
from collections import OrderedDict
import pandas as pd
import math
import tempfile

app = Flask(__name__)
app.secret_key = "tb_984_992_secret_key"


# -------------------------------------------------
# Ruta de test
# -------------------------------------------------
@app.route("/test")
def test():
    return "FLASK FUNCIONA"


# -------------------------------------------------
# Cálculo de la edad (parte entera diferencia fechas)
# -------------------------------------------------
def calcular_edad(fecha_matriculacion):
    hoy = date.today()
    diferencia = hoy - fecha_matriculacion
    return math.floor(diferencia.days / 365)


# -------------------------------------------------
# Leer modelos desde tablas.xlsx
# -------------------------------------------------
def obtener_modelos_desde_tablas():
    df = pd.read_excel("tablas.xlsx")
    df.columns = df.columns.astype(str).str.strip()

    modelos = []

    for _, fila in df.iterrows():
        modelo = str(fila["MODELO"]).strip()
        desde = fila.get("DESDE")
        hasta = fila.get("HASTA")
        potencia = fila.get("POTENCIA")

        partes = [modelo]

        if not pd.isna(desde) and not pd.isna(hasta):
            partes.append(f"{int(desde)}–{int(hasta)}")

        if not pd.isna(potencia):
            partes.append(f"{int(potencia)} CV")

        modelos.append(" | ".join(partes))

    modelos = sorted(set(modelos))
    modelos.append("NO ESTÁ EN TABLAS")

    return modelos


# -------------------------------------------------
# Obtener valor fiscal según edad
# -------------------------------------------------
def obtener_valor_fiscal(modelo_texto, edad):
    df = pd.read_excel("tablas.xlsx")
    df.columns = df.columns.astype(str).str.strip()

    modelo_base = modelo_texto.split("|")[0].strip()
    fila = df[df["MODELO"] == modelo_base]

    if fila.empty:
        raise ValueError("Modelo no encontrado")

    edad_tabla = min(edad, 12)
    columna = str(edad_tabla)

    return float(fila[columna].values[0])


# -------------------------------------------------
# Construcción de salida (epígrafe 3)
# -------------------------------------------------
def construir_salida(
    precio_origen,
    gestion_origen,
    revision_origen,
    transporte,
    otros_origen,
    impuesto,
    proceso_matriculacion,
    revision_post,
    honorarios
):
    salida = OrderedDict()

    salida["PRECIO EN ORIGEN"] = precio_origen
    salida["COSTE DE GESTIÓN EN ORIGEN"] = gestion_origen
    salida["COSTE DE REVISIÓN EN ORIGEN"] = revision_origen
    salida["COSTE DE TRANSPORTE"] = transporte
    salida["OTROS COSTES EN ORIGEN"] = otros_origen
    salida["IMPUESTO DE MATRICULACIÓN"] = impuesto
    salida["PROCESO DE MATRICULACIÓN"] = proceso_matriculacion
    salida["REVISION POST COMPRA"] = revision_post
    salida["HONORARIOS DE GESTIÓN"] = honorarios

    total_costes = (
        gestion_origen + revision_origen + transporte +
        otros_origen + impuesto +
        proceso_matriculacion + revision_post +
        honorarios
    )

    salida["TOTAL DE COSTES DE IMPORTACIÓN"] = total_costes
    salida["TOTAL DE COSTE PARA CLIENTE"] = math.ceil(
        (precio_origen + total_costes) / 100
    ) * 100

    return salida


# -------------------------------------------------
# Generar Excel de salida
# -------------------------------------------------
def generar_excel(salida):
    df = pd.DataFrame(
        list(salida.items()),
        columns=["Concepto", "Importe (€)"]
    )
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df.to_excel(tmp.name, index=False)
    return tmp.name


# -------------------------------------------------
# Ruta principal
# -------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def inicio():
    salida = None
    modelos = obtener_modelos_desde_tablas()

    if request.method == "POST":
        modelo = request.form["MODELO"]
        precio_origen = float(request.form["PRECIO EN ORIGEN"])

        fecha = datetime.strptime(
            request.form["FECHA DE PRIMERA MATRICULACIÓN"], "%Y-%m-%d"
        ).date()
        edad = calcular_edad(fecha)

        gestion = float(request.form["COSTE DE GESTIÓN EN ORIGEN"])
        revision = float(request.form["COSTE DE REVISIÓN EN ORIGEN"])
        transporte = float(request.form["COSTE DE TRANSPORTE"])
        otros = float(request.form["OTROS COSTES EN ORIGEN"])
        proceso = float(request.form["PROCESO DE MATRICULACIÓN"])
        post = float(request.form["REVISION POST COMPRA"])
        honorarios = float(request.form["HONORARIOS DE GESTIÓN"])

        if modelo == "NO ESTÁ EN TABLAS":
            base = precio_origen
        else:
            base = obtener_valor_fiscal(modelo, edad)

        impuesto = base * 0.1475

        salida = construir_salida(
            precio_origen,
            gestion,
            revision,
            transporte,
            otros,
            impuesto,
            proceso,
            post,
            honorarios
        )

        session["salida"] = salida

    response = make_response(
        render_template("index.html", salida=salida, modelos=modelos)
    )
    response.headers["Content-Type"] = "text/html; charset=utf-8"
    return response


# -------------------------------------------------
# Descargar Excel
# -------------------------------------------------
@app.route("/descargar_excel", methods=["POST"])
def descargar_excel():
    salida = session.get("salida")
    if not salida:
        return "No hay datos", 400

    ruta = generar_excel(salida)
    return send_file(
        ruta,
        as_attachment=True,
        download_name="coste_importacion_porsche.xlsx"
    )


if __name__ == "__main__":
    app.run()
