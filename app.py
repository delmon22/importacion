from flask import Flask, render_template, request, send_file, session
from datetime import datetime, date
from collections import OrderedDict
import pandas as pd
import math
import tempfile

app = Flask(__name__)
app.secret_key = "tb_984_992_secret_key"


# -------------------------------------------------
# Test rápido
# -------------------------------------------------
@app.route("/test")
def test():
    return "FLASK FUNCIONA"


# -------------------------------------------------
# Cálculo de edad (parte entera)
# -------------------------------------------------
def calcular_edad(fecha_matriculacion):
    hoy = date.today()
    return math.floor((hoy - fecha_matriculacion).days / 365)


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

        if pd.notna(desde) and pd.notna(hasta):
            partes.append(f"{int(desde)}–{int(hasta)}")

        if pd.notna(potencia):
            partes.append(f"{int(potencia)} CV")

        modelos.append(" | ".join(partes))

    return sorted(set(modelos))


# -------------------------------------------------
# Obtener valor fiscal
# -------------------------------------------------
def obtener_valor_fiscal(modelo_texto, edad):
    df = pd.read_excel("tablas.xlsx")
    df.columns = df.columns.astype(str).str.strip()

    modelo_base = modelo_texto.split("|")[0].strip()
    fila = df[df["MODELO"] == modelo_base]

    if fila.empty:
        raise ValueError("Modelo no encontrado en tablas")

    edad = min(edad, 12)
    return float(fila[str(edad)].values[0])


# -------------------------------------------------
# Construir salida
# -------------------------------------------------
def construir_salida(precio, base, impuesto):
    salida = OrderedDict()
    salida["PRECIO EN ORIGEN"] = precio
    salida["BASE IMPONIBLE"] = base
    salida["IMPUESTO DE MATRICULACIÓN"] = impuesto
    salida["TOTAL COSTE CLIENTE"] = math.ceil((precio + impuesto) / 100) * 100
    return salida


# -------------------------------------------------
# Ruta principal
# -------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def inicio():
    salida = None
    modelos = obtener_modelos_desde_tablas()

    if request.method == "POST":
        modelo = request.form["MODELO"]
        precio = float(request.form["PRECIO EN ORIGEN"])

        fecha = datetime.strptime(
            request.form["FECHA DE PRIMERA MATRICULACIÓN"], "%Y-%m-%d"
        ).date()

        edad = calcular_edad(fecha)

        if modelo == "NO ESTÁ EN TABLAS":
            base = precio
        else:
            base = obtener_valor_fiscal(modelo, edad)

        impuesto = base * 0.1475
        salida = construir_salida(precio, base, impuesto)
        session["salida"] = salida

    return render_template("index.html", modelos=modelos, salida=salida)


# -------------------------------------------------
# Descargar Excel
# -------------------------------------------------
@app.route("/descargar_excel", methods=["POST"])
def descargar_excel():
    salida = session.get("salida")
    if not salida:
        return "No hay datos", 400

    df = pd.DataFrame(list(salida.items()), columns=["Concepto", "Importe (€)"])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df.to_excel(tmp.name, index=False)

    return send_file(
        tmp.name,
        as_attachment=True,
        download_name="coste_importacion_porsche.xlsx"
    )
