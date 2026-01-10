import os
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from docx.oxml import OxmlElement, ns
from datetime import datetime
import pytz
import zipfile
import io

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLA_WORD = os.path.join(BASE_DIR, "templates_word", "AUTOR.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --------------------------------------------------
# UTILIDADES WORD
# --------------------------------------------------

def limpiar_documento(doc):
    body = doc._element.body
    for element in list(body):
        body.remove(element)

def aplicar_bordes_dobles(tabla):
    tbl = tabla._element
    tblPr = tbl.tblPr

    borders = OxmlElement("w:tblBorders")
    for borde in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        elem = OxmlElement(f"w:{borde}")
        elem.set(ns.qn("w:val"), "double")
        elem.set(ns.qn("w:sz"), "8")       # grosor
        elem.set(ns.qn("w:space"), "0")
        elem.set(ns.qn("w:color"), "000000")
        borders.append(elem)

    tblPr.append(borders)

def negrita(run):
    run.bold = True

# --------------------------------------------------
# FICHA
# --------------------------------------------------

def agregar_ficha(doc, registro):
    tiene_edicion = bool(registro.get("EDICION", "").strip())
    filas = 7 if tiene_edicion else 6

    tabla = doc.add_table(rows=filas, cols=2)
    tabla.alignment = 1  # centrada
    aplicar_bordes_dobles(tabla)

    ancho_total = Cm(12)  # 20 - 4 - 4
    tabla.columns[0].width = Cm(4)
    tabla.columns[1].width = Cm(8)

    fila_actual = 0

    def fila(titulo, valor, titulo_negrita=True, valor_negrita=False):
        nonlocal fila_actual
        c1, c2 = tabla.rows[fila_actual].cells

        r1 = c1.paragraphs[0].add_run(titulo)
        if titulo_negrita:
            negrita(r1)

        r2 = c2.paragraphs[0].add_run(valor or "")
        if valor_negrita:
            negrita(r2)

        fila_actual += 1

    # ORDEN CORRECTO
    fila("SIGNATURA TOPOGRAFICA", registro.get("SIGNATURA", ""), True, True)
    fila("AUTOR PRINCIPAL", registro.get("AUTOR", ""), True, True)
    fila("TITULO / SUBTITULO", registro.get("TITULO", ""), True, False)
    fila("MENCION RESPONSABILIDAD", registro.get("MENCION", ""), True, False)

    if tiene_edicion:
        fila("EDICION", registro.get("EDICION", ""), True, False)

    fila("IMPRENTA", registro.get("IMPRENTA", ""), True, False)
    fila("DESCRIPCION FISICA", registro.get("DESCRIPCION", ""), True, False)

    doc.add_page_break()

# --------------------------------------------------
# CREAR WORD
# --------------------------------------------------

def crear_word(registros):
    tz = pytz.timezone("America/Bogota")
    ahora = datetime.now(tz).strftime("%Y-%m-%d_%H-%M-%S")

    mfns = [r["MFN"] for r in registros if r.get("MFN")]
    mfn_inicio = mfns[0] if mfns else "00000"
    mfn_fin = mfns[-1] if mfns else "00000"

    nombre = f"fichas_{ahora}_MFN_{mfn_inicio}_{mfn_fin}.docx"
    ruta = os.path.join(OUTPUT_DIR, nombre)

    doc = Document(PLANTILLA_WORD)
    limpiar_documento(doc)

    for r in registros:
        agregar_ficha(doc, r)

    doc.save(ruta)
    return ruta, nombre

# --------------------------------------------------
# WEB
# --------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        registros = request.get_json()

        ruta, nombre = crear_word(registros)
        return send_file(ruta, as_attachment=True, download_name=nombre)

    return render_template("index.html")

# --------------------------------------------------
# RENDER
# --------------------------------------------------

if __name__ == "__main__":
    app.run()
    #app.run(host="0.0.0.0", port=10000)

