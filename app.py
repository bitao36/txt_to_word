import os
import re
from datetime import datetime
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, ns

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLA_WORD = os.path.join(BASE_DIR, "templates_word", "AUTOR.docx")
SALIDA_DIR = os.path.join(BASE_DIR, "salida")
os.makedirs(SALIDA_DIR, exist_ok=True)

app = Flask(__name__)

# --------------------------------------------------
# WORD – utilidades
# --------------------------------------------------
def copiar_margenes(origen, destino):
    s1 = origen.sections[0]
    s2 = destino.sections[0]
    s2.top_margin = s1.top_margin
    s2.bottom_margin = s1.bottom_margin
    s2.left_margin = s1.left_margin
    s2.right_margin = s1.right_margin


def aplicar_bordes_tabla(tabla):
    tbl = tabla._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borde = OxmlElement(f"w:{lado}")
        borde.set(ns.qn("w:val"), "single")
        borde.set(ns.qn("w:sz"), "8")
        borde.set(ns.qn("w:space"), "0")
        borde.set(ns.qn("w:color"), "000000")
        borders.append(borde)

    tblPr.append(borders)


def extraer_estilo_base(modelo):
    """
    Extrae fuente y tamaño reales desde la primera celda del modelo
    """
    tabla = modelo.tables[0]
    p = tabla.cell(0, 0).paragraphs[0]
    run = p.runs[0]

    return {
        "font_name": run.font.name,
        "font_size": run.font.size,
    }


def escribir_texto(parrafo, texto, estilo, negrita=False):
    run = parrafo.add_run(texto)
    run.font.name = estilo["font_name"]
    run.font.size = estilo["font_size"]
    run.bold = negrita


# --------------------------------------------------
# TXT → registros
# --------------------------------------------------
def procesar_texto_a_registros(texto):
    bloques = re.split(r"\nMFN:\d+", texto)
    mfns = re.findall(r"MFN:(\d+)", texto)

    registros = []
    for i, bloque in enumerate(bloques):
        if not bloque.strip():
            continue

        r = {"MFN": mfns[i] if i < len(mfns) else ""}

        for linea in bloque.splitlines():
            if "\t" in linea:
                k, v = linea.split("\t", 1)
                r[k.strip()] = v.strip()

        registros.append(r)

    return registros


# --------------------------------------------------
# FICHA
# --------------------------------------------------
def agregar_ficha(doc, modelo, r, estilo):
    copiar_margenes(modelo, doc)

    tiene_edicion = bool(r.get("EDICION"))
    filas = 7 if tiene_edicion else 6

    tabla = doc.add_table(rows=filas, cols=2)
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
    aplicar_bordes_tabla(tabla)

    # Anchos exactos (13 cm útiles)
    tabla.columns[0].width = Cm(4.55)
    tabla.columns[1].width = Cm(8.45)

    campos = [
        ("SIGNATURA TOPOGRAFICA", r.get("SIGNATURA TOPOGRAFICA", "")),
        ("AUTOR PRINCIPAL", r.get("AUTOR PRINCIPAL", "")),
        ("TITULO/SUBTITULO", r.get("TITULO/SUBTITULO", "")),
        ("MENCION RESPONSABILIDAD", r.get("MENCION RESPONSABILIDAD", "")),
    ]

    if tiene_edicion:
        campos.append(("EDICION", r.get("EDICION", "")))

    campos.extend([
        ("IMPRENTA", r.get("IMPRENTA", "")),
        ("DESCRIPCION FISICA", r.get("DESCRIPCION FISICA", "")),
    ])

    for i, (k, v) in enumerate(campos):
        p1 = tabla.cell(i, 0).paragraphs[0]
        p2 = tabla.cell(i, 1).paragraphs[0]

        escribir_texto(p1, k, estilo, negrita=True)
        escribir_texto(p2, v, estilo, negrita=i < 2)

    doc.add_page_break()


# --------------------------------------------------
# WORD final
# --------------------------------------------------
def crear_word_con_fichas(registros):
    modelo = Document(PLANTILLA_WORD)
    estilo = extraer_estilo_base(modelo)

    doc = Document()
    copiar_margenes(modelo, doc)

    for r in registros:
        agregar_ficha(doc, modelo, r, estilo)

    mfns = [r["MFN"] for r in registros if r.get("MFN")]
    mfn_i = mfns[0] if mfns else "X"
    mfn_f = mfns[-1] if mfns else "Y"

    fecha = datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
    nombre = f"fichas_{fecha}_MFN_{mfn_i}_{mfn_f}.docx"

    ruta = os.path.join(SALIDA_DIR, nombre)
    doc.save(ruta)
    return ruta


# --------------------------------------------------
# FLASK
# --------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        archivo = request.files.get("archivo")
        if not archivo:
            return "No se subió ningún archivo", 400

        texto = archivo.read().decode("utf-8", errors="ignore")
        registros = procesar_texto_a_registros(texto)

        ruta = crear_word_con_fichas(registros)
        return send_file(ruta, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)

