import io
import os
from collections import defaultdict
from copy import copy
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "equivalencia-app-secret")

DATA_DIR = os.environ.get("DATA_DIR", "/app/data")
os.makedirs(DATA_DIR, exist_ok=True)

GRANDE_CACHE_PATH_XLS = os.path.join(DATA_DIR, "planilla_grande.xls")
GRANDE_CACHE_PATH_XLSX = os.path.join(DATA_DIR, "planilla_grande.xlsx")
GRANDE_CACHE_META = os.path.join(DATA_DIR, "planilla_grande.meta")

ALLOWED_EXT = {".xls", ".xlsx"}
BARCODE_HEADER = "Código de Barra"


def _ext(filename):
    return os.path.splitext(filename)[1].lower()


def _read_grande(path):
    """Lee la planilla grande (Equivalencia) sea .xls viejo (Crystal Reports) o .xlsx.
    Devuelve un dict: 'SKU COLOR' -> lista de (codigo_barra, talle)
    """
    ext = _ext(path)
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xl = pd.ExcelFile(path, engine=engine)

    lookup = defaultdict(list)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None, dtype=str)
        if df.empty:
            continue
        # La primera hoja trae una fila de encabezado ("Código", "Artículo", ...)
        if str(df.iloc[0, 0]).strip() == "Código":
            df = df.iloc[1:]
        for _, row in df.iterrows():
            codigo = row.get(0)
            articulo = row.get(1)
            talle = row.get(5) if len(row) > 5 else None
            if pd.isna(codigo) or pd.isna(articulo):
                continue
            codigo = str(codigo).strip()
            articulo = str(articulo).strip()
            talle = "" if pd.isna(talle) else str(talle).strip()
            if not codigo or not articulo:
                continue
            lookup[articulo].append((codigo, talle))
    return lookup


def _buscar_codigo(lookup, sku, color):
    if sku is None or color is None or str(sku).strip() == "" or str(color).strip() == "":
        return "", 0
    sku_str = str(sku).strip()
    if sku_str.endswith(".0"):
        sku_str = sku_str[:-2]
    color_str = str(color).strip()
    key = f"{sku_str} {color_str}"
    candidatos = lookup.get(key, [])
    if not candidatos:
        return "SIN COINCIDENCIA", 0
    if len(candidatos) == 1:
        return candidatos[0][0], 1
    texto = ", ".join(f"{codigo} ({talle})" if talle else codigo for codigo, talle in candidatos)
    return texto, len(candidatos)


def _copy_cell_style(src, dst):
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def _procesar(grande_path, pedido_path):
    """Abre el excel de pedido tal cual fue subido (misma planilla, mismo estilo) e
    inserta una columna nueva con el código de barra justo al lado de la columna SKU
    (columna A), corriendo el resto de las columnas una posición a la derecha.
    """
    lookup = _read_grande(grande_path)

    wb = openpyxl.load_workbook(pedido_path)
    ws = wb.worksheets[0]

    if ws.max_column < 2:
        raise ValueError("La planilla de pedido necesita al menos dos columnas (SKU y Color).")

    max_row = ws.max_row

    # 1) Calcular los resultados ANTES de tocar la estructura de la hoja
    #    (columna A = SKU, columna B = Color, tal cual vienen en el archivo original)
    resultados = []
    for row in range(2, max_row + 1):
        sku = ws.cell(row=row, column=1).value
        color = ws.cell(row=row, column=2).value
        texto, _cantidad = _buscar_codigo(lookup, sku, color)
        resultados.append(texto)

    # 2) Guardar anchos de columnas originales para reubicarlos después de insertar
    anchos_originales = {}
    for col_letter, dim in ws.column_dimensions.items():
        if dim.width:
            anchos_originales[column_index_from_string(col_letter)] = dim.width

    # 3) Insertar la columna nueva justo después del SKU (columna A -> nueva columna B)
    ws.insert_cols(2)

    # Reubicar anchos: todo lo que estaba desde la columna B en adelante corrió +1
    for idx, width in sorted(anchos_originales.items(), reverse=True):
        if idx == 1:
            continue
        ws.column_dimensions[get_column_letter(idx + 1)].width = width
    # Ancho razonable para la columna nueva de código de barra
    ws.column_dimensions[get_column_letter(2)].width = 18

    # 4) Encabezado de la columna nueva, con el mismo estilo que el resto del encabezado
    header_cell = ws.cell(row=1, column=2, value=BARCODE_HEADER)
    _copy_cell_style(ws.cell(row=1, column=3), header_cell)

    # 5) Cargar los valores calculados, con el mismo estilo de celda que el resto de la fila
    for i, texto in enumerate(resultados):
        row = i + 2
        cell = ws.cell(row=row, column=2, value=texto)
        _copy_cell_style(ws.cell(row=row, column=3), cell)
        cell.number_format = "@"  # texto plano, para no perder ceros a la izquierda

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/", methods=["GET"])
def index():
    grande_cached = os.path.exists(GRANDE_CACHE_META)
    grande_info = None
    if grande_cached:
        with open(GRANDE_CACHE_META) as f:
            grande_info = f.read().strip()
    return render_template("index.html", grande_cached=grande_cached, grande_info=grande_info)


@app.route("/procesar", methods=["POST"])
def procesar():
    pedido_file = request.files.get("pedido")
    grande_file = request.files.get("grande")

    if not pedido_file or pedido_file.filename == "":
        flash("Subí la planilla de pedido (la que hay que resolver).")
        return redirect(url_for("index"))

    if _ext(pedido_file.filename) != ".xlsx":
        flash("La planilla de pedido debe ser .xlsx (para poder mantener el mismo formato/estilo al insertar la columna).")
        return redirect(url_for("index"))

    # Resolver planilla grande: la subida ahora, o la que quedó cacheada
    if grande_file and grande_file.filename != "":
        if _ext(grande_file.filename) not in ALLOWED_EXT:
            flash("La planilla grande debe ser .xls o .xlsx")
            return redirect(url_for("index"))
        grande_ext = _ext(grande_file.filename)
        grande_path = GRANDE_CACHE_PATH_XLS if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLSX
        # Limpiar el otro formato cacheado para no usar una versión vieja por error
        other_path = GRANDE_CACHE_PATH_XLSX if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLS
        if os.path.exists(other_path):
            os.remove(other_path)
        grande_file.save(grande_path)
        with open(GRANDE_CACHE_META, "w") as f:
            f.write(f"{secure_filename(grande_file.filename)} — cargada {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    else:
        if os.path.exists(GRANDE_CACHE_PATH_XLS):
            grande_path = GRANDE_CACHE_PATH_XLS
        elif os.path.exists(GRANDE_CACHE_PATH_XLSX):
            grande_path = GRANDE_CACHE_PATH_XLSX
        else:
            flash("No hay planilla grande cargada todavía. Subí la planilla de Equivalencia primero.")
            return redirect(url_for("index"))

    pedido_bytes = io.BytesIO(pedido_file.read())
    pedido_bytes.name = pedido_file.filename

    try:
        output = _procesar(grande_path, pedido_bytes)
    except Exception as e:
        flash(f"Error procesando los archivos: {e}")
        return redirect(url_for("index"))

    nombre_salida = f"{secure_filename(os.path.splitext(pedido_file.filename)[0])} - con equivalencia.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=nombre_salida,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))

