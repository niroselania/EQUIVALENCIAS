import io
import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
import openpyxl
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


def _dedupe_codigos(codigos):
    """A veces la planilla grande tiene el mismo código de barra duplicado con
    distinto padding de ceros a la izquierda (ej. '00191743848643' y '191743848643').
    Los trata como el mismo código y se queda con la representación más corta."""
    vistos = {}
    for c in codigos:
        try:
            clave = int(c)
        except (TypeError, ValueError):
            clave = c
        if clave not in vistos or len(c) < len(vistos[clave]):
            vistos[clave] = c
    return list(vistos.values())


def _buscar_codigo(lookup, sku, color, talle):
    """Busca el código de barra exacto para SKU + Color + Talle."""
    if sku is None or color is None or str(sku).strip() == "" or str(color).strip() == "":
        return ""
    sku_str = str(sku).strip()
    if sku_str.endswith(".0"):
        sku_str = sku_str[:-2]
    color_str = str(color).strip()
    talle_str = "" if talle is None else str(talle).strip()

    key = f"{sku_str} {color_str}"
    candidatos = lookup.get(key, [])
    coincidencias = [
        codigo for codigo, t in candidatos if t.strip().upper() == talle_str.upper()
    ]
    coincidencias = _dedupe_codigos(coincidencias)

    if not coincidencias:
        return "SIN COINCIDENCIA"
    if len(coincidencias) == 1:
        return coincidencias[0]
    # Conflicto real de datos en la planilla grande: mismo SKU+Color+Talle con más
    # de un código de barra distinto. Se muestran todos para que se revise a mano.
    return " / ".join(coincidencias)


def _buscar_columna(headers, *nombres_posibles):
    for nombre in nombres_posibles:
        if nombre in headers:
            return headers[nombre]
    return None


def _procesar(grande_path, pedido_path):
    """Abre el excel de pedido tal cual fue subido (mismo formato, mismo estilo) y
    completa la columna EQUIVALENCIA que ya viene en la planilla, buscando el
    código de barra exacto por SKU + Color + Talle. No inserta columnas nuevas ni
    toca el resto de la planilla."""
    lookup = _read_grande(grande_path)

    wb = openpyxl.load_workbook(pedido_path)
    ws = wb.worksheets[0]

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None:
            headers[str(v).strip().upper()] = c

    col_equiv = _buscar_columna(headers, "EQUIVALENCIA")
    col_sku = _buscar_columna(headers, "SKU")
    col_color = _buscar_columna(headers, "COLOR")
    col_talle = _buscar_columna(headers, "TALLE")

    faltantes = [
        nombre
        for nombre, col in [
            ("EQUIVALENCIA", col_equiv),
            ("SKU", col_sku),
            ("COLOR", col_color),
            ("TALLE", col_talle),
        ]
        if col is None
    ]
    if faltantes:
        raise ValueError(
            "No encontré la columna "
            + ", ".join(faltantes)
            + " en la fila 1 de la planilla. Los encabezados deben llamarse exactamente así."
        )

    for row in range(2, ws.max_row + 1):
        sku = ws.cell(row=row, column=col_sku).value
        color = ws.cell(row=row, column=col_color).value
        talle = ws.cell(row=row, column=col_talle).value
        if sku is None and color is None:
            continue
        resultado = _buscar_codigo(lookup, sku, color, talle)
        cell = ws.cell(row=row, column=col_equiv, value=resultado)
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

