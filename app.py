import io
import os
import unicodedata
from collections import defaultdict
from datetime import datetime

import pandas as pd
import openpyxl
import barcode as barcode_lib
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont
from flask import Flask, jsonify, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "equivalencia-app-secret")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR") or os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

GRANDE_CACHE_PATH_XLS = os.path.join(DATA_DIR, "planilla_grande.xls")
GRANDE_CACHE_PATH_XLSX = os.path.join(DATA_DIR, "planilla_grande.xlsx")
GRANDE_CACHE_META = os.path.join(DATA_DIR, "planilla_grande.meta")
STOCK_CACHE_PATH = os.path.join(DATA_DIR, "stock.xlsx")
STOCK_CACHE_META = os.path.join(DATA_DIR, "stock.meta")

ALLOWED_EXT = {".xls", ".xlsx"}
FONT_BOLD = os.path.join(BASE_DIR, "fonts", "DejaVuSans-Bold.ttf")
FONT_REGULAR = os.path.join(BASE_DIR, "fonts", "DejaVuSans.ttf")


def _ext(filename):
    return os.path.splitext(filename)[1].lower()


def _encabezado(valor):
    texto = unicodedata.normalize("NFKD", str(valor).strip()).encode("ascii", "ignore").decode()
    return " ".join(texto.upper().split())


def _texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    texto = str(valor).strip()
    return texto[:-2] if texto.endswith(".0") else texto


def _clave_sku(sku):
    return _texto(sku).upper().replace(" ", "")


def _clave_barra(codigo):
    texto = _texto(codigo).replace(" ", "")
    return str(int(texto)) if texto.isdigit() else texto.upper()


def _buscar_columna(headers, *nombres_posibles):
    for nombre in nombres_posibles:
        if nombre in headers:
            return headers[nombre]
    return None


def _read_grande(path):
    """Lee equivalencias y arma índices por artículo, barra y SKU."""
    ext = _ext(path)
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xl = pd.ExcelFile(path, engine=engine)

    lookup = defaultdict(list)
    barcode_index = {}
    sku_index = defaultdict(list)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None, dtype=str)
        if df.empty:
            continue
        encabezados = {
            _encabezado(valor): columna
            for columna, valor in enumerate(df.iloc[0])
            if not pd.isna(valor) and str(valor).strip()
        }
        col_codigo = _buscar_columna(
            encabezados, "CODIGO", "CODIGOS", "CODIGO DE BARRA", "CODIGOS DE BARRA", "BARRA", "BARRAS", "BARCODE"
        )
        col_articulo = _buscar_columna(encabezados, "ARTICULO", "ARTICULOS", "ARTIC", "SKU COLOR", "CONCAT")
        col_sku = _buscar_columna(encabezados, "SKU")
        col_color = _buscar_columna(encabezados, "COLOR")
        col_talle = _buscar_columna(encabezados, "TALLE")
        col_descripcion = _buscar_columna(encabezados, "DESCRIPCION", "NOMBRE", "PRODUCTO")

        tiene_encabezados = col_codigo is not None and (
            col_articulo is not None or (col_sku is not None and col_color is not None)
        )
        if tiene_encabezados:
            df = df.iloc[1:]
        else:
            col_codigo, col_articulo, col_sku, col_color = 0, 1, None, None
            col_descripcion, col_talle = 2, 5

        for _, row in df.iterrows():
            codigo = row.get(col_codigo)
            if pd.isna(codigo):
                continue
            codigo = _texto(codigo)
            if not codigo:
                continue

            sku = _texto(row.get(col_sku)) if col_sku is not None else ""
            color = _texto(row.get(col_color)) if col_color is not None else ""
            if (not sku or not color) and col_articulo is not None and not pd.isna(row.get(col_articulo)):
                partes = str(row.get(col_articulo)).strip().split(maxsplit=1)
                if len(partes) == 2:
                    sku, color = partes
            if not sku or not color:
                continue

            descripcion = _texto(row.get(col_descripcion)) if col_descripcion is not None else ""
            talle = _texto(row.get(col_talle)) if col_talle is not None else ""
            articulo = f"{sku} {color}"
            lookup[articulo].append((codigo, talle, descripcion))
            producto = {
                "codigo": codigo,
                "sku": sku,
                "color": color,
                "talle": talle,
                "descripcion": descripcion,
            }
            barcode_index.setdefault(_clave_barra(codigo), producto)
            sku_index[_clave_sku(sku)].append(producto)
    return lookup, barcode_index, dict(sku_index)


_LOOKUP_CACHE = {"path": None, "mtime": None, "lookup": None, "barcode": None, "by_sku": None}


def _get_lookup(path):
    global _LOOKUP_CACHE
    mtime = os.path.getmtime(path)
    if _LOOKUP_CACHE["path"] == path and _LOOKUP_CACHE["mtime"] == mtime:
        return _LOOKUP_CACHE["lookup"], _LOOKUP_CACHE["barcode"], _LOOKUP_CACHE["by_sku"]
    lookup, barcode_index, sku_index = _read_grande(path)
    _LOOKUP_CACHE = {
        "path": path,
        "mtime": mtime,
        "lookup": lookup,
        "barcode": barcode_index,
        "by_sku": sku_index,
    }
    return lookup, barcode_index, sku_index


_STOCK_CACHE = {"path": None, "mtime": None, "by_sku": None}


def _get_stock_index(path):
    """Lee SKU, COLOR y TALLE de la planilla de stock (sin precios)."""
    global _STOCK_CACHE
    mtime = os.path.getmtime(path)
    if _STOCK_CACHE["path"] == path and _STOCK_CACHE["mtime"] == mtime:
        return _STOCK_CACHE["by_sku"]

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    hoja = workbook["STOCK"] if "STOCK" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    filas = hoja.iter_rows(values_only=True)
    try:
        encabezado = next(filas)
    except StopIteration:
        workbook.close()
        raise ValueError("La planilla de stock está vacía.")

    headers = {_encabezado(valor): i for i, valor in enumerate(encabezado) if valor is not None}
    col_sku = _buscar_columna(headers, "SKU", "ARTICULO", "ARTICULOS")
    col_color = _buscar_columna(headers, "COLOR")
    col_talle = _buscar_columna(headers, "TALLE")
    if col_sku is None:
        col_sku = 1
    if col_color is None:
        col_color = 2
    if col_talle is None:
        col_talle = 3

    by_sku = defaultdict(list)
    vistos = set()
    for row in filas:
        if row is None or len(row) <= max(col_sku, col_color, col_talle):
            continue
        sku = _clave_sku(row[col_sku])
        color = _texto(row[col_color])
        talle = _texto(row[col_talle])
        clave = (sku, color.upper(), talle.upper())
        if not sku or clave in vistos:
            continue
        vistos.add(clave)
        by_sku[sku].append({"sku": sku, "color": color, "talle": talle})
    workbook.close()
    _STOCK_CACHE = {"path": path, "mtime": mtime, "by_sku": dict(by_sku)}
    return _STOCK_CACHE["by_sku"]


def _resolver_grande_cache():
    if os.path.exists(GRANDE_CACHE_PATH_XLS):
        return GRANDE_CACHE_PATH_XLS
    if os.path.exists(GRANDE_CACHE_PATH_XLSX):
        return GRANDE_CACHE_PATH_XLSX
    return None


def _resolver_stock_cache():
    return STOCK_CACHE_PATH if os.path.exists(STOCK_CACHE_PATH) else None


def _leer_meta(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def _guardar_meta(path, filename):
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"{secure_filename(filename)} - cargada {datetime.now().strftime('%d/%m/%Y %H:%M')}")


def _dedupe_codigos(codigos):
    vistos = {}
    for c in codigos:
        try:
            clave = int(c)
        except (TypeError, ValueError):
            clave = c
        if clave not in vistos or len(c) < len(vistos[clave]):
            vistos[clave] = c
    return list(vistos.values())


def _buscar_producto(lookup, sku, color, talle):
    if not sku or not color:
        return None, None
    sku_str = _texto(sku)
    color_str = _texto(color)
    talle_str = _texto(talle)
    key = f"{sku_str} {color_str}"
    candidatos = lookup.get(key, [])
    coincidencias = [
        (codigo, desc) for codigo, t, desc in candidatos if t.strip().upper() == talle_str.upper()
    ]
    if not coincidencias:
        return None, None
    codigo = _dedupe_codigos([c for c, _d in coincidencias])[0]
    return codigo, coincidencias[0][1]


def _generar_imagen_codigo_barra(codigo, module_height_mm=9.0, module_width_mm=0.28):
    codigo = str(codigo).strip()
    solo_digitos = codigo.isdigit()
    if solo_digitos and len(codigo) == 13:
        clase = "ean13"
    elif solo_digitos and len(codigo) == 12:
        clase = "upca"
    elif solo_digitos and len(codigo) == 8:
        clase = "ean8"
    else:
        clase = "code128"

    def _render(tipo):
        BarcodeClass = barcode_lib.get_barcode_class(tipo)
        bc = BarcodeClass(codigo, writer=ImageWriter())
        buf = io.BytesIO()
        bc.write(
            buf,
            options={
                "write_text": False,
                "quiet_zone": 1.0,
                "module_height": module_height_mm,
                "module_width": module_width_mm,
            },
        )
        buf.seek(0)
        return Image.open(buf).convert("L")

    try:
        return _render(clase)
    except Exception:
        return _render("code128")


def _texto_ajustado(draw, texto, font, max_width_px, max_lineas=2):
    palabras = texto.split()
    lineas = []
    actual = ""
    for palabra in palabras:
        prueba = (actual + " " + palabra).strip()
        ancho = draw.textbbox((0, 0), prueba, font=font)[2]
        if ancho <= max_width_px or not actual:
            actual = prueba
        else:
            lineas.append(actual)
            actual = palabra
            if len(lineas) == max_lineas - 1:
                break
    if actual:
        lineas.append(actual)
    lineas = lineas[:max_lineas]
    usado = " ".join(lineas)
    if len(usado) < len(texto):
        ultima = lineas[-1]
        while draw.textbbox((0, 0), ultima + "…", font=font)[2] > max_width_px and len(ultima) > 1:
            ultima = ultima[:-1]
        lineas[-1] = ultima.rstrip() + "…"
    return lineas


def generar_etiqueta(descripcion, color, talle, codigo, sku="", dpi=300):
    """Genera la etiqueta de 4cm x 2cm como imagen PNG (devuelve BytesIO)."""
    px_mm = dpi / 25.4
    ancho_mm, alto_mm = 40.0, 20.0
    W = round(ancho_mm * px_mm)
    H = round(alto_mm * px_mm)

    img = Image.new("L", (W, H), color=255)
    draw = ImageDraw.Draw(img)

    margen = round(1.3 * px_mm)
    col_izq_ancho = round(21.5 * px_mm)

    font_desc = ImageFont.truetype(FONT_REGULAR, size=round(2.3 * px_mm))
    font_color = ImageFont.truetype(FONT_REGULAR, size=round(2.6 * px_mm))
    font_talle = ImageFont.truetype(FONT_BOLD, size=round(4.6 * px_mm))
    font_sku = ImageFont.truetype(FONT_REGULAR, size=round(2.2 * px_mm))
    font_digitos = ImageFont.truetype(FONT_REGULAR, size=round(2.0 * px_mm))

    y = margen
    max_w = col_izq_ancho - margen

    for linea in _texto_ajustado(draw, str(descripcion).upper(), font_desc, max_w, max_lineas=2):
        draw.text((margen, y), linea, font=font_desc, fill=0)
        y += draw.textbbox((0, 0), linea, font=font_desc)[3] + round(0.6 * px_mm)

    y += round(0.8 * px_mm)
    draw.text((margen, y), str(color).upper(), font=font_color, fill=0)
    y += draw.textbbox((0, 0), str(color).upper(), font=font_color)[3] + round(1.0 * px_mm)

    draw.text((margen, y), str(talle).upper(), font=font_talle, fill=0)
    y += draw.textbbox((0, 0), str(talle).upper(), font=font_talle)[3] + round(0.8 * px_mm)

    if sku:
        draw.text((margen, y), str(sku).strip(), font=font_sku, fill=0)

    codigo_str = str(codigo).strip()
    col_der_x = margen + col_izq_ancho + round(1.0 * px_mm)
    col_der_ancho = W - col_der_x - margen

    bc_img = _generar_imagen_codigo_barra(codigo_str)
    escala = col_der_ancho / bc_img.width
    bc_alto = min(round(bc_img.height * escala), round(13 * px_mm))
    bc_img = bc_img.resize((col_der_ancho, bc_alto))

    bc_y = margen
    img.paste(bc_img, (col_der_x, bc_y))

    digitos_y = bc_y + bc_alto + round(0.5 * px_mm)
    bbox = draw.textbbox((0, 0), codigo_str, font=font_digitos)
    digitos_x = col_der_x + max(0, (col_der_ancho - (bbox[2] - bbox[0])) // 2)
    draw.text((digitos_x, digitos_y), codigo_str, font=font_digitos, fill=0)

    salida = io.BytesIO()
    img.save(salida, format="PNG", dpi=(dpi, dpi))
    salida.seek(0)
    return salida


def _error_etiqueta(mensaje, status=400):
    if request.headers.get("X-Requested-With") == "fetch" or "application/json" in request.headers.get("Accept", ""):
        return jsonify({"error": mensaje}), status
    flash(mensaje)
    return redirect(url_for("index"))


def _producto_desde_stock(sku, color, talle, lookup, barcode_index, by_sku_grande, by_sku_stock):
    sku = _clave_sku(sku)
    color = _texto(color)
    talle = _texto(talle)

    if color and talle:
        codigo, descripcion = _buscar_producto(lookup, sku, color, talle)
        if not codigo:
            return None, _error_etiqueta(f"No encontré SKU {sku}, color {color} y talle {talle} en equivalencia.")
        return {
            "sku": sku,
            "color": color,
            "talle": talle,
            "codigo": codigo,
            "descripcion": descripcion or "",
        }, None

    opciones = []
    vistos = set()
    candidatos = list(by_sku_grande.get(sku) or [])
    if by_sku_stock and sku in by_sku_stock:
        stock_set = {
            (_texto(op["color"]).upper(), _texto(op["talle"]).upper())
            for op in by_sku_stock[sku]
        }
        filtrados = [
            op for op in candidatos
            if (_texto(op["color"]).upper(), _texto(op["talle"]).upper()) in stock_set
        ]
        if filtrados:
            candidatos = filtrados

    if color:
        candidatos = [op for op in candidatos if _texto(op["color"]).upper() == color.upper()]
    if talle:
        candidatos = [op for op in candidatos if _texto(op["talle"]).upper() == talle.upper()]

    for op in candidatos:
        clave = (op["color"].upper(), op["talle"].upper(), _clave_barra(op["codigo"]))
        if clave in vistos:
            continue
        vistos.add(clave)
        opciones.append({
            "sku": op["sku"],
            "color": op["color"],
            "talle": op["talle"],
            "codigo": op["codigo"],
            "descripcion": op.get("descripcion") or "",
        })

    if not opciones:
        return None, _error_etiqueta(f"No encontré el SKU {sku} en equivalencia" + (" / stock." if by_sku_stock else "."))
    if len(opciones) == 1:
        return opciones[0], None
    return None, jsonify({"opciones": opciones})


def _resolver_producto(datos):
    barcode_scanned = _texto(datos.get("barcode") or "").replace("\r", "").replace("\n", "").replace("\t", "")
    sku = _clave_sku(datos.get("sku") or "")
    color = _texto(datos.get("color") or "")
    talle = _texto(datos.get("talle") or "")

    grande_path = _resolver_grande_cache()
    if not grande_path:
        return None, _error_etiqueta("No hay planilla de equivalencia cargada. Subila primero.")

    try:
        lookup, barcode_index, by_sku_grande = _get_lookup(grande_path)
    except Exception as e:
        return None, _error_etiqueta(f"Error leyendo equivalencia: {e}", 500)

    by_sku_stock = None
    stock_path = _resolver_stock_cache()
    if stock_path:
        try:
            by_sku_stock = _get_stock_index(stock_path)
        except Exception as e:
            return None, _error_etiqueta(f"Error leyendo el stock: {e}", 500)

    if barcode_scanned:
        producto = barcode_index.get(_clave_barra(barcode_scanned))
        if not producto:
            return None, _error_etiqueta(f"No encontré el código {barcode_scanned}. Probá buscarlo por SKU.")
        return producto, None

    if not sku:
        return None, _error_etiqueta("Escaneá un código de barras o ingresá SKU, color y talle.")

    return _producto_desde_stock(sku, color, talle, lookup, barcode_index, by_sku_grande, by_sku_stock)


@app.route("/etiqueta")
def etiqueta():
    return redirect(url_for("index"))


@app.route("/etiqueta/generar", methods=["GET", "POST"])
def etiqueta_generar():
    producto, error = _resolver_producto(request.values)
    if error is not None:
        return error

    try:
        imagen = generar_etiqueta(
            producto.get("descripcion") or "",
            producto["color"],
            producto["talle"],
            producto["codigo"],
            sku=producto["sku"],
        )
    except Exception as e:
        return _error_etiqueta(f"No pude generar la etiqueta: {e}", 500)

    descargar = request.values.get("descargar") == "1"
    nombre = f"etiqueta_{secure_filename(producto['codigo'])}.png"
    return send_file(
        imagen,
        mimetype="image/png",
        as_attachment=descargar,
        download_name=nombre,
    )


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        grande_cached=_resolver_grande_cache() is not None,
        grande_info=_leer_meta(GRANDE_CACHE_META),
        stock_cached=_resolver_stock_cache() is not None,
        stock_info=_leer_meta(STOCK_CACHE_META),
    )


@app.route("/subir-equivalencia", methods=["POST"])
def subir_equivalencia():
    grande_file = request.files.get("grande")
    if not grande_file or grande_file.filename == "":
        flash("Subí la planilla de equivalencia.")
        return redirect(url_for("index"))
    if _ext(grande_file.filename) not in ALLOWED_EXT:
        flash("La planilla de equivalencia debe ser .xls o .xlsx.")
        return redirect(url_for("index"))

    grande_ext = _ext(grande_file.filename)
    grande_path = GRANDE_CACHE_PATH_XLS if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLSX
    other_path = GRANDE_CACHE_PATH_XLSX if grande_ext == ".xls" else GRANDE_CACHE_PATH_XLS
    if os.path.exists(other_path):
        os.remove(other_path)
    grande_file.save(grande_path)

    global _LOOKUP_CACHE
    _LOOKUP_CACHE = {"path": None, "mtime": None, "lookup": None, "barcode": None, "by_sku": None}
    try:
        _get_lookup(grande_path)
    except Exception as e:
        flash(f"La planilla se guardó, pero no pude leerla: {e}")
        return redirect(url_for("index"))

    _guardar_meta(GRANDE_CACHE_META, grande_file.filename)
    flash("Planilla de equivalencia cargada correctamente.")
    return redirect(url_for("index"))


@app.route("/subir-stock", methods=["POST"])
def subir_stock():
    stock_file = request.files.get("stock")
    if not stock_file or stock_file.filename == "":
        flash("Subí la planilla de stock.")
        return redirect(url_for("index"))
    if _ext(stock_file.filename) != ".xlsx":
        flash("La planilla de stock debe ser un archivo .xlsx.")
        return redirect(url_for("index"))

    stock_file.save(STOCK_CACHE_PATH)
    global _STOCK_CACHE
    _STOCK_CACHE = {"path": None, "mtime": None, "by_sku": None}
    try:
        _get_stock_index(STOCK_CACHE_PATH)
    except Exception as e:
        flash(f"La planilla se guardó, pero no pude leerla: {e}")
        return redirect(url_for("index"))

    _guardar_meta(STOCK_CACHE_META, stock_file.filename)
    flash("Planilla de stock cargada correctamente.")
    return redirect(url_for("index"))


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), threaded=True)
