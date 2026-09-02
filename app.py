import io
import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
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


def _read_pedido(path):
    ext = _ext(path)
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df = pd.read_excel(path, engine=engine)
    return df


def _procesar(grande_path, pedido_path):
    lookup = _read_grande(grande_path)
    pedido = _read_pedido(pedido_path)

    if pedido.shape[1] < 2:
        raise ValueError("La planilla de pedido necesita al menos dos columnas (SKU y Color).")

    codigos_col = []
    cantidad_col = []
    for _, row in pedido.iterrows():
        sku = row.iloc[0]
        color = row.iloc[1]
        if pd.isna(sku) or pd.isna(color):
            codigos_col.append("")
            cantidad_col.append(0)
            continue
        sku_str = str(sku).strip()
        # Si el SKU vino como número con .0 (float), lo normalizamos a entero
        try:
            if sku_str.endswith(".0"):
                sku_str = sku_str[:-2]
        except Exception:
            pass
        color_str = str(color).strip()
        key = f"{sku_str} {color_str}"
        candidatos = lookup.get(key, [])
        cantidad_col.append(len(candidatos))
        if not candidatos:
            codigos_col.append("SIN COINCIDENCIA")
        elif len(candidatos) == 1:
            codigos_col.append(candidatos[0][0])
        else:
            texto = ", ".join(
                f"{codigo} ({talle})" if talle else codigo for codigo, talle in candidatos
            )
            codigos_col.append(texto)

    resultado = pedido.copy()
    resultado["Código(s) de Barra"] = codigos_col
    resultado["Cantidad de Coincidencias"] = cantidad_col
    return resultado


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

    if _ext(pedido_file.filename) not in ALLOWED_EXT:
        flash("La planilla de pedido debe ser .xls o .xlsx")
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
        resultado = _procesar(grande_path, pedido_bytes)
    except Exception as e:
        flash(f"Error procesando los archivos: {e}")
        return redirect(url_for("index"))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultado.to_excel(writer, index=False, sheet_name="Equivalencia")
    output.seek(0)

    nombre_salida = f"Equivalencia_{secure_filename(os.path.splitext(pedido_file.filename)[0])}.xlsx"
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
