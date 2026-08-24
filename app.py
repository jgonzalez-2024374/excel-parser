from flask import Flask, request, jsonify, send_file, after_this_request
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime
import os
import tempfile
import unicodedata
import re

app = Flask(__name__)


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
TEMPLATE_FILE = os.path.join(TEMPLATES_DIR, "plantilla.xlsx")


# ============================================================
# INICIO
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "ok",
        "message": "Excel Parser funcionando",
        "template_exists": os.path.exists(TEMPLATE_FILE),
        "endpoint": "/process-excel"
    })


# ============================================================
# LIMPIAR TEXTO
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")

    text = unicodedata.normalize(
        "NFKD",
        text
    ).encode(
        "ascii",
        "ignore"
    ).decode(
        "ascii"
    )

    return text.strip().lower()


# ============================================================
# LIMPIAR NÚMEROS
# ============================================================

def clean_number(value):

    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    text = text.replace("Q", "")
    text = text.replace("$", "")
    text = text.replace(" ", "")

    if "," in text and "." in text:

        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)

    except Exception:
        return 0


# ============================================================
# CONVERTIR FECHA
# ============================================================

def clean_date(value):

    if value is None or value == "":
        return ""

    if isinstance(value, datetime):
        return value

    return value


# ============================================================
# DETECTAR COLUMNAS
# ============================================================

def detectar_columnas(rows):

    for row_index, row in enumerate(rows):

        columnas = {}

        for index, value in enumerate(row):

            texto = clean_text(value)

            if not texto:
                continue

            # FECHA
            if (
                texto == "fecha"
                or "fecha de transaccion" in texto
                or "fecha transaccion" in texto
                or texto == "date"
            ):
                columnas.setdefault("fecha", index)

            # REFERENCIA
            if (
                texto == "referencia"
                or "referencia de transaccion" in texto
                or "referencia transaccion" in texto
                or "no doc" in texto
                or "no. doc" in texto
                or "numero documento" in texto
                or "numero de documento" in texto
                or texto == "documento"
            ):
                columnas.setdefault("referencia", index)

            # CODIGO
            if (
                texto == "codigo"
                or "codigo de transaccion" in texto
                or texto == "tt"
                or texto == "secuencial"
                or "tipo transaccion" in texto
                or "tipo de transaccion" in texto
            ):
                columnas.setdefault("codigo", index)

            # DESCRIPCION
            if (
                texto == "descripcion"
                or "descripcion de transaccion" in texto
                or texto == "description"
                or texto == "detalle"
                or texto == "concepto"
                or texto == "movimiento"
            ):
                columnas.setdefault("descripcion", index)

            # DEBITO
            if (
                texto == "debito"
                or texto == "debito (-)"
                or texto == "debe"
                or "debito de transaccion" in texto
                or texto == "debit"
                or texto == "cargo"
                or texto == "retiro"
            ):
                columnas.setdefault("debito", index)

            # CREDITO
            if (
                texto == "credito"
                or texto == "credito (+)"
                or texto == "haber"
                or "credito de transaccion" in texto
                or texto == "credit"
                or texto == "abono"
                or texto == "deposito"
            ):
                columnas.setdefault("credito", index)

            # SALDO
            if (
                texto == "saldo"
                or texto == "balance"
                or texto == "saldo contable"
                or texto == "saldo disponible"
                or "balance de transaccion" in texto
            ):

                if texto == "saldo contable":
                    columnas["saldo"] = index

                elif "saldo" not in columnas:
                    columnas["saldo"] = index

        # ====================================================
        # VALIDAR ENCABEZADO
        # ====================================================

        tiene_descripcion = "descripcion" in columnas
        tiene_debito = "debito" in columnas
        tiene_credito = "credito" in columnas
        tiene_saldo = "saldo" in columnas
        tiene_referencia = "referencia" in columnas
        tiene_fecha = "fecha" in columnas

        if (
            tiene_descripcion
            and tiene_debito
            and tiene_credito
            and tiene_saldo
            and (
                tiene_fecha
                or tiene_referencia
            )
        ):
            return row_index, columnas

    return None, {}


# ============================================================
# OBTENER VALOR DE COLUMNA
# ============================================================

def get_column(row, columns, name):

    index = columns.get(name)

    if index is None:
        return ""

    if index >= len(row):
        return ""

    return row[index]


# ============================================================
# EXTRAER TRANSACCIÓN
# ============================================================

def extraer_transaccion(row, columns):

    fecha = get_column(row, columns, "fecha")
    referencia = get_column(row, columns, "referencia")
    codigo = get_column(row, columns, "codigo")
    descripcion = get_column(row, columns, "descripcion")
    debito = get_column(row, columns, "debito")
    credito = get_column(row, columns, "credito")
    saldo = get_column(row, columns, "saldo")

    if not descripcion and not referencia:
        return None

    texto = clean_text(descripcion)

    palabras_excluir = [
        "total",
        "resumen",
        "saldo inicial",
        "saldo anterior",
        "saldo final",
        "totales"
    ]

    if any(palabra in texto for palabra in palabras_excluir):
        return None

    return {
        "fecha": clean_date(fecha),
        "referencia": str(referencia) if referencia else "",
        "codigo": str(codigo) if codigo else "",
        "descripcion": str(descripcion) if descripcion else "",
        "debito": clean_number(debito),
        "credito": clean_number(credito),
        "saldo": clean_number(saldo)
    }


# ============================================================
# PROCESAR HOJA
# ============================================================

def procesar_hoja(worksheet):

    rows = []

    for row in worksheet.iter_rows(values_only=True):

        values = [
            value if value is not None else ""
            for value in row
        ]

        if any(value != "" for value in values):
            rows.append(values)

    if not rows:
        return []

    header_index, columns = detectar_columnas(rows)

    if header_index is None:
        return []

    transactions = []

    for row in rows[header_index + 1:]:

        transaction = extraer_transaccion(
            row,
            columns
        )

        if transaction:
            transactions.append(transaction)

    return transactions


# ============================================================
# INSERTAR DATOS EN PLANTILLA
# ============================================================

def insertar_en_plantilla(transactions, output_path):

    if not os.path.exists(TEMPLATE_FILE):

        raise Exception(
            "No existe la plantilla: "
            + TEMPLATE_FILE
        )

    # Abrir plantilla
    workbook = load_workbook(TEMPLATE_FILE)

    # ========================================================
    # USAR LA PRIMERA HOJA DE LA PLANTILLA
    # ========================================================

    worksheet = workbook.active

    # ========================================================
    # BUSCAR ENCABEZADOS DE LA PLANTILLA
    # ========================================================

    headers = {}

    for row in worksheet.iter_rows():

        for cell in row:

            texto = clean_text(cell.value)

            if not texto:
                continue

            if texto == "fecha":
                headers["fecha"] = cell.column

            elif texto == "referencia":
                headers["referencia"] = cell.column

            elif texto == "codigo" or texto == "código":
                headers["codigo"] = cell.column

            elif texto == "descripcion" or texto == "descripción":
                headers["descripcion"] = cell.column

            elif (
                texto == "debito"
                or texto == "débito"
                or texto == "debito (-)"
            ):
                headers["debito"] = cell.column

            elif (
                texto == "credito"
                or texto == "crédito"
                or texto == "credito (+)"
            ):
                headers["credito"] = cell.column

            elif texto == "saldo":
                headers["saldo"] = cell.column

    # ========================================================
    # DETERMINAR FILA DE INICIO
    # ========================================================

    if not headers:

        raise Exception(
            "No se encontraron encabezados reconocibles "
            "en la plantilla."
        )

    header_row = 1

    for row in worksheet.iter_rows():

        found = False

        for cell in row:

            texto = clean_text(cell.value)

            if texto in [
                "fecha",
                "referencia",
                "descripcion",
                "descripción",
                "debito",
                "débito",
                "credito",
                "crédito",
                "saldo"
            ]:
                found = True
                break

        if found:
            header_row = row[0].row
        else:
            if found:
                break

    start_row = header_row + 1

    # ========================================================
    # INSERTAR TRANSACCIONES
    # ========================================================

    for transaction in transactions:

        row_number = start_row

        # Buscar siguiente fila disponible
        while any(
            worksheet.cell(
                row=row_number,
                column=column
            ).value not in (None, "")
            for column in headers.values()
        ):
            row_number += 1

        if "fecha" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["fecha"]
            ).value = transaction["fecha"]

        if "referencia" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["referencia"]
            ).value = transaction["referencia"]

        if "codigo" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["codigo"]
            ).value = transaction["codigo"]

        if "descripcion" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["descripcion"]
            ).value = transaction["descripcion"]

        if "debito" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["debito"]
            ).value = transaction["debito"]

        if "credito" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["credito"]
            ).value = transaction["credito"]

        if "saldo" in headers:
            worksheet.cell(
                row=row_number,
                column=headers["saldo"]
            ).value = transaction["saldo"]

    # ========================================================
    # GUARDAR
    # ========================================================

    workbook.save(output_path)

    workbook.close()


# ============================================================
# PROCESAR EXCEL COMPLETO
# ============================================================

@app.route("/process-excel", methods=["POST"])
def process_excel():

    if "file" not in request.files:

        return jsonify({
            "success": False,
            "error": "No se recibió ningún archivo"
        }), 400

    file = request.files["file"]

    if not file.filename:

        return jsonify({
            "success": False,
            "error": "El archivo no tiene nombre"
        }), 400

    if not file.filename.lower().endswith(".xlsx"):

        return jsonify({
            "success": False,
            "error": "El archivo debe ser .xlsx"
        }), 400

    if not os.path.exists(TEMPLATE_FILE):

        return jsonify({
            "success": False,
            "error": "La plantilla no existe en el servidor",
            "template": TEMPLATE_FILE
        }), 500

    input_path = None
    output_path = None

    try:

        # ====================================================
        # GUARDAR EXCEL RECIBIDO
        # ====================================================

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            file.save(temp.name)
            input_path = temp.name

        # ====================================================
        # ABRIR EXCEL
        # ====================================================

        workbook = load_workbook(
            input_path,
            data_only=True,
            read_only=True
        )

        all_transactions = []

        processed_sheets = []

        # ====================================================
        # PROCESAR TODAS LAS HOJAS
        # ====================================================

        for worksheet in workbook.worksheets:

            transactions = procesar_hoja(
                worksheet
            )

            if transactions:

                all_transactions.extend(
                    transactions
                )

                processed_sheets.append({
                    "name": worksheet.title,
                    "transactions": len(transactions)
                })

        workbook.close()

        # ====================================================
        # VALIDAR
        # ====================================================

        if not all_transactions:

            return jsonify({
                "success": False,
                "error": (
                    "No se encontraron movimientos "
                    "bancarios compatibles"
                ),
                "sheets": [
                    sheet["name"]
                    for sheet in processed_sheets
                ]
            }), 400

        # ====================================================
        # CREAR ARCHIVO TEMPORAL FINAL
        # ====================================================

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            output_path = temp.name

        # ====================================================
        # INSERTAR EN PLANTILLA
        # ====================================================

        insertar_en_plantilla(
            all_transactions,
            output_path
        )

        # ====================================================
        # ELIMINAR INPUT
        # ====================================================

        if input_path and os.path.exists(input_path):

            os.remove(input_path)
            input_path = None

        # ====================================================
        # ELIMINAR OUTPUT DESPUÉS DE ENVIARLO
        # ====================================================

        @after_this_request
        def cleanup(response):

            try:

                if output_path and os.path.exists(output_path):
                    os.remove(output_path)

            except Exception:
                pass

            return response

        # ====================================================
        # DEVOLVER EXCEL
        # ====================================================

        return send_file(
            output_path,
            as_attachment=True,
            download_name="resultado_normalizado.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

    finally:

        if input_path and os.path.exists(input_path):

            try:
                os.remove(input_path)
            except Exception:
                pass


# ============================================================
# SERVIDOR
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get("PORT", 5000)
        ),
        debug=False
    )