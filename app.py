from flask import Flask, request, jsonify
from openpyxl import load_workbook
import os
import tempfile
import unicodedata
import re

app = Flask(__name__)


# ============================================================
# 1. INICIO
# ============================================================

@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "ok",
        "message": "Excel Parser funcionando"
    })


# ============================================================
# 2. LEER EL EXCEL COMPLETO
# ============================================================

@app.route("/parse-excel", methods=["POST"])
def parse_excel():

    if "file" not in request.files:
        return jsonify({
            "success": False,
            "error": "No se recibió ningún archivo"
        }), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({
            "success": False,
            "error": "El archivo no tiene nombre"
        }), 400

    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({
            "success": False,
            "error": "El archivo debe ser .xlsx"
        }), 400

    temp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            file.save(temp.name)
            temp_path = temp.name

        workbook = load_workbook(
            temp_path,
            data_only=True,
            read_only=True
        )

        sheets = []

        for worksheet in workbook.worksheets:

            rows = []

            for row in worksheet.iter_rows(values_only=True):

                rows.append([
                    value if value is not None else ""
                    for value in row
                ])

            sheets.append({
                "name": worksheet.title,
                "rows": rows
            })

        workbook.close()

        return jsonify({
            "success": True,
            "filename": file.filename,
            "sheet_count": len(sheets),
            "sheets": sheets
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

    finally:

        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


# ============================================================
# 3. NORMALIZAR AUTOMÁTICAMENTE
# ============================================================

@app.route("/normalize-bank", methods=["POST"])
def normalize_bank():

    try:

        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "error": "No se recibió JSON"
            }), 400

        rows = data.get("rows", [])

        if not isinstance(rows, list):

            return jsonify({
                "success": False,
                "error": "rows debe ser un arreglo"
            }), 400

        if not rows:

            return jsonify({
                "success": False,
                "error": "La hoja no contiene datos"
            }), 400


        # ----------------------------------------------------
        # BUSCAR AUTOMÁTICAMENTE EL ENCABEZADO
        # ----------------------------------------------------

        header_index, columns = detectar_columnas(rows)


        if header_index is None:

            return jsonify({
                "success": False,
                "error": "No se pudo detectar automáticamente el encabezado de movimientos"
            }), 400


        # ----------------------------------------------------
        # NORMALIZAR
        # ----------------------------------------------------

        transactions = []

        for row in rows[header_index + 1:]:

            transaction = extraer_transaccion(
                row,
                columns
            )

            if transaction is not None:

                transactions.append(transaction)


        return jsonify({

            "success": True,

            "formato_detectado": {
                "fila_encabezado": header_index,
                "columnas": columns
            },

            "transaction_count": len(transactions),

            "transactions": transactions

        })


    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# ============================================================
# 4. LIMPIAR TEXTO
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
# 5. LIMPIAR NÚMEROS
# ============================================================

def clean_number(value):

    if value is None or value == "":
        return 0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    text = text.replace("Q", "")
    text = text.replace("$", "")
    text = text.replace(",", "")
    text = text.replace(" ", "")

    try:
        return float(text)

    except:
        return 0


# ============================================================
# 6. NORMALIZAR TRANSACCIÓN
# ============================================================

def normalize_transaction(
    fecha="",
    referencia="",
    codigo="",
    descripcion="",
    debito=0,
    credito=0,
    saldo=0
):

    return {

        "fecha": str(fecha) if fecha else "",

        "referencia": str(referencia)
        if referencia else "",

        "codigo": str(codigo)
        if codigo else "",

        "descripcion": str(descripcion)
        if descripcion else "",

        "debito": clean_number(debito),

        "credito": clean_number(credito),

        "saldo": clean_number(saldo)

    }


# ============================================================
# 7. DETECTAR COLUMNAS AUTOMÁTICAMENTE
# ============================================================

def detectar_columnas(rows):

    for i, row in enumerate(rows):

        columnas = {}

        for index, value in enumerate(row):

            texto = clean_text(value)

            if not texto:
                continue


            # -------------------------------
            # FECHA
            # -------------------------------

            if any(palabra in texto for palabra in [

                "fecha",
                "date",
                "fecha transaccion",
                "fecha de transaccion"

            ]):

                columnas.setdefault("fecha", index)


            # -------------------------------
            # REFERENCIA
            # -------------------------------

            if any(palabra in texto for palabra in [

                "referencia",
                "reference",
                "no doc",
                "numero documento",
                "numero de documento"

            ]):

                columnas.setdefault("referencia", index)


            # -------------------------------
            # CÓDIGO
            # -------------------------------

            if any(palabra in texto for palabra in [

                "codigo",
                "code",
                "tt",
                "tipo transaccion",
                "tipo de transaccion",
                "secuencial"

            ]):

                columnas.setdefault("codigo", index)


            # -------------------------------
            # DESCRIPCIÓN
            # -------------------------------

            if any(palabra in texto for palabra in [

                "descripcion",
                "description",
                "detalle",
                "concepto",
                "movimiento"

            ]):

                columnas.setdefault("descripcion", index)


            # -------------------------------
            # DÉBITO
            # -------------------------------

            if any(palabra in texto for palabra in [

                "debito",
                "debe",
                "debit",
                "cargo",
                "retiro"

            ]):

                columnas.setdefault("debito", index)


            # -------------------------------
            # CRÉDITO
            # -------------------------------

            if any(palabra in texto for palabra in [

                "credito",
                "haber",
                "credit",
                "abono",
                "deposito"

            ]):

                columnas.setdefault("credito", index)


            # -------------------------------
            # SALDO
            # -------------------------------

            if any(palabra in texto for palabra in [

                "saldo",
                "balance",
                "saldo contable",
                "saldo disponible"

            ]):

                # Preferimos saldo contable
                if "saldo contable" in texto:

                    columnas["saldo"] = index

                elif "saldo" not in columnas:

                    columnas["saldo"] = index


        # ----------------------------------------------------
        # VALIDAR QUE SEA UN ENCABEZADO REAL
        # ----------------------------------------------------

        tiene_descripcion = "descripcion" in columnas

        tiene_saldo = "saldo" in columnas

        tiene_movimiento = (
            "debito" in columnas or
            "credito" in columnas
        )

        tiene_referencia = "referencia" in columnas

        tiene_fecha = "fecha" in columnas


        if (
            tiene_descripcion
            and tiene_saldo
            and tiene_movimiento
            and (
                tiene_fecha
                or tiene_referencia
            )
        ):

            return i, columnas


    return None, {}


# ============================================================
# 8. EXTRAER TRANSACCIÓN
# ============================================================

def extraer_transaccion(row, columns):

    def get_column(nombre):

        index = columns.get(nombre)

        if index is None:
            return ""

        if index >= len(row):
            return ""

        return row[index]


    fecha = get_column("fecha")

    referencia = get_column("referencia")

    codigo = get_column("codigo")

    descripcion = get_column("descripcion")

    debito = get_column("debito")

    credito = get_column("credito")

    saldo = get_column("saldo")


    # --------------------------------------------------------
    # EVITAR FILAS VACÍAS
    # --------------------------------------------------------

    if not descripcion and not referencia:

        return None


    # --------------------------------------------------------
    # EVITAR FILAS QUE NO SEAN TRANSACCIONES
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # CREAR TRANSACCIÓN NORMALIZADA
    # --------------------------------------------------------

    return normalize_transaction(

        fecha=fecha,

        referencia=referencia,

        codigo=codigo,

        descripcion=descripcion,

        debito=debito,

        credito=credito,

        saldo=saldo

    )


# ============================================================
# 9. INICIAR SERVIDOR
# ============================================================

if __name__ == "__main__":

    app.run(

        host="0.0.0.0",

        port=5000,

        debug=True

    )