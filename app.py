from flask import Flask, request, jsonify, send_file, after_this_request
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy
from datetime import datetime, date
import os
import tempfile
import unicodedata
import re

app = Flask(__name__)


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATES_DIR = os.path.join(
    BASE_DIR,
    "templates"
)

TEMPLATE_FILE = os.path.join(
    TEMPLATES_DIR,
    "Consolidación Estados Financieros e Ingresos - Plantilla.xlsx"
)

OUTPUT_FILENAME = "Consolidación Estados Financieros e Ingresos.xlsx"

SHEET_SALDOS = "SALDOS POR CUENTA"
SHEET_ESTADOS = "ESTADOS CONSOLIDADOS"
SHEET_REPORTE = "REPORTE CREDITOS DIARIOS"
SHEET_TABLERO = "TABLERO CREDITOS"


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
# UTILIDADES PARA CELDAS COMBINADAS
# ============================================================

def es_celda_combinada(cell):
    """
    Devuelve True si la celda es una MergedCell.
    """
    return isinstance(cell, MergedCell)


def escribir_celda_segura(worksheet, row, column, value):
    """
    Escribe en una celda solamente si NO es una MergedCell.

    Esto evita:
        'MergedCell' object attribute error
    """
    cell = worksheet.cell(
        row=row,
        column=column
    )

    if es_celda_combinada(cell):
        return False

    cell.value = value
    return True


def limpiar_celda_segura(worksheet, row, column):
    """
    Limpia una celda solamente si es editable.
    """
    cell = worksheet.cell(
        row=row,
        column=column
    )

    if es_celda_combinada(cell):
        return False

    cell.value = None
    return True


def obtener_celda_segura(worksheet, row, column):
    """
    Devuelve la celda si es editable.
    """
    cell = worksheet.cell(
        row=row,
        column=column
    )

    if es_celda_combinada(cell):
        return None

    return cell


def obtener_column_letter_seguro(column):
    """
    IMPORTANTE:
    No usar:

        worksheet.cell(...).column_letter

    porque puede devolver MergedCell.

    Usamos directamente get_column_letter().
    """
    return get_column_letter(column)


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
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    text = text.replace("Q", "")
    text = text.replace("$", "")
    text = text.replace("€", "")
    text = text.replace(" ", "")

    negative = False

    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    if "," in text and "." in text:

        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    elif "," in text:

        text = text.replace(",", ".")

    try:

        number = float(text)

        if negative:
            number *= -1

        return number

    except Exception:

        return 0.0


# ============================================================
# CONVERTIR FECHA
# ============================================================

def clean_date(value):

    if value is None or value == "":
        return ""

    if isinstance(value, datetime):
        return value

    if isinstance(value, date):

        return datetime(
            value.year,
            value.month,
            value.day
        )

    if isinstance(value, (int, float)):

        try:

            from openpyxl.utils.datetime import from_excel

            return from_excel(value)

        except Exception:

            return value

    text = str(value).strip()

    formatos = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
        "%d-%m-%y",
        "%m/%d/%Y",
        "%m-%d-%Y"
    ]

    for formato in formatos:

        try:

            return datetime.strptime(
                text,
                formato
            )

        except Exception:

            pass

    return value


# ============================================================
# DETECTAR BANCO
# ============================================================

def detectar_banco(sheet_name, rows):

    texto_hoja = clean_text(sheet_name)

    bancos = [

        (
            [
                "bi",
                "banco industrial",
                "industrial"
            ],
            "BANCO INDUSTRIAL"
        ),

        (
            [
                "g&t",
                "g t",
                "gyt",
                "g&t continental",
                "gyt continental"
            ],
            "G&T"
        ),

        (
            [
                "bac",
                "bac credomatic"
            ],
            "BAC"
        ),

        (
            [
                "ban",
                "banrural",
                "ban rural"
            ],
            "BANRURAL"
        ),

        (
            [
                "agricola",
                "agrícola",
                "banco agricola",
                "banco agrícola"
            ],
            "BANCO AGRÍCOLA"
        ),

        (
            [
                "cuscatlan",
                "cuscatlán"
            ],
            "CUSCATLÁN"
        ),

        (
            [
                "promerica"
            ],
            "PROMERICA"
        ),

        (
            [
                "davivienda"
            ],
            "DAVIVIENDA"
        )
    ]

    # --------------------------------------------------------
    # Primero por nombre de hoja
    # --------------------------------------------------------

    for aliases, nombre in bancos:

        for alias in aliases:

            if alias == "bi":

                if re.search(
                    r"\bbi\b",
                    texto_hoja
                ):
                    return nombre

            elif alias in texto_hoja:

                return nombre

    # --------------------------------------------------------
    # Buscar en primeras filas
    # --------------------------------------------------------

    for row in rows[:15]:

        for value in row:

            texto = clean_text(value)

            for aliases, nombre in bancos:

                for alias in aliases:

                    if alias == "bi":

                        if re.search(
                            r"\bbi\b",
                            texto
                        ):
                            return nombre

                    elif alias in texto:

                        return nombre

    return sheet_name.strip()


# ============================================================
# DETECTAR CUENTA
# ============================================================

def detectar_cuenta(
    sheet_name,
    rows,
    header_index=None
):

    patrones = [

        r"cuenta\s*(?:no\.?|numero|número)?\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",

        r"no\.?\s*cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",

        r"numero\s+de\s+cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",

        r"número\s+de\s+cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})"
    ]

    limite = 20

    if header_index is not None:

        limite = min(
            header_index + 1,
            len(rows)
        )

    for row in rows[:limite]:

        texto_row = " ".join(
            str(value)
            for value in row
            if value not in (None, "")
        )

        texto_normalizado = clean_text(
            texto_row
        )

        for patron in patrones:

            match = re.search(
                patron,
                texto_normalizado,
                re.IGNORECASE
            )

            if match:

                cuenta = match.group(1)

                cuenta = re.sub(
                    r"[^0-9]",
                    "",
                    cuenta
                )

                if cuenta:
                    return cuenta

    # Buscar número en nombre de hoja

    match = re.search(
        r"(?<!\d)(\d{3,20})(?!\d)",
        str(sheet_name)
    )

    if match:

        return match.group(1)

    return ""


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
                or texto == "fecha movimiento"
                or texto == "fecha de movimiento"
            ):

                columnas.setdefault(
                    "fecha",
                    index
                )

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
                or texto == "no documento"
                or texto == "no. documento"
            ):

                columnas.setdefault(
                    "referencia",
                    index
                )

            # CODIGO

            if (
                texto == "codigo"
                or "codigo de transaccion" in texto
                or texto == "tt"
                or texto == "secuencial"
                or "tipo transaccion" in texto
                or "tipo de transaccion" in texto
            ):

                columnas.setdefault(
                    "codigo",
                    index
                )

            # DESCRIPCION

            if (
                texto == "descripcion"
                or "descripcion de transaccion" in texto
                or texto == "description"
                or texto == "detalle"
                or texto == "concepto"
                or texto == "movimiento"
                or "descripcion movimiento" in texto
            ):

                columnas.setdefault(
                    "descripcion",
                    index
                )

            # DEBITO

            if (
                texto == "debito"
                or texto == "debito (-)"
                or texto == "debe"
                or "debito de transaccion" in texto
                or texto == "debit"
                or texto == "cargo"
                or texto == "retiro"
                or texto == "cargos"
                or texto == "cargo / debito"
            ):

                columnas.setdefault(
                    "debito",
                    index
                )

            # CREDITO

            if (
                texto == "credito"
                or texto == "credito (+)"
                or texto == "haber"
                or "credito de transaccion" in texto
                or texto == "credit"
                or texto == "abono"
                or texto == "deposito"
                or texto == "abonos"
                or texto == "abono / credito"
            ):

                columnas.setdefault(
                    "credito",
                    index
                )

            # SALDO

            if (
                texto == "saldo"
                or texto == "balance"
                or texto == "saldo contable"
                or texto == "saldo disponible"
                or texto == "saldo disponible / saldo / balance"
                or "balance de transaccion" in texto
            ):

                if texto == "saldo contable":

                    columnas["saldo"] = index

                elif "saldo" not in columnas:

                    columnas["saldo"] = index

        tiene_descripcion = (
            "descripcion" in columnas
        )

        tiene_debito = (
            "debito" in columnas
        )

        tiene_credito = (
            "credito" in columnas
        )

        tiene_saldo = (
            "saldo" in columnas
        )

        tiene_referencia = (
            "referencia" in columnas
        )

        tiene_fecha = (
            "fecha" in columnas
        )

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

def extraer_transaccion(
    row,
    columns,
    banco,
    cuenta
):

    fecha = get_column(
        row,
        columns,
        "fecha"
    )

    referencia = get_column(
        row,
        columns,
        "referencia"
    )

    codigo = get_column(
        row,
        columns,
        "codigo"
    )

    descripcion = get_column(
        row,
        columns,
        "descripcion"
    )

    debito = get_column(
        row,
        columns,
        "debito"
    )

    credito = get_column(
        row,
        columns,
        "credito"
    )

    saldo = get_column(
        row,
        columns,
        "saldo"
    )

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

    if any(
        palabra in texto
        for palabra in palabras_excluir
    ):
        return None

    fecha = clean_date(fecha)

    if (
        fecha == ""
        and not referencia
        and not descripcion
    ):
        return None

    return {
        "banco": banco,
        "cuenta": cuenta,
        "fecha": fecha,
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

    for row in worksheet.iter_rows(
        values_only=True
    ):

        values = [
            value if value is not None else ""
            for value in row
        ]

        if any(
            value != ""
            for value in values
        ):

            rows.append(values)

    if not rows:
        return []

    header_index, columns = detectar_columnas(
        rows
    )

    if header_index is None:
        return []

    banco = detectar_banco(
        worksheet.title,
        rows
    )

    cuenta = detectar_cuenta(
        worksheet.title,
        rows,
        header_index
    )

    transactions = []

    for row in rows[
        header_index + 1:
    ]:

        transaction = extraer_transaccion(
            row,
            columns,
            banco,
            cuenta
        )

        if transaction:

            transactions.append(
                transaction
            )

    return transactions


# ============================================================
# COPIAR ESTILO DE FILA
# ============================================================

def copiar_estilo_fila(
    worksheet,
    source_row,
    target_row
):

    for column in range(
        1,
        worksheet.max_column + 1
    ):

        source = worksheet.cell(
            row=source_row,
            column=column
        )

        target = worksheet.cell(
            row=target_row,
            column=column
        )

        # No tocar MergedCell

        if (
            es_celda_combinada(source)
            or es_celda_combinada(target)
        ):
            continue

        if source.has_style:

            target._style = copy(
                source._style
            )

        if source.number_format:

            target.number_format = (
                source.number_format
            )

        if source.alignment:

            target.alignment = copy(
                source.alignment
            )

        if source.protection:

            target.protection = copy(
                source.protection
            )


# ============================================================
# LIMPIAR CONTENIDO
# ============================================================

def limpiar_filas(
    worksheet,
    start_row,
    end_row,
    start_column=1,
    end_column=None
):

    if end_column is None:

        end_column = worksheet.max_column

    for row in range(
        start_row,
        end_row + 1
    ):

        for column in range(
            start_column,
            end_column + 1
        ):

            limpiar_celda_segura(
                worksheet,
                row,
                column
            )


# ============================================================
# ESCRIBIR ESTADOS CONSOLIDADOS
# ============================================================

def escribir_estados_consolidados(
    workbook,
    transactions
):

    if SHEET_ESTADOS not in workbook.sheetnames:

        raise Exception(
            f"No existe la hoja '{SHEET_ESTADOS}'"
        )

    ws = workbook[
        SHEET_ESTADOS
    ]

    start_row = 2

    limpiar_filas(
        ws,
        start_row,
        ws.max_row,
        1,
        8
    )

    required_last_row = (
        start_row
        + len(transactions)
        - 1
    )

    if required_last_row > ws.max_row:

        old_max_row = ws.max_row

        for row in range(
            old_max_row + 1,
            required_last_row + 1
        ):

            copiar_estilo_fila(
                ws,
                start_row,
                row
            )

            ws.row_dimensions[
                row
            ].height = ws.row_dimensions[
                start_row
            ].height

    for index, transaction in enumerate(
        transactions,
        start=start_row
    ):

        escribir_celda_segura(
            ws,
            index,
            1,
            transaction["banco"]
        )

        escribir_celda_segura(
            ws,
            index,
            2,
            transaction["cuenta"]
        )

        cell = obtener_celda_segura(
            ws,
            index,
            3
        )

        if cell:

            cell.value = transaction["fecha"]

            cell.number_format = "dd/mm/yyyy"

        escribir_celda_segura(
            ws,
            index,
            4,
            transaction["referencia"]
        )

        escribir_celda_segura(
            ws,
            index,
            5,
            transaction["descripcion"]
        )

        escribir_celda_segura(
            ws,
            index,
            6,
            transaction["debito"]
        )

        escribir_celda_segura(
            ws,
            index,
            7,
            transaction["credito"]
        )

        escribir_celda_segura(
            ws,
            index,
            8,
            transaction["saldo"]
        )


# ============================================================
# NORMALIZAR NOMBRE BANCO
# ============================================================

def banco_corto(banco):

    texto = clean_text(banco)

    if "agricola" in texto:
        return "AGRÍCOLA"

    if "cuscatlan" in texto:
        return "CUSCATLÁN"

    if "bac" in texto:
        return "BAC"

    if "promerica" in texto:
        return "PROMERICA"

    if "davivienda" in texto:
        return "DAVIVIENDA"

    if "industrial" in texto:
        return "INDUSTRIAL"

    if "g&t" in texto or "gyt" in texto:
        return "G&T"

    if "banrural" in texto or texto == "ban":
        return "BANRURAL"

    return str(banco).strip()


# ============================================================
# ESCRIBIR SALDOS POR CUENTA
# ============================================================

def escribir_saldos_por_cuenta(
    workbook,
    transactions
):

    if SHEET_SALDOS not in workbook.sheetnames:

        raise Exception(
            f"No existe la hoja '{SHEET_SALDOS}'"
        )

    ws = workbook[
        SHEET_SALDOS
    ]

    start_row = 5

    cuentas = {}

    for transaction in transactions:

        banco = transaction["banco"]
        cuenta = transaction["cuenta"]

        clave = (
            banco,
            cuenta
        )

        if clave not in cuentas:

            cuentas[clave] = []

        cuentas[clave].append(
            transaction
        )

    limpiar_filas(
        ws,
        start_row,
        ws.max_row,
        1,
        4
    )

    required_last_row = (
        start_row
        + len(cuentas)
        - 1
    )

    if required_last_row > ws.max_row:

        old_max_row = ws.max_row

        for row in range(
            old_max_row + 1,
            required_last_row + 1
        ):

            copiar_estilo_fila(
                ws,
                start_row,
                row
            )

    row_number = start_row

    for (
        banco,
        cuenta
    ), movimientos in cuentas.items():

        movimientos = sorted(
            movimientos,
            key=lambda x: (
                x["fecha"]
                if isinstance(
                    x["fecha"],
                    datetime
                )
                else datetime.min
            )
        )

        if not movimientos:
            continue

        saldo_inicial = movimientos[0]["saldo"]
        saldo_final = movimientos[-1]["saldo"]

        escribir_celda_segura(
            ws,
            row_number,
            1,
            banco_corto(banco)
        )

        escribir_celda_segura(
            ws,
            row_number,
            2,
            cuenta
        )

        escribir_celda_segura(
            ws,
            row_number,
            3,
            saldo_inicial
        )

        escribir_celda_segura(
            ws,
            row_number,
            4,
            saldo_final
        )

        row_number += 1


# ============================================================
# ESCRIBIR REPORTE CREDITOS DIARIOS
# ============================================================

def escribir_reporte_creditos(
    workbook,
    transactions
):

    if SHEET_REPORTE not in workbook.sheetnames:

        raise Exception(
            f"No existe la hoja '{SHEET_REPORTE}'"
        )

    ws = workbook[
        SHEET_REPORTE
    ]

    data = {}
    cuentas = {}

    for transaction in transactions:

        fecha = transaction["fecha"]

        if not isinstance(
            fecha,
            (datetime, date)
        ):
            continue

        fecha_key = (
            fecha.date()
            if isinstance(
                fecha,
                datetime
            )
            else fecha
        )

        banco = banco_corto(
            transaction["banco"]
        )

        cuenta = str(
            transaction["cuenta"]
        )

        clave = (
            banco,
            cuenta
        )

        cuentas[clave] = True

        if fecha_key not in data:

            data[fecha_key] = {}

        if banco not in data[fecha_key]:

            data[fecha_key][banco] = {}

        if cuenta not in data[
            fecha_key
        ][banco]:

            data[
                fecha_key
            ][banco][cuenta] = 0.0

        data[
            fecha_key
        ][banco][cuenta] += (
            transaction["credito"]
        )

    cuentas_ordenadas = sorted(
        cuentas.keys(),
        key=lambda x: (
            x[0],
            x[1]
        )
    )

    header_bank_row = 3
    header_account_row = 4
    start_row = 5

    limpiar_filas(
        ws,
        header_bank_row,
        ws.max_row,
        1,
        max(
            ws.max_column,
            11
        )
    )

    # --------------------------------------------------------
    # IMPORTANTE:
    # Descombinar únicamente el área dinámica del reporte.
    #
    # Esto evita que intentemos escribir sobre MergedCell.
    # --------------------------------------------------------

    merges_to_remove = []

    for merged_range in list(
        ws.merged_cells.ranges
    ):

        min_col = merged_range.min_col
        max_col = merged_range.max_col
        min_row = merged_range.min_row
        max_row = merged_range.max_row

        # Rangos que afectan filas 3-4 del reporte

        if (
            min_row <= header_account_row
            and max_row >= header_bank_row
            and max_col >= 2
        ):

            merges_to_remove.append(
                str(merged_range)
            )

    for merged_range in merges_to_remove:

        try:

            ws.unmerge_cells(
                merged_range
            )

        except Exception:

            pass

    # Restaurar encabezados

    escribir_celda_segura(
        ws,
        header_bank_row,
        1,
        "BANCO"
    )

    escribir_celda_segura(
        ws,
        header_account_row,
        1,
        "Fecha"
    )

    # --------------------------------------------------------
    # Crear columnas
    # --------------------------------------------------------

    column_map = {}

    current_column = 2

    bancos_en_orden = []

    for banco, cuenta in cuentas_ordenadas:

        if banco not in bancos_en_orden:

            bancos_en_orden.append(
                banco
            )

    for banco in bancos_en_orden:

        cuentas_banco = [
            cuenta
            for (
                banco_actual,
                cuenta
            ) in cuentas_ordenadas
            if banco_actual == banco
        ]

        for cuenta in cuentas_banco:

            column_map[
                (
                    banco,
                    cuenta
                )
            ] = current_column

            escribir_celda_segura(
                ws,
                header_bank_row,
                current_column,
                banco
            )

            escribir_celda_segura(
                ws,
                header_account_row,
                current_column,
                cuenta
            )

            current_column += 1

    total_column = current_column

    cuentas_abono_column = (
        current_column + 1
    )

    escribir_celda_segura(
        ws,
        header_account_row,
        total_column,
        "TOTAL CRÉDITOS"
    )

    escribir_celda_segura(
        ws,
        header_account_row,
        cuentas_abono_column,
        "CUENTAS CON ABONO"
    )

    fechas = sorted(
        data.keys()
    )

    if not fechas:
        return

    required_last_row = (
        start_row
        + len(fechas)
        - 1
    )

    if required_last_row > ws.max_row:

        old_max_row = ws.max_row

        for row in range(
            old_max_row + 1,
            required_last_row + 1
        ):

            copiar_estilo_fila(
                ws,
                start_row,
                row
            )

    # --------------------------------------------------------
    # Escribir datos
    # --------------------------------------------------------

    for index, fecha in enumerate(
        fechas,
        start=start_row
    ):

        fecha_cell = obtener_celda_segura(
            ws,
            index,
            1
        )

        if fecha_cell:

            fecha_cell.value = datetime(
                fecha.year,
                fecha.month,
                fecha.day
            )

            fecha_cell.number_format = (
                "dd/mm/yyyy"
            )

        total = 0.0
        cuentas_con_abono = 0

        for clave, column in column_map.items():

            banco, cuenta = clave

            valor = data.get(
                fecha,
                {}
            ).get(
                banco,
                {}
            ).get(
                cuenta,
                0.0
            )

            if valor != 0:

                escribir_celda_segura(
                    ws,
                    index,
                    column,
                    valor
                )

                cuentas_con_abono += 1

                total += valor

            else:

                limpiar_celda_segura(
                    ws,
                    index,
                    column
                )

        escribir_celda_segura(
            ws,
            index,
            total_column,
            total
        )

        escribir_celda_segura(
            ws,
            index,
            cuentas_abono_column,
            cuentas_con_abono
        )


# ============================================================
# ACTUALIZAR TABLERO
# ============================================================

def actualizar_tablero(
    workbook,
    transactions
):

    if SHEET_TABLERO not in workbook.sheetnames:
        return

    if SHEET_REPORTE not in workbook.sheetnames:
        return

    ws = workbook[
        SHEET_TABLERO
    ]

    reporte = workbook[
        SHEET_REPORTE
    ]

    # --------------------------------------------------------
    # Fechas
    # --------------------------------------------------------

    fechas = [
        transaction["fecha"]
        for transaction in transactions
        if isinstance(
            transaction["fecha"],
            (datetime, date)
        )
    ]

    if fechas:

        fechas = sorted(fechas)

        fecha_min = fechas[0]
        fecha_max = fechas[-1]

        if (
            isinstance(fecha_min, date)
            and not isinstance(
                fecha_min,
                datetime
            )
        ):

            fecha_min = datetime(
                fecha_min.year,
                fecha_min.month,
                fecha_min.day
            )

        if (
            isinstance(fecha_max, date)
            and not isinstance(
                fecha_max,
                datetime
            )
        ):

            fecha_max = datetime(
                fecha_max.year,
                fecha_max.month,
                fecha_max.day
            )

        cell = obtener_celda_segura(
            ws,
            5,
            1
        )

        if cell:

            cell.value = fecha_min
            cell.number_format = "dd/mm/yyyy"

        cell = obtener_celda_segura(
            ws,
            5,
            3
        )

        if cell:

            cell.value = fecha_max
            cell.number_format = "dd/mm/yyyy"

    escribir_celda_segura(
        ws,
        5,
        5,
        "TODAS"
    )

    escribir_celda_segura(
        ws,
        5,
        7,
        "TODOS"
    )

    # --------------------------------------------------------
    # Evolución
    # --------------------------------------------------------

    report_start = 5

    report_end = reporte.max_row

    tablero_start = 15

    for row in range(
        tablero_start,
        max(
            ws.max_row,
            tablero_start + 100
        ) + 1
    ):

        for col in range(1, 5):

            limpiar_celda_segura(
                ws,
                row,
                col
            )

    max_evolution_rows = max(
        19,
        report_end - report_start + 1
    )

    required_end = (
        tablero_start
        + max_evolution_rows
        - 1
    )

    if required_end > ws.max_row:

        old_max = ws.max_row

        for row in range(
            old_max + 1,
            required_end + 1
        ):

            copiar_estilo_fila(
                ws,
                15,
                row
            )

    for i in range(
        max_evolution_rows
    ):

        row = (
            tablero_start + i
        )

        report_row = (
            report_start + i
        )

        if report_row <= report_end:

            # Fecha

            formula_fecha = (
                f'=IF(AND('
                f"'{SHEET_REPORTE}'!A{report_row}>=$A$5,"
                f"'{SHEET_REPORTE}'!A{report_row}<=$C$5),"
                f"'{SHEET_REPORTE}'!A{report_row},"
                f'""'
                f')'
            )

            escribir_celda_segura(
                ws,
                row,
                1,
                formula_fecha
            )

            # Crédito seleccionado

            formula_credito = (
                f'=IF(A{row}="",0,'
                f'IF($E$5<>"TODAS",'
                f'IFERROR('
                f'INDEX('
                f"'{SHEET_REPORTE}'!"
                f'$B$5:$I${report_end},'
                f'MATCH('
                f'A{row},'
                f"'{SHEET_REPORTE}'!"
                f'$A$5:$A${report_end},'
                f'0),'
                f'MATCH('
                f'$E$5,'
                f"'{SHEET_REPORTE}'!"
                f'$B$4:$I$4,'
                f'0)'
                f'),'
                f'0),'
                f'IF($G$5="TODOS",'
                f'IFERROR('
                f'SUMIF('
                f"'{SHEET_REPORTE}'!"
                f'$A$5:$A${report_end},'
                f'A{row},'
                f"'{SHEET_REPORTE}'!"
                f'$J$5:$J${report_end}'
                f'),'
                f'0),'
                f'0)'
                f')'
                f')'
            )

            escribir_celda_segura(
                ws,
                row,
                2,
                formula_credito
            )

            # Variación diaria

            if i == 0:

                escribir_celda_segura(
                    ws,
                    row,
                    3,
                    0
                )

            else:

                escribir_celda_segura(
                    ws,
                    row,
                    3,
                    f'=B{row}-B{row - 1}'
                )

            # Cuentas con abono

            formula_cuentas = (
                f'=IF(A{row}="",0,'
                f'IF($G$5="TODOS",'
                f'IFERROR('
                f'INDEX('
                f"'{SHEET_REPORTE}'!"
                f'$K$5:$K${report_end},'
                f'MATCH('
                f'A{row},'
                f"'{SHEET_REPORTE}'!"
                f'$A$5:$A${report_end},'
                f'0)'
                f'),'
                f'0),'
                f'IF(B{row}>0,1,0)'
                f')'
                f')'
            )

            escribir_celda_segura(
                ws,
                row,
                4,
                formula_cuentas
            )

    # --------------------------------------------------------
    # CRÉDITOS POR CUENTA
    # --------------------------------------------------------

    cuentas = []

    for transaction in transactions:

        clave = (
            banco_corto(
                transaction["banco"]
            ),
            str(
                transaction["cuenta"]
            )
        )

        if clave not in cuentas:

            cuentas.append(
                clave
            )

    cuentas = sorted(
        cuentas,
        key=lambda x: (
            x[0],
            x[1]
        )
    )

    cuenta_start = 15

    for row in range(
        cuenta_start,
        max(
            ws.max_row,
            cuenta_start + 100
        ) + 1
    ):

        for col in range(
            6,
            9
        ):

            limpiar_celda_segura(
                ws,
                row,
                col
            )

    cuenta_end = (
        cuenta_start
        + len(cuentas)
        - 1
    )

    if cuenta_end > ws.max_row:

        old_max = ws.max_row

        for row in range(
            old_max + 1,
            cuenta_end + 1
        ):

            copiar_estilo_fila(
                ws,
                15,
                row
            )

    # --------------------------------------------------------
    # MAPA DE COLUMNAS DEL REPORTE
    # --------------------------------------------------------

    report_columns = {}

    for col in range(
        2,
        reporte.max_column + 1
    ):

        banco_cell = reporte.cell(
            row=3,
            column=col
        )

        cuenta_cell = reporte.cell(
            row=4,
            column=col
        )

        # Si alguno está combinado, ignorar

        if (
            es_celda_combinada(banco_cell)
            or es_celda_combinada(cuenta_cell)
        ):
            continue

        banco = banco_cell.value
        cuenta = cuenta_cell.value

        if banco and cuenta:

            report_columns[
                (
                    clean_text(banco),
                    str(cuenta).strip()
                )
            ] = col

    # --------------------------------------------------------
    # Escribir cuentas
    # --------------------------------------------------------

    for index, (
        banco,
        cuenta
    ) in enumerate(
        cuentas,
        start=cuenta_start
    ):

        escribir_celda_segura(
            ws,
            index,
            6,
            banco
        )

        escribir_celda_segura(
            ws,
            index,
            7,
            cuenta
        )

        report_col = report_columns.get(
            (
                clean_text(banco),
                cuenta
            )
        )

        if report_col:

            # ==================================================
            # CORRECCIÓN PRINCIPAL
            #
            # ANTES:
            # reporte.cell(...).column_letter
            #
            # AHORA:
            # get_column_letter(report_col)
            #
            # Esto evita:
            # 'MergedCell' object has no attribute
            # 'column_letter'
            # ==================================================

            col_letter = get_column_letter(
                report_col
            )

            formula = (
                f'=IF(AND('
                f'OR('
                f'$G$5="TODOS",'
                f'TRIM($G$5)=TRIM(F{index})'
                f'),'
                f'OR('
                f'$E$5="TODAS",'
                f'TEXT($E$5,"0")=TEXT(G{index},"0")'
                f')'
                f'),'
                f'SUMIFS('
                f"'{SHEET_REPORTE}'!"
                f'${col_letter}$5:${col_letter}${report_end},'
                f"'{SHEET_REPORTE}'!"
                f'$A$5:$A${report_end},'
                f'">="&$A$5,'
                f"'{SHEET_REPORTE}'!"
                f'$A$5:$A${report_end},'
                f'"<="&$C$5'
                f'),'
                f'0)'
            )

            escribir_celda_segura(
                ws,
                index,
                8,
                formula
            )


# ============================================================
# PROCESAR PLANTILLA COMPLETA
# ============================================================

def insertar_en_plantilla(
    transactions,
    output_path
):

    if not os.path.exists(
        TEMPLATE_FILE
    ):

        raise Exception(
            "No existe la plantilla: "
            + TEMPLATE_FILE
        )

    workbook = load_workbook(
        TEMPLATE_FILE
    )

    try:

        # 1
        escribir_estados_consolidados(
            workbook,
            transactions
        )

        # 2
        escribir_saldos_por_cuenta(
            workbook,
            transactions
        )

        # 3
        escribir_reporte_creditos(
            workbook,
            transactions
        )

        # 4
        actualizar_tablero(
            workbook,
            transactions
        )

        workbook.save(
            output_path
        )

    finally:

        workbook.close()


# ============================================================
# PROCESAR EXCEL
# ============================================================

@app.route(
    "/process-excel",
    methods=["POST"]
)
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

    if not file.filename.lower().endswith(
        ".xlsx"
    ):

        return jsonify({
            "success": False,
            "error": "El archivo debe ser .xlsx"
        }), 400

    if not os.path.exists(
        TEMPLATE_FILE
    ):

        return jsonify({
            "success": False,
            "error": "La plantilla no existe en el servidor",
            "template": TEMPLATE_FILE
        }), 500

    input_path = None
    output_path = None

    try:

        # ====================================================
        # GUARDAR ARCHIVO
        # ====================================================

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            file.save(
                temp.name
            )

            input_path = temp.name

        # ====================================================
        # ABRIR EXCEL DE ENTRADA
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

            try:

                transactions = procesar_hoja(
                    worksheet
                )

                if transactions:

                    all_transactions.extend(
                        transactions
                    )

                    processed_sheets.append({
                        "name": worksheet.title,
                        "transactions": len(
                            transactions
                        )
                    })

            except Exception as sheet_error:

                # Una hoja problemática no debe tumbar
                # todo el procesamiento.

                processed_sheets.append({
                    "name": worksheet.title,
                    "transactions": 0,
                    "error": str(sheet_error)
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
                "processed_sheets": processed_sheets
            }), 400

        # ====================================================
        # ARCHIVO FINAL TEMPORAL
        # ====================================================

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            output_path = temp.name

        # ====================================================
        # CREAR RESULTADO
        # ====================================================

        insertar_en_plantilla(
            all_transactions,
            output_path
        )

        # ====================================================
        # ELIMINAR INPUT
        # ====================================================

        if (
            input_path
            and os.path.exists(input_path)
        ):

            try:

                os.remove(
                    input_path
                )

                input_path = None

            except Exception:

                pass

        # ====================================================
        # LIMPIAR OUTPUT DESPUÉS DE ENVIAR
        # ====================================================

        @after_this_request
        def cleanup(response):

            try:

                if (
                    output_path
                    and os.path.exists(
                        output_path
                    )
                ):

                    os.remove(
                        output_path
                    )

            except Exception:

                pass

            return response

        # ====================================================
        # DEVOLVER EXCEL
        # ====================================================

        return send_file(
            output_path,
            as_attachment=True,
            download_name=OUTPUT_FILENAME,
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

        if (
            input_path
            and os.path.exists(
                input_path
            )
        ):

            try:

                os.remove(
                    input_path
                )

            except Exception:

                pass


# ============================================================
# SERVIDOR
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),
        debug=False
    )