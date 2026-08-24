```python
from flask import Flask, request, jsonify, send_file, after_this_request
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from copy import copy
from datetime import datetime, date
import os
import tempfile
import unicodedata
import re
import shutil

app = Flask(__name__)


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

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

    # Quitar símbolos
    text = text.replace("Q", "")
    text = text.replace("$", "")
    text = text.replace("€", "")
    text = text.replace(" ", "")

    # Paréntesis = negativo
    negative = False

    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    # Manejar formatos:
    # 1,234.56
    # 1.234,56
    # 1234,56
    # 1234.56

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

        # Excel serial date
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(value)
        except Exception:
            return value

    text = str(value).strip()

    # Intentar fechas comunes
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
            return datetime.strptime(text, formato)
        except Exception:
            pass

    return value


# ============================================================
# DETECTAR BANCO
# ============================================================

def detectar_banco(sheet_name, rows):

    texto_hoja = clean_text(sheet_name)

    # Primero intentar por nombre de hoja
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

    # Evitar que "bi" coincida accidentalmente dentro
    # de otra palabra.
    for aliases, nombre in bancos:

        for alias in aliases:

            if alias == "bi":

                if re.search(r"\bbi\b", texto_hoja):
                    return nombre

            elif alias in texto_hoja:
                return nombre

    # Buscar en las primeras filas
    for row in rows[:15]:

        for value in row:

            texto = clean_text(value)

            for aliases, nombre in bancos:

                for alias in aliases:

                    if alias == "bi":

                        if re.search(r"\bbi\b", texto):
                            return nombre

                    elif alias in texto:
                        return nombre

    # Si no se detecta, usar nombre de hoja
    return sheet_name.strip()


# ============================================================
# DETECTAR CUENTA
# ============================================================

def detectar_cuenta(sheet_name, rows, header_index=None):

    patrones = [
        r"cuenta\s*(?:no\.?|numero|número)?\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",
        r"no\.?\s*cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",
        r"numero\s+de\s+cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",
        r"número\s+de\s+cuenta\s*[:#-]?\s*([0-9][0-9\- ]{2,30})",
    ]

    limite = 20

    if header_index is not None:
        limite = min(header_index + 1, len(rows))

    for row in rows[:limite]:

        texto_row = " ".join(
            str(value)
            for value in row
            if value not in (None, "")
        )

        texto_normalizado = clean_text(texto_row)

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

    # Intentar encontrar un número de cuenta en el nombre
    # de la hoja.
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
                or texto == "no documento"
                or texto == "no. documento"
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
                or "descripcion movimiento" in texto
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
                or texto == "cargos"
                or texto == "cargo / debito"
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
                or texto == "abonos"
                or texto == "abono / credito"
            ):
                columnas.setdefault("credito", index)

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

        tiene_descripcion = "descripcion" in columnas
        tiene_debito = "debito" in columnas
        tiene_credito = "credito" in columnas
        tiene_saldo = "saldo" in columnas
        tiene_referencia = "referencia" in columnas
        tiene_fecha = "fecha" in columnas

        # Condición principal
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

    # Ignorar filas que realmente no tengan
    # información de movimiento.
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
# LIMPIAR CONTENIDO DE FILAS
# SIN ELIMINAR FORMATO
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

            worksheet.cell(
                row=row,
                column=column
            ).value = None


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

    # La plantilla tiene encabezados en fila 1.
    header_row = 1
    start_row = 2

    # Eliminar datos anteriores.
    # No eliminamos filas ni estilos.
    limpiar_filas(
        ws,
        start_row,
        ws.max_row,
        1,
        8
    )

    # Si hay más movimientos que filas
    # existentes, crear nuevas filas copiando
    # el formato de la fila 2.
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

            # Altura de fila
            ws.row_dimensions[
                row
            ].height = ws.row_dimensions[
                start_row
            ].height

    # Escribir movimientos
    for index, transaction in enumerate(
        transactions,
        start=start_row
    ):

        ws.cell(
            index,
            1
        ).value = transaction[
            "banco"
        ]

        ws.cell(
            index,
            2
        ).value = transaction[
            "cuenta"
        ]

        ws.cell(
            index,
            3
        ).value = transaction[
            "fecha"
        ]

        ws.cell(
            index,
            4
        ).value = transaction[
            "referencia"
        ]

        ws.cell(
            index,
            5
        ).value = transaction[
            "descripcion"
        ]

        ws.cell(
            index,
            6
        ).value = transaction[
            "debito"
        ]

        ws.cell(
            index,
            7
        ).value = transaction[
            "credito"
        ]

        ws.cell(
            index,
            8
        ).value = transaction[
            "saldo"
        ]

    # Formato de fecha
    for row in range(
        start_row,
        required_last_row + 1
    ):

        ws.cell(
            row,
            3
        ).number_format = "dd/mm/yyyy"


# ============================================================
# NORMALIZAR NOMBRE DE BANCO
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

    header_row = 4
    start_row = 5

    # Agrupar por banco + cuenta
    cuentas = {}

    for transaction in transactions:

        banco = transaction[
            "banco"
        ]

        cuenta = transaction[
            "cuenta"
        ]

        clave = (
            banco,
            cuenta
        )

        if clave not in cuentas:
            cuentas[clave] = []

        cuentas[clave].append(
            transaction
        )

    # Limpiar contenido existente.
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

    # Escribir una fila por cuenta
    row_number = start_row

    for (
        banco,
        cuenta
    ), movimientos in cuentas.items():

        # Ordenar por fecha
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

        saldo_inicial = movimientos[0][
            "saldo"
        ]

        saldo_final = movimientos[-1][
            "saldo"
        ]

        ws.cell(
            row_number,
            1
        ).value = banco_corto(
            banco
        )

        ws.cell(
            row_number,
            2
        ).value = cuenta

        ws.cell(
            row_number,
            3
        ).value = saldo_inicial

        ws.cell(
            row_number,
            4
        ).value = saldo_final

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

    # --------------------------------------------------------
    # Agrupar:
    #
    # fecha -> banco -> cuenta -> crédito
    # --------------------------------------------------------

    data = {}

    cuentas = {}

    for transaction in transactions:

        fecha = transaction[
            "fecha"
        ]

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
            transaction[
                "banco"
            ]
        )

        cuenta = str(
            transaction[
                "cuenta"
            ]
        )

        clave = (
            banco,
            cuenta
        )

        cuentas[
            clave
        ] = True

        if fecha_key not in data:
            data[
                fecha_key
            ] = {}

        if banco not in data[
            fecha_key
        ]:
            data[
                fecha_key
            ][banco] = {}

        if cuenta not in data[
            fecha_key
        ][banco]:
            data[
                fecha_key
            ][banco][cuenta] = 0.0

        data[
            fecha_key
        ][banco][cuenta] += (
            transaction[
                "credito"
            ]
        )

    # Ordenar cuentas
    cuentas_ordenadas = sorted(
        cuentas.keys(),
        key=lambda x: (
            x[0],
            x[1]
        )
    )

    # --------------------------------------------------------
    # Encabezados existentes
    # --------------------------------------------------------

    header_bank_row = 3
    header_account_row = 4
    start_row = 5

    # Limpiar área
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

    # Restaurar textos principales
    ws.cell(
        header_bank_row,
        1
    ).value = "BANCO"

    ws.cell(
        header_account_row,
        1
    ).value = "Fecha"

    # --------------------------------------------------------
    # Crear columnas dinámicamente
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

            ws.cell(
                header_bank_row,
                current_column
            ).value = banco

            ws.cell(
                header_account_row,
                current_column
            ).value = cuenta

            current_column += 1

    total_column = current_column
    cuentas_abono_column = (
        current_column + 1
    )

    ws.cell(
        header_account_row,
        total_column
    ).value = "TOTAL CRÉDITOS"

    ws.cell(
        header_account_row,
        cuentas_abono_column
    ).value = "CUENTAS CON ABONO"

    # --------------------------------------------------------
    # Fechas
    # --------------------------------------------------------

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
    # Escribir datos diarios
    # --------------------------------------------------------

    for index, fecha in enumerate(
        fechas,
        start=start_row
    ):

        ws.cell(
            index,
            1
        ).value = datetime(
            fecha.year,
            fecha.month,
            fecha.day
        )

        ws.cell(
            index,
            1
        ).number_format = "dd/mm/yyyy"

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

            # Si hubo crédito, escribirlo.
            # Si no hubo, dejar vacío.
            if valor != 0:
                ws.cell(
                    index,
                    column
                ).value = valor

                cuentas_con_abono += 1

                total += valor

            else:
                ws.cell(
                    index,
                    column
                ).value = None

        ws.cell(
            index,
            total_column
        ).value = total

        ws.cell(
            index,
            cuentas_abono_column
        ).value = cuentas_con_abono


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
    # Valores de selección
    # --------------------------------------------------------

    # Mantener las fechas de la plantilla si existen.
    # Si no existen, utilizar primera/última fecha.
    fechas = [
        transaction["fecha"]
        for transaction in transactions
        if isinstance(
            transaction["fecha"],
            (datetime, date)
        )
    ]

    if fechas:

        fechas = sorted(
            fechas
        )

        fecha_min = fechas[0]
        fecha_max = fechas[-1]

        if isinstance(
            fecha_min,
            date
        ) and not isinstance(
            fecha_min,
            datetime
        ):
            fecha_min = datetime(
                fecha_min.year,
                fecha_min.month,
                fecha_min.day
            )

        if isinstance(
            fecha_max,
            date
        ) and not isinstance(
            fecha_max,
            datetime
        ):
            fecha_max = datetime(
                fecha_max.year,
                fecha_max.month,
                fecha_max.day
            )

        ws["A5"] = fecha_min
        ws["C5"] = fecha_max

        ws["A5"].number_format = "dd/mm/yyyy"
        ws["C5"].number_format = "dd/mm/yyyy"

    ws["E5"] = "TODAS"
    ws["G5"] = "TODOS"

    # --------------------------------------------------------
    # Detectar filas de evolución
    # --------------------------------------------------------

    report_start = 5
    report_end = (
        reporte.max_row
    )

    tablero_start = 15

    # Limpiar valores/fórmulas de evolución
    for row in range(
        tablero_start,
        max(
            ws.max_row,
            tablero_start + 100
        ) + 1
    ):

        for col in range(
            1,
            5
        ):

            ws.cell(
                row,
                col
            ).value = None

    # --------------------------------------------------------
    # Fórmulas de evolución
    # --------------------------------------------------------

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
            ws.cell(
                row,
                1
            ).value = (
                f'=IF(AND(\'{SHEET_REPORTE}\'!A{report_row}>=$A$5,'
                f'\'{SHEET_REPORTE}\'!A{report_row}<=$C$5),'
                f'\'{SHEET_REPORTE}\'!A{report_row},"")'
            )

            # Crédito seleccionado
            ws.cell(
                row,
                2
            ).value = (
                f'=IF(A{row}="",0,'
                f'IF($E$5<>"TODAS",'
                f'IFERROR(INDEX(\'{SHEET_REPORTE}\'!$B$5:$I${report_end},'
                f'MATCH(A{row},\'{SHEET_REPORTE}\'!$A$5:$A${report_end},0),'
                f'MATCH($E$5,\'{SHEET_REPORTE}\'!$B$4:$I$4,0)),0),'
                f'IF($G$5="TODOS",'
                f'IFERROR(SUMIF(\'{SHEET_REPORTE}\'!$A$5:$A${report_end},'
                f'A{row},\'{SHEET_REPORTE}\'!$J$5:$J${report_end}),0),0)))'
            )

            # Variación diaria
            if i == 0:

                ws.cell(
                    row,
                    3
                ).value = 0

            else:

                ws.cell(
                    row,
                    3
                ).value = (
                    f'=B{row}-B{row - 1}'
                )

            # Cuentas con abono
            ws.cell(
                row,
                4
            ).value = (
                f'=IF(A{row}="",0,'
                f'IF($G$5="TODOS",'
                f'IFERROR(INDEX(\'{SHEET_REPORTE}\'!$K$5:$K${report_end},'
                f'MATCH(A{row},\'{SHEET_REPORTE}\'!$A$5:$A${report_end},0)),0),'
                f'IF(B{row}>0,1,0)))'
            )

    # --------------------------------------------------------
    # CRÉDITOS POR CUENTA
    # --------------------------------------------------------

    # Obtener cuentas únicas
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

    # Limpiar columnas F:H
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

            ws.cell(
                row,
                col
            ).value = None

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

    # Las columnas del reporte empiezan en B.
    # Construir mapa.
    report_columns = {}

    for col in range(
        2,
        reporte.max_column + 1
    ):

        banco = reporte.cell(
            3,
            col
        ).value

        cuenta = reporte.cell(
            4,
            col
        ).value

        if banco and cuenta:

            report_columns[
                (
                    clean_text(banco),
                    str(cuenta).strip()
                )
            ] = col

    for index, (
        banco,
        cuenta
    ) in enumerate(
        cuentas,
        start=cuenta_start
    ):

        ws.cell(
            index,
            6
        ).value = banco

        ws.cell(
            index,
            7
        ).value = cuenta

        report_col = report_columns.get(
            (
                clean_text(banco),
                cuenta
            )
        )

        if report_col:

            col_letter = (
                reporte.cell(
                    1,
                    report_col
                ).column_letter
            )

            ws.cell(
                index,
                8
            ).value = (
                f'=IF(AND('
                f'OR($G$5="TODOS",TRIM($G$5)=TRIM(F{index})),'
                f'OR($E$5="TODAS",TEXT($E$5,"0")=TEXT(G{index},"0"))'
                f'),'
                f'SUMIFS(\'{SHEET_REPORTE}\'!'
                f'${col_letter}$5:${col_letter}${report_end},'
                f'\'{SHEET_REPORTE}\'!$A$5:$A${report_end},">="&$A$5,'
                f'\'{SHEET_REPORTE}\'!$A$5:$A${report_end},"<="&$C$5),'
                f'0)'
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

    # Abrir plantilla real
    workbook = load_workbook(
        TEMPLATE_FILE
    )

    try:

        # ----------------------------------------------------
        # 1. ESTADOS CONSOLIDADOS
        # ----------------------------------------------------

        escribir_estados_consolidados(
            workbook,
            transactions
        )

        # ----------------------------------------------------
        # 2. SALDOS POR CUENTA
        # ----------------------------------------------------

        escribir_saldos_por_cuenta(
            workbook,
            transactions
        )

        # ----------------------------------------------------
        # 3. REPORTE CREDITOS DIARIOS
        # ----------------------------------------------------

        escribir_reporte_creditos(
            workbook,
            transactions
        )

        # ----------------------------------------------------
        # 4. TABLERO CREDITOS
        # ----------------------------------------------------

        actualizar_tablero(
            workbook,
            transactions
        )

        # ----------------------------------------------------
        # Guardar
        # ----------------------------------------------------

        workbook.save(
            output_path
        )

    finally:

        workbook.close()


# ============================================================
# PROCESAR EXCEL COMPLETO
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

    file = request.files[
        "file"
    ]

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
        # GUARDAR EXCEL RECIBIDO
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
        # CREAR ARCHIVO TEMPORAL FINAL
        # ====================================================

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as temp:

            output_path = temp.name

        # ====================================================
        # CREAR RESULTADO DESDE LA PLANTILLA
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

            os.remove(
                input_path
            )

            input_path = None

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
```
