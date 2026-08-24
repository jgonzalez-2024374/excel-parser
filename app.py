from flask import Flask, request, jsonify, send_file, after_this_request
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy
from datetime import datetime, date
from difflib import SequenceMatcher
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
        "parser_version": PARSER_VERSION if "PARSER_VERSION" in globals() else "legacy",
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
# PARSER UNIVERSAL / NORMALIZACIÓN
# ============================================================

PARSER_VERSION = "universal-2.0"

# Este parser NO depende de un banco concreto. Las listas siguientes son
# vocabulario contable para reconocer columnas, no formatos rígidos por banco.
HEADER_ALIASES = {
    "fecha": [
        "fecha", "date", "fecha movimiento", "fecha de movimiento",
        "fecha transaccion", "fecha de transaccion", "transaction date",
        "fecha operacion", "fecha de operacion", "posting date", "value date",
    ],
    "referencia": [
        "referencia", "reference", "ref", "no doc", "no. doc", "nro doc",
        "documento", "numero documento", "numero de documento", "comprobante",
        "numero operacion", "operacion", "transaction id", "trace", "secuencial",
    ],
    "codigo": [
        "codigo", "code", "codigo transaccion", "codigo de transaccion",
        "transaction code", "tt", "tipo transaccion", "tipo de transaccion",
    ],
    "descripcion": [
        "descripcion", "description", "detalle", "concepto", "glosa", "memo",
        "movimiento", "detalle movimiento", "descripcion movimiento",
        "descripcion de transaccion", "transaction description", "narrative",
    ],
    "debito": [
        "debito", "debit", "debe", "cargo", "cargos", "retiro", "retiros",
        "egreso", "egresos", "withdrawal", "withdrawals", "valor -",
        "monto debito", "debito de transaccion", "cargo debito",
    ],
    "credito": [
        "credito", "credit", "haber", "abono", "abonos", "deposito",
        "depositos", "ingreso", "ingresos", "deposit", "valor +",
        "monto credito", "credito de transaccion", "abono credito",
    ],
    "saldo": [
        "saldo", "balance", "saldo contable", "saldo disponible",
        "available balance", "running balance", "balance de transaccion",
        "saldo final", "book balance",
    ],
    "monto": [
        "monto", "amount", "importe", "valor transaccion", "valor de transaccion",
        "monto transaccion", "transaction amount", "importe movimiento",
    ],
    "tipo": [
        "tipo", "naturaleza", "d/c", "dc", "dr/cr", "dr cr", "deb cred",
        "signo", "transaction type", "tipo movimiento",
    ],
}

GENERIC_SHEET_NAMES = {
    "hoja", "hoja1", "hoja 1", "sheet", "sheet1", "sheet 1",
    "movimientos", "transacciones", "estado de cuenta", "estado cuenta",
}

SUMMARY_PREFIXES = (
    "resumen", "total", "totales", "subtotal", "saldo inicial", "saldo anterior",
    "saldo final", "cantidad deb", "cantidad cred", "monto deb", "monto cred",
    "numero de transacciones", "no debitos", "no creditos", "summary",
    "beginning balance", "ending balance", "opening balance", "closing balance",
)

CREDIT_HINTS = (
    "credito", "credit", "abono", "deposito", "deposit", "ingreso", "haber",
    "transferencia recibida", "tef de", "ach credito", "cr", "c",
)

DEBIT_HINTS = (
    "debito", "debit", "cargo", "retiro", "egreso", "debe", "pago",
    "transferencia enviada", "ach debito", "dr", "d",
)

KNOWN_BANKS = [
    # Guatemala
    (["banco industrial", "bi"], "BANCO INDUSTRIAL"),
    (["g&t continental", "gyt continental", "g&t", "gyt"], "G&T"),
    (["bac credomatic", "banco de america central", "bac"], "BAC"),
    (["banrural", "banco de desarrollo rural", "ban"], "BANRURAL"),
    (["banco agricola mercantil", "bam"], "BAM"),
    (["banco promerica", "promerica"], "PROMERICA"),
    (["banco davivienda", "davivienda"], "DAVIVIENDA"),
    (["banco ficohsa", "ficohsa"], "FICOHSA"),
    # El Salvador / regionales
    (["banco agricola", "banco agrícola"], "BANCO AGRÍCOLA"),
    (["banco cuscatlan", "banco cuscatlán", "cuscatlan", "cuscatlán"], "CUSCATLÁN"),
    (["banco hipotecario"], "BANCO HIPOTECARIO"),
    (["banco azul", "abank"], "BANCO AZUL"),
    (["banco atlantida", "banco atlántida", "atlantida", "atlántida"], "BANCO ATLÁNTIDA"),
    (["banco promete", "banco integral"], "BANCO INTEGRAL"),
]


def clean_text(value):
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ").replace("\u200b", " ")

    # Elimina acentos y también tolera encabezados con codificación dañada
    # (ej.: "Descripci髇", "D閎ito", "Cr閐ito", "D�bito").
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def compact_text(value):
    text = clean_text(value)
    return re.sub(r"[^a-z0-9+\-/& ]+", "", text).strip()


def parse_number(value):
    """Convierte números bancarios sin confundir 1,000 con 1.00.

    Devuelve None si la celda no contiene un número utilizable.
    """
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text or text in {"-", "--", "—", "–"}:
        return None

    original = clean_text(text)
    negative = False

    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    if text.endswith("-"):
        negative = True
        text = text[:-1]

    # Indicadores comunes de débito al lado del importe.
    if re.search(r"\b(db|dr|debit|debito)\b", original):
        negative = True

    text = text.replace("Q", "").replace("$", "").replace("€", "")
    text = re.sub(r"\b(GTQ|USD|EUR|SVC)\b", "", text, flags=re.I)
    text = text.replace("'", "").replace(" ", "")
    text = re.sub(r"[^0-9,\.\-+]", "", text)

    if not re.search(r"\d", text):
        return None

    sign = -1 if negative else 1
    if text.startswith("-"):
        sign = -1
    text = text.lstrip("+-")

    # Determina separador decimal según la posición del último separador.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    elif "," in text:
        parts = text.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            text = parts[0] + "." + parts[1]
        elif len(parts) > 2 and len(parts[-1]) in (1, 2):
            text = "".join(parts[:-1]) + "." + parts[-1]
        else:
            text = "".join(parts)

    elif "." in text:
        parts = text.split(".")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            pass
        elif len(parts) > 2 and len(parts[-1]) in (1, 2):
            text = "".join(parts[:-1]) + "." + parts[-1]
        elif len(parts) == 2 and len(parts[1]) == 3:
            text = "".join(parts)

    try:
        return sign * abs(float(text))
    except Exception:
        return None


def clean_number(value):
    parsed = parse_number(value)
    return parsed if parsed is not None else 0.0


def _parse_text_date(text, expected_month=None, expected_year=None):
    text = str(text).strip()
    if not text:
        return None

    # Limpia hora cuando viene junto a la fecha.
    text = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?.*$", "", text)

    # ISO primero (no ambiguo).
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            pass

    # dd/mm es el estándar principal para Guatemala y El Salvador.
    ddmm = None
    mmdd = None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            ddmm = datetime.strptime(text, fmt)
            break
        except Exception:
            pass

    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%m-%d-%y"):
        try:
            mmdd = datetime.strptime(text, fmt)
            break
        except Exception:
            pass

    if expected_month:
        candidates = [d for d in (ddmm, mmdd) if d is not None]
        for d in candidates:
            if d.month == expected_month and (not expected_year or d.year == expected_year):
                return d

    return ddmm or mmdd


def clean_date(value, expected_month=None, expected_year=None):
    if value is None or value == "":
        return ""

    parsed = None

    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime(value.year, value.month, value.day)
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        # Los seriales de Excel razonables están aprox. entre 20,000 y 80,000.
        if 20000 <= float(value) <= 80000:
            try:
                from openpyxl.utils.datetime import from_excel
                parsed = from_excel(value)
            except Exception:
                parsed = None
    else:
        parsed = _parse_text_date(value, expected_month, expected_year)

    if not isinstance(parsed, datetime):
        return ""

    # Algunos exports bancarios latinoamericanos guardan 04/08 como serial
    # equivalente a 08/04. Si conocemos el mes del estado de cuenta, se corrige
    # de forma genérica (sin depender del banco).
    if expected_month and parsed.month != expected_month and parsed.day == expected_month:
        try:
            year = expected_year or parsed.year
            parsed = datetime(year, expected_month, parsed.month)
        except Exception:
            pass

    if expected_year and parsed.year != expected_year:
        # Solo ajustar años abreviados/erróneos cuando el mes coincide y la
        # diferencia es claramente de interpretación del periodo.
        if parsed.month == expected_month and abs(parsed.year - expected_year) <= 1:
            try:
                parsed = parsed.replace(year=expected_year)
            except Exception:
                pass

    return parsed


def _similarity(a, b):
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if b in a or a in b:
        short = min(len(a), len(b))
        long = max(len(a), len(b))
        return max(0.86, short / max(long, 1))
    return SequenceMatcher(None, a, b).ratio()


def _header_field(value):
    text = compact_text(value)
    if not text:
        return None, 0.0

    best_field = None
    best_score = 0.0

    for field, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            alias_clean = compact_text(alias)
            score = _similarity(text, alias_clean)

            # Coincidencias por raíces robustas a caracteres dañados.
            if field == "descripcion" and ("descrip" in text or "concept" in text or "detalle" in text):
                score = max(score, 0.98)
            elif field == "debito" and (
                any(root in text for root in ("deb", "debe", "cargo", "reti", "egres", "dbit"))
                or text.startswith("dito")
            ):
                score = max(score, 0.96)
            elif field == "credito" and (
                any(root in text for root in ("cred", "haber", "abono", "deposit", "ingres", "crdit"))
                or text.startswith("crito")
            ):
                score = max(score, 0.96)
            elif field == "fecha" and (text == "fecha" or text.startswith("fecha ") or " date" in f" {text}"):
                score = max(score, 0.98)
            elif field == "saldo" and ("saldo" in text or "balance" in text):
                score = max(score, 0.98)
            elif field == "referencia" and any(root in text for root in ("refer", "no doc", "document", "secuencial", "trace")):
                score = max(score, 0.97)
            elif field == "codigo" and any(root in text for root in ("codigo", "cigo", "transaction code")):
                score = max(score, 0.95)
            elif field == "monto" and any(root in text for root in ("monto", "importe", "amount")):
                score = max(score, 0.95)

            if score > best_score:
                best_field = field
                best_score = score

    if best_score >= 0.78:
        return best_field, best_score
    return None, best_score


def detectar_periodo(rows, sheet_name=""):
    """Devuelve (mes, año) del estado de cuenta cuando puede inferirse."""
    month = None
    year = None

    title = clean_text(sheet_name)

    # Ej.: "BAC 08.26", "Cuenta_08-2026", "2026-08".
    m = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[.\-_/ ](20\d{2}|\d{2})(?!\d)", title)
    if m:
        month = int(m.group(1))
        y = int(m.group(2))
        year = 2000 + y if y < 100 else y

    if month is None:
        m = re.search(r"(?<!\d)(20\d{2})[.\-_/ ](0?[1-9]|1[0-2])(?!\d)", title)
        if m:
            year = int(m.group(1))
            month = int(m.group(2))

    # Busca periodos explícitos en las primeras filas.
    for row in rows[:25]:
        joined = " | ".join(str(v) for v in row if v not in (None, ""))
        normalized = clean_text(joined)
        if any(k in normalized for k in ("fecha inicial", "fecha final", "periodo", "desde", "hasta", "corte")):
            dates = re.findall(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", joined)
            for dtext in dates:
                d = _parse_text_date(dtext)
                if d:
                    month = month or d.month
                    year = year or d.year
                    if month == d.month:
                        year = d.year
                        break

    return month, year


def detectar_banco(sheet_name, rows):
    title = clean_text(sheet_name)

    # 1) Catálogo opcional. Sirve para estandarizar nombres, pero el parser no
    # depende de que el banco esté en esta lista.
    for aliases, canonical in KNOWN_BANKS:
        for alias in aliases:
            a = clean_text(alias)
            if a in {"bi", "ban"}:
                if re.search(rf"\b{re.escape(a)}\b", title):
                    return canonical
            elif a and a in title:
                return canonical

    # 2) Buscar nombre de institución en el encabezado.
    for row in rows[:20]:
        for value in row:
            raw = str(value).strip() if value not in (None, "") else ""
            text = clean_text(raw)
            if not text:
                continue

            for aliases, canonical in KNOWN_BANKS:
                for alias in aliases:
                    a = clean_text(alias)
                    if a in {"bi", "ban"}:
                        if re.search(rf"\b{re.escape(a)}\b", text):
                            return canonical
                    elif a and a in text:
                        return canonical

            if re.search(r"\bbanco\b|\bbank\b", text) and len(text) <= 80:
                # Quita etiquetas genéricas y conserva un nombre legible.
                candidate = re.sub(r"(?i)^.*?(?:banco|bank)\s*[:\-]?\s*", "BANCO ", raw).strip()
                if candidate:
                    return candidate.upper()

    # 3) Banco desconocido: usar el nombre de la hoja si es informativo.
    if title not in GENERIC_SHEET_NAMES and title:
        cleaned = re.sub(r"\s+(0?[1-9]|1[0-2])[.\-_/ ](20\d{2}|\d{2})\s*$", "", str(sheet_name), flags=re.I).strip()
        return cleaned or str(sheet_name).strip()

    return "BANCO NO IDENTIFICADO"


def _account_candidate(value):
    if value in (None, ""):
        return ""

    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()

    # Evitar saldos y fechas.
    if re.search(r"[.,]\d{1,2}$", text):
        return ""
    if re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", text):
        return ""

    digits = re.sub(r"[^0-9]", "", text)
    if 5 <= len(digits) <= 30:
        return digits
    return ""


def detectar_cuenta(sheet_name, rows, header_index=None):
    limit = min((header_index + 1) if header_index is not None else 25, len(rows))

    # A) Formato etiqueta/valor en la misma fila o celda.
    regexes = [
        r"(?:numero|nro|no\.?|#)?\s*(?:de\s+)?[#:]?\s*cuenta\s*[:#\-]?\s*([0-9][0-9\- ]{3,30})",
        r"account(?:\s+number|\s+no\.?)?\s*[:#\-]?\s*([0-9][0-9\- ]{3,30})",
        r"iban\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\- ]{5,34})",
    ]

    for row in rows[:limit]:
        joined = " | ".join(str(v) for v in row if v not in (None, ""))
        normalized = clean_text(joined)
        for pattern in regexes:
            match = re.search(pattern, normalized, re.I)
            if match:
                digits = re.sub(r"[^0-9]", "", match.group(1))
                if len(digits) >= 5:
                    return digits

    # B) Encabezado tipo tabla: Cuenta / Producto / Account en una fila y valor
    # debajo (muy común en exports de BAC y otros bancos regionales).
    for r in range(max(0, limit - 1)):
        row = rows[r]
        next_row = rows[r + 1] if r + 1 < len(rows) else []
        for c, value in enumerate(row):
            label = compact_text(value)
            if not label:
                continue
            if (
                "cuenta" in label
                or label in {"account", "account number", "producto", "product"}
                or "numero de cuenta" in label
            ):
                # Misma fila: algunos exports dejan una celda vacía entre
                # etiqueta y valor por celdas combinadas. Revisar las próximas 3.
                for cc in range(c + 1, min(len(row), c + 4)):
                    candidate = _account_candidate(row[cc])
                    if candidate:
                        return candidate
                # Fila siguiente: misma columna y vecinas cercanas.
                for cc in range(c, min(len(next_row), c + 3)):
                    candidate = _account_candidate(next_row[cc])
                    if candidate:
                        return candidate

    # C) Nombre de hoja. Solo como último recurso.
    candidates = re.findall(r"(?<!\d)(\d{3,30})(?!\d)", str(sheet_name))
    # Ignora secuencias que parecen periodo 0826 / 2026.
    candidates = [x for x in candidates if x not in {"2024", "2025", "2026", "2027", "2028"}]
    if candidates:
        return max(candidates, key=len)

    return ""


def detectar_columnas(rows):
    best = None

    for row_index, row in enumerate(rows):
        columnas = {}
        confidences = {}

        for index, value in enumerate(row):
            field, score = _header_field(value)
            if not field:
                continue

            # "valor" sin contexto puede ser monto, pero no debe desplazar
            # débito/crédito/saldo si ya existe un match más fuerte.
            if field not in columnas or score > confidences.get(field, 0):
                columnas[field] = index
                confidences[field] = score

        has_date = "fecha" in columnas
        has_desc_or_ref = "descripcion" in columnas or "referencia" in columnas
        has_pair = "debito" in columnas and "credito" in columnas
        has_amount = "monto" in columnas
        has_financial = has_pair or has_amount or (
            "saldo" in columnas and ("debito" in columnas or "credito" in columnas)
        )

        if not (has_date and has_desc_or_ref and has_financial):
            continue

        score = 0
        weights = {
            "fecha": 4, "descripcion": 3, "referencia": 2, "codigo": 1,
            "debito": 3, "credito": 3, "monto": 3, "saldo": 3, "tipo": 1,
        }
        for field in columnas:
            score += weights.get(field, 1)
        score += sum(confidences.values())

        candidate = (score, row_index, columnas)
        if best is None or candidate[0] > best[0]:
            best = candidate

    if best:
        return best[1], best[2]
    return None, {}


def get_column(row, columns, name):
    index = columns.get(name)
    if index is None or index >= len(row):
        return ""
    return row[index]


def _looks_like_summary(row):
    values = [str(v) for v in row if v not in (None, "")]
    if not values:
        return True
    joined = clean_text(" | ".join(values))
    first = compact_text(values[0])
    first_words = re.sub(r"[^a-z0-9 ]+", " ", first)
    first_words = re.sub(r"\s+", " ", first_words).strip()

    normalized_prefixes = [
        re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]+", " ", compact_text(p))).strip()
        for p in SUMMARY_PREFIXES
    ]
    if any(first_words.startswith(prefix) for prefix in normalized_prefixes if prefix):
        return True
    if (
        first_words.startswith("no credito")
        or first_words.startswith("no debito")
        or first_words.startswith("cantidad credito")
        or first_words.startswith("cantidad debito")
        or first_words.startswith("monto credito")
        or first_words.startswith("monto debito")
    ):
        return True
    if "resumen de estado" in joined or "summary of" in joined:
        return True
    return False


def _movement_direction(text):
    t = compact_text(text)
    if not t:
        return None

    # Match de palabras completas cuando son códigos cortos.
    tokens = set(re.findall(r"[a-z]+", t))
    if tokens.intersection({"cr", "credit", "credito", "abono", "deposito", "deposit", "haber", "ingreso"}):
        return "credit"
    if tokens.intersection({"dr", "db", "debit", "debito", "cargo", "retiro", "debe", "egreso"}):
        return "debit"

    if any(h in t for h in CREDIT_HINTS if len(h) > 2):
        return "credit"
    if any(h in t for h in DEBIT_HINTS if len(h) > 2):
        return "debit"
    return None


def detectar_saldo_inicial(rows, header_index):
    limit = min(header_index + 1, len(rows)) if header_index is not None else min(25, len(rows))

    for r in range(limit):
        row = rows[r]
        for c, value in enumerate(row):
            label = compact_text(value)
            if not label:
                continue

            if any(key in label for key in (
                "saldo inicial", "saldo anterior", "balance inicial",
                "beginning balance", "opening balance", "previous balance",
            )):
                # Número embebido en la misma celda.
                numbers = re.findall(r"[-+]?\d[\d., ]*\d|[-+]?\d", str(value))
                for n in reversed(numbers):
                    parsed = parse_number(n)
                    if parsed is not None:
                        return parsed

                # Celda siguiente.
                if c + 1 < len(row):
                    parsed = parse_number(row[c + 1])
                    if parsed is not None:
                        return parsed

                # Misma columna, fila siguiente.
                if r + 1 < len(rows) and c < len(rows[r + 1]):
                    parsed = parse_number(rows[r + 1][c])
                    if parsed is not None:
                        return parsed

    return None


def extraer_transaccion(
    row,
    columns,
    banco,
    cuenta,
    expected_month=None,
    expected_year=None,
    previous_balance=None,
    previous_date=None,
    saldo_inicial_cuenta=None,
):
    raw_date = get_column(row, columns, "fecha")
    referencia = get_column(row, columns, "referencia")
    codigo = get_column(row, columns, "codigo")
    descripcion = get_column(row, columns, "descripcion")
    tipo = get_column(row, columns, "tipo")

    fecha = clean_date(raw_date, expected_month, expected_year)

    # Permite estados donde la fecha aparece una sola vez y las siguientes
    # líneas pertenecen al mismo día, pero evita usar esto en resúmenes.
    if not isinstance(fecha, datetime) and previous_date and raw_date in (None, ""):
        fecha = previous_date

    if not isinstance(fecha, datetime):
        return None

    if _looks_like_summary(row):
        return None

    debit_raw = get_column(row, columns, "debito")
    credit_raw = get_column(row, columns, "credito")
    amount_raw = get_column(row, columns, "monto")
    balance_raw = get_column(row, columns, "saldo")

    debit_val = parse_number(debit_raw)
    credit_val = parse_number(credit_raw)
    amount_val = parse_number(amount_raw)
    balance_val = parse_number(balance_raw)

    debito = -abs(debit_val) if debit_val not in (None, 0) else 0.0
    credito = abs(credit_val) if credit_val not in (None, 0) else 0.0

    # Formato de una sola columna de monto.
    if debito == 0 and credito == 0 and amount_val not in (None, 0):
        direction_text = " ".join(str(v) for v in (tipo, codigo, descripcion) if v not in (None, ""))
        direction = _movement_direction(direction_text)

        if amount_val < 0:
            direction = "debit"

        # El saldo corrido permite resolver formatos que muestran el monto
        # siempre positivo pero no indican explícitamente D/C.
        if direction is None and previous_balance is not None and balance_val is not None:
            delta = balance_val - previous_balance
            if abs(delta) > 0.000001:
                direction = "credit" if delta > 0 else "debit"

        if direction == "debit":
            debito = -abs(amount_val)
        else:
            # Para columnas firmadas, un positivo normalmente es crédito.
            credito = abs(amount_val)

    # Formato sin monto explícito, pero con saldo corrido.
    if debito == 0 and credito == 0 and balance_val is not None and previous_balance is not None:
        delta = balance_val - previous_balance
        if abs(delta) > 0.000001:
            if delta > 0:
                credito = abs(delta)
            else:
                debito = -abs(delta)

    texto_desc = str(descripcion).strip() if descripcion not in (None, "") else ""
    texto_ref = str(referencia).strip() if referencia not in (None, "") else ""
    texto_codigo = str(codigo).strip() if codigo not in (None, "") else ""

    # Una transacción real necesita monto/movimiento y alguna identificación.
    has_money = debito != 0 or credito != 0 or balance_val is not None
    has_identity = bool(texto_desc or texto_ref or texto_codigo)
    if not (has_money and has_identity):
        return None

    return {
        "banco": banco,
        "cuenta": cuenta,
        "fecha": fecha,
        "referencia": texto_ref or texto_codigo,
        "codigo": texto_codigo,
        "descripcion": texto_desc or texto_codigo or texto_ref,
        "debito": debito,
        "credito": credito,
        "saldo": balance_val if balance_val is not None else 0.0,
        "saldo_inicial_cuenta": saldo_inicial_cuenta,
    }


def procesar_hoja(worksheet):
    rows = []

    for row in worksheet.iter_rows(values_only=True):
        values = [value if value is not None else "" for value in row]
        if any(value != "" for value in values):
            rows.append(values)

    if not rows:
        return []

    header_index, columns = detectar_columnas(rows)
    if header_index is None:
        return []

    banco = detectar_banco(worksheet.title, rows)
    cuenta = detectar_cuenta(worksheet.title, rows, header_index)
    expected_month, expected_year = detectar_periodo(rows, worksheet.title)
    saldo_inicial = detectar_saldo_inicial(rows, header_index)

    transactions = []
    previous_balance = saldo_inicial
    previous_date = None
    invalid_streak = 0

    for row in rows[header_index + 1:]:
        transaction = extraer_transaccion(
            row,
            columns,
            banco,
            cuenta,
            expected_month=expected_month,
            expected_year=expected_year,
            previous_balance=previous_balance,
            previous_date=previous_date,
            saldo_inicial_cuenta=saldo_inicial,
        )

        if transaction:
            transactions.append(transaction)
            previous_date = transaction["fecha"]
            if transaction["saldo"] != 0 or get_column(row, columns, "saldo") not in (None, ""):
                previous_balance = transaction["saldo"]
            invalid_streak = 0
        else:
            invalid_streak += 1

            # No corta inmediatamente al ver un resumen: algunos bancos ponen
            # subtotales entre grupos. Solo deja de recorrer después de un bloque
            # largo sin movimientos reales.
            if transactions and invalid_streak >= 25:
                break

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
        raise Exception(f"No existe la hoja '{SHEET_SALDOS}'")

    ws = workbook[SHEET_SALDOS]
    start_row = 5
    cuentas = {}

    for transaction in transactions:
        clave = (transaction["banco"], transaction["cuenta"])
        cuentas.setdefault(clave, []).append(transaction)

    limpiar_filas(ws, start_row, ws.max_row, 1, 4)

    required_last_row = start_row + len(cuentas) - 1
    if required_last_row > ws.max_row:
        old_max_row = ws.max_row
        for row in range(old_max_row + 1, required_last_row + 1):
            copiar_estilo_fila(ws, start_row, row)

    row_number = start_row

    for (banco, cuenta), movimientos in cuentas.items():
        movimientos = sorted(
            movimientos,
            key=lambda x: x["fecha"] if isinstance(x["fecha"], datetime) else datetime.min
        )
        if not movimientos:
            continue

        primero = movimientos[0]
        ultimo = movimientos[-1]

        saldo_explicito = primero.get("saldo_inicial_cuenta")
        if saldo_explicito is not None:
            saldo_inicial = saldo_explicito
        elif primero.get("saldo") is not None:
            # saldo_nuevo = saldo_anterior + debito + credito
            saldo_inicial = primero["saldo"] - primero["debito"] - primero["credito"]
        else:
            saldo_inicial = 0.0

        saldo_final = ultimo.get("saldo", 0.0)

        escribir_celda_segura(ws, row_number, 1, banco_corto(banco))
        escribir_celda_segura(ws, row_number, 2, cuenta)
        escribir_celda_segura(ws, row_number, 3, saldo_inicial)
        escribir_celda_segura(ws, row_number, 4, saldo_final)
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
    if SHEET_TABLERO not in workbook.sheetnames or SHEET_REPORTE not in workbook.sheetnames:
        return

    ws = workbook[SHEET_TABLERO]
    reporte = workbook[SHEET_REPORTE]

    fechas = [
        t["fecha"] for t in transactions
        if isinstance(t.get("fecha"), (datetime, date))
    ]

    if fechas:
        fechas = sorted(fechas)
        fecha_min, fecha_max = fechas[0], fechas[-1]
        if isinstance(fecha_min, date) and not isinstance(fecha_min, datetime):
            fecha_min = datetime(fecha_min.year, fecha_min.month, fecha_min.day)
        if isinstance(fecha_max, date) and not isinstance(fecha_max, datetime):
            fecha_max = datetime(fecha_max.year, fecha_max.month, fecha_max.day)

        cell = obtener_celda_segura(ws, 5, 1)
        if cell:
            cell.value = fecha_min
            cell.number_format = "dd/mm/yyyy"
        cell = obtener_celda_segura(ws, 5, 3)
        if cell:
            cell.value = fecha_max
            cell.number_format = "dd/mm/yyyy"

    escribir_celda_segura(ws, 5, 5, "TODAS")
    escribir_celda_segura(ws, 5, 7, "TODOS")

    # Detectar dinámicamente dónde quedaron TOTAL CRÉDITOS y CUENTAS CON ABONO.
    total_col = None
    count_col = None
    for col in range(2, reporte.max_column + 1):
        header = clean_text(reporte.cell(row=4, column=col).value)
        if "total" in header and ("credito" in header or "credit" in header):
            total_col = col
        if "cuentas" in header and "abono" in header:
            count_col = col

    if total_col is None:
        # escribir_reporte_creditos siempre crea el total justo después de cuentas.
        total_col = max(2, reporte.max_column - 1)
    if count_col is None:
        count_col = total_col + 1

    last_account_col = max(2, total_col - 1)
    first_account_letter = get_column_letter(2)
    last_account_letter = get_column_letter(last_account_col)
    total_letter = get_column_letter(total_col)
    count_letter = get_column_letter(count_col)

    report_start = 5
    report_end = reporte.max_row
    tablero_start = 15

    for row in range(tablero_start, max(ws.max_row, tablero_start + 100) + 1):
        for col in range(1, 5):
            limpiar_celda_segura(ws, row, col)

    max_evolution_rows = max(19, report_end - report_start + 1)
    required_end = tablero_start + max_evolution_rows - 1

    if required_end > ws.max_row:
        old_max = ws.max_row
        for row in range(old_max + 1, required_end + 1):
            copiar_estilo_fila(ws, 15, row)

    for i in range(max_evolution_rows):
        row = tablero_start + i
        report_row = report_start + i

        if report_row <= report_end:
            formula_fecha = (
                f'=IF(AND(\'{SHEET_REPORTE}\'!A{report_row}>=$A$5,'
                f'\'{SHEET_REPORTE}\'!A{report_row}<=$C$5),'
                f'\'{SHEET_REPORTE}\'!A{report_row},"")'
            )
            escribir_celda_segura(ws, row, 1, formula_fecha)

            formula_credito = (
                f'=IF(A{row}="",0,'
                f'IF($E$5<>"TODAS",'
                f'IFERROR(INDEX(\'{SHEET_REPORTE}\'!${first_account_letter}$5:${last_account_letter}${report_end},'
                f'MATCH(A{row},\'{SHEET_REPORTE}\'!$A$5:$A${report_end},0),'
                f'MATCH($E$5,\'{SHEET_REPORTE}\'!${first_account_letter}$4:${last_account_letter}$4,0)),0),'
                f'IF($G$5="TODOS",'
                f'IFERROR(SUMIF(\'{SHEET_REPORTE}\'!$A$5:$A${report_end},A{row},'
                f'\'{SHEET_REPORTE}\'!${total_letter}$5:${total_letter}${report_end}),0),0)))'
            )
            escribir_celda_segura(ws, row, 2, formula_credito)

            if i == 0:
                escribir_celda_segura(ws, row, 3, 0)
            else:
                escribir_celda_segura(ws, row, 3, f'=B{row}-B{row - 1}')

            formula_cuentas = (
                f'=IF(A{row}="",0,'
                f'IF($G$5="TODOS",IFERROR(INDEX(\'{SHEET_REPORTE}\'!${count_letter}$5:${count_letter}${report_end},'
                f'MATCH(A{row},\'{SHEET_REPORTE}\'!$A$5:$A${report_end},0)),0),IF(B{row}>0,1,0)))'
            )
            escribir_celda_segura(ws, row, 4, formula_cuentas)

    # Créditos por cuenta
    cuentas = []
    for transaction in transactions:
        clave = (banco_corto(transaction["banco"]), str(transaction["cuenta"]))
        if clave not in cuentas:
            cuentas.append(clave)
    cuentas = sorted(cuentas, key=lambda x: (x[0], x[1]))

    cuenta_start = 15
    for row in range(cuenta_start, max(ws.max_row, cuenta_start + 100) + 1):
        for col in range(6, 9):
            limpiar_celda_segura(ws, row, col)

    cuenta_end = cuenta_start + len(cuentas) - 1
    if cuenta_end > ws.max_row:
        old_max = ws.max_row
        for row in range(old_max + 1, cuenta_end + 1):
            copiar_estilo_fila(ws, 15, row)

    report_columns = {}
    for col in range(2, total_col):
        banco_cell = reporte.cell(row=3, column=col)
        cuenta_cell = reporte.cell(row=4, column=col)
        if es_celda_combinada(banco_cell) or es_celda_combinada(cuenta_cell):
            continue
        banco = banco_cell.value
        cuenta = cuenta_cell.value
        if banco and cuenta:
            report_columns[(clean_text(banco), str(cuenta).strip())] = col

    for index, (banco, cuenta) in enumerate(cuentas, start=cuenta_start):
        escribir_celda_segura(ws, index, 6, banco)
        escribir_celda_segura(ws, index, 7, cuenta)

        report_col = report_columns.get((clean_text(banco), cuenta))
        if report_col:
            col_letter = get_column_letter(report_col)
            formula = (
                f'=IF(AND(OR($G$5="TODOS",TRIM($G$5)=TRIM(F{index})),'
                f'OR($E$5="TODAS",TEXT($E$5,"0")=TEXT(G{index},"0"))),'
                f'SUMIFS(\'{SHEET_REPORTE}\'!${col_letter}$5:${col_letter}${report_end},'
                f'\'{SHEET_REPORTE}\'!$A$5:$A${report_end},">="&$A$5,'
                f'\'{SHEET_REPORTE}\'!$A$5:$A${report_end},"<="&$C$5),0)'
            )
            escribir_celda_segura(ws, index, 8, formula)


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

        response = send_file(
            output_path,
            as_attachment=True,
            download_name=OUTPUT_FILENAME,
            mimetype=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )
        response.headers["X-Parser-Version"] = PARSER_VERSION
        response.headers["X-Parser-Transactions"] = str(len(all_transactions))
        response.headers["X-Parser-Sheets"] = "; ".join(
            f"{item.get('name')}={item.get('transactions', 0)}"
            for item in processed_sheets
        )[:1500]
        return response

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