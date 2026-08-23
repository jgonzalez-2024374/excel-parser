from flask import Flask, request, jsonify
from openpyxl import load_workbook
import os
import tempfile

app = Flask(__name__)


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "ok",
        "message": "Excel Parser funcionando"
    })


@app.route("/parse-excel", methods=["POST"])
def parse_excel():

    if "file" not in request.files:
        return jsonify({
            "error": "No se recibió ningún archivo"
        }), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({
            "error": "El archivo no tiene nombre"
        }), 400

    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({
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


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )