"""
app.py — Flask web-сервис трансформации ВОР → ЛЗК.

Запуск: python app.py
Порт:   6511

Зависимости: flask>=3.0, openpyxl>=3.1
"""

from io import BytesIO

from flask import Flask, request, jsonify, send_file, render_template
import openpyxl
from openpyxl.utils import get_column_letter

from vor_core import (
    col_letter_to_index, detect_sheet, detect_qty_col, detect_data_start,
    transform,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/detect", methods=["POST"])
def detect():
    """
    Принимает Excel-файл (multipart), возвращает JSON:
    {
        "sheets": ["Лист1", "Лист2", ...],
        "detected_sheet": "Лист1" | null,
        "qty_cols": {"Лист1": "G", "Лист2": null, ...}
    }
    """
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Файл не передан"}), 400

    try:
        data = file.read()
        wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    except Exception as e:
        return jsonify({"error": f"Не удалось открыть файл: {e}"}), 422

    sheets = wb.sheetnames
    detected = detect_sheet(wb)

    qty_cols = {}
    col_num_idx = 0  # default: column A
    for name in sheets:
        try:
            ws = wb[name]
            idx = detect_qty_col(ws, col_num_idx)
            qty_cols[name] = get_column_letter(idx + 1) if idx is not None else None
        except Exception:
            qty_cols[name] = None

    wb.close()
    return jsonify({"sheets": sheets, "detected_sheet": detected, "qty_cols": qty_cols})


@app.route("/transform", methods=["POST"])
def do_transform():
    """
    Принимает multipart/form-data:
      file       — Excel-файл ВОР
      sheet      — имя листа
      col_num    — буква/номер колонки № п.п.
      col_name   — буква/номер колонки Наименование
      col_unit   — буква/номер колонки Ед. изм.
      col_qty    — буква/номер колонки Кол-во

    Возвращает Excel-файл результата (attachment).
    Ничего не пишет на диск.
    """
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Файл не передан"}), 400

    sheet_name = request.form.get("sheet", "").strip()
    if not sheet_name:
        return jsonify({"error": "Не указан лист"}), 400

    try:
        col_num  = col_letter_to_index(request.form.get("col_num",  "A"))
        col_name = col_letter_to_index(request.form.get("col_name", "B"))
        col_unit = col_letter_to_index(request.form.get("col_unit", "C"))
        col_qty  = col_letter_to_index(request.form.get("col_qty",  "G"))
    except Exception as e:
        return jsonify({"error": f"Неверное обозначение колонки: {e}"}), 400

    try:
        data = file.read()
        count, mode, buf = transform(
            BytesIO(data), sheet_name,
            col_num, col_name, col_unit, col_qty,
            output_target=None,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 422

    original_name = file.filename or "vor"
    import os
    base = os.path.splitext(original_name)[0]
    out_name = f"{base}_ЛЗК.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=6511, debug=False)
