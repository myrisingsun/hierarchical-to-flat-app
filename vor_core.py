"""
vor_core.py — Ядро трансформации ВОР → ЛЗК (без GUI, без дисковых зависимостей).

Используется как transform_vor.py (desktop GUI), так и app.py (Flask web).
"""

import re
from io import BytesIO

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

HIERARCHY_RE = re.compile(r'^\d+(\.\d+)*$')


def col_letter_to_index(letter: str) -> int:
    """Буква колонки (A-Z) → 0-based индекс. Принимает также числовые строки."""
    letter = letter.strip().upper()
    if letter.isdigit():
        return int(letter) - 1
    if len(letter) == 1 and letter.isalpha():
        return ord(letter) - ord('A')
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def is_hierarchy_num(value) -> bool:
    """Проверить, является ли значение номером в иерархии ('1', '1.1', '1.1.1')."""
    if value is None:
        return False
    return bool(HIERARCHY_RE.match(str(value).strip()))


def detect_sheet(wb) -> str | None:
    """Найти лист с иерархическими данными ВОР. Возвращает имя листа или None."""
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        count = 0
        for row in ws.iter_rows(min_row=1, max_row=min(100, ws.max_row), values_only=True):
            if row and is_hierarchy_num(row[0]):
                count += 1
                if count >= 3:
                    return sheet_name
    return None


def detect_data_start(ws, col_num: int) -> int:
    """Найти первую строку данных: строка, где col_num = номер уровня 1 ('1', '2', ...)."""
    first_level_re = re.compile(r'^\d+$')
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        val = row[col_num] if col_num < len(row) else None
        if val is not None and first_level_re.match(str(val).strip()):
            return i
    return 1


def detect_qty_col(ws, col_num: int) -> int | None:
    """
    Найти колонку с итоговым количеством материала/работы.
    Ищем в заголовочных строках ключевые слова, а также смотрим на строки данных.
    """
    data_start = detect_data_start(ws, col_num)
    priority_keywords = ('итого', 'расход', 'списан', 'списание')
    fallback_keywords = ('кол-во', 'количество')
    best_priority = None
    best_fallback = None
    for row in ws.iter_rows(min_row=max(1, data_start - 5), max_row=data_start - 1, values_only=True):
        for ci, cell in enumerate(row):
            if not cell:
                continue
            cell_low = str(cell).lower()
            if best_priority is None and any(k in cell_low for k in priority_keywords):
                best_priority = ci
            if best_fallback is None and any(k in cell_low for k in fallback_keywords):
                best_fallback = ci
    if best_priority is not None:
        return best_priority
    if best_fallback is not None:
        return best_fallback
    # Попытка 2: первая строка материала (col_num=None)
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        num = row[col_num] if col_num < len(row) else None
        if num is None:
            for ci in range(col_num + 3, min(len(row), col_num + 10)):
                if row[ci] is not None and isinstance(row[ci], (int, float)):
                    return ci
            break
    # Попытка 3: первая числовая ячейка в строке работы
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        num = row[col_num] if col_num < len(row) else None
        if is_hierarchy_num(num) and str(num).count('.') >= 2:
            for ci in range(col_num + 2, min(len(row), col_num + 8)):
                if row[ci] is not None and isinstance(row[ci], (int, float)):
                    return ci
            break
    return None


def sheet_has_material_rows(ws, col_num: int, col_qty: int, data_start: int) -> bool:
    """Проверить, есть ли на листе строки материалов (col_num=None с числовым кол-вом)."""
    checked = 0
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        num = row[col_num] if col_num < len(row) else None
        if num is None:
            qty = row[col_qty] if col_qty < len(row) else None
            name = row[col_num + 1] if col_num + 1 < len(row) else None
            if qty is not None and name:
                return True
        checked += 1
        if checked > 200:
            break
    return False


# ---------------------------------------------------------------------------
# Трансформация
# ---------------------------------------------------------------------------

def transform(
    input_source,
    sheet_name: str,
    col_num: int,
    col_name: int,
    col_unit: int,
    col_qty: int,
    output_target=None,
    progress_callback=None,
) -> tuple[int, str, BytesIO | None]:
    """
    Трансформирует иерархический ВОР в плоский список.

    input_source  — путь к файлу (str) или файлоподобный объект (BytesIO)
    output_target — путь к файлу (str), файлоподобный объект, или None.
                    Если None — возвращает BytesIO с результатом (диск не задействован).

    Возвращает (количество_строк, режим, buf_или_None).
    """
    wb = openpyxl.load_workbook(input_source, data_only=True)
    ws = wb[sheet_name]

    data_start = detect_data_start(ws, col_num)
    has_materials = sheet_has_material_rows(ws, col_num, col_qty, data_start)
    mode = "с материалами" if has_materials else "только работы"

    current_section = ""
    current_work_num = ""
    current_work_name = ""
    rows_out = []

    total_rows = ws.max_row
    processed = 0

    for row in ws.iter_rows(min_row=data_start, values_only=True):
        processed += 1
        if progress_callback and processed % 100 == 0:
            progress_callback(processed, total_rows)

        if not row:
            continue

        num = row[col_num] if col_num < len(row) else None
        name = row[col_name] if col_name < len(row) else None

        if num is None:
            if not has_materials:
                continue
            qty_val = row[col_qty] if col_qty < len(row) else None
            unit_val = row[col_unit] if col_unit < len(row) else None
            if qty_val is not None and name:
                rows_out.append((
                    current_work_num,
                    current_section,
                    current_work_name,
                    str(name).strip(),
                    str(unit_val).strip() if unit_val else "",
                    qty_val,
                    None, None,
                ))
        elif is_hierarchy_num(num):
            num_str = str(num).strip()
            depth = num_str.count('.')
            if depth == 0:
                current_section = str(name).strip() if name else ""
            if depth >= 2:
                current_work_num = num_str
                current_work_name = str(name).strip() if name else ""
                if not has_materials:
                    qty_val = row[col_qty] if col_qty < len(row) else None
                    unit_val = row[col_unit] if col_unit < len(row) else None
                    if current_work_name:
                        rows_out.append((
                            current_work_num,
                            current_section,
                            current_work_name,
                            None,
                            str(unit_val).strip() if unit_val else "",
                            qty_val,
                            None, None,
                        ))
        else:
            if rows_out:
                break

    buf = _write_output(rows_out, output_target)
    return len(rows_out), mode, buf


def _write_output(rows: list, output_target) -> BytesIO | None:
    """
    Записать плоский список в Excel.
    output_target=None  → вернуть BytesIO (без диска)
    output_target=str   → сохранить в файл, вернуть None
    """
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "ЛЗК"

    headers = [
        "№ работ",
        "Раздел сметы",
        "Наименование работ",
        "Наименование материалов",
        "Ед.\nизм. (материалы)",
        "Кол-во (материалы)",
        "Цена",
        "Сумма",
    ]

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_out.append(headers)
    for ci, _ in enumerate(headers, start=1):
        cell = ws_out.cell(row=1, column=ci)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    ws_out.row_dimensions[1].height = 30

    cell_align_wrap = Alignment(vertical="top", wrap_text=True)
    cell_align_center = Alignment(horizontal="center", vertical="top")

    for row in rows:
        ws_out.append(list(row))
        ri = ws_out.max_row
        ws_out.cell(ri, 1).alignment = cell_align_center
        ws_out.cell(ri, 2).alignment = cell_align_wrap
        ws_out.cell(ri, 3).alignment = cell_align_wrap
        ws_out.cell(ri, 4).alignment = cell_align_wrap
        ws_out.cell(ri, 5).alignment = cell_align_center
        ws_out.cell(ri, 6).alignment = cell_align_center

    col_widths = [10, 25, 50, 50, 12, 12, 12, 12]
    for ci, w in enumerate(col_widths, start=1):
        ws_out.column_dimensions[get_column_letter(ci)].width = w

    ws_out.freeze_panes = "A2"

    if output_target is None:
        buf = BytesIO()
        wb_out.save(buf)
        buf.seek(0)
        return buf
    else:
        wb_out.save(output_target)
        return None
