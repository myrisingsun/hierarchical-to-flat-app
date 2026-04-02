"""
transform_vor.py — Универсальный инструмент трансформации иерархического ВОР в плоский список ЛЗК.

Запуск: python transform_vor.py
Зависимости: openpyxl (pip install openpyxl), tkinter (стандартная библиотека Python)
"""

import re
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

HIERARCHY_RE = re.compile(r'^\d+(\.\d+)*$')


def col_letter_to_index(letter: str) -> int:
    """Буква колонки (A-Z) → 0-based индекс. Принимает также числовые строки."""
    letter = letter.strip().upper()
    if letter.isdigit():
        return int(letter) - 1  # если пользователь ввёл цифру
    if len(letter) == 1 and letter.isalpha():
        return ord(letter) - ord('A')
    # многобуквенные (AA, AB...) — редко нужно, но пусть будет
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def is_hierarchy_num(value) -> bool:
    """Проверить, является ли значение номером в иерархии (строка типа '1', '1.1', '1.1.1')."""
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
    """Найти первую строку данных: строка, где col_num содержит номер уровня 1 ('1', '2', ...)."""
    first_level_re = re.compile(r'^\d+$')
    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        val = row[col_num] if col_num < len(row) else None
        if val is not None and first_level_re.match(str(val).strip()):
            return i
    return 1


def detect_qty_col(ws, col_num: int) -> int | None:
    """
    Найти колонку с итоговым количеством материала/работы.
    Ищем в заголовочных строках ключевые слова, а также смотрим на первую строку
    материала: первая колонка с числовым значением после col_num+3.
    """
    data_start = detect_data_start(ws, col_num)
    # Попытка 1: ищем в заголовках по приоритетным ключевым словам
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
    # Попытка 2: смотрим на первую строку материала (col[0]=None)
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
            name = row[col_num + 1] if col_num + 1 < len(row) else None  # обычно col B
            if qty is not None and name:
                return True
        checked += 1
        if checked > 200:
            break
    return False


# ---------------------------------------------------------------------------
# Трансформация
# ---------------------------------------------------------------------------

def transform(input_file: str, sheet_name: str,
              col_num: int, col_name: int, col_unit: int, col_qty: int,
              output_file: str,
              progress_callback=None) -> tuple[int, str]:
    """
    Трансформирует иерархический ВОР в плоский список.
    Поддерживает два режима (авто-определение):
      - «С материалами»: под каждой работой есть строки материалов (col_num=None)
      - «Только работы»: строк материалов нет → каждая работа сама становится строкой
    Возвращает (количество строк, описание режима).
    """
    wb = openpyxl.load_workbook(input_file, data_only=True)
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
                continue  # пустые строки-разделители — пропускаем
            # Режим «с материалами»: строка материала
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
                    # Режим «только работы»: столбец материала остаётся пустым
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
            # Вне иерархии (сводная таблица и т.п.) — останавливаем, если уже есть данные
            if rows_out:
                break

    # Записать результат
    _write_output(rows_out, output_file)
    return len(rows_out), mode


def _write_output(rows: list, output_file: str):
    """Записать плоский список в Excel-файл."""
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

    # Стиль заголовка
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

    # Данные
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

    # Ширина колонок
    col_widths = [10, 25, 50, 50, 12, 12, 12, 12]
    for ci, w in enumerate(col_widths, start=1):
        ws_out.column_dimensions[get_column_letter(ci)].width = w

    ws_out.freeze_panes = "A2"

    wb_out.save(output_file)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Трансформация ВОР → Плоский список (ЛЗК)")
        self.resizable(False, False)
        self.configure(padx=16, pady=16)

        self._input_path = tk.StringVar()
        self._sheet_var = tk.StringVar()
        self._col_num = tk.StringVar(value="A")
        self._col_name = tk.StringVar(value="B")
        self._col_unit = tk.StringVar(value="C")
        self._col_qty = tk.StringVar(value="G")
        self._output_path = tk.StringVar()
        self._status = tk.StringVar(value="Выберите входной файл ВОР")
        self._sheets = []

        self._build_ui()
        self._check_openpyxl()

    # ------------------------------------------------------------------
    def _check_openpyxl(self):
        if not OPENPYXL_OK:
            messagebox.showerror(
                "Отсутствует зависимость",
                "Библиотека openpyxl не установлена.\n\n"
                "Установите её командой:\n  pip install openpyxl\n\n"
                "Затем перезапустите программу."
            )
            self.destroy()

    # ------------------------------------------------------------------
    def _build_ui(self):
        PAD = {"padx": 6, "pady": 4}

        # --- Входной файл ---
        frame_in = ttk.LabelFrame(self, text="Входной файл ВОР")
        frame_in.grid(row=0, column=0, sticky="ew", **PAD)
        ttk.Entry(frame_in, textvariable=self._input_path, width=58,
                  state="readonly").grid(row=0, column=0, padx=6, pady=6)
        ttk.Button(frame_in, text="Открыть…", command=self._browse_input).grid(
            row=0, column=1, padx=6, pady=6)

        # --- Лист ---
        frame_sheet = ttk.LabelFrame(self, text="Лист")
        frame_sheet.grid(row=1, column=0, sticky="ew", **PAD)
        self._sheet_cb = ttk.Combobox(frame_sheet, textvariable=self._sheet_var,
                                       width=40, state="readonly")
        self._sheet_cb.grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Label(frame_sheet, text="(авто-определяется при открытии файла)",
                  foreground="gray").grid(row=0, column=1, padx=6)

        # --- Настройка колонок ---
        frame_cols = ttk.LabelFrame(self, text="Настройка колонок (буква A-Z или номер)")
        frame_cols.grid(row=2, column=0, sticky="ew", **PAD)

        col_defs = [
            ("№ п.п.:",     self._col_num,  0),
            ("Наименование:", self._col_name, 1),
            ("Ед. изм.:",    self._col_unit, 2),
            ("Кол-во итого:", self._col_qty,  3),
        ]
        for label, var, ci in col_defs:
            r, c = divmod(ci, 2)
            ttk.Label(frame_cols, text=label).grid(row=r, column=c * 2, padx=6, pady=4, sticky="e")
            ttk.Entry(frame_cols, textvariable=var, width=6).grid(
                row=r, column=c * 2 + 1, padx=6, pady=4, sticky="w")

        # --- Выходной файл ---
        frame_out = ttk.LabelFrame(self, text="Выходной файл")
        frame_out.grid(row=3, column=0, sticky="ew", **PAD)
        ttk.Entry(frame_out, textvariable=self._output_path, width=58).grid(
            row=0, column=0, padx=6, pady=6)
        ttk.Button(frame_out, text="Изменить…", command=self._browse_output).grid(
            row=0, column=1, padx=6, pady=6)

        # --- Прогресс + кнопка ---
        frame_act = ttk.Frame(self)
        frame_act.grid(row=4, column=0, sticky="ew", **PAD)
        self._progress = ttk.Progressbar(frame_act, length=400, mode="determinate")
        self._progress.grid(row=0, column=0, padx=6, pady=6)
        self._btn_run = ttk.Button(frame_act, text="Преобразовать",
                                    command=self._run, width=18)
        self._btn_run.grid(row=0, column=1, padx=6, pady=6)

        # --- Статус ---
        ttk.Label(self, textvariable=self._status,
                  relief="sunken", anchor="w", width=70).grid(
            row=5, column=0, sticky="ew", padx=6, pady=(4, 0))

        # --- Копирайт ---
        ttk.Label(
            self,
            text="Разработка: Руководитель отдела архитектуры и проектирования — Пахарев Кирилл  |  kpakharev@afid.ru",
            foreground="gray",
            font=("TkDefaultFont", 8),
            anchor="center",
        ).grid(row=6, column=0, sticky="ew", padx=6, pady=(6, 2))

    # ------------------------------------------------------------------
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Выберите файл ВОР",
            filetypes=[("Excel файлы", "*.xlsx *.xlsm"), ("Все файлы", "*.*")]
        )
        if not path:
            return
        self._input_path.set(path)
        self._status.set("Загрузка файла…")
        self.update_idletasks()
        self._load_file_info(path)

    def _load_file_info(self, path: str):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
            self._status.set("Ошибка открытия файла")
            return

        self._sheets = sheets
        self._sheet_cb["values"] = sheets

        # Авто-определение листа
        try:
            wb2 = openpyxl.load_workbook(path, data_only=True)
            best_sheet = detect_sheet(wb2)
            wb2.close()
        except Exception:
            best_sheet = None

        if best_sheet:
            self._sheet_var.set(best_sheet)
            self._status.set(f"Лист авто-определён: «{best_sheet}»")
            self._auto_detect_qty_col(path, best_sheet)
        elif sheets:
            self._sheet_var.set(sheets[0])
            self._status.set("Лист авто-определить не удалось — выберите вручную")
        else:
            self._status.set("В файле нет листов")

        # Выходной файл по умолчанию
        base = os.path.splitext(path)[0]
        self._output_path.set(base + "_ЛЗК.xlsx")

        # Обновление при смене листа вручную
        self._sheet_cb.bind("<<ComboboxSelected>>", self._on_sheet_changed)

    def _on_sheet_changed(self, _event=None):
        path = self._input_path.get()
        sheet = self._sheet_var.get()
        if path and sheet:
            self._auto_detect_qty_col(path, sheet)

    def _auto_detect_qty_col(self, path: str, sheet_name: str):
        """Попытаться авто-определить колонку Кол-во итого."""
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb[sheet_name]
            col_num_idx = col_letter_to_index(self._col_num.get() or "A")
            idx = detect_qty_col(ws, col_num_idx)
            wb.close()
            if idx is not None:
                letter = get_column_letter(idx + 1)
                self._col_qty.set(letter)
        except Exception:
            pass

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
        )
        if path:
            self._output_path.set(path)

    # ------------------------------------------------------------------
    def _run(self):
        input_file = self._input_path.get()
        sheet_name = self._sheet_var.get()
        output_file = self._output_path.get()

        if not input_file:
            messagebox.showwarning("Нет файла", "Выберите входной файл ВОР.")
            return
        if not sheet_name:
            messagebox.showwarning("Нет листа", "Выберите лист для обработки.")
            return
        if not output_file:
            messagebox.showwarning("Нет пути", "Укажите путь для выходного файла.")
            return

        # Проверка колонок
        try:
            col_num  = col_letter_to_index(self._col_num.get())
            col_name = col_letter_to_index(self._col_name.get())
            col_unit = col_letter_to_index(self._col_unit.get())
            col_qty  = col_letter_to_index(self._col_qty.get())
        except Exception as e:
            messagebox.showerror("Ошибка колонок", f"Неверное обозначение колонки:\n{e}")
            return

        self._btn_run.config(state="disabled")
        self._progress["value"] = 0
        self._status.set("Обработка…")

        def worker():
            try:
                count, mode = transform(
                    input_file, sheet_name,
                    col_num, col_name, col_unit, col_qty,
                    output_file,
                    progress_callback=self._update_progress,
                )
                self.after(0, lambda c=count, md=mode, f=output_file: self._on_done(c, md, f))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda m=msg: self._on_error(m))

        threading.Thread(target=worker, daemon=True).start()

    def _update_progress(self, current: int, total: int):
        pct = int(current / total * 100) if total else 0
        self.after(0, lambda v=pct: self._progress.__setitem__("value", v))

    def _on_done(self, count: int, mode: str, output_file: str):
        self._progress["value"] = 100
        self._btn_run.config(state="normal")
        self._status.set(f"Готово! Режим: {mode} | Записано строк: {count}")
        if messagebox.askyesno(
            "Готово",
            f"Трансформация завершена.\nРежим: {mode}\nЗаписано строк: {count}\n\nОткрыть выходной файл?"
        ):
            os.startfile(output_file)

    def _on_error(self, msg: str):
        self._progress["value"] = 0
        self._btn_run.config(state="normal")
        self._status.set("Ошибка!")
        messagebox.showerror("Ошибка при обработке", msg)


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
