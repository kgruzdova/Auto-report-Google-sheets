"""
Простое приложение на Tkinter: ввод параметров отчёта, симуляция случайных данных,
запись в Google Таблицу в виде оформленного «документного» листа.
"""

from __future__ import annotations

import os
import random
import re
import sys
import threading
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, List, Sequence, Tuple

from dotenv import load_dotenv

from google_sheets_crud import GoogleSheetsClient

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError as e:
    print("Требуется Tkinter (обычно входит в установку Python).", file=sys.stderr)
    raise SystemExit(1) from e

# --- генерация «фейковых» данных ---

_DEPARTMENTS = ("Офис", "Производство", "Логистика", "Продажи", "Склад")
_REPORT_TYPES = (
    "Ежедневный",
    "Еженедельный",
    "Ежемесячный",
    "Квартальный",
    "Годовой",
)
_RESPONSIBLE_OPTIONS = (
    "Иванов А.А.",
    "Петрова С.В.",
    "Сидоров М.К.",
    "Козлова Е.Н.",
    "Орлов Д.П.",
    "Михайлова Т.Ю.",
    "Николаев И.Г.",
    "Фёдорова А.Р.",
    "Главный бухгалтер",
    "Начальник отдела снабжения",
    "Руководитель проекта",
)
_ITEMS = (
    "Болт М8",
    "Гайка М10",
    "Профиль 40×40",
    "Кабель ВВГ 3×2.5",
    "Краска RAL 9010",
    "Подшипник 6205",
    "Ремень приводной",
    "Фильтр масляный",
    "Шайба DIN 125",
    "Электроды МР-3",
)
_STATUSES = ("ОК", "В ожидании", "Частично", "Отмена")


def _parse_date(s: str) -> date:
    s = s.strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError("Дата: используйте ДД.ММ.ГГГГ или ГГГГ-ММ-ДД")


def _random_date_between(d0: date, d1: date) -> date:
    if d1 < d0:
        d0, d1 = d1, d0
    delta = (d1 - d0).days
    return d0 + timedelta(days=random.randint(0, max(delta, 0)))


def generate_rows(
    date_from: date,
    date_to: date,
    n: int,
) -> List[Tuple[str, str, int, str, float, str]]:
    """Строки: дата, наименование, кол-во, ед., сумма, статус."""
    rows: List[Tuple[str, str, int, str, float, str]] = []
    for i in range(n):
        d = _random_date_between(date_from, date_to)
        qty = random.randint(1, 500)
        unit = random.choice(("шт.", "кг", "м", "уп."))
        price = round(random.uniform(50.0, 150_000.0), 2)
        status = random.choices(_STATUSES, weights=(70, 15, 10, 5), k=1)[0]
        item = random.choice(_ITEMS)
        rows.append(
            (
                d.strftime("%d.%m.%Y"),
                item,
                qty,
                unit,
                price,
                status,
            )
        )
    return rows


def _sheet_title() -> str:
    return "Отчёт_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def _quote_sheet(name: str) -> str:
    """A1-имя листа с экранированием для API."""
    if re.search(r"[^A-Za-z0-9_]", name):
        return "'" + name.replace("'", "''") + "'"
    return name


def build_value_grid(
    date_from: date,
    date_to: date,
    department: str,
    report_type: str,
    responsible: str,
    comment: str,
    data_rows: Sequence[Tuple[str, str, int, str, float, str]],
) -> Tuple[List[List[Any]], int, int]:
    """
    Возвращает матрицу значений, индекс строки заголовка таблицы и индекс строки «Итого».
    """
    cols = 7
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    grid: List[List[Any]] = []

    title = f"ОТЧЁТ: {report_type.upper()}"
    grid.append([title] + [""] * (cols - 1))
    grid.append([""] * cols)

    grid.append(["Период с", date_from.strftime("%d.%m.%Y"), "по", date_to.strftime("%d.%m.%Y"), "", "", ""])
    grid.append(["Подразделение", department, "", "", "", "", ""])
    grid.append(["Тип отчета", report_type, "", "", "", "", ""])
    grid.append(["Сформировано", now_str, "", "", "", "", ""])
    grid.append(["Ответственное лицо", responsible, "", "", "", "", ""])
    grid.append(["Комментарий", comment, "", "", "", "", ""])
    grid.append([""] * cols)

    header_row = len(grid)
    grid.append(["№", "Дата", "Наименование", "Кол-во", "Ед.", "Сумма, ₽", "Статус"])

    total_sum = 0.0
    for i, (d, name, qty, unit, amount, status) in enumerate(data_rows, start=1):
        total_sum += amount
        grid.append([i, d, name, qty, unit, amount, status])

    total_row = len(grid)
    grid.append(["Итого по разделу", "", "", "", "", total_sum, ""])

    return grid, header_row, total_row


def _color(r: float, g: float, b: float) -> dict:
    return {"red": r, "green": g, "blue": b, "alpha": 1.0}


def format_report_requests(sheet_id: int, header_row: int, total_row: int) -> List[dict]:
    """Запросы batchUpdate: объединения, стили, границы, ширины колонок."""
    cols = 7
    # total_row — индекс строки «Итого»; строки данных — [header_row+1, total_row)
    # Строка «Комментарий» — перед пустой строкой и шапкой таблицы
    comment_row = header_row - 2

    requests: List[dict] = []

    # Заголовок документа — объединение и стиль
    requests.append(
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": cols,
                },
                "mergeType": "MERGE_ALL",
            }
        }
    )
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": cols,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _color(0.93, 0.93, 0.93),
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                        "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": _color(0.15, 0.15, 0.2)},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)",
            }
        }
    )

    # Комментарий: объединить ячейки значения B–G
    requests.append(
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": comment_row,
                    "endRowIndex": comment_row + 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": cols,
                },
                "mergeType": "MERGE_ALL",
            }
        }
    )

    # Метаданные (период … комментарий): подписи в колонке A
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 2,
                    "endRowIndex": 8,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True, "fontSize": 10},
                        "horizontalAlignment": "LEFT",
                    }
                },
                "fields": "userEnteredFormat(textFormat,horizontalAlignment)",
            }
        }
    )

    # Текст комментария: перенос по строкам, вертикально сверху
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": comment_row,
                    "endRowIndex": comment_row + 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": cols,
                },
                "cell": {
                    "userEnteredFormat": {
                        "verticalAlignment": "TOP",
                        "wrapStrategy": "WRAP",
                        "textFormat": {"fontSize": 10},
                    }
                },
                "fields": "userEnteredFormat(verticalAlignment,wrapStrategy,textFormat)",
            }
        }
    )

    # Шапка таблицы
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": header_row,
                    "endRowIndex": header_row + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": cols,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _color(0.85, 0.88, 0.92),
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                        "textFormat": {"bold": True, "fontSize": 10},
                        "borders": {
                            "top": {"style": "SOLID", "width": 1, "color": _color(0.6, 0.6, 0.6)},
                            "bottom": {"style": "SOLID", "width": 1, "color": _color(0.6, 0.6, 0.6)},
                            "left": {"style": "SOLID", "width": 1, "color": _color(0.6, 0.6, 0.6)},
                            "right": {"style": "SOLID", "width": 1, "color": _color(0.6, 0.6, 0.6)},
                        },
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat,borders)",
            }
        }
    )

    # Тело таблицы (без строки «Итого»): границы
    if header_row + 1 < total_row:
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": header_row + 1,
                        "endRowIndex": total_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": cols,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP",
                            "borders": {
                                "top": {"style": "SOLID", "width": 1, "color": _color(0.85, 0.85, 0.85)},
                                "bottom": {"style": "SOLID", "width": 1, "color": _color(0.85, 0.85, 0.85)},
                                "left": {"style": "SOLID", "width": 1, "color": _color(0.85, 0.85, 0.85)},
                                "right": {"style": "SOLID", "width": 1, "color": _color(0.85, 0.85, 0.85)},
                            },
                        }
                    },
                    "fields": "userEnteredFormat(verticalAlignment,wrapStrategy,borders)",
                }
            }
        )

    # Числовой формат для колонки «Сумма» (F = index 5)
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": header_row + 1,
                    "endRowIndex": total_row,
                    "startColumnIndex": 5,
                    "endColumnIndex": 6,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"},
                        "horizontalAlignment": "RIGHT",
                    }
                },
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
            }
        }
    )

    # Строка итого
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": total_row,
                    "endRowIndex": total_row + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": cols,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _color(0.96, 0.97, 0.99),
                        "textFormat": {"bold": True, "fontSize": 10},
                        "horizontalAlignment": "LEFT",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        }
    )
    requests.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": total_row,
                    "endRowIndex": total_row + 1,
                    "startColumnIndex": 5,
                    "endColumnIndex": 6,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"},
                        "horizontalAlignment": "RIGHT",
                    }
                },
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
            }
        }
    )

    # Ширины колонок (пиксели)
    widths = (44, 96, 220, 72, 56, 110, 100)
    for i, w in enumerate(widths):
        requests.append(
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i + 1,
                    },
                    "properties": {"pixelSize": w},
                    "fields": "pixelSize",
                }
            }
        )

    # Закрепить верхние строки до конца шапки таблицы (заголовок таблицы виден при прокрутке)
    requests.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": header_row + 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    )

    return requests


def push_report_to_sheets(
    client: GoogleSheetsClient,
    date_from: date,
    date_to: date,
    department: str,
    report_type: str,
    responsible: str,
    comment: str,
    row_count: int,
) -> Tuple[str, str]:
    """
    Создать новый лист, записать данные и форматирование.
    Возвращает (имя листа, URL таблицы).
    """
    random.seed()
    data_rows = generate_rows(date_from, date_to, row_count)
    stitle = _sheet_title()
    grid, header_row, total_row = build_value_grid(
        date_from, date_to, department, report_type, responsible, comment, data_rows
    )

    sheet_id = client.add_sheet(stitle)
    nrows = len(grid)
    ncols = len(grid[0]) if grid else 7
    end_col = chr(ord("A") + ncols - 1)
    range_a1 = f"{_quote_sheet(stitle)}!A1:{end_col}{nrows}"

    client.update_range(range_a1, grid)

    fmt = format_report_requests(sheet_id, header_row, total_row)
    client.batch_update(fmt)

    url = f"https://docs.google.com/spreadsheets/d/{client.spreadsheet_id}/edit#gid={sheet_id}"
    return stitle, url


class ReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор отчётов → Google Таблица")
        self.geometry("520x480")
        self.minsize(480, 420)

        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Период (даты в формате ДД.ММ.ГГГГ)").grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(frm, text="С даты:").grid(row=1, column=0, sticky="e", **pad)
        self.entry_from = ttk.Entry(frm, width=18)
        self.entry_from.grid(row=1, column=1, sticky="w", **pad)
        self.entry_from.insert(0, (date.today().replace(day=1)).strftime("%d.%m.%Y"))

        ttk.Label(frm, text="По дату:").grid(row=2, column=0, sticky="e", **pad)
        self.entry_to = ttk.Entry(frm, width=18)
        self.entry_to.grid(row=2, column=1, sticky="w", **pad)
        self.entry_to.insert(0, date.today().strftime("%d.%m.%Y"))

        ttk.Label(frm, text="Подразделение:").grid(row=3, column=0, sticky="e", **pad)
        self.combo_dept = ttk.Combobox(frm, values=_DEPARTMENTS, width=28, state="readonly")
        self.combo_dept.grid(row=3, column=1, sticky="w", **pad)
        self.combo_dept.current(0)

        ttk.Label(frm, text="Тип отчета:").grid(row=4, column=0, sticky="e", **pad)
        self.combo_type = ttk.Combobox(frm, values=_REPORT_TYPES, width=36, state="readonly")
        self.combo_type.grid(row=4, column=1, sticky="ew", **pad)
        self.combo_type.current(0)

        ttk.Label(frm, text="Число строк (симуляция):").grid(row=5, column=0, sticky="e", **pad)
        self.spin_rows = ttk.Spinbox(frm, from_=3, to=100, width=10)
        self.spin_rows.grid(row=5, column=1, sticky="w", **pad)
        self.spin_rows.delete(0, tk.END)
        self.spin_rows.insert(0, "12")

        ttk.Label(frm, text="Ответственное лицо:").grid(row=6, column=0, sticky="e", **pad)
        self.combo_responsible = ttk.Combobox(
            frm, values=_RESPONSIBLE_OPTIONS, width=36, state="readonly"
        )
        self.combo_responsible.grid(row=6, column=1, sticky="ew", **pad)
        self.combo_responsible.current(0)

        ttk.Label(frm, text="Комментарий:").grid(row=7, column=0, sticky="ne", **pad)
        self.text_comment = tk.Text(frm, width=36, height=4, wrap="word", font=("Segoe UI", 9))
        self.text_comment.grid(row=7, column=1, sticky="ew", **pad)

        frm.columnconfigure(1, weight=1)

        self.btn = ttk.Button(frm, text="Сгенерировать отчет", command=self._on_generate)
        self.btn.grid(row=8, column=0, columnspan=2, pady=16)

        self.status = tk.StringVar(value="Укажите параметры и нажмите кнопку.")
        ttk.Label(frm, textvariable=self.status, wraplength=480, justify=tk.LEFT).grid(
            row=9, column=0, columnspan=2, sticky="w"
        )

        ##ttk.Label(frm, text="Нужны GOOGLE_SPREADSHEET_ID в .env и доступ сервисного аккаунта к таблице.", font=("Segoe UI", 8)).grid(
        ##   row=10, column=0, columnspan=2, sticky="w", pady=(12, 0)
        ##)

    def _on_generate(self) -> None:
        self.btn.configure(state=tk.DISABLED)
        self.status.set("Отправка в Google Sheets…")

        try:
            d0 = _parse_date(self.entry_from.get())
            d1 = _parse_date(self.entry_to.get())
        except ValueError as e:
            self.status.set(str(e))
            self.btn.configure(state=tk.NORMAL)
            messagebox.showerror("Дата", str(e))
            return

        dept = self.combo_dept.get().strip()
        rtype = self.combo_type.get().strip()
        responsible = self.combo_responsible.get().strip()
        comment = self.text_comment.get("1.0", tk.END).strip()
        try:
            n = int(self.spin_rows.get())
        except ValueError:
            self.status.set("Укажите целое число строк.")
            self.btn.configure(state=tk.NORMAL)
            messagebox.showerror("Параметры", "Число строк должно быть целым.")
            return
        n = max(3, min(100, n))

        load_dotenv(Path(__file__).resolve().parent / ".env")
        spreadsheet_id = os.environ.get("GOOGLE_SPREADSHEET_ID", "").strip()
        if not spreadsheet_id:
            self.btn.configure(state=tk.NORMAL)
            messagebox.showerror(
                "Конфигурация",
                "В .env задайте GOOGLE_SPREADSHEET_ID (ID таблицы из URL).",
            )
            self.status.set("Нет GOOGLE_SPREADSHEET_ID.")
            return

        def work() -> None:
            try:
                client = GoogleSheetsClient(spreadsheet_id)
                stitle, url = push_report_to_sheets(
                    client, d0, d1, dept, rtype, responsible, comment, n
                )

                def ok() -> None:
                    self.status.set(f"Готово. Лист «{stitle}» создан. Открыть ссылку можно из диалога.")
                    self.btn.configure(state=tk.NORMAL)
                    if messagebox.askyesno("Готово", f"Лист «{stitle}» записан.\n\nОткрыть таблицу в браузере?"):
                        webbrowser.open(url)

                self.after(0, ok)
            except Exception as ex:
                def err() -> None:
                    self.status.set(f"Ошибка: {ex}")
                    self.btn.configure(state=tk.NORMAL)
                    messagebox.showerror("Ошибка", str(ex))

                self.after(0, err)

        threading.Thread(target=work, daemon=True).start()


def main() -> None:
    load_dotenv(Path(__file__).resolve().parent / ".env")
    app = ReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
