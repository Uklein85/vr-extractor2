import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# ---------------------- Вспомогательные функции ----------------------

def clean_text(value):
    """Убирает переносы строк, множественные пробелы, None."""
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def is_row_number(s):
    """Проверяет, что строка — это номер строки (возможно с точкой)."""
    if not s:
        return False
    return s.strip().rstrip(".").isdigit()


def is_section_header(row):
    """
    Заголовок раздела: одна непустая ячейка, начинающаяся с числа и текста,
    например «1 Устройство утепления балочной клетки».
    """
    non_empty = [c for c in row if c]
    if len(non_empty) != 1:
        return False
    return bool(re.match(r"^\d+\s+[А-ЯЁA-Zа-яё]", non_empty[0]))


def is_skip_row(row_text):
    """Служебные строки (шапка, подвал, обозначение документа)."""
    markers = [
        "Наименование вида работ",
        "Код ГЭСН",
        "Изм.", "Кол.уч", "Лист", "№док", "Подпись", "Дата",
        "Ведомость объемов",
        "0349.032",
        "ФОРМАТ",
        "Разраб.", "Пров.", "Нач.отд.", "Н.контр.",
        "Главный инженер",
    ]
    return any(m in row_text for m in markers)


# ---------------------- Основная обработка ----------------------

def process_pdf(pdf_path, excel_path, root):
    extracted = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    if not table:
                        continue
                    for row in table:
                        if not row:
                            continue

                        cleaned = [clean_text(c) for c in row]
                        row_text = " ".join(c for c in cleaned if c).strip()
                        if not row_text:
                            continue
                        if is_skip_row(row_text):
                            continue

                        # --- Заголовок раздела ---
                        if is_section_header(cleaned):
                            m = re.match(r"^(\d+)\s+(.+)$", cleaned[0])
                            if m:
                                extracted.append({
                                    "№строки по ВР": m.group(1),
                                    "Наименование вида работ": m.group(2),
                                    "Ед. изм.": "",
                                    "Кол-во": "",
                                })
                            continue

                        # --- Обычная строка с номером ---
                        if is_row_number(cleaned[0]):
                            num  = cleaned[0].rstrip(".")
                            desc = cleaned[1] if len(cleaned) > 1 else ""
                            unit = ""
                            qty  = ""

                            if len(cleaned) >= 5:
                                unit = cleaned[2]
                                qty  = cleaned[4]
                            elif len(cleaned) == 4:
                                unit = cleaned[2]
                                qty  = cleaned[3]
                            elif len(cleaned) == 3:
                                qty  = cleaned[2]

                            extracted.append({
                                "№строки по ВР": num,
                                "Наименование вида работ": desc,
                                "Ед. изм.": unit,
                                "Кол-во": qty,
                            })
                            continue

                        # --- Продолжение предыдущей строки ---
                        if not cleaned[0] and extracted:
                            continuation = " ".join(c for c in cleaned[1:] if c)
                            if continuation:
                                for i in range(len(extracted) - 1, -1, -1):
                                    if extracted[i]["Наименование вида работ"]:
                                        extracted[i]["Наименование вида работ"] += " " + continuation
                                        break
                            # Иногда в продолжении бывает кол-во (5-я колонка)
                            if len(cleaned) >= 5 and cleaned[4] and not extracted[-1]["Кол-во"]:
                                extracted[-1]["Кол-во"] = cleaned[4]
                            continue
    except Exception as e:
        messagebox.showerror("Ошибка чтения PDF", str(e), parent=root)
        return False

    if not extracted:
        messagebox.showwarning(
            "Внимание",
            "В PDF не найдено табличных данных. Проверьте, что файл не защищён и содержит текст (не скан).",
            parent=root,
        )
        return False

    # --- Сборка итоговой таблицы ---
    df = pd.DataFrame(extracted)
    for col in [
        "примечание", "Уровень структуры*", "Шифр РД",
        "Код ВиКР", "Наменование ВиКР", "ед. изм. ВиКР",
        "Кол-во ВиКР", "Кприв/Плечо возки",
    ]:
        df[col] = ""

    df = df[[
        "№строки по ВР", "Наименование вида работ", "Ед. изм.", "Кол-во",
        "примечание", "Уровень структуры*", "Шифр РД",
        "Код ВиКР", "Наменование ВиКР", "ед. изм. ВиКР",
        "Кол-во ВиКР", "Кприв/Плечо возки",
    ]]

    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="ВР исходный")
            ws = writer.sheets["ВР исходный"]
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_len = max(
                    (len(str(c.value)) for c in column_cells if c.value), default=0
                )
                ws.column_dimensions[letter].width = min(max_len + 2, 80)
    except Exception as e:
        messagebox.showerror("Ошибка записи Excel", str(e), parent=root)
        return False

    messagebox.showinfo("Готово", f"Файл сохранён:\n{excel_path}", parent=root)
    return True


# ---------------------- Точка входа ----------------------

def main():
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="Выберите PDF-файл ведомости объёмов работ",
        filetypes=[("PDF files", "*.pdf")],
    )
    if not pdf_path:
        root.destroy()
        return

    excel_path = filedialog.asksaveasfilename(
        title="Куда сохранить выборку",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="Выборка_из_PDF.xlsx",
    )
    if not excel_path:
        root.destroy()
        return

    process_pdf(pdf_path, excel_path, root)
    root.destroy()


if __name__ == "__main__":
    main()
