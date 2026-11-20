import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def extract_data_from_pdf(pdf_path, root):
    extracted = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Извлекаем отдельные слова с их координатами
                words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=False)
                if not words:
                    continue
                
                # Сортируем слова: сверху вниз, слева направо
                words.sort(key=lambda w: (round(w['top'], 1), w['x0']))
                
                # Группируем слова в строки
                lines = []
                current_line = []
                current_top = None

                for word in words:
                    if current_top is None:
                        current_top = word['top']
                        current_line.append(word)
                    elif abs(word['top'] - current_top) < 5: # Допуск по высоте строки
                        current_line.append(word)
                    else:
                        if current_line:
                            lines.append(current_line)
                        current_line = [word]
                        current_top = word['top']
                if current_line:
                    lines.append(current_line)

                for line_words in lines:
                    line_text = " ".join([w['text'] for w in line_words]).strip()
                    if not line_text:
                        continue
                    
                    # Отсеиваем мусорные строки (шапки, подвалы, примечания)
                    if any(m in line_text for m in [
                        "Наименование вида работ", "Код ГЭСН", "Изм.", "Кол.уч", 
                        "Лист", "№док", "Подпись", "Дата", "Ведомость", 
                        "0349.032", "ФОРМАТ", "Разраб.", "Пров.", "Нач.отд.", 
                        "Н.контр.", "Главный инженер", "Взам.инв"
                    ]):
                        continue
                    
                    first_word = line_words[0]
                    # Проверяем, начинается ли строка с номера (от 1 до 100)
                    num_match = re.match(r'^(\d+)$', first_word['text'].strip('.'))
                    
                    if num_match and int(num_match.group(1)) <= 100:
                        num = num_match.group(1)
                        rest_words = line_words[1:]
                        rest_text = " ".join([w['text'] for w in rest_words]).strip()
                        
                        qty = ""
                        unit = ""
                        desc = rest_text
                        
                        # Ищем количество (число) в правой части страницы (x0 > 600)
                        qty_word = None
                        for w in reversed(rest_words):
                            if w['x0'] > 600 and re.match(r'^\d+[\.,]?\d*$', w['text']):
                                qty_word = w
                                break
                        
                        if qty_word:
                            qty = qty_word['text']
                            # Удаляем количество из текста описания
                            desc = rest_text.replace(qty, "").strip()
                            desc = re.sub(r'\s+', ' ', desc)
                            
                            # Ищем единицу измерения в конце описания
                            unit_match = re.search(r'\b(м3|м2|м|т|шт\.?|кг)\s*$', desc, re.IGNORECASE)
                            if unit_match:
                                unit = unit_match.group(1)
                                desc = desc[:unit_match.start()].strip()
                            elif desc.endswith("-"):
                                unit = "-"
                                desc = desc[:-1].strip()
                        else:
                            # Если количества нет, проверяем на наличие "-" в конце
                            if rest_text.endswith("-"):
                                unit = "-"
                                qty = "-"
                                desc = rest_text[:-1].strip()
                        
                        if desc or unit or qty:
                            extracted.append({
                                "№строки по ВР": num,
                                "Наименование вида работ": desc,
                                "Ед. изм.": unit,
                                "Кол-во": qty
                            })
                    else:
                        # Обработка заголовков разделов (например, "1 Устройство утепления...")
                        sec_match = re.match(r'^(\d+)\s+([А-Яа-яЁё].*)$', line_text)
                        if sec_match:
                            extracted.append({
                                "№строки по ВР": sec_match.group(1),
                                "Наименование вида работ": sec_match.group(2),
                                "Ед. изм.": "",
                                "Кол-во": ""
                            })
                        elif extracted:
                            # Если это продолжение предыдущей строки, прикрепляем к ней
                            extracted[-1]["Наименование вида работ"] += " " + line_text
                            # Если в продолжении есть количество, обновляем
                            if not extracted[-1]["Кол-во"]:
                                qty_match = re.search(r'([\d,\.]+)\s*$', line_text)
                                if qty_match:
                                    extracted[-1]["Кол-во"] = qty_match.group(1)
                                    extracted[-1]["Наименование вида работ"] = extracted[-1]["Наименование вида работ"].replace(qty_match.group(1), "").strip()

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при чтении PDF: {e}", parent=root)
        return None

    return extracted

def process_pdf(pdf_path, excel_path, root):
    extracted = extract_data_from_pdf(pdf_path, root)
    
    if extracted is None:
        return False
    if not extracted:
        messagebox.showwarning(
            "Внимание", 
            "В PDF не найдено табличных данных. Проверьте, что файл не защищён и содержит текстовый слой (не просто картинку).", 
            parent=root
        )
        return False

    # Сборка DataFrame
    df = pd.DataFrame(extracted)
    for col in [
        "примечание", "Уровень структуры*", "Шифр РД",
        "Код ВиКР", "Наменование ВиКР", "ед. изм. ВиКР",
        "Кол-во ВиКР", "Кприв/Плечо возки",
    ]:
        df[col] = ""

    columns_order = [
        "№строки по ВР", "Наименование вида работ", "Ед. изм.", "Кол-во",
        "примечание", "Уровень структуры*", "Шифр РД",
        "Код ВиКР", "Наменование ВиКР", "ед. изм. ВиКР",
        "Кол-во ВиКР", "Кприв/Плечо возки",
    ]
    df = df[columns_order]

    # Запись в Excel
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="ВР исходный")
            ws = writer.sheets["ВР исходный"]
            
            # Автоподбор ширины колонок
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_len = max((len(str(c.value)) for c in column_cells if c.value), default=0)
                ws.column_dimensions[letter].width = min(max_len + 2, 80)
    except Exception as e:
        messagebox.showerror("Ошибка записи", f"Не удалось сохранить Excel: {e}", parent=root)
        return False

    messagebox.showinfo("Готово", f"Файл успешно сохранён:\n{excel_path}", parent=root)
    return True

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
