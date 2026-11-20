import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

def clean_text(text):
    """Очистка текста от переносов строк и лишних пробелов"""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

def is_valid_row(row_data):
    """Проверка, является ли строка валидной записью (имеет номер и описание)"""
    # Обычно номер строки - это первая колонка, но в PDF она может быть смещена
    if len(row_data) < 2:
        return False
    
    # Проверяем, есть ли в первой или второй колонке число (номер строки)
    col0 = clean_text(row_data[0])
    if not col0.isdigit():
        return False
    return True

def process_pdf(pdf_path, excel_path):
    extracted_data = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    
                    for row in table:
                        if not row:
                            continue
                            
                        # Пропускаем заголовки таблиц
                        row_text = " ".join([clean_text(cell) for cell in row if cell])
                        if "Наименование вида работ" in row_text or "Код ГЭСН" in row_text:
                            continue
                        
                        # Очищаем ячейки
                        cleaned_row = [clean_text(cell) for cell in row]
                        
                        # Обработка строк, где номер строки есть, а остальное может быть разбито
                        if is_valid_row(cleaned_row):
                            # Формат PDF: [№ строки, Наименование, Ед.изм, Код ГЭСН, Кол-во]
                            # Иногда колонки могут съезжать, поэтому берем по индексам
                            row_num = cleaned_row[0]
                            description = cleaned_row[1] if len(cleaned_row) > 1 else ""
                            unit = cleaned_row[2] if len(cleaned_row) > 2 else ""
                            quantity = cleaned_row[4] if len(cleaned_row) > 4 else ""
                            
                            # Если кол-во пустое, проверяем 3-ю колонку
                            if not quantity and len(cleaned_row) > 3:
                                quantity = cleaned_row[3]
                            
                            extracted_data.append({
                                '№строки по ВР': row_num,
                                'Наименование вида работ': description,
                                'Ед. изм.': unit,
                                'Кол-во': quantity
                            })
                        else:
                            # Обработка текстовых примечаний, которые идут без номера, 
                            # но относятся к предыдущей строке (или просто являются примечаниями)
                            # В вашем примере они иногда имеют номера, иногда нет.
                            # Если номера нет, пропускаем, чтобы не засорять таблицу.
                            pass
                            
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать PDF: {e}")
        return False

    if not extracted_data:
        messagebox.showwarning("Внимание", "Не удалось найти табличные данные в PDF.")
        return False

    # Создаем DataFrame
    df = pd.DataFrame(extracted_data)
    
    # Добавляем пустые колонки согласно требованиям (ВиКР не заполняем)
    df['примечание'] = ""
    df['Уровень структуры*'] = ""
    df['Шифр РД'] = ""
    df['Код ВиКР'] = ""
    df['Наменование ВиКР'] = "" # Опечатка из примера сохранена
    df['ед. изм. ВиКР'] = ""
    df['Кол-во ВиКР'] = ""
    df['Кприв/Плечо возки'] = ""
    
    # Переупорядочиваем колонки как в примере
    columns_order = [
        '№строки по ВР', 'Наименование вида работ', 'Ед. изм.', 'Кол-во', 
        'примечание', 'Уровень структуры*', 'Шифр РД', 
        'Код ВиКР', 'Наменование ВиКР', 'ед. изм. ВиКР', 'Кол-во ВиКР', 'Кприв/Плечо возки'
    ]
    df = df[columns_order]

    # Сохраняем в Excel
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Выборка')
            
            # Автоподбор ширины колонок
            worksheet = writer.sheets['Выборка']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80) # Ограничиваем макс. ширину
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
        messagebox.showinfo("Успех", f"Файл успешно сохранен:\n{excel_path}")
        return True
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить Excel: {e}")
        return False

def select_files():
    root = tk.Tk()
    root.withdraw() # Скрываем главное окно
    
    pdf_path = filedialog.askopenfilename(
        title="Выберите исходный PDF файл",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return

    excel_path = filedialog.asksaveasfilename(
        title="Сохранить Excel как",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="Выборка_из_PDF.xlsx"
    )
    if not excel_path:
        return

    process_pdf(pdf_path, excel_path)

if __name__ == "__main__":
    select_files()
