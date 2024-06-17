import zipfile
import os
from tkinter import Tk, filedialog
import openpyxl
from openpyxl import Workbook

def list_pdfs_in_zip(zip_path):
    pdf_files = []
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            if file_info.filename.endswith('.pdf'):
                pdf_files.append(file_info.filename)
    return pdf_files

def write_list_to_xlsx(file_list, xlsx_path):
    workbook = Workbook()
    sheet = workbook.active
    for i, file_name in enumerate(file_list, start=1):
        # Удаление первых 11 и последних 4 символов
        modified_file_name = file_name[11:-4]
        sheet.cell(row=i, column=1).value = modified_file_name
    workbook.save(xlsx_path)

def main():
    # Инициализация Tkinter
    root = Tk()
    root.withdraw()  # Скрыть основное окно

    # Запрос пути к zip-архиву
    zip_path = filedialog.askopenfilename(title="Выберите zip-архив", filetypes=[("Zip Files", "*.zip")])
    if not zip_path:
        print("Файл не выбран")
        return

    # Запрос пути для сохранения Excel файла
    xlsx_path = filedialog.asksaveasfilename(title="Сохранить как", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if not xlsx_path:
        print("Путь для сохранения не выбран")
        return

    # Получение списка PDF файлов и запись их в Excel файл
    pdf_files = list_pdfs_in_zip(zip_path)
    write_list_to_xlsx(pdf_files, xlsx_path)

    print(f"Названия PDF файлов записаны в {xlsx_path}")

if __name__ == "__main__":
    main()
