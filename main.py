import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import tkinter as tk
from tkinter import filedialog
import os
import random

def generate_badges():
    """Главная функция для выбора файла, чтения данных и генерации пропусков."""
    
    # 1. Настройка окружения
    root = tk.Tk()
    root.withdraw()

    # Проверка папок
    templates_dir = 'templates'
    output_dir = 'output'
    
    if not os.path.exists(templates_dir):
        print(f"ОШИБКА: Папка '{templates_dir}' не найдена! Создайте её и поместите туда шаблоны.")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Путь к фото-заглушке
    default_photo_path = 'default_photo.jpg'
    if not os.path.exists(default_photo_path):
        print(f"ОШИБКА: Файл-заглушка '{default_photo_path}' не найден. Добавьте любое фото в папку с программой.")
        return

    # 2. Выбор типа шаблона
    print("Выберите тип пропуска:")
    print("1 - Студент (template-student.docx)")
    print("2 - Сотрудник (template-worker.docx)")
    
    choice = input("Введите 1 или 2: ").strip()
    
    if choice == '1':
        template_name = 'template-student.docx'
        type_name = 'student'
    elif choice == '2':
        template_name = 'template-worker.docx'
        type_name = 'worker'
    else:
        print("Неверный выбор. Программа завершена.")
        return

    template_path = os.path.join(templates_dir, template_name)
    
    if not os.path.exists(template_path):
        print(f"ОШИБКА: Шаблон '{template_name}' не найден в папке '{templates_dir}'!")
        return

    # 3. Выбор Excel/CSV файла
    print("\nПожалуйста, выберите файл с данными (.xlsx или .csv) в открывшемся окне...")
    file_path = filedialog.askopenfilename(
        title="Выберите файл с данными",
        filetypes=[("Data files", "*.xlsx *.csv")]
    )

    if not file_path:
        print("Файл не выбран. Программа завершена.")
        return

    # Определяем тип файла и читаем его (с универсальной логикой для CSV)
    df = None
    
    if file_path.lower().endswith('.csv'):
        # Логика для CSV: пробуем разные разделители и кодировки
        delimiters = [',', ';', '\t'] 
        encodings = ['cp1251', 'utf-8']

        for sep in delimiters:
            for enc in encodings:
                try:
                    df = pd.read_csv(file_path, header=0, encoding=enc, sep=sep, dtype=str)
                    # Проверяем, что колонок больше одной (признак правильного разделителя)
                    if df.shape[1] > 1:
                        print(f"Файл успешно прочитан как CSV. Разделитель: '{sep}', Кодировка: '{enc}'")
                        break # Выходим из цикла encodings
                except Exception:
                    continue
            if df is not None and df.shape[1] > 1:
                break # Выходим из цикла delimiters
        
        if df is None or df.shape[1] <= 1:
            print("\nОШИБКА: Не удалось прочитать CSV файл. Пожалуйста, проверьте разделитель (должен быть ',' или ';').")
            return

    elif file_path.lower().endswith('.xlsx'):
        try:
            df = pd.read_excel(file_path, header=0, dtype=str)
            print("Файл успешно прочитан как XLSX.")
        except Exception as e:
            print(f"Ошибка при чтении Excel: {e}")
            return
    else:
        print("Выбранный файл не является CSV или XLSX. Программа завершена.")
        return

    # Отладочная информация: выводим заголовки
    print("\n[DEBUG] Прочитанные заголовки (Колонки A-E):")
    # Обрезаем до первых 5 колонок, так как нам нужно только A, B, C, D, E
    if len(df.columns) >= 5:
        print(df.columns[:5].tolist())
    else:
        print(df.columns.tolist())
        print("ВНИМАНИЕ: Прочитано менее 5 колонок. Проверьте правильность разделителя в CSV.")
        
    
    # Подготовка данных:
    # 1. Обрезаем до 5 нужных колонок (A, B, C, D, E)
    # 2. Удаляем строки, где Фамилия (первая колонка) пуста
    # 3. Заполняем остальные пустые ячейки пустой строкой ''
    df = df.iloc[:, 0:5].dropna(subset=[df.columns[0]]).fillna('')
    
    print(f"\nНайдено {len(df)} записей для обработки.")
    
    # Отладочная информация: выводим первую строку данных
    if not df.empty:
        print("\n[DEBUG] Первая запись сотрудника (данные для проверки):")
        print(df.iloc[0].tolist())
    
    # 4. Обработка каждой строки и генерация документов
    count = 0
    for index, row in df.iterrows():
        try:
            # Извлечение данных по индексу колонки (0-индексирование)
            surname = str(row.iloc[0]).strip()        # Колонка A - Фамилия ({{ firstname }})
            first_name = str(row.iloc[1]).strip()     # Колонка B - Имя ({{ name }})
            patronymic = str(row.iloc[2]).strip()     # Колонка C - Отчество ({{ lastname }})
            department = str(row.iloc[3]).strip()     # Колонка D - Отдел ({{ department }})
            position = str(row.iloc[4]).strip()       # Колонка E - Должность ({{ position }})
            
            # Генерация кода работника/студента (8 цифр)
            unique_code = f"{random.randint(0, 99999999):08}"

            # Загружаем шаблон
            doc = DocxTemplate(template_path)

            # Подготовка изображения (ширина 30мм)
            img_obj = InlineImage(doc, default_photo_path, width=Mm(30))

            # Словарь контекста (ВСЕ переменные из обоих шаблонов)
            context = {
                'photo': img_obj,
                'firstname': surname,
                'name': first_name,
                'lastname': patronymic,
                'department': department,
                
                # Поля, которые могут быть в одном из шаблонов:
                'position': position,          # Для сотрудника
                'codeWorker': unique_code,     # Для сотрудника
                
                # Поля для студента (используем заглушки или данные из Excel, где возможно)
                'codeStudent': unique_code,
                'middlename': patronymic,      # Используем Отчество для middlename
                'specialty': department,       # Используем Отдел/Департамент как Специальность
                'group_date_start': '01.09.2025', # Заглушка
                'form_name': 'Очная'           # Заглушка
            }

            # Рендеринг
            doc.render(context)

            # Сохранение файла
            output_filename = f"Пропуск_{type_name}_{surname}_{first_name}.docx"
            output_path = os.path.join(output_dir, output_filename)
            
            doc.save(output_path)
            print(f"Создан: {output_filename} | Код: {unique_code}")
            count += 1
            
        except Exception as e:
            # Логгирование ошибки с конкретной строкой
            print(f"\nКРИТИЧЕСКАЯ ОШИБКА при обработке {index+2}-й строки (Фамилия: {surname}): {e}")
            print("Проверьте шаблон на наличие опечаток или невидимых символов вокруг тегов.")
            
    print(f"\n--- ПРОЦЕСС ЗАВЕРШЕН ---")
    print(f"Создано карточек: {count}")
    print(f"Файлы находятся в папке: {os.path.abspath(output_dir)}")
    
if __name__ == "__main__":
    generate_badges()