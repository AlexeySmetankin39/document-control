from tkinter import filedialog
import sys
import tkinter as tk
import pandas as pd
import os
import re
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                             QInputDialog, QDialog, QVBoxLayout, QHBoxLayout,
                             QTableWidget, QTableWidgetItem, QDialogButtonBox, QFileDialog)
from PyQt6.QtCore import Qt

# Ключевые слова для поиска столбцов
NAME_KEYWORDS = ['Наименование', 'NAME', 'Наименование товара']
COST_KEYWORDS = ['Цена, RUB', 'RBL_PRICE', 'Цена', 'Цена за 1 шт., руб.','Стоимость']
ARTICLE_KEYWORDS = ['Номенклатурный номер', 'NOM_N', 'Код товара','Артикул']
QUANTITY_KEYWORDS = ['Кол-во', 'NUMBER_OF', 'Кол-во, шт.', 'Количество' ]

def select_files(main_window):
    """Открывает диалоговое окно выбора файлов с помощью QFileDialog (PyQt6)"""
    file_dialog = QFileDialog(main_window)
    file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
    file_dialog.setNameFilters([
        "Все файлы (*.*)",
        "CSV файлы (*.csv)",
        "Excel файлы (*.xlsx *.xls)"
    ])
    if file_dialog.exec():
        return file_dialog.selectedFiles()
    return None

def find_header_row(df, keywords):
    """Находит строку с заголовками в DataFrame"""
    for i in range(min(10, len(df))):
        for cell in df.iloc[i]:
            if any(keyword in str(cell) for keyword in keywords):
                return i
    return None

def read_data(file_path):
    """Читает данные из файла с обработкой различных форматов"""
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path, sep=';', encoding='utf-8-sig')
    elif file_path.endswith(('.xlsx', '.xls')):
        # Первоначальное чтение без заголовка
        df = pd.read_excel(file_path, header=None)
        
        # Поиск строки с заголовками для названий
        header_row = find_header_row(df, NAME_KEYWORDS + COST_KEYWORDS + ARTICLE_KEYWORDS + QUANTITY_KEYWORDS)
        
        if header_row is not None:
            # Перечитываем с правильным заголовком
            df = pd.read_excel(file_path, header=header_row)
        return df
    else:
        raise ValueError("Unsupported file format")

def detect_shop(file_path, df, main_window):
    """Определяет магазин по имени файла и содержимому"""
    filename = os.path.basename(file_path)
    name_no_ext = os.path.splitext(filename)[0]
    
    # Правило 1: Если имя файла состоит только из цифр - Минимакс
    if re.fullmatch(r'^\d+$', name_no_ext):
        return "Минимакс"
    
    # Правило 2: Если в имени есть "basket" - Платан
    if 'basket' in name_no_ext.lower():
        return "Платан"
    
    # Правило 3: Если в имени есть "chipdip" - Чип и Дип
    if 'chipdip' in name_no_ext.lower():
        return "Чип и Дип"
    
    # Правило 4: Если в имени есть '№' и 'от' - Все инструменты
    if '№' in name_no_ext and 'от' in name_no_ext:
        return "Все инструменты"

    # Проверка содержимого на ЭТМ (по артикулам)
    for col in df.columns:
        if any(str(x).startswith('ETM') for x in df[col]):
            return "ЭТМ"
    
    # Если ни одно правило не сработало - показать диалоговое окно
    return get_shop_name_from_user(main_window, filename)[0] or name_no_ext

def get_shop_name_from_user(main_window, filename):
    """Запрашивает название магазина у пользователя"""
    shop_name, ok = QInputDialog.getText(
        main_window, 
        "Не удалось определить магазин",
        f"Введите название магазина для файла:\n{filename}",
        text=os.path.splitext(filename)[0]
    )
    return shop_name, ok

def format_number(value):
    """Форматирует число, убирая .0 для целых чисел"""
    try:
        # Пробуем преобразовать в число
        num = float(value)
        if num.is_integer():
            return str(int(num))
        return str(num)
    except ValueError:
        return str(value)

def process_files_with_shops(file_paths, main_window, file_shop_map):
    """Обрабатывает файлы с использованием предопределенных магазинов"""
    all_products = []
    all_costs = []
    all_articles = []
    all_quantities = []
    all_sum = []
    all_shops = []
    all_numbers = []
    
    delete_substrings = ['nan', 'Итого', 'Сумма товаров в заказе', 'Кол-во товаров в заказе']

    # Обновляем статус в главном окне
    main_window.label.setText(f"Обработка {len(file_paths)} файлов...")
    QApplication.processEvents()  # Обновляем интерфейс
    
    for i, file_path in enumerate(file_paths):
        try:
            filename = os.path.basename(file_path)
            shop_name = file_shop_map.get(filename, filename)
            
            # Обновляем статус для каждого файла
            main_window.label.setText(f"Обработка файла {i+1}/{len(file_paths)}: {filename}")
            QApplication.processEvents()  # Обновляем интерфейс
            
            df = read_data(file_path)
            print(f"\nОбработка файла: {file_path}")
            print(f"Используется магазин: {shop_name}")
            
            # Поиск столбцов по ключевым словам
            name_col = None
            cost_col = None
            article_col = None
            qountity_col = None
            
            # Проверяем все ячейки в DataFrame
            for col in df.columns:
                col_str = str(col)
                if any(keyword in col_str for keyword in NAME_KEYWORDS):
                    name_col = col
                if any(keyword in col_str for keyword in COST_KEYWORDS):
                    cost_col = col
                if any(keyword in col_str for keyword in ARTICLE_KEYWORDS):
                    article_col = col
                if any(keyword in col_str for keyword in QUANTITY_KEYWORDS):
                    qountity_col = col
            
            if name_col is None or cost_col is None or article_col is None or qountity_col is None:
                print(f"  Столбцы не найдены: name_col={name_col}, article_col={article_col}, qountity_col={qountity_col} , cost_col={cost_col}")
                continue
            
            print(f"  Найденные столбцы: '{name_col}', '{article_col}', '{qountity_col}', '{cost_col}'")
            
            # Сбор данных с обработкой
            for idx in range(len(df)):
                name_val = str(df.loc[idx, name_col])
                article_val = str(df.loc[idx, article_col])
                qountity_val = str(df.loc[idx, qountity_col])
                cost_val = str(df.loc[idx, cost_col])
                
                # Пропускаем нежелательные записи
                if any(sub in name_val or sub in cost_val or sub in article_val or sub in qountity_val for sub in delete_substrings):
                    continue
                
                # Преобразуем цену и количество в числовой формат
                try:
                    cost_val_float = float(cost_val.replace(',', '.').replace(' ', ''))
                    quantity = float(str(qountity_val).replace(',', '.').replace(' ', ''))
                    
                    # Форматируем артикул и количество
                    article_val = format_number(article_val)
                    qountity_val = format_number(quantity)
                    
                    all_products.append(name_val)
                    all_costs.append(cost_val_float)
                    all_articles.append(article_val)
                    all_quantities.append(qountity_val)
                    all_sum.append(cost_val_float * quantity)
                    all_shops.append(shop_name)
                    all_numbers.append(len(all_products))
                except ValueError:
                    continue
                        
        except Exception as e:
            print(f"Ошибка при обработке {file_path}: {str(e)}")
            # Обновляем статус об ошибке
            main_window.label.setText(f"Ошибка при обработке файла: {os.path.basename(file_path)}")
            QApplication.processEvents()  # Обновляем интерфейс
    
    return all_numbers, all_products, all_articles, all_quantities, all_costs, all_sum, all_shops