import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton
from datetime import datetime

# Ключевые слова для поиска столбцов
NAME_KEYWORDS = ['Наименование', 'NAME', 'Наименование товара']
COST_KEYWORDS = ['Цена, RUB', 'RBL_PRICE', 'Цена', 'Цена за 1 шт., руб.','Стоимость']
ARTICLE_KEYWORDS = ['Номенклатурный номер', 'NOM_N', 'Код товара','Артикул']
QUANTITY_KEYWORDS = ['Кол-во', 'NUMBER_OF', 'Кол-во, шт.', 'Количество' ]

def select_files():
    """Открывает диалоговое окно выбора файлов"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_paths = filedialog.askopenfilenames(
        title='Выберите файлы',
        filetypes=[
            ("CSV файлы", "*.csv"),
            ("Excel файлы", "*.xlsx *.xls"),
            ("Все файлы", "*.*")
        ]
    )
    return file_paths if file_paths else None

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

def process_files(file_paths):
    """Обрабатывает все выбранные файлы и извлекает данные"""
    all_products = []
    all_costs = []
    all_articles = []
    all_quantities = []
    all_sum = []
    delete_substrings = ['nan', 'Итого', 'Сумма товаров в заказе', 'Кол-во товаров в заказе']

    for file_path in file_paths:
        try:
            df = read_data(file_path)
            print(f"\nОбработка файла: {file_path}")
            
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
                if any(sub in name_val or sub in cost_val or sub in article_val for sub in delete_substrings or sub in qountity_val):
                    continue
                
                # Преобразуем цену в числовой формат
                try:
                    cost_val = float(cost_val.replace(',', '.').replace(' ', ''))
                    all_products.append(name_val)
                    all_costs.append(cost_val)
                    all_articles.append(article_val)
                    all_quantities.append(qountity_val)
                    all_sum.append(cost_val * float(str(qountity_val).replace(',', '.').replace(' ', '')))
                except ValueError:
                    continue
                        
        except Exception as e:
            print(f"Ошибка при обработке {file_path}: {str(e)}")
    
    # Вывод результатов
    print("\nРезультаты обработки:")
    print(f"Всего позиций: {len(all_products)}")
    for name, article, qoutity, cost, sum in zip(all_products, all_articles, all_quantities, all_costs, all_sum ):
        print(f"{name}  ({article}) - {qoutity} шт. : {cost} руб. = {sum} руб.")
    return all_products, all_articles, all_quantities, all_costs, all_sum

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер цен")
        self.setGeometry(100, 100, 300, 200)
        
        self.button = QPushButton("Выбрать файлы", self)
        self.button.setGeometry(100, 80, 100, 30)
        self.button.clicked.connect(self.process_files)

    def process_files(self):
        file_paths = select_files()
        if file_paths:
            all_products, all_articles, all_quantities, all_costs, all_sum = process_files(file_paths)
            df_result = pd.DataFrame({
                'Наименование': all_products,
                'Артикул': all_articles,
                'Количество': all_quantities,
                'Цена': all_costs,
                'Сумма': all_sum
            })
            filename = f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_result.to_excel(filename, index=False)
            print(f"\nРезультаты сохранены в файл: {filename}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())