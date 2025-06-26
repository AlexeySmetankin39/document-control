import sys
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def select_files():
    """
    Открывает диалоговое окно выбора файлов
    Возвращает список выбранных путей или None, если выбор отменен
    """
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно
    root.attributes('-topmost', True)  # Поверх других окон
    
    file_paths = filedialog.askopenfilenames(
        title='Выберите файлы',
        filetypes=[
        ("CSV файлы", "*.csv"),
        ("Excel файлы", "*.xlsx *.xls"),
        ("Все файлы", "*.*")
    ]
    )
    return file_paths if file_paths else None

def read_data(file_path):
    """
    Читает данные из файла и возвращает DataFrame
    """
    if file_path.endswith('.csv'):

        return pd.read_csv(file_path, sep=None, engine='python')  # Чтение CSV файла с заголовком и кодировкой UTF-8, указание разделителя
    elif file_path.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format")

selected_files_name = [] # Список для хранения выбранных файлов
extensions = [] # Список для хранения расширений файлов
data_in_files = [] # Список для хранения данных из файлов
name_products = [] # Список для хранения названий продуктов
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
   
        self.setWindowTitle("my app") # Устанавливаем заголовок окна
        self.setGeometry(100, 100, 300, 200) # Устанавливаем размеры окна

        self.button = QPushButton("download file", self) # Создаем кнопку
        self.button.setGeometry(100, 80, 100, 30) # Устанавливаем размеры и позицию кнопки
        self.button.clicked.connect(self.on_button_click) # Подключаем обработчик нажатия кнопки

    def on_button_click(self): # Обработчик нажатия кнопки
        selected_files_name = select_files() # Вызываем функцию выбора файлов
        if selected_files_name:# Проверяем, что файлы выбраны
            print("Выбраны файлы:", selected_files_name)


        for file_path in selected_files_name:  # Проходим по каждому выбранному файлу
            try:
                data_in_files.append(read_data(file_path))  # Читаем данные из файла
            except Exception as e:
                print(f"Ошибка при чтении {file_path}: {e}")

        
        delete_substring = ['nan','Итого']     
        j = 0
        while j < len(data_in_files): # Проходим по каждому DataFrame в списке
            if data_in_files: # Проверяем, что список не пуст
                column_names = [df.columns.tolist() for df in data_in_files] # Получаем названия столбцов из каждого DataFrame
                idx = 0
                found = False
                if column_names[j]:
                    for idx, col in enumerate(column_names[j]):
                        if "Наименование" in col:
                            found = True
                            break
                if found:
                    name_products.append(data_in_files[j].iloc[:, idx])
                    j += 1
                else:
                    k = 0
                    while k < len(data_in_files[j]):
                        row = data_in_files[j].loc[k]
                        found_inner = False
                        for col_idx, cell in enumerate(row):
                            if "Наименование" in str(cell):
                                name_products.append(data_in_files[j].iloc[k + 1:, col_idx])
                                found_inner = True
                                break
                        if found_inner:
                            break
                        k += 1
                    j += 1 
            else:
                break
        # Собираем все названия продуктов в один список
        all_products = []
        for prod_series in name_products:
            all_products.extend(prod_series.tolist())
        # Удаляем значения, совпадающие со списком delete_substring
        all_products = [prod for prod in all_products if str(prod) not in delete_substring]
        print("Все названия продуктов:", all_products)


app = QApplication(sys.argv)

window = MainWindow()
window.show()
  
app.exec()