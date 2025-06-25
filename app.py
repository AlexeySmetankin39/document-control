import sys
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton
import tkinter as tk
from tkinter import filedialog


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

selected_files_name = []
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
   
        self.setWindowTitle("my app")
        self.setGeometry(100, 100, 300, 200)

        self.button = QPushButton("download file", self)
        self.button.setGeometry(100, 80, 100, 30)
        self.button.clicked.connect(self.on_button_click)

    def on_button_click(self):
        selected_files_name = select_files()
        if selected_files_name:
            print("Выбраны файлы:", selected_files_name)
        extensions = []
        for file_path in selected_files_name:
            if '.' in file_path:
                ext = file_path.split('.')[-1]
                extensions.append(ext)
        print("Расширения файлов:", extensions)
            


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()