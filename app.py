import excel_parsing as ep
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel
from datetime import datetime
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер цен")
        self.setGeometry(300, 300, 900, 800)
        
        self.label = QLabel("Выберите файлы для обработки", self)
        self.label.setGeometry(40, 120, 800, 30)
        self.label.setStyleSheet("font-size: 16px;")

        self.button = QPushButton("Выбрать файлы", self)
        self.button.setGeometry(40, 40, 150, 40)
        self.button.clicked.connect(self.process_files)

    def process_files(self):
        file_paths = ep.select_files()
        if file_paths:
            all_numbers, all_products, all_articles, all_quantities, all_costs, all_sum, all_shops = ep.process_files(file_paths, self)  
            df_result = pd.DataFrame({
                '№': all_numbers,
                'Наименование': all_products,
                'Магазин': all_shops,
                'Артикул': all_articles,
                'Кол-во': all_quantities,
                'Цена': all_costs,
                'Сумма': all_sum
            })
            out_dir = "out"
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)
            filename = f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(out_dir, filename)
            df_result.to_excel(filepath, index=False)
            try:
                # Открываем созданный файл с помощью openpyxl
                wb = load_workbook(filepath)
                ws = wb.active
                
                # Устанавливаем ширину столбцов
                ws.column_dimensions['A'].width = 4  # Столбец "№"
                ws.column_dimensions['B'].width = 40  # Столбец "Наименование"
                ws.column_dimensions['C'].width = 12  # Столбец "Магазин"
                ws.column_dimensions['D'].width = 14  # Столбец "Артикул"
                ws.column_dimensions['E'].width = 8   # Столбец "Кол-во"
                ws.column_dimensions['F'].width = 8  # Столбец "Цена"
                ws.column_dimensions['G'].width = 8  # Столбец "Сумма"
                
                # Выравнивание по центру для всех столбцов, кроме "Наименование" (столбец B)
                center_alignment = Alignment(horizontal='center', vertical='center')
                for col in [1, 3, 4, 5, 6, 7]:  # A, C, D, E, F, G
                    for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                        for cell in row:
                            cell.alignment = center_alignment

                # Включаем перенос текста для столбца "Наименование"
                for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
                    for cell in row:
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                
                # Сохраняем изменения
                wb.save(filepath)
                
            except Exception as e:
                print(f"Ошибка при настройке Excel: {str(e)}")
            
        self.label.setText(f"Результаты сохранены в файл: {filepath}")


if __name__ == "__main__":  
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())