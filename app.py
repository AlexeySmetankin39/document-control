import excel_parsing as ep
import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                             QDialog, QVBoxLayout,
                             QTableWidget, QTableWidgetItem, QDialogButtonBox)
from PyQt6.QtCore import Qt
from datetime import datetime
import pandas as pd
import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

class ShopConfirmationDialog(QDialog):
    def __init__(self, file_shop_map, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Подтверждение магазинов")
        self.setGeometry(200, 200, 600, 400)
        
        layout = QVBoxLayout()
        
        # Создаем таблицу для отображения файлов и магазинов
        self.table = QTableWidget(len(file_shop_map), 2)
        self.table.setHorizontalHeaderLabels(["Файл", "Магазин"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        
        # Заполняем таблицу данными
        for row, (filename, shop) in enumerate(file_shop_map.items()):
            self.table.setItem(row, 0, QTableWidgetItem(filename))
            self.table.setItem(row, 1, QTableWidgetItem(shop))
        
        layout.addWidget(self.table)
        
        # Кнопки подтверждения
        button_box = QDialogButtonBox()
        self.confirm_button = button_box.addButton("Верно", QDialogButtonBox.ButtonRole.AcceptRole)
        self.reject_button = button_box.addButton("Неверно", QDialogButtonBox.ButtonRole.RejectRole)
        
        self.confirm_button.clicked.connect(self.accept)
        self.reject_button.clicked.connect(self.reject)
            
        layout.addWidget(button_box)
        self.setLayout(layout)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Парсер цен")
        self.setGeometry(300, 300, 900, 800)
        
        self.button = QPushButton("Выбрать файлы", self)
        self.button.setGeometry(40, 40, 150, 40)
        self.button.clicked.connect(self.process_files)

        self.label = QLabel("Выберите файлы для обработки", self)
        self.label.setGeometry(40, 100, 800, 30)
        self.label.setStyleSheet("font-size: 16px;")

        self.label_defined_shop = QLabel("", self)
        self.label_defined_shop.setGeometry(40, 140, 800, 200)
        self.label_defined_shop.setStyleSheet("font-size: 14px;")
        self.label_defined_shop.setAlignment(Qt.AlignmentFlag.AlignTop)

    def process_files(self):
        file_paths = ep.select_files()
        if not file_paths:
            return
            
        # Сначала определяем магазины для всех файлов
        file_shop_map = {}
        for file_path in file_paths:
            try:
                df = ep.read_data(file_path)
                shop_name = ep.detect_shop(file_path, df, self)
                filename = os.path.basename(file_path)
                file_shop_map[filename] = shop_name
            except Exception as e:
                print(f"Ошибка при определении магазина для {file_path}: {str(e)}")
                filename = os.path.basename(file_path)
                file_shop_map[filename] = "Ошибка определения"
        
        # Показываем пользователю таблицу соответствия
        dialog = ShopConfirmationDialog(file_shop_map, self)
        result = dialog.exec()
        
        # Если пользователь сказал "Неверно", запрашиваем магазины вручную
        if result == QDialog.DialogCode.Rejected:
            for file_path in file_paths:
                filename = os.path.basename(file_path)
                shop_name, ok = ep.get_shop_name_from_user(self, filename)
                if ok and shop_name.strip():
                    file_shop_map[filename] = shop_name
        
        # Обновляем статус в интерфейсе
        status_text = "Определённые магазины:\n"
        for filename, shop in file_shop_map.items():
            status_text += f"{filename}: {shop}\n"
        self.label_defined_shop.setText(status_text)
        QApplication.processEvents()
        
        # Теперь обрабатываем файлы с подтвержденными магазинами
        all_numbers, all_products, all_articles, all_quantities, all_costs, all_sum, all_shops = ep.process_files_with_shops(
            file_paths, self, file_shop_map
        )
        
        # Формируем результат
        df_result = pd.DataFrame({
            '№': all_numbers,
            'Наименование': all_products,
            'Магазин': all_shops,
            'Артикул': all_articles,
            'Кол-во': all_quantities,
            'Цена': all_costs,
            'Сумма': all_sum
        })
        
        # Сохраняем результат в Excel
        out_dir = "out"
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)
        filename = f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(out_dir, filename)
        
        # Создаем временный файл для данных
        temp_filepath = os.path.join(out_dir, "temp_" + filename)
        df_result.to_excel(temp_filepath, index=False)
        
        try:
            # Создаем новую книгу для итогового результата
            wb_result = Workbook()
            ws_result = wb_result.active
            
            # Копируем header.xlsx если он существует
            header_path = "header.xlsx"
            header_exists = os.path.exists(header_path)
            header_row_count = 0
            
            if header_exists:
                wb_header = load_workbook(header_path)
                ws_header = wb_header.active
                
                # Копируем все ячейки
                for row in ws_header.iter_rows():
                        for cell in row:
                            ws_result.cell(
                                row=cell.row, 
                                column=cell.column, 
                                value=cell.value
                            )
                        
                
                # Копируем объединенные ячейки
                for merged_range in ws_header.merged_cells.ranges:
                    ws_result.merge_cells(str(merged_range))
                
                header_row_count = ws_header.max_row
            
            # Копируем основную таблицу с отступом
            wb_temp = load_workbook(temp_filepath)
            ws_temp = wb_temp.active
            
            start_row = header_row_count + 2 if header_exists else 1
            
            for row in ws_temp.iter_rows():
                for cell in row:
                    ws_result.cell(
                        row=cell.row + start_row - 1,
                        column=cell.column,
                        value=cell.value
                    )
            
            # Удаляем временный файл
            os.remove(temp_filepath)
            
            # Устанавливаем ширину столбцов
            ws_result.column_dimensions['A'].width = 4
            ws_result.column_dimensions['B'].width = 40
            ws_result.column_dimensions['C'].width = 12
            ws_result.column_dimensions['D'].width = 14
            ws_result.column_dimensions['E'].width = 8
            ws_result.column_dimensions['F'].width = 8
            ws_result.column_dimensions['G'].width = 8
            
            # Выравнивание данных
            center_alignment = Alignment(horizontal='center', vertical='center')
            data_start_row = start_row + 1
            data_end_row = start_row + len(df_result)
            
            # Центрирование для столбцов A, C-G
            for col_idx in [1, 3, 4, 5, 6, 7]:  # A, C, D, E, F, G
                for row_idx in range(data_start_row, data_end_row + 1):
                    cell = ws_result.cell(row=row_idx, column=col_idx)
                    cell.alignment = center_alignment
            
            # Перенос текста для столбца B
            for row_idx in range(data_start_row, data_end_row + 1):
                cell = ws_result.cell(row=row_idx, column=2)
                cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            # Вставляем подвал из bottom.xlsx
            bottom_path = "bottom.xlsx"
            bottom_exists = os.path.exists(bottom_path)
            if bottom_exists:
                wb_bottom = load_workbook(bottom_path)
                ws_bottom = wb_bottom.active
                
                # Начальная строка для подвала: после основной таблицы + 2 пустые строки
                start_bottom_row = data_end_row + 3
                
                # Копируем ячейки
                for row in ws_bottom.iter_rows():
                        for cell in row:
                            ws_result.cell(
                                row=cell.row + start_bottom_row - 1,
                                column=cell.column,
                                value=cell.value
                            )  
                
                # Копируем объединенные ячейки
                for merged_range in ws_bottom.merged_cells.ranges:
                    # Сдвигаем диапазон
                    new_range = f"{get_column_letter(merged_range.min_col)}{merged_range.min_row + start_bottom_row - 1}:{get_column_letter(merged_range.max_col)}{merged_range.max_row + start_bottom_row - 1}"
                    ws_result.merge_cells(new_range)
            
            # Сохраняем итоговый файл
            wb_result.save(filepath)
            
        except Exception as e:
            print(f"Ошибка при формировании файла: {str(e)}")
            # Если возникла ошибка - сохраняем без шапки и подвала
            df_result.to_excel(filepath, index=False)
            if os.path.exists(temp_filepath):
                os.remove(temp_filepath)
        
        self.label.setText(f"Результаты сохранены в файл: {filepath}")


if __name__ == "__main__":  
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())