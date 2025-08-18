import pandas as pd
from openpyxl import load_workbook
import os


class ExcelWriter:
    def __init__(self, file_path):
        """
        Инициализация класса для работы с Excel файлом

        :param file_path: Полный путь к Excel файлу
        """
        self.file_path = file_path

    def write_to_sheet(self, data, sheet_name, start_row=1, start_col=1):
        """
        Запись данных в указанный лист Excel файла

        :param data: Данные для записи (список списков или DataFrame)
        :param sheet_name: Имя листа
        :param start_row: Начальная строка (по умолчанию 1)
        :param start_col: Начальный столбец (по умолчанию 1)
        """
        try:
            # Если файл существует, загружаем его
            if os.path.exists(self.file_path):
                # Загружаем существующую книгу
                book = load_workbook(self.file_path)

                # Если лист существует, удаляем его
                if sheet_name in book.sheetnames:
                    del book[sheet_name]

                # Создаем новый лист
                sheet = book.create_sheet(sheet_name)

                # Записываем данные
                if isinstance(data, pd.DataFrame):
                    # Если данные в формате DataFrame
                    for r_idx, row in enumerate(data.values, start=start_row):
                        for c_idx, value in enumerate(row, start=start_col):
                            sheet.cell(row=r_idx, column=c_idx, value=value)
                else:
                    # Если данные в формате списка списков
                    for r_idx, row in enumerate(data, start=start_row):
                        for c_idx, value in enumerate(row, start=start_col):
                            sheet.cell(row=r_idx, column=c_idx, value=value)

                # Сохраняем книгу
                book.save(self.file_path)

            else:
                # Если файл не существует, создаем новый
                if isinstance(data, pd.DataFrame):
                    with pd.ExcelWriter(self.file_path, engine='openpyxl') as writer:
                        data.to_excel(writer, sheet_name=sheet_name, startrow=start_row - 1, startcol=start_col - 1,
                                      index=False)
                else:
                    # Преобразуем список списков в DataFrame для удобства
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(self.file_path, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name=sheet_name, startrow=start_row - 1, startcol=start_col - 1,
                                    index=False)

            print(f"Данные успешно записаны в лист '{sheet_name}' файла {self.file_path}")

        except Exception as e:
            print(f"Ошибка при записи в Excel: {e}")

    def append_to_sheet(self, data, sheet_name):
        """
        Добавление данных в конец существующего листа

        :param data: Данные для добавления
        :param sheet_name: Имя листа
        """
        try:
            if os.path.exists(self.file_path):
                book = load_workbook(self.file_path)

                if sheet_name in book.sheetnames:
                    sheet = book[sheet_name]
                    # Определяем следующую строку
                    start_row = sheet.max_row + 1
                else:
                    sheet = book.create_sheet(sheet_name)
                    start_row = 1

                # Записываем данные
                if isinstance(data, pd.DataFrame):
                    for r_idx, row in enumerate(data.values, start=start_row):
                        for c_idx, value in enumerate(row, start=1):
                            sheet.cell(row=r_idx, column=c_idx, value=value)
                else:
                    for r_idx, row in enumerate(data, start=start_row):
                        for c_idx, value in enumerate(row, start=1):
                            sheet.cell(row=r_idx, column=c_idx, value=value)

                book.save(self.file_path)
                print(f"Данные успешно добавлены в лист '{sheet_name}' файла {self.file_path}")
            else:
                # Если файла нет, просто создаем его
                self.write_to_sheet(data, sheet_name)

        except Exception as e:
            print(f"Ошибка при добавлении данных в Excel: {e}")

# Пример использования:
# writer = ExcelWriter("C:/path/to/your/file.xlsx")
# writer.write_to_sheet(result_data, "Sheet1")