from datetime import date, timedelta
from tkinter import messagebox

import pandas as pd

import JsonWork


class SearchJP:
    def __init__(self):
        self.config = JsonWork.JsonConfig()

        self.file_path = self.config.getJPPathFile_input()
        self.data = None

        self.load_excel()
        self.get_dict_all_data()

    def get_dict_all_data(self):
        """
        Возвращает весь словарь данных из self.data.

        :return: словарь в формате {индекс: {столбец: значение, ...}}
        """
        if self.data is None:
            print("Данные не загружены.")
            return {}

        result_dict = {}
        for index, row in self.data.iterrows():
            result_dict[index] = row.to_dict()
        return result_dict

    def load_excel(self):
        """Загружает данные из Excel файла."""
        try:
            self.data = pd.read_excel(self.file_path)
            print("Файл успешно загружен.")
        except Exception as e:
            print(f"Ошибка при загрузке файла: {e}")


class SearchCz:
    def __init__(self):
        self.columns_save = []
        self.filtered_data = None
        self.headers = None
        self.config = JsonWork.JsonConfig()

        self.file_path = self.config.getCzPathFile_input()
        self.data = None

        self.load_excel()
        self.get_dict_all_data()

    def get_dict_all_data(self):
        """
        Возвращает весь словарь данных из self.data.

        :return: словарь в формате {индекс: {столбец: значение, ...}}
        """
        if self.data is None:
            print("Данные не загружены.")
            return {}

        result_dict = {}
        for index, row in self.data.iterrows():
            result_dict[index] = row.to_dict()
        return result_dict

    def get_headers(self):
        return self.headers

    def load_excel(self):
        """Загружает данные из Excel файла."""
        try:
            # Укажите имя нужного листа
            self.data = pd.read_excel(self.file_path, sheet_name=self.config.getCzColumnName("Table of contents: List"))
            self.headers = list(self.data.columns)
            print("Файл успешно загружен.")
        except Exception as e:
            print(f"Ошибка при загрузке файла: {e}")

    def filter_and_save_columns(self, columns_to_save: list):
        """
        Сохраняет значения из указанных столбцов в словарь.
        Можно передать функцию фильтрации.
        :param columns_to_save: список названий столбцов для сохранения.
        :param filter_func: функция для фильтрации строк (опционально).
        """
        if self.data is None:
            print("Данные не загружены.")
            return
        self.columns_save = columns_to_save
        # Сохранение данных по строкам
        self.filtered_data = {}
        for index, row in self.data.iterrows():
            self.filtered_data[index] = {col: row[col] for col in columns_to_save if col in row}

    def get_filtered_data(self):
        """Возвращает отфильтрованные данные."""
        return self.filtered_data

    def get_colum(self, get_colum: str, foc_mode: float = 0):

        # Проверка столбца
        if get_colum not in self.columns_save:
            messagebox.showerror("Ошибка", f"Название столбцов не совпадают.\n{get_colum} в {self.columns_save}")
            return None
        result = []

        # === 1. Подготовка: один раз до цикла ===
        # Получаем словарь данных один раз
        all_data_dict = self.get_dict_all_data() if foc_mode != 0 else None

        # Получаем имена колонок один раз
        done_col = self.config.getCzColumnName("Table of contents: Done") if foc_mode != 0 else None
        close_col = self.config.getCzColumnName("Table of contents: Close") if foc_mode != 0 else None
        dse_col = self.config.getCzColumnName("Table of contents: DCE") if foc_mode != 0 else None

        # Вспомогательная функция проверки пустоты
        def is_empty(val):
            return (val is None or
                    (isinstance(val, str) and val.strip() in ("", "nan", "n/a", "none", "-", "null")) or
                    (hasattr(pd, 'isna') and pd.isna(val)))

        # === 2. Основной цикл ===
        for key in self.filtered_data:
            # --- FOC MODE: фильтрация ---
            if foc_mode != 0:
                row_data = all_data_dict.get(key, {})
                if not (is_empty(row_data.get(done_col)) and
                        is_empty(row_data.get(close_col)) and
                        not is_empty(row_data.get(dse_col))):
                    continue  # не подходит под условие — пропускаем

            # --- Обработка значения ---
            value = self.filtered_data[key].get(get_colum)
            if value is None or (hasattr(pd, 'isna') and pd.isna(value)):
                continue

            # Преобразуем в строку и разбиваем
            value_str = str(value).strip()
            if not value_str:
                continue

            # Заменяем ", " на "|", разбиваем
            items = value_str.replace(", ", "|").split("|")

            for item in items:
                item = item.strip()
                if not item:
                    continue

                # Обрезка для сложных имён
                if ":" in item and "-" in item:
                    item = item[:len(item) // 2 + 1]

                if item not in result:
                    result.append(item)

        # === 3. Финальная сортировка ===
        result.sort(reverse=True)
        result.append("")
        return result

    # def get_colum(self, get_colum: str, foc_mode:float = 0):
    #     print( get_colum , self.columns_save)
    #     if get_colum in self.columns_save:
    #         result = []
    #         for i in self.filtered_data:
    #             valueResult = self.filtered_data[i][f"{get_colum}"]
    #             if not valueResult in result:
    #                 valueResult = str(valueResult).replace(", ","|")
    #                 valueResult = valueResult.split("|")
    #                 for i in valueResult:
    #                     if not i in result:
    #                         i = str(i)
    #                         if ":" in i and  "-" in i :
    #                             i = i[:len(i)//2+1]
    #                         if i != "":
    #                             result.append(str(i))
    #
    #
    #
    #         result.sort()
    #         result.reverse()
    #         return result
    #     else:
    #         messagebox.showerror("Ошибка", f"Название столбцов не совпадают.\n {get_colum} в {self.columns_save} ")
    #         return None


"""def get_sheet_names(self):
    # "Получает список всех листов в Excel файле."
    try:
        sheet_names = pd.ExcelFile(self.file_path).sheet_names
        print("Доступные листы:", sheet_names)
        return sheet_names
    except Exception as e:
        print(f"Ошибка при получении списка листов: {e}")
        return []"""
