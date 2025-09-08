from tkinter import messagebox

import pandas as pd

import JsonWork


class SearchBam:
    def __init__(self,sheet_name):
        self.columns_save = []
        self.config = JsonWork.JsonConfig()
        self.file_path = self.config.getBAMPathFile_input()
        self.sheet_name = sheet_name
        self.data = None
        self.load_excel()
        self.filtered_data = None
        self.filter_and_save_columns(self.config.getBAMColumnName("Table of contents: Date"))

    def return_data(self):
        return self.data

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
        """Загружает данные из Excel файла с указанного листа."""
        try:
            # Читаем Excel, используя sheet_name (может быть str, int, или None)
            self.data = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

            # Если sheet_name не задан (None), по умолчанию берём первый лист
            if isinstance(self.data, dict):
                sheet_names = list(self.data.keys())
                if sheet_names:
                    first_sheet = sheet_names[0]
                    print(f"Лист не указан. Загружаем первый лист: {first_sheet}")
                    self.data = self.data[first_sheet]
                else:
                    raise ValueError("Файл Excel не содержит листов.")
            else:
                # Успешно загружен один лист
                print(f"Лист успешно загружен: {self.sheet_name if self.sheet_name else 'первый лист'}")

        except Exception as e:
            print(f"Ошибка при загрузке файла или листа: {e}")
            self.data = None

    def get_colum(self, get_colum: str, foc_mode: float = 0):
        # Проверка столбца
        if get_colum not in self.columns_save:
            messagebox.showerror("Ошибка", f"Название столбцов не совпадают.\n{get_colum} в {self.columns_save}")
            return None
        result = []

        # === 1. Подготовка: один раз до цикла ===
        all_data_dict = self.get_dict_all_data() if foc_mode != 0 else None
        done_col = self.config.getBAMColumnName("Table of contents: Date") if foc_mode != 0 else None

        def is_empty(val):
            return (val is None or
                    (isinstance(val, str) and val.strip() in ("", "nan", "n/a", "none", "-", "null")) or
                    (hasattr(pd, 'isna') and pd.isna(val)))

        # === 2. Основной цикл ===
        for key in self.filtered_data:
            if foc_mode != 0:
                row_data = all_data_dict.get(key, {})
                if not (is_empty(row_data.get(done_col))):
                    continue  # не подходит под условие — пропускаем

            value = self.filtered_data[key].get(get_colum)
            if value is None or (hasattr(pd, 'isna') and pd.isna(value)):
                continue

            value_str = str(value).strip()
            if not value_str:
                continue

            items = value_str.replace(", ", "|").split("|")

            for item in items:
                item = item.strip()
                if not item:
                    continue

                if ":" in item and "-" in item:
                    item = item[:len(item) // 2 + 1]

                if item not in result:
                    result.append(item)

        # === 3. Финальная сортировка ===
        result.sort(reverse=True)
        result.append("")
        return result

    def filter_and_save_columns(self, columns_to_save: list):
        """
        Сохраняет значения из указанных столбцов в словарь.
        :param columns_to_save: список названий столбцов для сохранения.
        """
        if self.data is None:
            print("Данные не загружены.")
            return

        # Убедимся, что columns_to_save — это список
        if isinstance(columns_to_save, str):
            columns_to_save = [columns_to_save]
        elif not isinstance(columns_to_save, list):
            columns_to_save = list(columns_to_save)

        self.columns_save = columns_to_save

        self.filtered_data = {}
        for index, row in self.data.iterrows():
            self.filtered_data[index] = {col: row[col] for col in columns_to_save if col in row}


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

class SearchGe:
    def __init__(self):
        self.data_jp = None
        self.data_cz = None
        self.data_bam = None

        self.data_ge = None

        self.config = JsonWork.JsonConfig()
        self.file_path = self.config.getJPColumnName("Path for output excel")
    def sheet_name_list(self):
        sheet_name_list = self.config.getSheetNameList()
        for sheet_name in sheet_name_list:
            data_and_name = self.load_excel(sheet_name=sheet_name)
            if sheet_name == sheet_name_list[0]:self.data_jp = data_and_name
            elif sheet_name == sheet_name_list[1]: self.data_cz = data_and_name
            elif sheet_name == sheet_name_list[2]: self.data_bam = data_and_name
            # elif data_and_name[0] == sheet_name_list[3]: self.data_ge = data_and_name[1]

        result_data_jp = {}
        for index, row in self.data_jp.iterrows():
            result_data_jp[index] = row.to_dict()

        result_data_cz = {}
        for index, row in self.data_cz.iterrows():
            result_data_cz[index] = row.to_dict()

        result_data_bam = {}
        for index, row in self.data_bam.iterrows():
            result_data_bam[index] = row.to_dict()


        return result_data_jp, result_data_cz, result_data_bam


    def load_excel(self, sheet_name: str = None):
        """Загружает данные из Excel файла с указанного листа."""
        try:
            # Читаем Excel, используя sheet_name (может быть str, int, или None)
            data = pd.read_excel(self.file_path, sheet_name=sheet_name)



        except Exception as e:
            print(f"Ошибка при загрузке файла или листа: {e}")
            data = None


        return data




"""def get_sheet_names(self):
    # "Получает список всех листов в Excel файле."
    try:
        sheet_names = pd.ExcelFile(self.file_path).sheet_names
        print("Доступные листы:", sheet_names)
        return sheet_names
    except Exception as e:
        print(f"Ошибка при получении списка листов: {e}")
        return []"""
