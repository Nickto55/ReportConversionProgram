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
