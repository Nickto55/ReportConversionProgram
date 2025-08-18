import json
import os
import tkinter as tk
from tkinter import messagebox, filedialog

import Config


def updateInfoConfig(fileOrDir: int):
    root_select_file = tk.Tk()
    root_select_file.withdraw()  # скрываем главное окно
    if fileOrDir:
        file_path = filedialog.askopenfilename(title="Выберите файл")
    else:
        file_path = filedialog.askdirectory(title="Выберите папку")
    return file_path


class JsonConfig:
    def __init__(self):
        self.CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".ReportConversionProgram")
        self.file_path = os.path.join(self.CONFIG_DIR, "ConfigFile.json")
        self.data = Config.configProgram

        self._ensure_file_exists()
        self.load()

        try:
            if not os.path.exists(self.getJPPathFile_input()):
                self.setJPFilePathAndName()
        except:
            messagebox.showerror("Ошибка", "Произошла ошибка при проверке наличия файлов")

    def save(self):
        """Сохраняет текущие данные в файл."""
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=4)
        self.load()

    def _ensure_file_exists(self):
        """Создаёт файл и структуру данных, если их нет."""
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        if not os.path.exists(self.file_path):
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4)

    def load(self):
        """Загружает данные из файла."""
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            print("JsonSave файл пуст")

    """
    
    Для JP
    
    """

    def setJPFilePathAndName(self):
        def updatePath():
            messagebox.showwarning("Внимание!",
                                   f"Файла ({self.getJPNameFile_input()}), для ЖП не обнаружено\nОткройте его заново")
            return updateInfoConfig(1)

        full_path = updatePath()
        self.data["ЖП"]["Path for input excel"] = full_path
        self.data["ЖП"]["Input excel file"] = os.path.basename(full_path)

        if str(self.data["ЖП"].get("Path for output excel", "")) == "":
            self.setJPPathFile_output()
        self.save()

    def setJPColumnName(self, columnName: str, columnValue: str):
        """
        Установка новых иминований столбцов, на вход ключ и знначение
        :param columnName: ключ
        :param columnValue: значение
        :return:
        """
        self.data["ЖП"][columnName] = columnValue
        self.save()
    def setJPNameFile_output(self, name: str):
        """
        Установка нового имени
        :param name: имя файла
        :return:
        """
        self.data["ЖП"]["Name for output excel"] = name
        self.save()

    def setJPPathFile_output(self):
        """
        Установка нового пути
        """
        print(f"{os.getcwd()}/{self.getJPNameFile_output()}")
        self.data["ЖП"]["Path for output excel"] = f"{os.getcwd()}/{self.getJPNameFile_output()}"
        self.save()

    def getJPNameFile_input(self):
        data = self.data["ЖП"]
        return str(data.get("Input excel file", ""))

    def getJPNameFile_output(self):
        data = self.data["ЖП"]
        return str(data.get("Name for output excel", ""))

    def getJPPathFile_input(self):
        data = self.data["ЖП"]
        return str(data.get("Path for input excel", ""))
    def getJPPathFile_output(self):
        data = self.data["ЖП"]
        return str(data.get("Path for output excel", ""))

    def getJPColumnName(self, columnName: str):
        data = self.data["ЖП"]
        return str(data.get(columnName, ""))


if __name__ == "__main__":
    run = JsonConfig()
    # print("{",run.getJPPathFile_output(),"}")
    # print("{",run.getJPPathFile_input(),"}")
