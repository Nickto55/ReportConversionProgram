import json
import os
import tkinter as tk
from tkinter import messagebox, filedialog

import static.config as Config


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
            if not os.path.exists((self.getCzPathFile_input())):
                self.setCzFilePathAndName()
            if not os.path.exists((self.getBAMPathFile_input())):
                self.setBAMFilePathAndName()
        except Exception as e:
            messagebox.showerror("Ошибка", "Произошла ошибка при проверке наличия файлов")
            print("Произошла ошибка при проверке наличия файлов",e)

        try:
            if self.getConfigVersionConfig() < float(
                    Config.configProgram["Program:"].get("Config version", "").replace(":", "")):
                messagebox.showwarning("Внимание",
                                       "Версия файла конфига ниже чем версия программы, конфиг может быть не совместим")
                return
        except Exception as e:
            messagebox.showwarning("Внимание", "Версия файла конфига ниже чем версия программы, конфиг не совместим")
            print("Версия файла конфига ниже чем версия программы, конфиг не совместим",e)
            return

    def save(self):
        """Сохраняет текущие данные в файл."""
        with open(self.file_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=4,ensure_ascii=False)
        self.load()

    def _ensure_file_exists(self):
        """Создаёт файл и структуру данных, если их нет."""
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        if not os.path.exists(self.file_path):
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4,ensure_ascii=False)

    def load(self):
        """Загружает данные из файла."""
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            print("JsonSave файл пуст")

    def getData(self):
        return self.data

    def recoveryLoseKeyAndValue(self, dictionary: str, key_for_dictionary: str, value_for_key: str):
        print("JsonWork, config, recoveryLoseKeyAndValue", dictionary, key_for_dictionary, value_for_key)
        self.data[dictionary][key_for_dictionary] = value_for_key
        self.save()

    """
    
    Конфиг основной программы
    
    """

    def setConfigMainProgram(self):

        if str(self.data["Program:"].get("Assembly name", "")) == "":
            self.setConfigNameAssemblyProgram()
        self.save()

    def setConfigNameAssemblyProgram(self):
        self.data["Program:"]["Assembly name"] = "ReportConversionProgram"
        self.save()

    def getConfigNameAssemblyProgram(self):
        return str(self.data["Program:"].get("Assembly name", ""))

    def getConfigsimbaProgram(self):
        return int(self.data["Program:"].get("Cimb plus", ""))

    def getConfigSizeXProgram(self):
        return int(self.data["Program:"].get("Size by X", ""))

    def getConfigSizeYProgram(self):
        return int(self.data["Program:"].get("Size by Y", ""))

    def getConfigVersionProgram(self):
        return float(self.data["Program:"].get("Version", "").replace(":", ""))

    def getConfigVersionConfig(self):
        return float(self.data["Program:"].get("Config version", "").replace(":", ""))

    def setConfigSizeXProgram(self, modification):
        self.data["Program:"]["Size by X"] = modification
        self.save()

    def setConfigSizeYProgram(self, modification):
        self.data["Program:"]["Size by Y"] = modification
        self.save()

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

    def setJPFilePathAndName_input(self, path):
        full_path = path
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

    """
    
    Для CZ
    
    """

    def setCzFilePathAndName(self):
        def updatePath():
            messagebox.showwarning("Внимание!",
                                   f"Файла ({self.getCzNameFile_input()}), для СЗ не обнаружено\nОткройте его заново")
            return updateInfoConfig(1)

        full_path = updatePath()
        self.data["СЗ"]["Path for input excel"] = full_path
        self.data["СЗ"]["Input excel file"] = os.path.basename(full_path)

        if str(self.data["СЗ"].get("Path for output excel", "")) == "":
            self.setCzPathFile_output()
        self.save()

    def setCzFilePathAndName_input(self, path):
        full_path = path
        self.data["СЗ"]["Path for input excel"] = full_path
        self.data["СЗ"]["Input excel file"] = os.path.basename(full_path)

        if str(self.data["СЗ"].get("Path for output excel", "")) == "":
            self.setJPPathFile_output()
        self.save()

    def setCzColumnName(self, columnName: str, columnValue):
        """
        Установка новых иминований столбцов, на вход ключ и знначение
        :param columnName: ключ
        :param columnValue: значение
        :return:
        """
        self.data["СЗ"][columnName] = columnValue
        self.save()

    def setCzNameFile_output(self, name: str):
        """
        Установка нового имени
        :param name: имя файла
        :return:
        """
        self.data["СЗ"]["Name for output excel"] = name
        self.save()

    def setCzPathFile_output(self):
        """
        Установка нового пути
        """
        self.data["СЗ"]["Path for output excel"] = f"{os.getcwd()}/{self.getCzNameFile_output()}"
        self.save()

    def getCzNameFile_input(self):
        data = self.data["СЗ"]
        return str(data.get("Input excel file", ""))

    def getCzNameFile_output(self):
        data = self.data["СЗ"]
        return str(data.get("Name for output excel", ""))

    def getCzPathFile_input(self):
        data = self.data["СЗ"]
        return str(data.get("Path for input excel", ""))

    def getCzPathFile_output(self):
        data = self.data["СЗ"]
        return str(data.get("Path for output excel", ""))

    def getCzColumnName(self, columnName: str, neeed_list: int = 0):
        data = self.data["СЗ"]
        if neeed_list:
            return data.get(columnName, "")
        return str(data.get(columnName, ""))

    """
    
    Для BAM
    
    """

    def setBAMFilePathAndName_input(self, path):
        full_path = path
        self.data["BAM"]["Path for input excel"] = full_path
        self.data["BAM"]["Name input excel file"] = os.path.basename(full_path)
        print(os.path.basename(full_path), "setBAMFilePathAndName_input")

        if str(self.data["BAM"].get("Path for output excel", "")) == "":
            self.setBAMPathFile_output()
        self.save()

    def setBAMFilePathAndName(self):
        def updatePath():
            messagebox.showwarning("Внимание!",
                                   f"Файла ({self.getBAMNameFile_input()}), для BAM не обнаружено\nОткройте его заново")
            return updateInfoConfig(1)

        full_path = updatePath()
        self.data["BAM"]["Path for input excel"] = full_path
        self.data["BAM"]["Name input excel file"] = os.path.basename(full_path)

        if str(self.data["BAM"].get("Path for output excel", "")) == "":
            self.setBAMPathFile_output()
        self.save()

    def setBAMColumnName(self, columnName: str, columnValue: str):
        """
        Установка новых иминований столбцов, на вход ключ и знначение
        :param columnName: ключ
        :param columnValue: значение
        :return:
        """
        self.data["BAM"][columnName] = columnValue
        self.save()

    def setBAMNameFile_output(self, name: str):
        """
        Установка нового имени
        :param name: имя файла
        :return:
        """
        self.data["BAM"]["Name output excel"] = name
        self.save()

    def setBAMPathFile_output(self):
        """
        Установка нового пути
        """
        self.data["BAM"]["Path for output excel"] = f"{os.getcwd()}/{self.getBAMNameFile_output()}"
        self.save()

    def getBAMNameFile_input(self):
        data = self.data["BAM"]
        return str(data.get("Name input excel file", ""))

    def getBAMNameFile_output(self):
        data = self.data["BAM"]
        return str(data.get("Name output excel", ""))

    def getBAMPathFile_input(self):
        data = self.data["BAM"]
        return str(data.get("Path for input excel", ""))

    def getBAMPathFile_output(self):
        data = self.data["BAM"]
        return str(data.get("Path for output excel", ""))

    def getBAMColumnName(self, columnName: str, intOrlist: int = 0):
        data = self.data["BAM"]
        if intOrlist:
            return data.get(columnName, "")
        return str(data.get(columnName, ""))

    """
    
    Для Ge
    
    """

    def getSheetNameList(self):
        data = self.data["Ge"]
        return data.get("Sheet name list", "")

    def get_ge_last_days(self):
        data = self.data["Ge"]
        return data.get("last 3 day in mounth", "")

    def set_ge_last_days(self, Mounth, data):


        self.data["Ge"]["last 3 day in mounth"][Mounth] = data
        self.save()


if __name__ == "__main__":
    run = JsonConfig()
