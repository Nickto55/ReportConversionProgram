import argparse
import calendar
import ctypes
import os
import sys
import webbrowser
from datetime import datetime as dt
from tkinter import messagebox, filedialog

import plyer
import tkinter as tk
import static.config as Config
import scripts.excel_enter as excel_enter
import scripts.handlings.handling_json as handling_json

from handlings.program_jp import JpMain
from handlings.program_bam import BamMain
from handlings.program_cz import CzMain
from handlings.program_generalization import GeneProg
from handlings.program_for_tracking import TrackMain

from scripts.handlings.handling_log import logger, attempt_recover


def send_notification(title, message, settime=15, file_path=""):
    plyer.notification.notify(
        title=title, 
        message=message, 
        app_name="Good Morning (JP)", 
        timeout=settime, 
        app_icon=resource_path("static/icons/new_icon_report-conversion-program.ico")
    )


def updateInfoConfig(fileOrDir: int):
    root_select_file = tk.Tk()
    root_select_file.withdraw()
    if fileOrDir:
        file_path = filedialog.askopenfilename(title="Выберите файл")
    else:
        file_path = filedialog.askdirectory(title="Выберите папку")
    return file_path


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ReportConversion:
    """
    Основной класс бизнес-логики.
    Содержит всю логику обработки данных, конфигурации и управления процессами.
    Не содержит GUI-кода.
    """
    
    def __init__(self):
        self.start_no_gui_var = False
        self.start_button_var = False
        self.config = handling_json.JsonConfig()
        
        # Размеры окон
        self.distance_y_root = self.config.getConfigSizeYProgram()
        self.distance_x_root = self.config.getConfigSizeXProgram()
        
        # Переменные состояния
        self.modification = 1
        self.time_and_day_now = str(dt.now())
        self.ubroutine_Jp_var = True
        self.ubroutine_Cz_var = True
        self.dop_date_Cz_var = False
        self.ubroutine_Bam_var = True
        self.availability_of_required_data = True
        self.parent_label_column_jp_bool = False
        self.parent_label_column_cz_bool = False
        self.parent_label_column_BAM_bool = False
        
        # Расчет даты прошлого месяца
        current_year = int(str(dt.now())[:4])
        current_month = int(str(dt.now())[5:7])
        if current_month == 1:
            last_month_num = 12
            last_year = current_year - 1
        else:
            last_month_num = current_month - 1
            last_year = current_year
        self.count_day_last_mounth = calendar.monthrange(last_year, last_month_num)[1]
        self.last_mouns_last_day = f"{last_year}-{last_month_num}-{self.count_day_last_mounth}"
        
        # Парсер аргументов командной строки
        self.args = None
        self._parse_args()
    
    def _parse_args(self):
        parser = argparse.ArgumentParser()
        parser.add_argument("--console", action="store_true", help="Запустить с консолью")
        parser.add_argument("--cfile", action="store_true", help="Конфиг файла")
        parser.add_argument("--pstart", action="store_true", help="Моментальный старт")
        parser.add_argument("--nog", action="store_true", help="Без граф. оболочки")
        self.args = parser.parse_args()
    
    def setup_console(self):
        """Настраивает консольный вывод если нужно."""
        if self.args.console:
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            hStdOut = kernel32.GetStdHandle(-11)
            if hStdOut == -1 or hStdOut == 0:
                kernel32.AllocConsole()
                sys.stdout = open('CONOUT$', 'w')
                sys.stderr = open('CONOUT$', 'w')
    
    def should_open_config(self):
        """Проверяет, нужно ли открыть папку с конфигом."""
        return self.args.cfile
    
    def open_config_dir(self):
        """Открывает папку с конфигурацией."""
        CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".ReportConversionProgram")
        webbrowser.open(CONFIG_DIR)
    
    def should_start_immediately(self):
        """Проверяет, нужен ли моментальный старт."""
        return self.args.pstart or self.args.nog
    
    def is_no_gui_mode(self):
        """Проверяет, запущен ли режим без GUI."""
        return self.args.nog
    
    def update_last_mouns(self, progress_callback=None):
        """
        Обновляет данные за прошлый месяц.
        
        :param progress_callback: функция обратного вызова для обновления прогресса (value)
        """
        total_steps = int(self.ubroutine_Jp_var) + int(self.ubroutine_Cz_var) + int(self.ubroutine_Bam_var)
        step_value = 100 // total_steps if total_steps > 0 else 100
        
        if self.ubroutine_Jp_var:
            jp_prog = JpMain(mask_date=self.last_mouns_last_day)
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Jp")
            excelPr.write_to_sheet(jp_prog.main(), "ЖП")
            if progress_callback:
                progress_callback(step_value)
        
        if self.ubroutine_Cz_var:
            if self.dop_date_Cz_var:
                cz_prog = CzMain(None, dop_date=1, mask_date=self.last_mouns_last_day)
            else:
                cz_prog = CzMain(None, mask_date=self.last_mouns_last_day)
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Cz")
            excelPr.write_to_sheet(cz_prog.main(), "СЗ")
            if progress_callback:
                progress_callback(step_value)
        
        if self.ubroutine_Bam_var:
            bam_prog = BamMain(mask_date=self.last_mouns_last_day)
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="BAM")
            excelPr.write_to_sheet(bam_prog.main(), "Бам по УП")
            if progress_callback:
                progress_callback(step_value)
        
        if progress_callback:
            progress_callback(100 - (progress_callback.current if hasattr(progress_callback, 'current') else 0))
        
        if self.ubroutine_Bam_var and self.ubroutine_Cz_var and self.ubroutine_Jp_var:
            ge_prog = GeneProg(mask_date=self.last_mouns_last_day)
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Ge")
            excelPr.write_to_sheet(ge_prog.main(), "Общая информация")

        send_notification("Программа завершена", "Программа завершена, проверте файл", 16)
        
        return True
    
    def start_processing(self, progress_callback=None, completion_callback=None):
        """
        Запускает основную обработку данных.
        
        :param progress_callback: функция обратного вызова для обновления прогресса
        :param completion_callback: функция обратного вызова по завершении
        """
        # Проверка на обновление прошлого месяца
        if int(self.time_and_day_now[8:10]) <= 5:
            # Это должно быть обработано GUI (messagebox), 
            # но логика остается здесь для no-gui режима
            pass  # GUI обработает это через callback
        
        total_steps = int(self.ubroutine_Jp_var) + int(self.ubroutine_Cz_var) + int(self.ubroutine_Bam_var)
        step_value = 100 // total_steps if total_steps > 0 else 100
        
        if self.ubroutine_Jp_var:
            jp_prog = JpMain()
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Jp")
            excelPr.write_to_sheet(jp_prog.main(), "ЖП")
            if progress_callback:
                progress_callback(step_value)
        
        if self.ubroutine_Cz_var:
            if self.dop_date_Cz_var:
                cz_prog = CzMain(None, dop_date=1)
            else:
                cz_prog = CzMain(None)
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Cz")
            excelPr.write_to_sheet(cz_prog.main(), "СЗ")
            if progress_callback:
                progress_callback(step_value)
        
        if self.ubroutine_Bam_var:
            bam_prog = BamMain()
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="BAM")
            excelPr.write_to_sheet(bam_prog.main(), "Бам по УП")
            if progress_callback:
                progress_callback(step_value)
        
        if progress_callback:
            progress_callback(100 - (getattr(progress_callback, 'current', 0)))
        
        if self.ubroutine_Bam_var and self.ubroutine_Cz_var and self.ubroutine_Jp_var:
            ge_prog = GeneProg()
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Ge")
            excelPr.write_to_sheet(ge_prog.main(), "Общая информация")

        print('======================================================')
        print(self.ubroutine_Bam_var)
        if self.ubroutine_Bam_var:
            tr_prog = TrackMain()
            excelPr = excel_enter.ExcelWriter(self.config.getJPPathFile_output(), min_prog='Tr')
            excelPr.write_to_sheet(tr_prog.main(), "Отслеживание")

        send_notification("Программа завершена", "Программа завершена, проверте файл", 16)

        if completion_callback:
            completion_callback()
        
        if self.start_no_gui_var:
            sys.exit()
        
        return True
    
    def checking_data_inputs(self):
        """
        Проверяет наличие всех необходимых данных в конфиге.
        Возвращает словарь с недостающими данными.
        """
        recovery_data_inputs = {"Program:": [], "ЖП": [], "СЗ": [], "BAM": [], "Ge": []}
        data = {}
        
        def _load_config_data():
            return self.config.getData()
        
        def _reload_config():
            try:
                self.config = handling_json.JsonConfig()
                logger.info("JsonConfig reinitialized during checking_data_inputs recovery")
            except Exception:
                logger.exception("Failed to reinitialize JsonConfig during recovery")
        
        try:
            data = attempt_recover(_load_config_data, recover_funcs=[_reload_config], attempts=2)
        except Exception:
            logger.exception("checking_data_inputs: cannot load config data; proceeding with empty data")
            data = {}
        
        # Проверка структуры конфига
        try:
            lose_key_list = Config.configProgram.keys() - data.keys()
        except Exception:
            logger.exception("Failed to determine missing top-level config keys")
            lose_key_list = Config.configProgram.keys()
        
        if len(lose_key_list) != 0:
            logger.warning("Config structure mismatch, missing keys: %s", lose_key_list)
            return {"error": "structure_mismatch", "missing_keys": list(lose_key_list)}
        
        # Проверка локальных ключей
        for local_key in list(data.keys()):
            try:
                local_lose_key_list = Config.configProgram[local_key].keys() - data[local_key].keys()
                if len(local_lose_key_list) != 0:
                    return {"error": "local_keys_mismatch", "section": local_key, 
                            "missing_keys": list(local_lose_key_list)}
            except Exception:
                logger.exception("Error while checking local config keys for %s", local_key)
        
        # Проверка значений
        for key_program in list(data.keys()):
            try:
                for key_value in data[key_program].keys():
                    value_config = data[key_program].get(key_value, "")
                    if key_value in Config.keys_for_unnecessary_data.get(key_program, []):
                        if value_config == "" or value_config == []:
                            if key_program == "Program:":
                                try:
                                    self.config.recoveryLoseKeyAndValue(
                                        key_program, key_value,
                                        Config.configProgram[key_program].get(key_value, "")
                                    )
                                except Exception:
                                    logger.exception("Failed to recover Program: %s", key_value)
                            else:
                                self.availability_of_required_data = False
                                recovery_data_inputs.setdefault(key_program, []).append(key_value)
            except Exception:
                logger.exception("Error while checking values in config section %s", key_program)
        
        # Автоматическое восстановление
        if any(recovery_data_inputs.values()):
            logger.info("Attempting automatic recovery for missing config values: %s", recovery_data_inputs)
            return {"error": "missing_values", "data": recovery_data_inputs}
        
        return {"error": None}
    
    def recovery_data_inputs(self, dic_recovery: dict):
        """
        Восстанавливает недостающие данные конфигурации.
        Возвращает True если успешно, False если нужен GUI ввод.
        """
        if not self.availability_of_required_data:
            return {"need_gui": True, "data": dic_recovery}
        return {"need_gui": False}
    
    # === Методы для работы с настройками ===
    
    def get_jp_settings(self):
        """Возвращает текущие настройки ЖП."""
        return {
            "path": self.config.getJPPathFile_input(),
            "name": self.config.getJPNameFile_input(),
            "list_date": self.config.getJPColumnName("Table of contents: List_date"),
            "columns": {
                "date": self.config.getJPColumnName("Table of contents: Date"),
                "date_removed": self.config.getJPColumnName("Table of contents: Date removed"),
                "dse": self.config.getJPColumnName("Table of contents: DCE"),
                "accepted_name": self.config.getJPColumnName("Table of contents: Question accepted Full name"),
                "removed_name": self.config.getJPColumnName("Table of contents: Question removed Full name"),
                "number": self.config.getJPColumnName("Table of contents: Namber"),
                "name": self.config.getJPColumnName("Table of contents: Name"),
                "translation": self.config.getJPColumnName("Table of contents: Translation"),
            }
        }
    
    def set_jp_path(self, path: str):
        """Устанавливает путь к файлу ЖП."""
        self.config.setJPFilePathAndName_input(path)
    
    def set_jp_settings(self, path: str, list_date: str, columns: dict):
        """Сохраняет настройки ЖП."""
        self.config.setJPFilePathAndName_input(path)
        self.config.setJPColumnName("Table of contents: List_date", list_date)
        for key, value in columns.items():
            self.config.setJPColumnName(key, value)
    
    def reset_jp_columns(self):
        """Сбрасывает колонки ЖП на значения по умолчанию."""
        defaults = {
            "Table of contents: Date": Config.configProgram["ЖП"].get("Table of contents: Date", ""),
            "Table of contents: Date removed": Config.configProgram["ЖП"].get("Table of contents: Date removed", ""),
            "Table of contents: DCE": Config.configProgram["ЖП"].get("Table of contents: DCE", ""),
            "Table of contents: Question accepted Full name": Config.configProgram["ЖП"].get(
                "Table of contents: Question accepted Full name", ""),
            "Table of contents: Question removed Full name": Config.configProgram["ЖП"].get(
                "Table of contents: Question removed Full name", ""),
            "Table of contents: Namber": Config.configProgram["ЖП"].get("Table of contents: Namber", ""),
            "Table of contents: Name": Config.configProgram["ЖП"].get("Table of contents: Name", ""),
            "Table of contents: Translation": Config.configProgram["ЖП"].get("Table of contents: Translation", ""),
        }
        for key, value in defaults.items():
            self.config.setJPColumnName(key, value)
        return defaults
    
    def get_bam_settings(self):
        """Возвращает текущие настройки BAM."""
        return {
            "path": self.config.getBAMPathFile_input(),
            "name": self.config.getBAMNameFile_input(),
            "list_date": self.config.getBAMColumnName("Table of contents: List_date"),
            "columns": {
                "date": self.config.getBAMColumnName("Table of contents: Date"),
                "listes_excel": self.config.getBAMColumnName("Table of contents: listes_excel"),
            }
        }
    
    def set_bam_path(self, path: str):
        """Устанавливает путь к файлу BAM."""
        self.config.setBAMFilePathAndName_input(path)
    
    def set_bam_settings(self, path: str, list_date: str, columns: dict):
        """Сохраняет настройки BAM."""
        self.config.setBAMFilePathAndName_input(path)
        self.config.setBAMColumnName("Table of contents: List_date", list_date)
        for key, value in columns.items():
            self.config.setBAMColumnName(key, value)
    
    def reset_bam_columns(self):
        """Сбрасывает колонки BAM на значения по умолчанию."""
        defaults = {
            "Table of contents: Date": Config.configProgram["BAM"].get("Table of contents: Date", ""),
            "Table of contents: listes_excel": Config.configProgram["BAM"].get("Table of contents: listes_excel", ""),
        }
        for key, value in defaults.items():
            self.config.setBAMColumnName(key, value)
        return defaults
    
    def get_cz_settings(self):
        """Возвращает текущие настройки CZ."""
        return {
            "path": self.config.getCzPathFile_input(),
            "name": self.config.getCzNameFile_input(),
            "list_date": self.config.getCzColumnName("Table of contents: List_date"),
            "columns": {
                "list": self.config.getCzColumnName("Table of contents: List"),
                "accepted": self.config.getCzColumnName("Table of contents: Accepted"),
                "dse": self.config.getCzColumnName("Table of contents: DCE"),
                "rc": self.config.getCzColumnName("Table of contents: RC"),
                "close": self.config.getCzColumnName("Table of contents: Close"),
                "done": self.config.getCzColumnName("Table of contents: Done"),
                "date_writer": self.config.getCzColumnName("Table of contents: Date writer"),
            },
            "rc_params": {
                "foc": self.config.getCzColumnName("RC search parameters: Rc foc value"),
                "toc": self.config.getCzColumnName("RC search parameters: Rc toc value"),
                "poc": self.config.getCzColumnName("RC search parameters: Rc poc value"),
            }
        }
    
    def set_cz_path(self, path: str):
        """Устанавливает путь к файлу CZ."""
        self.config.setCzFilePathAndName_input(path)
    
    def set_cz_settings(self, path: str, list_date: str, columns: dict, rc_params: dict):
        """Сохраняет настройки CZ."""
        self.config.setCzFilePathAndName_input(path)
        self.config.setCzColumnName("Table of contents: List_date", list_date)
        for key, value in columns.items():
            self.config.setCzColumnName(key, value)
        for key, value in rc_params.items():
            self.config.setCzColumnName(key, value)
    
    def reset_cz_columns(self):
        """Сбрасывает колонки CZ на значения по умолчанию."""
        defaults = {
            "Table of contents: List": Config.configProgram["СЗ"].get("Table of contents: List", ""),
            "Table of contents: Accepted": Config.configProgram["СЗ"].get("Table of contents: Accepted", ""),
            "Table of contents: DCE": Config.configProgram["СЗ"].get("Table of contents: DCE", ""),
            "Table of contents: RC": Config.configProgram["СЗ"].get("Table of contents: RC", ""),
            "Table of contents: Close": Config.configProgram["СЗ"].get("Table of contents: Close", ""),
            "Table of contents: Done": Config.configProgram["СЗ"].get("Table of contents: Done", ""),
            "Table of contents: Date writer": Config.configProgram["СЗ"].get("Table of contents: Date writer", ""),
        }
        for key, value in defaults.items():
            self.config.setCzColumnName(key, value)
        return defaults
    
    def reset_cz_rc_params(self):
        """Сбрасывает параметры РЦ CZ на значения по умолчанию."""
        defaults = {
            "RC search parameters: Rc foc value": Config.configProgram["СЗ"].get("RC search parameters: Rc foc value", ""),
            "RC search parameters: Rc toc value": Config.configProgram["СЗ"].get("RC search parameters: Rc toc value", ""),
            "RC search parameters: Rc poc value": Config.configProgram["СЗ"].get("RC search parameters: Rc poc value", ""),
        }
        for key, value in defaults.items():
            self.config.setCzColumnName(key, value)
        return defaults
    
    # === Методы для работы с размерами окна ===
    
    def get_window_size(self):
        """Возвращает текущий размер окна."""
        return {"x": self.distance_x_root, "y": self.distance_y_root}
    
    def set_window_size(self, x: int, y: int):
        """Устанавливает размер окна."""
        self.distance_x_root = x
        self.distance_y_root = y
        self.config.setConfigSizeXProgram(x)
        self.config.setConfigSizeYProgram(y)
    
    def modify_window_size(self, axis: str, increment: int):
        """Изменяет размер окна."""
        if axis == "x":
            self.distance_x_root += increment
            self.config.setConfigSizeXProgram(self.distance_x_root)
            return self.distance_x_root
        elif axis == "y":
            self.distance_y_root += increment
            self.config.setConfigSizeYProgram(self.distance_y_root)
            return self.distance_y_root
    
    def reset_window_size(self):
        """Сбрасывает размер окна на значения по умолчанию."""
        self.distance_x_root = int(Config.configProgram["Program:"].get("Size by X", ""))
        self.distance_y_root = int(Config.configProgram["Program:"].get("Size by Y", ""))
        self.config.setConfigSizeXProgram(self.distance_x_root)
        self.config.setConfigSizeYProgram(self.distance_y_root)
        return {"x": self.distance_x_root, "y": self.distance_y_root}
    
    def get_modification_step(self):
        """Возвращает текущий шаг модификации."""
        return self.modification
    
    def set_modification_step(self, value: int):
        """Устанавливает шаг модификации."""
        self.modification = value
    
    # === Методы для получения информации ===
    
    def get_current_date(self):
        """Возвращает текущую дату."""
        return str(dt.now())[:10]
    
    def get_file_modification_date(self):
        """Возвращает дату последнего изменения файла."""
        try:
            modification_time = os.path.getmtime(f"{os.getcwd()}/ReportConversionProgram.xlsx")
            last_date_start = str(dt.fromtimestamp(modification_time))
            return {"date": last_date_start[:10], "time": last_date_start[11:16]}
        except:
            return {"date": "-", "time": ""}
    
    def get_program_title(self):
        """Возвращает заголовок программы."""
        return self.config.getConfigNameAssemblyProgram()
    
    def get_last_month_info(self):
        """Возвращает информацию о прошлом месяце."""
        return {
            "date": self.last_mouns_last_day,
            "day_count": self.count_day_last_mounth
        }
    
    # === Геттеры/сеттеры для флагов подпрограмм ===
    
    def set_jp_enabled(self, value: bool):
        self.ubroutine_Jp_var = value
    
    def set_cz_enabled(self, value: bool):
        self.ubroutine_Cz_var = value
    
    def set_bam_enabled(self, value: bool):
        self.ubroutine_Bam_var = value
    
    def set_cz_dop_date(self, value: bool):
        self.dop_date_Cz_var = value
    
    def is_jp_enabled(self):
        return self.ubroutine_Jp_var
    
    def is_cz_enabled(self):
        return self.ubroutine_Cz_var
    
    def is_bam_enabled(self):
        return self.ubroutine_Bam_var
    
    def is_cz_dop_date(self):
        return self.dop_date_Cz_var


if __name__ == "__main__":
    app = ReportConversion()
    
    if app.should_open_config():
        app.open_config_dir()
        sys.exit()
    
    if app.is_no_gui_mode():
        app.start_processing()
    else:
        # Импортируем и запускаем GUI
        from Body import MainGUI
        import tkinter as tk
        root = tk.Tk()
        gui = MainGUI(root, app)
        root.mainloop()