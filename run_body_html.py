import os
import random
import sys
import webview
import threading
import time
from pathlib import Path

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent

STATIC_DIR = BASE_DIR / "static"
ICON_PATH = BASE_DIR / "static" / "icons" / "new_icon_report-conversion-program.ico"

from report_conversion import ReportConversion, resource_path, updateInfoConfig


class Api:
    def __init__(self):
        self.app = ReportConversion()
        self.app.ubroutine_Jp_var = True
        self.app.ubroutine_Cz_var = True
        self.app.ubroutine_Bam_var = True

    def select_folder(self):
        try:
            path = updateInfoConfig(fileOrDir=0)
            return path if path else ""
        except Exception as e:
            print(f"Error selecting folder: {e}")
            return ""

    def select_file(self):
        try:
            path = updateInfoConfig(fileOrDir=1)
            return path if path else ""
        except Exception as e:
            print(f"Error selecting file: {e}")
            return ""

    def load_settings(self):
        try:
            jp = self.app.get_jp_settings()
            bam = self.app.get_bam_settings()
            cz = self.app.get_cz_settings()
            size = self.app.get_window_size()
            file_info = self.app.get_file_modification_date()

            return {
                "programs": {
                    "jp": self.app.is_jp_enabled(),
                    "cz": self.app.is_cz_enabled(),
                    "bam": self.app.is_bam_enabled()
                },
                "jp": {
                    "path": jp.get("path", ""),
                    "name": jp.get("name", ""),
                    "days": jp.get("list_date", "1"),
                    "columns": jp.get("columns", {})
                },
                "cz": {
                    "path": cz.get("path", ""),
                    "name": cz.get("name", ""),
                    "days": cz.get("list_date", "1"),
                    "columns": cz.get("columns", {}),
                    "rc": cz.get("rc_params", {}),
                    "dop_sheets": self.app.is_cz_dop_date()
                },
                "bam": {
                    "path": bam.get("path", ""),
                    "name": bam.get("name", ""),
                    "days": bam.get("list_date", "1"),
                    "columns": bam.get("columns", {})
                },
                "window_size": size,
                "file_date": file_info,
                "title": self.app.get_program_title()
            }
        except Exception as e:
            print(f"Error loading settings: {e}")
            return {}

    def save_jp_settings(self, path, days, columns):
        try:
            self.app.set_jp_settings(path, days, columns)
            return {
                "success": True,
                "path": self.app.config.getJPPathFile_input(),
                "name": self.app.config.getJPNameFile_input(),
                "days": self.app.config.getJPColumnName("Table of contents: List_date")
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_cz_settings(self, path, days, columns, rc_params, dop_sheets):
        try:
            self.app.set_cz_dop_date(dop_sheets)
            self.app.set_cz_settings(path, days, columns, rc_params)
            return {
                "success": True,
                "path": self.app.config.getCzPathFile_input(),
                "name": self.app.config.getCzNameFile_input(),
                "days": self.app.config.getCzColumnName("Table of contents: List_date")
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_bam_settings(self, path, days, columns):
        try:
            self.app.set_bam_settings(path, days, columns)
            return {
                "success": True,
                "path": self.app.config.getBAMPathFile_input(),
                "name": self.app.config.getBAMNameFile_input(),
                "days": self.app.config.getBAMColumnName("Table of contents: List_date")
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def reset_jp_columns(self):
        return self.app.reset_jp_columns()

    def reset_cz_columns(self):
        return self.app.reset_cz_columns()

    def reset_bam_columns(self):
        return self.app.reset_bam_columns()

    def reset_cz_rc_params(self):
        return self.app.reset_cz_rc_params()

    def set_jp_enabled(self, value):
        self.app.set_jp_enabled(value)
        return value

    def set_cz_enabled(self, value):
        self.app.set_cz_enabled(value)
        return value

    def set_bam_enabled(self, value):
        self.app.set_bam_enabled(value)
        return value

    def set_cz_dop_date(self, value):
        self.app.set_cz_dop_date(value)
        return value

    def get_window_size(self):
        return self.app.get_window_size()

    def modify_window_size(self, axis, increment):
        return self.app.modify_window_size(axis, increment)

    def reset_window_size(self):
        return self.app.reset_window_size()

    def get_file_modification_date(self):
        return self.app.get_file_modification_date()

    def process_data(self):
        try:
            self.app.start_processing()
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def update_last_month(self):
        try:
            self.app.update_last_mouns()
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def checking_data_inputs(self):
        try:
            result = self.app.checking_data_inputs()
            if result.get("error") == "missing_values":
                return result.get("data", {})
            elif result.get("error") == "structure_mismatch":
                return {"error": "structure_mismatch"}
            elif result.get("error") == "local_keys_mismatch":
                return {"error": "local_keys_mismatch", "section": result.get("section")}
            return {}
        except Exception as e:
            return {}

    def recovery_data_inputs(self, data):
        try:
            result = self.app.recovery_data_inputs(data)
            return result
        except Exception as e:
            return {"need_gui": True, "data": data}


def main():
    api = Api()
    splash_path = STATIC_DIR / "splash.html"
    main_path = STATIC_DIR / "body_html.html"

    # Создаём splash окно (маленькое, по центру)
    splash = webview.create_window(
        title='Good Morning',
        url=str(splash_path),
        width=400,
        height=500,
        resizable=False,
        frameless=True,
        on_top=True,     # Поверх всех окон
        confirm_close=False
    )

    # Создаём основное окно (скрытое пока)
    main_window = webview.create_window(
        title='Report Conversion',
        url=str(main_path),
        width=950,
        height=750,
        resizable=False,
        js_api=api,
        text_select=False,
        hidden=True  # Скрыто при создании
    )

    def on_loaded():
        """Вызывается когда основное окно загрузилось"""
        # Даём время на инициализацию JS
        time_sleep_main_window = random.randint(2,3)
        time.sleep(time_sleep_main_window)
        # Закрываем splash
        splash.destroy()
        # Показываем основное окно
        main_window.show()
        main_window.restore()

    # Подписываемся на событие загрузки основного окна
    main_window.events.loaded += on_loaded

    # Запускаем
    webview.start(
        icon=str(ICON_PATH),
        debug=False
    )


if __name__ == '__main__':
    main()