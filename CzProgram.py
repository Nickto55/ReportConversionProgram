import os
import sys
import tkinter as tk
from datetime import date, timedelta, datetime
from tkinter import messagebox, ttk

import pandas as pd

import ExcelPrint
import JsonWork
import Search


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_dates(len_date: int, len_date_value: int = 3, year: bool = False):
    """
    Возвращает список из трех лет: позавчерашней, вчерашней и сегодняшней.

    """
    result_list = []
    today = date.today()

    for i in range(0, len_date):
        if year:
            date_form = today - timedelta(days=i * 365)
        else:
            date_form = today - timedelta(days=i)
        if len_date_value == 1:
            format_str = "%Y"
        elif len_date_value == 2:
            format_str = "%Y-%m"
        else:
            format_str = "%Y-%m-%d"
        result_list.append(date.strftime(date_form, format_str))

    return result_list


class CzMain:

    def __init__(self, parent):
        self.parent = parent
        self.config = JsonWork.JsonConfig()

        self.search = Search.SearchCz()
        self.data = self.search.get_dict_all_data()
        self.hearders: list = self.search.get_headers()
        self.columns = []
        self.accepted_no_repit = None
        self.rc_no_repit = None


    def main(self):
        result = []
        modification_time = os.path.getmtime(self.config.getCzPathFile_input())
        file_date = datetime.fromtimestamp(modification_time)
        result_row = ["Дата:", file_date, ""]

        for year in get_dates(int(self.config.getCzColumnName("Table of contents: List_date")), 1, True):
            result_row.append(year)
        result.append(result_row)

        result_row = ["", "", "", "", "", "", ""]
        result.append(result_row)

        count_conversion = 0

        pos_done = 0
        pos_close = 0
        pos_dse = 0
        pos_rc = 0
        pos_accepted = 0
        pos_date = 0

        # Определяем позиции колонок
        for hearder in self.hearders:
            if hearder == self.config.getCzColumnName("Table of contents: Close"):
                pos_close = self.hearders.index(hearder)
            if hearder == self.config.getCzColumnName("Table of contents: RC"):
                self.columns.append(hearder)
                pos_rc = self.hearders.index(hearder)
            if hearder == self.config.getCzColumnName("Table of contents: Accepted"):
                self.columns.append(hearder)
                pos_accepted = self.hearders.index(hearder)
            if hearder == self.config.getCzColumnName("Table of contents: DCE"):
                pos_dse = self.hearders.index(hearder)
            if hearder == self.config.getCzColumnName("Table of contents: Done"):
                pos_done = self.hearders.index(hearder)
            if hearder == self.config.getCzColumnName("Table of contents: Date writer"):
                pos_date = self.hearders.index(hearder)

        if 0 in [pos_close, pos_rc, pos_accepted, pos_dse, pos_done, pos_date]:
            messagebox.showerror("Ошибка", "Произошла ошибка при проверке наличия заголовков")

        self.search.filter_and_save_columns(self.columns)

        rc_no_repit = self.search.get_colum(self.columns[0], foc_mode=1)
        accepted_no_repit = self.search.get_colum(self.columns[1], foc_mode=1)

        """ Если вдруг надо будет автоматизировать выбор
        if "[]" == self.config.getCzColumnName("Table of contents: Rc full value"):
            self.config.setCzColumnName("Table of contents: Rc full value", rc_no_repit)
        if "[]" == self.config.getCzColumnName("Table of contents: Accepted full value"):
            self.config.setCzColumnName("Table of contents: Accepted full value", rc_no_repit)"""

        self.rc_no_repit = rc_no_repit
        self.accepted_no_repit = accepted_no_repit

        """Надо разобраться"""

        gui = GuiManager(len(self.accepted_no_repit), len(self.rc_no_repit), self.parent)

        root = gui.create_gui(

            self.rc_no_repit,
            self.accepted_no_repit
        )

        root.mainloop()

        selected = gui.get_selected_values()
        self.config.setCzColumnName("Table of contents: Rc user chuse value", ["11102", "11403"])
        self.config.setCzColumnName("Table of contents: Accepted user chuse value", selected["accepted"])

        """Конец разбора"""

        def foc(date_search: int = None):
            count_foc = 0
            dce_list = []
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if n_rc in self.config.getCzColumnName("Table of contents: Rc user chuse value", 1):
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            dceLast = ""
                                            for dce in dce_list:
                                                if n_dse == dce:
                                                    dceLast = dce
                                                    continue
                                            if n_dse == dceLast:
                                                continue
                                            dce_list.append(n_dse)
                                            count_foc += 1
                                    else:
                                        dceLast = ""
                                        for dce in dce_list:
                                            if n_dse == dce:
                                                dceLast = dce
                                                continue
                                        if n_dse == dceLast:
                                            continue
                                        dce_list.append(n_dse)
                                        count_foc += 1

                                    break
                                break

            return count_foc

        result_row = ["ФОЦ", foc(), "", foc(2025), foc(2024), foc(2023), foc(2022)]
        result.append(result_row)

        def toc(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11402" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["ТОЦ", toc(), "", toc(2025), toc(2024), toc(2023), toc(2022)]
        result.append(result_row)

        def poc(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11404" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["ПОЦ", poc(), "", poc(2025), poc(2024), poc(2023), poc(2022)]
        result.append(result_row)

        result_row = ["", "", "", "", "", "", ""]
        result.append(result_row)
        result.insert(2, ["Итого", result[2][1] + result[3][1] + result[4][1], "",
                          result[2][3] + result[3][3] + result[4][3], result[2][4] + result[3][4] + result[4][4],
                          result[2][5] + result[3][5] + result[4][5], result[2][6] + result[3][6] + result[4][6]])

        def rc_11102(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11102" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["РЦ 11102", rc_11102(), "", rc_11102(2025), rc_11102(2024), rc_11102(2023), rc_11102(2022)]
        result.append(result_row)

        def rc_11402(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11402" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["РЦ 11402", rc_11402(), "", rc_11402(2025), rc_11402(2024), rc_11402(2023), rc_11402(2022)]
        result.append(result_row)

        def rc_11403(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11403" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["РЦ 11403", rc_11403(), "", rc_11403(2025), rc_11403(2024), rc_11403(2023), rc_11403(2022)]
        result.append(result_row)

        def rc_11404(date_search: int = None):
            count_foc = 0
            for i in self.data:
                data = self.data.get(i, "")
                n_done = data.get(self.config.getCzColumnName("Table of contents: Done"), "")
                n_close = data.get(self.config.getCzColumnName("Table of contents: Close"), "")
                n_dse = data.get(self.config.getCzColumnName("Table of contents: DCE"), "")
                if pd.isna(n_done) and pd.isna(n_close) and not pd.isna(n_dse):
                    n_rc_list = str(data.get(self.config.getCzColumnName("Table of contents: RC"), "")).replace(" ",
                                                                                                                "").split(
                        ",")
                    for n_rc in n_rc_list:
                        if "11404" in n_rc:
                            n_accepted_list = data.get(self.config.getCzColumnName("Table of contents: Accepted", 2),
                                                       "")
                            if not (isinstance(n_accepted_list, str)) and not (
                                    isinstance(n_accepted_list, int)) and not pd.isna(n_accepted_list):
                                try:
                                    n_accepted_list = str(n_accepted_list)[:len(str(n_accepted_list)) // 2 + 1]
                                except:
                                    print(n_accepted_list, """Ошибка с datetime""")
                            n_accepted_list1 = [n_accepted_list]

                            for n_accepted in n_accepted_list1:
                                if n_accepted in self.config.getCzColumnName(
                                        "Table of contents: Accepted user chuse value", 1):
                                    if not pd.isna(date_search):
                                        n_date = data.get(self.config.getCzColumnName("Table of contents: Date writer"),
                                                          "")
                                        if str(date_search) in str(n_date):
                                            count_foc += 1
                                    else:
                                        count_foc += 1

                                    break
                                break
            return count_foc

        result_row = ["РЦ 11404", rc_11404(), "", rc_11404(2025), rc_11404(2024), rc_11404(2023), rc_11404(2022)]
        result.append(result_row)

        # for i in result:
        #     print(i)
        return result


class GuiManager:
    def __init__(self, len_subscription, len_complite, parent):
        """
        Выводит графический интерфейс что отвечает за выбор пользователем нужных данных
        :param len_subscription: колличество согласованных данных (не повторяются)
        :param len_complite: колличество выполненых данных (не повторяются)
        :param parent: главное окно, к которому подвязывается создаваемое
        """
        self.root = tk.Toplevel(parent)
        self.root.title("Выбор параметров")
        self.max_sizeY = 500
        self.root.geometry("600x500")

        try:
            icon_path = resource_path("dirBook.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")

        # Переменные для хранения выбранных значений
        self.selected_subscribed = []
        self.selected_completed = []
        self.selected_rc = []

        self.len_subscription = len_subscription
        self.len_complite = len_complite
        self.max_len = max(len_complite, len_subscription)

        # Временные переменные для чекбоксов
        self.subscribed_vars = []
        self.completed_vars = []
        self.rc_vars = []

    def create_checkboxes(self, parent, items, title):
        """
        Создает фрейм с чекбоксами для списка элементов
        :param parent: вкладка в которой находятся чекбоксы
        :param items: все имена для чекбоксов
        :param title: Название(оглавление) вкладки
        :return vars_list: возвращает лист
        """
        frame = ttk.LabelFrame(parent, text=title)
        frame.pack(fill="x", padx=5, pady=10)

        # Создаем canvas и scrollbar для прокрутки

        canvas = tk.Canvas(frame, height=375)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        vars_list = []
        for item in items:
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(scrollable_frame, text=str(item), variable=var)
            checkbox.pack(anchor="w")
            vars_list.append((var, str(item)))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        return vars_list

    def create_gui(self, subscribed_list, rc_list):
        """Создает основной интерфейс"""
        # Создаем вкладки для каждого списка
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Вкладка для Подписанных
        subscribed_frame = ttk.Frame(notebook)
        notebook.add(subscribed_frame, text="Подписано")
        self.subscribed_vars = self.create_checkboxes(subscribed_frame, subscribed_list, "Выберите подписанные")

        # # Вкладка для РЦ
        # rc_frame = ttk.Frame(notebook)
        # notebook.add(rc_frame, text="РЦ")
        # self.rc_vars = self.create_checkboxes(rc_frame, rc_list, "Выберите РЦ")

        # Кнопка для подтверждения выбора
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(button_frame, text="Применить", command=self.on_confirm).pack(side="right", padx=5)
        ttk.Button(button_frame, text="Выбрать все", command=self.select_all).pack(side="right", padx=5)

        return self.root

    def select_all(self):
        """Выбирает все чекбоксы"""
        for var_list in [self.subscribed_vars, self.completed_vars, self.rc_vars]:
            for var, _ in var_list:
                var.set(True)

    def on_confirm(self):
        """Обработчик нажатия кнопки подтверждения"""
        # Собираем выбранные значения
        self.selected_subscribed = [item for var, item in self.subscribed_vars if var.get()]
        # self.selected_rc = [item for var, item in self.rc_vars if var.get()]

        # Проверяем, что хотя бы что-то выбрано
        if not any([self.selected_subscribed, self.selected_completed, self.selected_rc]):
            messagebox.showwarning("Предупреждение", "Выберите хотя бы одно значение")
            return

        self.root.quit()
        self.root.destroy()

    def get_selected_values(self):
        """Возвращает выбранные значения"""
        return {
            'accepted': self.selected_subscribed,
            'rc': self.selected_rc
        }


if __name__ == "__main__":
    root = tk.Tk()

    config = JsonWork.JsonConfig()
    excelPr = ExcelPrint.ExcelWriter(config.getJPPathFile_output())

    run = CzMain(root)
    excelPr.write_to_sheet(run.main(), "СЗ")

    root.mainloop()






