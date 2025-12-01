import argparse
import calendar
import ctypes
import os
import sys
import tkinter as tk
import webbrowser
from datetime import datetime as dt
from tkinter import *
from tkinter import ttk, messagebox, BOTH, filedialog
from tkinter.ttk import Progressbar

import plyer

import Config
import ExcelPrint
import JpProgram
import JsonWork
from BamProgram import BamMain
from CzProgram import CzMain
from GeneralizationProg import GeneProg


def send_notification(title, message, settime=15, file_path=""):
    plyer.notification.notify(title=title, message=message, app_name="Good Morning (JP)", timeout=settime)


def updateInfoConfig(fileOrDir: int):
    root_select_file = tk.Tk()
    root_select_file.withdraw()  # скрываем главное окно
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


class Main_gui:
    def __init__(self, root):
        """
        Создаёт главное окно приложения
        :param root: окно
        """

        self.start_no_gui_var = False
        self.start_button_var = False
        self.progresslabel_var = StringVar(value="...")
        self.config = JsonWork.JsonConfig()

        """
        
        Размеры окон
        
        """
        self.distance_y_root = self.config.getConfigSizeYProgram()
        self.distance_x_root = self.config.getConfigSizeXProgram()

        """
        
        Настройка главного окна
        
        """
        self.root = root
        self.root.title(self.config.getConfigNameAssemblyProgram())
        self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
        self.root.resizable(False, False)

        """
        
        Меню
        
        """
        main_menu = tk.Menu()
        """Подменю: Settings"""
        settings_menu = tk.Menu(tearoff=0)
        settings_menu.add_command(label="debug mode", command=self.gui_debug_mode)

        """Подменю: Run"""
        run_menu = tk.Menu(tearoff=0)
        # run_menu.add_cascade(label="Открыть программу с консолью", command=self.console_create_in_program)
        run_menu.add_cascade(label="Ручной запуск обновлений прошлого месяца",
                             command=lambda: self.update_last_mouns(progress_barbar=True))
        """Основное меню"""
        main_menu.add_cascade(label="Settings", menu=settings_menu)
        main_menu.add_cascade(label="Run", menu=run_menu)
        self.root.config(menu=main_menu)

        """
        
        Переменные
        
        """
        self.brogressbar_value_var = IntVar(value=0)

        self.modification = IntVar(value=1)
        self.time_and_day_now = str(dt.now())
        self.ubroutine_Jp_var = BooleanVar(value=True)
        self.ubroutine_Cz_var = BooleanVar(value=True)
        self.dop_date_Cz_var = BooleanVar(value=False)
        self.ubroutine_Bam_var = BooleanVar(value=True)
        self.availability_of_required_data = True
        self.parent_label_column_jp_bool = False
        self.parent_label_column_cz_bool = False
        self.parent_label_column_BAM_bool = False
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

        """
        
        
        
        
        Вызовы функций
        
        """
        try:
            icon_path = resource_path("iconca.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
        self.checking_data_inputs()
        if not self.start_no_gui_var:
            self.create_gui()

    def console_create_in_program(self):
        self.args.console = True
        self.create_gui()

    def update_last_mouns(self, progress_barbar: bool = False):

        if not self.start_no_gui_var and progress_barbar:
            self.brogressbar_value_var.set(0)
            self.progressbar.update()
            self.progresslabel_var.set("Работает...")

        if self.ubroutine_Jp_var.get():
            jp_prog = JpProgram.JpMain(mask_date=self.last_mouns_last_day)
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Jp")
            excelPr.write_to_sheet(jp_prog.main(), "ЖП")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))

        if self.ubroutine_Cz_var.get():
            if self.dop_date_Cz_var.get():
                cz_prog = CzMain(self.root, dop_date=1, mask_date=self.last_mouns_last_day)
            else:
                cz_prog = CzMain(self.root, mask_date=self.last_mouns_last_day)
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Cz")
            excelPr.write_to_sheet(cz_prog.main(), "СЗ")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))

        if self.ubroutine_Bam_var.get():
            bam_prog = BamMain(mask_date=self.last_mouns_last_day)
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="BAM")
            excelPr.write_to_sheet(bam_prog.main(), "Бам по УП")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))
        if not self.start_no_gui_var:
            self.change_value_progress_bar_var(100 - self.brogressbar_value_var.get())

        if self.ubroutine_Bam_var.get() and self.ubroutine_Cz_var.get() and self.ubroutine_Jp_var.get():
            ge_prog = GeneProg(mask_date=self.last_mouns_last_day)
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Ge")
            excelPr.write_to_sheet(ge_prog.main(), "Общая информация")

        if not self.start_no_gui_var:
            self.progresslabel_var.set("Программа завершена!")
            self.progresslabel.update()

        print("===================== Обработка старогго месяца завершена=====================================")
        return

    def start_button_command(self):
        if int(self.time_and_day_now[8:10]) <= 5:
            if messagebox.askyesno("Внимание", "Обновить данные за прошлый месяц?"):
                self.update_last_mouns()

        if not self.start_no_gui_var:
            self.brogressbar_value_var.set(0)
            self.progressbar.update()
            self.progresslabel_var.set("Работает...")

        if self.ubroutine_Jp_var.get():
            jp_prog = JpProgram.JpMain()
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Jp")
            excelPr.write_to_sheet(jp_prog.main(), "ЖП")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))

        if self.ubroutine_Cz_var.get():
            if self.dop_date_Cz_var.get():
                cz_prog = CzMain(self.root, dop_date=1)
            else:
                cz_prog = CzMain(self.root)
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Cz")
            excelPr.write_to_sheet(cz_prog.main(), "СЗ")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))

        if self.ubroutine_Bam_var.get():
            bam_prog = BamMain()
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="BAM")
            excelPr.write_to_sheet(bam_prog.main(), "Бам по УП")
            if not self.start_no_gui_var:
                self.change_value_progress_bar_var(
                    100 // (int(self.ubroutine_Jp_var.get()) + int(self.ubroutine_Cz_var.get()) +
                            int(self.ubroutine_Bam_var.get())))
        if not self.start_no_gui_var:
            self.change_value_progress_bar_var(100 - self.brogressbar_value_var.get())

        if self.ubroutine_Bam_var.get() and self.ubroutine_Cz_var.get() and self.ubroutine_Jp_var.get():
            ge_prog = GeneProg()
            excelPr = ExcelPrint.ExcelWriter(self.config.getJPPathFile_output(), min_prog="Ge")
            excelPr.write_to_sheet(ge_prog.main(), "Общая информация")

        if not self.start_no_gui_var:
            self.progresslabel_var.set("Программа завершена!")
            self.progresslabel.update()
        send_notification("Программа завершена", "Программа завершена, проверте файл", 16)

        if self.start_no_gui_var:
            sys.exit()

    def recovery_data_inputs(self, dic_recovery: dict):
        if not self.availability_of_required_data:
            count_recovery_data_inputs = 0

            def label_and_enntry(input_frame):
                nonlocal count_recovery_data_inputs

                label = Label(input_frame, text=f"{value} для {key}")
                label.grid(row=count_recovery_data_inputs, column=0, sticky="w")
                entryL = Entry(input_frame, width=71)
                entryL.grid(row=count_recovery_data_inputs, column=1)
                count_recovery_data_inputs += 1

                return input_frame

            def button_save_command():
                messagebox.showinfo("Недоступно", "В разработке")

            parent = tk.Toplevel(self.root)
            parent.title("Необходимо ввести данные")
            parent.geometry("600x180")
            try:
                icon_path = resource_path("dirBook.ico")
                parent.iconbitmap(icon_path)
            except Exception as e:
                print(f"Не удалось установить иконку: {e}")

            main_faim = LabelFrame(parent, text="Введите:", height=100)
            main_faim.pack(ipadx=5, ipady=5, fill=BOTH)

            for key, values in dic_recovery.items():
                for value in values:
                    main_faim = label_and_enntry(main_faim)

            button_save = Button(parent, text="Сохранить", command=button_save_command)
            button_save.pack(anchor="se")

            parent.mainloop()

    def checking_data_inputs(self):
        """
        Для определения какие данные в json файле отсутсдвуют для корректной работы подпрограмм:
        :return: None, если есть всё, или возвращения имяни не хватающего ключа
        """
        recovery_data_inputs = {"Program:": [], "ЖП": [], "СЗ": [], "BAM": [], "Ge": []}
        data = self.config.getData()
        """
        Проверка наличия всех ключей, глубина = 2
        """
        lose_key_list = Config.configProgram.keys() - data.keys()
        if len(lose_key_list) != 0:
            messagebox.showerror("Ошибка",
                                 "Структура файла конфига нарушена, или значительно обновлена в последнем обнавлении\nУдалите текущий файл конфига")
        for local_key in data.keys():
            local_lose_key_list = Config.configProgram[local_key].keys() - data[local_key].keys()
            if len(local_lose_key_list) != 0:
                if messagebox.askyesno("внимане", "Вайл конфига был изменён, хотите внести возможные изменения?"):
                    for i in local_lose_key_list:
                        print("Измменены ключи:", i)
                        self.config.recoveryLoseKeyAndValue(local_key, i, Config.configProgram[local_key].get(i, ""))

        """
        Проверка всех значений
        """

        for key_program in data.keys():
            for key_value in data[key_program].keys():
                value_config = data[key_program].get(key_value, "")
                if key_value in Config.keys_for_unnecessary_data[key_program]:
                    if value_config == "" or value_config == []:
                        if key_program == "Program:":
                            self.config.recoveryLoseKeyAndValue(key_program, key_value,
                                                                Config.configProgram[key_program].get(key_value, ""))
                        else:
                            self.availability_of_required_data = False
                            recovery_data_inputs[key_program].append(key_value)
        self.recovery_data_inputs(recovery_data_inputs)

    def setings_jp_notebook(self, setings_notebook):
        def button_puth_command():
            self.config.setJPFilePathAndName_input(updateInfoConfig(1))

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getJPPathFile_input())

            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getJPNameFile_input())

        def button_save_command():
            self.config.setJPColumnName("Table of contents: List_date", entry_date.get())
            self.config.setJPFilePathAndName_input(entry_puth.get())

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getJPPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getJPNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.config.getJPColumnName("Table of contents: List_date"))

        def label_column_command():
            def dismiss():
                self.parent_label_column_jp.grab_release()
                self.parent_label_column_jp.destroy()
                self.parent_label_column_jp_bool = False

            def command_save_button():
                self.config.setJPColumnName("Table of contents: Date", entry_Date.get())
                self.config.setJPColumnName("Table of contents: Date removed", entry_Date_removed.get())
                self.config.setJPColumnName("Table of contents: DCE", entry_Date_dse.get())
                self.config.setJPColumnName("Table of contents: Question accepted Full name",
                                            entry_Date_full_accepted_name.get())
                self.config.setJPColumnName("Table of contents: Question removed Full name",
                                            entry_Date_full_removed_name.get())
                self.config.setJPColumnName("Table of contents: Namber", entry_Date_namber.get())
                self.config.setJPColumnName("Table of contents: Name", entry_Date_name.get())
                self.config.setJPColumnName("Table of contents: Translation", entry_Date_translation.get())

                dismiss()

            def command_reset_button():
                self.config.setJPColumnName("Table of contents: Date",
                                            Config.configProgram["ЖП"].get("Table of contents: Date", ""))
                self.config.setJPColumnName("Table of contents: Date removed",
                                            Config.configProgram["ЖП"].get("Table of contents: Date removed", ""))
                self.config.setJPColumnName("Table of contents: DCE",
                                            Config.configProgram["ЖП"].get("Table of contents: DCE", ""))
                self.config.setJPColumnName("Table of contents: Question accepted Full name",
                                            Config.configProgram["ЖП"].get(
                                                "Table of contents: Question accepted Full name", ""))
                self.config.setJPColumnName("Table of contents: Question removed Full name",
                                            Config.configProgram["ЖП"].get(
                                                "Table of contents: Question removed Full name", ""))
                self.config.setJPColumnName("Table of contents: Namber",
                                            Config.configProgram["ЖП"].get("Table of contents: Namber", ""))
                self.config.setJPColumnName("Table of contents: Name",
                                            Config.configProgram["ЖП"].get("Table of contents: Name", ""))
                self.config.setJPColumnName("Table of contents: Translation",
                                            Config.configProgram["ЖП"].get("Table of contents: Translation", ""))

                entry_Date.delete(0, END)
                entry_Date.insert(0, self.config.getJPColumnName("Table of contents: Date"))
                entry_Date_removed.delete(0, END)
                entry_Date_removed.insert(0, self.config.getJPColumnName("Table of contents: Date removed"))
                entry_Date_dse.delete(0, END)
                entry_Date_dse.insert(0, self.config.getJPColumnName("Table of contents: DCE"))
                entry_Date_full_accepted_name.delete(0, END)
                entry_Date_full_accepted_name.insert(0, self.config.getJPColumnName(
                    "Table of contents: Question accepted Full name"))
                entry_Date_full_removed_name.delete(0, END)
                entry_Date_full_removed_name.insert(0, self.config.getJPColumnName(
                    "Table of contents: Question removed Full name"))
                entry_Date_namber.delete(0, END)
                entry_Date_namber.insert(0, self.config.getJPColumnName("Table of contents: Namber"))
                entry_Date_name.delete(0, END)
                entry_Date_name.insert(0, self.config.getJPColumnName("Table of contents: Name"))
                entry_Date_translation.delete(0, END)
                entry_Date_translation.insert(0, self.config.getJPColumnName("Table of contents: Translation"))

            if self.parent_label_column_jp_bool:
                self.parent_label_column_jp.destroy()
                self.parent_label_column_jp_bool = False
            else:
                self.parent_label_column_jp = tk.Toplevel(self.root)
                self.parent_label_column_jp_bool = True
                self.parent_label_column_jp.title("Изменение столбцов")
                self.parent_label_column_jp.geometry("300x197")
                self.parent_label_column_jp.protocol("WM_DELETE_WINDOW", lambda: dismiss())
                self.parent_label_column_jp.wm_attributes("-topmost", True)

                frame = Frame(self.parent_label_column_jp)
                frame.pack(fill=BOTH)

                label_Date = Label(frame, text=f'Дата: ')
                label_Date.grid(row=0, column=0, sticky="w")
                entry_Date = Entry(frame, width=28)
                entry_Date.grid(row=0, column=1)
                entry_Date.delete(0, END)
                entry_Date.insert(0, self.config.getJPColumnName("Table of contents: Date"))

                label_Date_removed = Label(frame, text=f'Дата выполнения:: ')
                label_Date_removed.grid(row=1, column=0, sticky="w")
                entry_Date_removed = Entry(frame, width=28)
                entry_Date_removed.grid(row=1, column=1)
                entry_Date_removed.delete(0, END)
                entry_Date_removed.insert(0, self.config.getJPColumnName("Table of contents: Date removed"))

                label_Date_dse = Label(frame, text=f'ДСЕ: ')
                label_Date_dse.grid(row=2, column=0, sticky="w")
                entry_Date_dse = Entry(frame, width=28)
                entry_Date_dse.grid(row=2, column=1)
                entry_Date_dse.delete(0, END)
                entry_Date_dse.insert(0, self.config.getJPColumnName("Table of contents: DCE"))

                label_Date_full_accepted_name = Label(frame, text=f'Вопрос принят ФИО: ')
                label_Date_full_accepted_name.grid(row=3, column=0, sticky="w")
                entry_Date_full_accepted_name = Entry(frame, width=28)
                entry_Date_full_accepted_name.grid(row=3, column=1)
                entry_Date_full_accepted_name.delete(0, END)
                entry_Date_full_accepted_name.insert(0, self.config.getJPColumnName(
                    "Table of contents: Question accepted Full name"))

                label_Date_full_removed_name = Label(frame, text=f'Вопрос снят ФИО: ')
                label_Date_full_removed_name.grid(row=4, column=0, sticky="w")
                entry_Date_full_removed_name = Entry(frame, width=28)
                entry_Date_full_removed_name.grid(row=4, column=1)
                entry_Date_full_removed_name.delete(0, END)
                entry_Date_full_removed_name.insert(0, self.config.getJPColumnName(
                    "Table of contents: Question removed Full name"))

                label_Date_namber = Label(frame, text=f'Номер: ')
                label_Date_namber.grid(row=5, column=0, sticky="w")
                entry_Date_namber = Entry(frame, width=28)
                entry_Date_namber.grid(row=5, column=1, sticky="w")
                entry_Date_namber.delete(0, END)
                entry_Date_namber.insert(0, self.config.getJPColumnName("Table of contents: Namber"))

                label_Date_name = Label(frame, text=f'Наименование: ')
                label_Date_name.grid(row=6, column=0, sticky="w")
                entry_Date_name = Entry(frame, width=28)
                entry_Date_name.grid(row=6, column=1)
                entry_Date_name.delete(0, END)
                entry_Date_name.insert(0, self.config.getJPColumnName("Table of contents: Name"))

                label_Date_translation = Label(frame, text=f'Перевод: ')
                label_Date_translation.grid(row=7, column=0, sticky="w")
                entry_Date_translation = Entry(frame, width=28)
                entry_Date_translation.grid(row=7, column=1)
                entry_Date_translation.delete(0, END)
                entry_Date_translation.insert(0, self.config.getJPColumnName("Table of contents: Translation"))

                button_reset_column = Button(frame, text="Сбросить всё", command=command_reset_button)
                button_reset_column.grid(row=8, column=0, sticky="w")

                button_save_column = Button(frame, text="Сохранить", command=command_save_button)
                button_save_column.grid(row=8, column=1, sticky="e")

        label_puth = Label(setings_notebook, text="Путь:")
        label_puth.grid(row=0, column=0)

        entry_puth = Entry(setings_notebook, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, self.config.getJPPathFile_input())

        button_puth = Button(setings_notebook, text="Изменить", command=button_puth_command)
        button_puth.grid(row=0, column=2, ipadx=5)

        label_name = Label(setings_notebook, text="Имя: ")
        label_name.grid(row=1, column=0)

        entry_name = Entry(setings_notebook, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, self.config.getJPNameFile_input())

        label_date = Label(setings_notebook, text="Количество дней отображения:")
        label_date.place(x=0, y=50)
        entry_date = Entry(setings_notebook, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)

        entry_date.insert(0, self.config.getJPColumnName("Table of contents: List_date"))

        label_column = Label(setings_notebook, text="Настройка столбцов:")
        label_column.place(x=0, y=75)

        button_column = Button(setings_notebook, text="Настроить", command=label_column_command)
        button_column.place(x=125, y=75)
        button_save = Button(setings_notebook, text="Сохранить", command=button_save_command)
        button_save.place(x=552, y=156)

        return setings_notebook

    def setings_bam_notebook(self, setings_notebook):
        def button_puth_command():
            self.config.setBAMFilePathAndName_input(updateInfoConfig(1))

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getBAMPathFile_input())

            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getBAMNameFile_input())

        def button_save_command():
            self.config.setBAMColumnName("Table of contents: List_date", entry_date.get())
            self.config.setBAMFilePathAndName_input(entry_puth.get())

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getBAMPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getBAMNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.config.getBAMColumnName("Table of contents: List_date"))

        def label_column_command():
            def dismiss(window):
                self.parent_label_column_BAM.grab_release()
                self.parent_label_column_BAM.destroy()
                self.parent_label_column_BAM_bool = False

            def command_save_button():
                self.config.setBAMColumnName("Table of contents: Date", entry_Date.get())
                self.config.setBAMColumnName("Table of contents: listes_excel", entry_listes_excel.get().split(","))

                dismiss(self.parent_label_column_BAM)

            def command_reset_button():
                self.config.setBAMColumnName("Table of contents: Date",
                                             Config.configProgram["BAM"].get("Table of contents: Date", ""))
                self.config.setBAMColumnName("Table of contents: listes_excel",
                                             Config.configProgram["BAM"].get("Table of contents: listes_excel", ""))

                entry_Date.delete(0, END)
                entry_Date.insert(0, self.config.getBAMColumnName("Table of contents: Date"))
                entry_listes_excel.delete(0, END)
                entry_listes_excel.insert(0,
                                          self.config.getBAMColumnName("Table of contents: listes_excel").replace("[",
                                                                                                                  "").replace(
                                              "]", ""))

            if self.parent_label_column_BAM_bool:
                self.parent_label_column_BAM.destroy()
                self.parent_label_column_BAM_bool = False
            else:
                self.parent_label_column_BAM = tk.Toplevel(self.root)
                self.parent_label_column_BAM_bool = True
                self.parent_label_column_BAM.title("Изменение столбцов")
                self.parent_label_column_BAM.geometry("269x68")
                self.parent_label_column_BAM.protocol("WM_DELETE_WINDOW", lambda: dismiss(
                    self.parent_label_column_BAM))  # перехватываем нажатие на крестик
                self.parent_label_column_BAM.wm_attributes("-topmost", True)

                frame = Frame(self.parent_label_column_BAM)
                frame.pack(fill=BOTH)

                label_Date = Label(frame, text=f'Дата: ')
                label_Date.grid(row=0, column=0, sticky="w")
                entry_Date = Entry(frame, width=28)
                entry_Date.grid(row=0, column=1)
                entry_Date.delete(0, END)
                entry_Date.insert(0, self.config.getBAMColumnName("Table of contents: Date"))

                label_listes_excel = Label(frame, text=f'Назван. листов: ')
                label_listes_excel.grid(row=1, column=0, sticky="w")
                entry_listes_excel = Entry(frame, width=28)
                entry_listes_excel.grid(row=1, column=1)
                entry_listes_excel.delete(0, END)
                entry_listes_excel.insert(0,
                                          self.config.getBAMColumnName("Table of contents: listes_excel").replace("[",
                                                                                                                  "").replace(
                                              "]", ""))

                button_reset_column = Button(frame, text="Сбросить всё", command=command_reset_button)
                button_reset_column.grid(row=8, column=0, sticky="w")

                button_save_column = Button(frame, text="Сохранить", command=command_save_button)
                button_save_column.grid(row=8, column=1, sticky="e")

        label_puth = Label(setings_notebook, text="Путь:")
        label_puth.grid(row=0, column=0)

        entry_puth = Entry(setings_notebook, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, self.config.getBAMPathFile_input())

        button_puth = Button(setings_notebook, text="Изменить", command=button_puth_command)
        button_puth.grid(row=0, column=2, ipadx=5)

        label_name = Label(setings_notebook, text="Имя: ")
        label_name.grid(row=1, column=0)

        entry_name = Entry(setings_notebook, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, self.config.getBAMNameFile_input())

        label_date = Label(setings_notebook, text="Количество дней отображения:")
        label_date.place(x=0, y=50)
        entry_date = Entry(setings_notebook, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)
        entry_date.insert(0, self.config.getBAMColumnName("Table of contents: List_date"))

        label_column = Label(setings_notebook, text="Настройка столбцов:")
        label_column.place(x=0, y=75)

        button_column = Button(setings_notebook, text="Настроить", command=label_column_command)
        button_column.place(x=125, y=75)
        button_save = Button(setings_notebook, text="Сохранить", command=button_save_command)
        button_save.place(x=552, y=156)

        return setings_notebook

    def setings_cz_notebook(self, setings_notebook):
        def button_puth_command():
            self.config.setCzFilePathAndName_input(updateInfoConfig(1))

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getCzPathFile_input())

            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getCzNameFile_input())

        def button_save_command():
            self.config.setCzColumnName("Table of contents: List_date", entry_date.get())
            self.config.setCzFilePathAndName_input(entry_puth.get())

            entry_puth.delete(0, END)
            entry_puth.insert(0, self.config.getCzPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.config.getCzNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.config.getCzColumnName("Table of contents: List_date"))

        def label_rc_command():
            def dismiss(window):
                self.parent_label_column_cz.grab_release()
                self.parent_label_column_cz.destroy()
                self.parent_label_column_cz_bool = False

            def command_save_button():
                self.config.setCzColumnName("RC search parameters: Rc foc value", entry_foc.get().split(","))
                self.config.setCzColumnName("RC search parameters: Rc toc value", entry_toc.get().split(","))
                self.config.setCzColumnName("RC search parameters: Rc poc value", entry_poc.get().split(","))

                dismiss(self.parent_label_column_cz)

            def command_reset_button():
                self.config.setCzColumnName("RC search parameters: Rc foc value",
                                            Config.configProgram["СЗ"].get("RC search parameters: Rc foc value", ""))
                self.config.setCzColumnName("RC search parameters: Rc toc value",
                                            Config.configProgram["СЗ"].get("RC search parameters: Rc toc value", ""))
                self.config.setCzColumnName("RC search parameters: Rc poc value",
                                            Config.configProgram["СЗ"].get("RC search parameters: Rc poc value", ""))

                entry_foc.delete(0, END)
                entry_toc.delete(0, END)
                entry_poc.delete(0, END)
                entry_foc.insert(0, self.config.getCzColumnName("RC search parameters: Rc foc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))
                entry_toc.insert(0, self.config.getCzColumnName("RC search parameters: Rc toc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))
                entry_poc.insert(0, self.config.getCzColumnName("RC search parameters: Rc poc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))

            if self.parent_label_column_cz_bool:
                self.parent_label_column_cz.destroy()
                self.parent_label_column_cz_bool = False
            else:
                self.parent_label_column_cz = tk.Toplevel(self.root)
                self.parent_label_column_cz_bool = True
                self.parent_label_column_cz.title("Изменение столбцов")
                self.parent_label_column_cz.geometry("326x110")
                self.parent_label_column_cz.protocol("WM_DELETE_WINDOW", lambda: dismiss(
                    self.parent_label_column_cz))  # перехватываем нажатие на крестик
                self.parent_label_column_cz.wm_attributes("-topmost", True)

                frame = Frame(self.parent_label_column_cz)
                frame.pack(fill=BOTH)

                Label(frame, text="").grid(row=0, column=0)
                label_info = Label(frame, text="Если нужно несколько значений пишите через запятую")
                label_info.place(x=0, y=0)

                label_foc = Label(frame, text=f'ФОЦ: ')
                label_foc.grid(row=1, column=0, sticky="w")
                entry_foc = Entry(frame, width=39)
                entry_foc.grid(row=1, column=1)
                entry_foc.delete(0, END)
                entry_foc.insert(0, self.config.getCzColumnName("RC search parameters: Rc foc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))

                label_toc = Label(frame, text=f'ТОЦ: ')
                label_toc.grid(row=2, column=0, sticky="w")
                entry_toc = Entry(frame, width=39)
                entry_toc.grid(row=2, column=1)
                entry_toc.delete(0, END)
                entry_toc.insert(0, self.config.getCzColumnName("RC search parameters: Rc toc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))

                label_poc = Label(frame, text=f'ПОЦ: ')
                label_poc.grid(row=3, column=0, sticky="w")
                entry_poc = Entry(frame, width=39)
                entry_poc.grid(row=3, column=1)
                entry_poc.delete(0, END)
                entry_poc.insert(0, self.config.getCzColumnName("RC search parameters: Rc poc value").replace("[",
                                                                                                              "").replace(
                    "]", ""))

                button_reset_column = Button(frame, text="Сбросить всё", command=command_reset_button)
                button_reset_column.grid(row=4, column=0, sticky="w")

                button_save_column = Button(frame, text="Сохранить", command=command_save_button)
                button_save_column.grid(row=4, column=1, sticky="e")

        def label_column_command():
            def dismiss(window):
                self.parent_label_column_cz.grab_release()
                self.parent_label_column_cz.destroy()
                self.parent_label_column_cz_bool = False

            def command_save_button():
                self.config.setCzColumnName("Table of contents: List", entry_List.get())
                self.config.setCzColumnName("Table of contents: Accepted", entry_Accepted.get())
                self.config.setCzColumnName("Table of contents: DCE", entry_Date_dse.get())
                self.config.setCzColumnName("Table of contents: RC", entry_Rc.get())
                self.config.setCzColumnName("Table of contents: Close", entry_Close.get())
                self.config.setCzColumnName("Table of contents: Done", entry_Done.get())
                self.config.setCzColumnName("Table of contents: Date writer", entry_Date_writer.get())

                dismiss(self.parent_label_column_cz)

            def command_reset_button():
                self.config.setCzColumnName("Table of contents: List",
                                            Config.configProgram["СЗ"].get("Table of contents: List", ""))
                self.config.setCzColumnName("Table of contents: Accepted",
                                            Config.configProgram["СЗ"].get("Table of contents: Accepted", ""))
                self.config.setCzColumnName("Table of contents: DCE",
                                            Config.configProgram["СЗ"].get("Table of contents: DCE", ""))
                self.config.setCzColumnName("Table of contents: RC",
                                            Config.configProgram["СЗ"].get("Table of contents: RC", ""))
                self.config.setCzColumnName("Table of contents: Close",
                                            Config.configProgram["СЗ"].get("Table of contents: Close", ""))
                self.config.setCzColumnName("Table of contents: Done",
                                            Config.configProgram["СЗ"].get("Table of contents: Done", ""))
                self.config.setCzColumnName("Table of contents: Date writer",
                                            Config.configProgram["СЗ"].get("Table of contents: Date writer", ""))

                entry_List.delete(0, END)
                entry_List.insert(0, self.config.getCzColumnName("Table of contents: List"))
                entry_Accepted.delete(0, END)
                entry_Accepted.insert(0, self.config.getCzColumnName("Table of contents: Accepted"))
                entry_Date_dse.delete(0, END)
                entry_Date_dse.insert(0, self.config.getCzColumnName("Table of contents: DCE"))
                entry_Rc.delete(0, END)
                entry_Rc.insert(0, self.config.getCzColumnName("Table of contents: RC"))
                entry_Close.delete(0, END)
                entry_Close.insert(0, self.config.getCzColumnName("Table of contents: Close"))
                entry_Done.delete(0, END)
                entry_Done.insert(0, self.config.getCzColumnName("Table of contents: Done"))
                entry_Date_writer.delete(0, END)
                entry_Date_writer.insert(0, self.config.getCzColumnName("Table of contents: Date writer"))

            if self.parent_label_column_cz_bool:
                self.parent_label_column_cz.destroy()
                self.parent_label_column_cz_bool = False
            else:
                self.parent_label_column_cz = tk.Toplevel(self.root)
                self.parent_label_column_cz_bool = True
                self.parent_label_column_cz.title("Изменение столбцов")
                self.parent_label_column_cz.geometry("270x173")
                self.parent_label_column_cz.protocol("WM_DELETE_WINDOW", lambda: dismiss(
                    self.parent_label_column_cz))  # перехватываем нажатие на крестик
                self.parent_label_column_cz.wm_attributes("-topmost", True)

                frame = Frame(self.parent_label_column_cz)
                frame.pack(fill=BOTH)

                label_List = Label(frame, text=f'Лист: ')
                label_List.grid(row=0, column=0, sticky="w")
                entry_List = Entry(frame, width=28)
                entry_List.grid(row=0, column=1)
                entry_List.delete(0, END)
                entry_List.insert(0, self.config.getCzColumnName("Table of contents: List"))

                label_Accepted = Label(frame, text="Подписано: ")
                label_Accepted.grid(row=1, column=0, sticky="w")
                entry_Accepted = Entry(frame, width=28)
                entry_Accepted.grid(row=1, column=1)
                entry_Accepted.delete(0, END)
                entry_Accepted.insert(0, self.config.getCzColumnName("Table of contents: Accepted"))

                label_Date_dse = Label(frame, text=f'ДСЕ: ')
                label_Date_dse.grid(row=2, column=0, sticky="w")
                entry_Date_dse = Entry(frame, width=28)
                entry_Date_dse.grid(row=2, column=1)
                entry_Date_dse.delete(0, END)
                entry_Date_dse.insert(0, self.config.getCzColumnName("Table of contents: DCE"))

                label_Rc = Label(frame, text=f'РЦ: ')
                label_Rc.grid(row=3, column=0, sticky="w")
                entry_Rc = Entry(frame, width=28)
                entry_Rc.grid(row=3, column=1)
                entry_Rc.delete(0, END)
                entry_Rc.insert(0, self.config.getCzColumnName("Table of contents: RC"))

                label_Close = Label(frame, text=f'Закрыто: ')
                label_Close.grid(row=4, column=0, sticky="w")
                entry_Close = Entry(frame, width=28)
                entry_Close.grid(row=4, column=1)
                entry_Close.delete(0, END)
                entry_Close.insert(0, self.config.getCzColumnName("Table of contents: Close"))

                label_Done = Label(frame, text=f'Выполнено: ')
                label_Done.grid(row=5, column=0, sticky="w")
                entry_Done = Entry(frame, width=28)
                entry_Done.grid(row=5, column=1, sticky="w")
                entry_Done.delete(0, END)
                entry_Done.insert(0, self.config.getCzColumnName("Table of contents: Done"))

                label_Date_writer = Label(frame, text=f'Наименование: ')
                label_Date_writer.grid(row=6, column=0, sticky="w")
                entry_Date_writer = Entry(frame, width=28)
                entry_Date_writer.grid(row=6, column=1)
                entry_Date_writer.delete(0, END)
                entry_Date_writer.insert(0, self.config.getCzColumnName("Table of contents: Date writer"))

                button_reset_column = Button(frame, text="Сбросить всё", command=command_reset_button)
                button_reset_column.grid(row=8, column=0, sticky="w")

                button_save_column = Button(frame, text="Сохранить", command=command_save_button)
                button_save_column.grid(row=8, column=1, sticky="e")

        label_puth = Label(setings_notebook, text="Путь:")
        label_puth.grid(row=0, column=0)

        entry_puth = Entry(setings_notebook, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, self.config.getCzPathFile_input())

        button_puth = Button(setings_notebook, text="Изменить", command=button_puth_command)
        button_puth.grid(row=0, column=2, ipadx=5)

        label_name = Label(setings_notebook, text="Имя: ")
        label_name.grid(row=1, column=0)

        entry_name = Entry(setings_notebook, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, self.config.getCzNameFile_input())

        label_date = Label(setings_notebook, text="Количество дней отображения:")
        label_date.place(x=0, y=50)
        entry_date = Entry(setings_notebook, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)
        entry_date.insert(0, self.config.getCzColumnName("Table of contents: List_date"))

        label_column = Label(setings_notebook, text="Настройка столбцов:")
        label_column.place(x=0, y=75)

        button_column = Button(setings_notebook, text="Настроить", command=label_column_command)
        button_column.place(x=125, y=75)

        label_rc = Label(setings_notebook, text="Настройка РЦ:")
        label_rc.place(x=0, y=100)

        button_rc = Button(setings_notebook, text="Настроить", command=label_rc_command)
        button_rc.place(x=125, y=100)

        label_dop_date = Label(setings_notebook, text="Дополнительные листы: ")
        label_dop_date.place(x=0, y=126)

        check_button_dop_date = Checkbutton(setings_notebook, variable=self.dop_date_Cz_var)
        check_button_dop_date.place(x=150, y=126)

        button_save = Button(setings_notebook, text="Сохранить", command=button_save_command)
        button_save.place(x=552, y=156)

        return setings_notebook

    def create_gui(self):
        """
        Рисует главное окно
        """

        parser = argparse.ArgumentParser()
        parser.add_argument("--console", action="store_true", help="Запустить с консолью")
        parser.add_argument("--cfile", action="store_true", help="Конфиг файла")
        parser.add_argument("--pstart", action="store_true", help="Моментальный старт")
        parser.add_argument("--nog", action="store_true", help="Без граф. оболочки")

        self.args = parser.parse_args()

        if self.args.cfile:
            CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".ReportConversionProgram")
            webbrowser.open(CONFIG_DIR)
            self.root.destroy()
            self.root.quit()
            return None
        if self.args.pstart:
            self.start_button_var = True
        if self.args.nog:
            self.start_button_var = True
            self.start_no_gui_var = True

            self.root.destroy()
            self.root.quit()

            self.start_button_command()
            return None
        if self.args.console:
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            hStdOut = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE

            if hStdOut == -1 or hStdOut == 0:
                # Если нет — создаём новую консоль
                kernel32.AllocConsole()

                # Перенаправляем stdout/stderr в новую консоль
                sys.stdout = open('CONOUT$', 'w')
                sys.stderr = open('CONOUT$', 'w')

        main_frame = Frame(self.root)
        main_frame.place(x=0, y=0, height=self.distance_y_root, width=self.distance_x_root)

        """
        
        Левое крыло
        
        """

        left_frame = LabelFrame(main_frame, text="Варианты программ:")
        left_frame.place(x=0, y=0, height=self.distance_y_root * 2 / 7, width=self.distance_x_root / 5)

        left_down_frame = LabelFrame(main_frame, text="Информация:")
        left_down_frame.place(x=0, y=100, height=self.distance_y_root * 2 / 3 - self.distance_y_root * 2 / 7 - 4,
                              width=self.distance_x_root / 5)

        """Получаем и выводим информацию"""
        label_time_now = Label(left_down_frame, text=f"Cейчас: {str(dt.now())[:10]}")
        label_time_now.place(x=0, y=0)

        try:
            modification_time = os.path.getmtime(f"{os.getcwd()}/ReportConversionProgram.xlsx")
            last_date_start = str(dt.fromtimestamp(modification_time))
        except:
            last_date_start = "-"

        """Таймеры"""
        label_time_last_start_program = Label(left_down_frame,
                                              text=f"Файл изменён: {last_date_start[:10]}\n                              {last_date_start[11:16]}")
        label_time_last_start_program.place(x=0, y=15)

        """Созаем чеки"""
        checkbut_jp = Checkbutton(left_frame, text="Jp", variable=self.ubroutine_Jp_var, onvalue=1)
        checkbut_jp.grid(row=0, column=0, sticky="w")
        checkbut_cz = Checkbutton(left_frame, text="Cz", variable=self.ubroutine_Cz_var, onvalue=1)
        checkbut_cz.grid(row=1, column=0, sticky="w")
        checkbut_bam = Checkbutton(left_frame, text="Bam", variable=self.ubroutine_Bam_var, onvalue=1)
        checkbut_bam.grid(row=2, column=0, sticky="w")

        """
        
        Правое крыло
        
        """

        right_frame = LabelFrame(main_frame, text="Настройки программ:")
        right_frame.place(x=self.distance_x_root / 5 + 10, y=0, height=self.distance_y_root * 2 / 3,
                          width=self.distance_x_root * 4 / 5 - 10)

        notebook = ttk.Notebook(right_frame)
        notebook.pack(expand=True, fill=BOTH)

        # создаем пару фреймвов
        setings_jp = ttk.Frame(notebook)
        setings_cz = ttk.Frame(notebook)
        setings_bam = ttk.Frame(notebook)

        # Пример size_set = self.create_size_se(size_set)
        setings_jp = self.setings_jp_notebook(setings_jp)
        setings_cz = self.setings_cz_notebook(setings_cz)
        setings_bam = self.setings_bam_notebook(setings_bam)

        setings_jp.pack(fill=BOTH, expand=True)
        setings_cz.pack(fill=BOTH, expand=True)
        setings_bam.pack(fill=BOTH, expand=True)

        # добавляем фреймы в качестве вкладок
        notebook.add(setings_jp, text="Jp")
        notebook.add(setings_cz, text="Cz")
        notebook.add(setings_bam, text="Bam")

        """
        
        Низ
        
        """
        down_frame = LabelFrame(main_frame, text="Процесс")
        down_frame.place(x=0, y=self.distance_y_root * 2 / 3, height=self.distance_y_root * 1 / 3,
                         width=self.distance_x_root)

        self.progressbar = Progressbar(down_frame, orient="horizontal", variable=self.brogressbar_value_var)
        self.progressbar.place(x=5, y=0, width=785)

        self.progresslabel = Label(down_frame, textvariable=self.progresslabel_var)
        self.progresslabel.place(x=0, y=25)

        button_start = Button(down_frame, text="Начать", command=self.start_button_command)
        button_start.place(x=self.distance_x_root - 55, y=25, width=50)

        if self.start_button_var:
            self.start_button_command()

    def change_value_progress_bar_var(self, value_pb):
        self.brogressbar_value_var.set(self.brogressbar_value_var.get() + value_pb)
        self.progressbar.update()

    def gui_debug_mode(self):
        parent = tk.Toplevel(self.root)
        parent.title("Выбор параметров")
        parent.geometry("400x200")

        # создаем набор вкладок
        notebook = ttk.Notebook(parent)
        notebook.pack(expand=True, fill=BOTH)

        # создаем пару фреймвов
        size_set = ttk.Frame(notebook)
        frame2 = ttk.Frame(notebook)

        size_set = self.create_size_se(size_set)

        size_set.pack(fill=BOTH, expand=True, padx=5, pady=5)
        frame2.pack(fill=BOTH, expand=True)

        # добавляем фреймы в качестве вкладок
        notebook.add(size_set, text="Размер")
        notebook.add(frame2, text="...")

        parent.mainloop()

    def create_size_se(self, size_set):
        strVarSizeX = StringVar(value=f"X: {self.config.getConfigSizeXProgram()}")
        strVarSizeY = StringVar(value=f"Y: {self.config.getConfigSizeYProgram()}")

        def button_x_command(sign: float):
            if sign:
                self.distance_x_root += self.modification.get()
            else:
                self.distance_x_root -= self.modification.get()
            self.config.setConfigSizeXProgram(self.distance_x_root)
            strVarSizeX.set(f"X: {self.config.getConfigSizeXProgram()}")
            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
            self.create_gui()

        def button_y_command(sign: float):
            if sign:
                self.distance_y_root += self.modification.get()
            else:
                self.distance_y_root -= self.modification.get()
            self.config.setConfigSizeYProgram(self.distance_y_root)
            strVarSizeY.set(f"Y: {self.config.getConfigSizeYProgram()}")
            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
            self.create_gui()

        def resset_command_button():
            self.distance_x_root = int(Config.configProgram["Program:"].get("Size by X", ""))
            self.distance_y_root = int(Config.configProgram["Program:"].get("Size by Y", ""))
            strVarSizeX.set(f"X: {self.distance_x_root}")
            strVarSizeY.set(f"Y: {self.distance_y_root}")

            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
            self.create_gui()

        button_x_revers = Button(size_set, text="-", command=lambda: button_x_command(False))
        button_x_revers.grid(row=0, column=0)
        label_x = Label(size_set, textvariable=strVarSizeX)
        label_x.grid(row=0, column=1)
        button_x = Button(size_set, text="+", command=lambda: button_x_command(True))
        button_x.grid(row=0, column=2)

        button_y_revers = Button(size_set, text="-", command=lambda: button_y_command(False))
        button_y_revers.grid(row=1, column=0)
        label_y = Label(size_set, textvariable=strVarSizeY)
        label_y.grid(row=1, column=1)
        button_y = Button(size_set, text="+", command=lambda: button_y_command(True))
        button_y.grid(row=1, column=2)

        buton_reset = Button(size_set, text="Reset", command=resset_command_button)
        buton_reset.grid(row=3, column=0)
        frame_right = Frame(size_set)
        frame_right.place(x=100, y=0)

        listmod = [1, 5, 10, 50, 100]

        for i in range(len(listmod)):
            Radiobutton(frame_right, value=listmod[i], text=listmod[i], variable=self.modification).grid(row=i,
                                                                                                         column=0,
                                                                                                         sticky="w")

        return size_set


if __name__ == "__main__":
    root = tk.Tk()
    run = Main_gui(root)
    root.mainloop()
