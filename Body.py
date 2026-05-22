import os
import sys
import tkinter as tk
from datetime import datetime as dt
from tkinter import (
    ttk, messagebox, BOTH, filedialog, END, StringVar, IntVar, BooleanVar,
    Frame, LabelFrame, Label, Entry, Button, Checkbutton, Radiobutton, Frame
)
from tkinter.ttk import Progressbar

from report_conversion import ReportConversion, resource_path, updateInfoConfig


class MainGUI:
    """
    Класс GUI интерфейса.
    Вся отрисовка и взаимодействие с пользователем.
    """
    
    def __init__(self, root, app: ReportConversion):
        self.app = app
        self.root = root
        
        # Переменные GUI
        self.progresslabel_var = StringVar(value="...")
        self.brogressbar_value_var = IntVar(value=0)
        self.modification_var = IntVar(value=1)
        self.ubroutine_Jp_var = BooleanVar(value=app.is_jp_enabled())
        self.ubroutine_Cz_var = BooleanVar(value=app.is_cz_enabled())
        self.dop_date_Cz_var = BooleanVar(value=app.is_cz_dop_date())
        self.ubroutine_Bam_var = BooleanVar(value=app.is_bam_enabled())
        
        # Синхронизация переменных с app
        self._sync_vars_to_app()
        
        self._setup_window()
        self._create_menu()
        self._create_gui()
        
        # Проверка данных при старте
        self._check_data_on_startup()
        
        # Мгновенный старт если нужен
        if app.should_start_immediately():
            self.start_button_command()
    
    def _sync_vars_to_app(self):
        """Синхронизирует переменные GUI с логикой."""
        self.ubroutine_Jp_var.trace_add("write", lambda *args: self.app.set_jp_enabled(self.ubroutine_Jp_var.get()))
        self.ubroutine_Cz_var.trace_add("write", lambda *args: self.app.set_cz_enabled(self.ubroutine_Cz_var.get()))
        self.ubroutine_Bam_var.trace_add("write", lambda *args: self.app.set_bam_enabled(self.ubroutine_Bam_var.get()))
        self.dop_date_Cz_var.trace_add("write", lambda *args: self.app.set_cz_dop_date(self.dop_date_Cz_var.get()))
    
    def _setup_window(self):
        """Настраивает главное окно."""
        size = self.app.get_window_size()
        self.root.title(self.app.get_program_title())
        self.root.geometry(f"{size['x']}x{size['y']}")
        self.root.resizable(False, False)
        
        try:
            icon_path = resource_path("static/icons/new_icon_report-conversion-program.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
    
    def _create_menu(self):
        """Создает меню приложения."""
        main_menu = tk.Menu()
        
        # Подменю: Settings
        settings_menu = tk.Menu(tearoff=0)
        settings_menu.add_command(label="debug mode", command=self.gui_debug_mode)
        
        # Подменю: Run
        run_menu = tk.Menu(tearoff=0)
        run_menu.add_cascade(
            label="Ручной запуск обновлений прошлого месяца",
            command=lambda: self.update_last_mouns(progress_barbar=True)
        )
        
        main_menu.add_cascade(label="Settings", menu=settings_menu)
        main_menu.add_cascade(label="Run", menu=run_menu)
        self.root.config(menu=main_menu)
    
    def _check_data_on_startup(self):
        """Проверяет данные при запуске."""
        result = self.app.checking_data_inputs()
        if result.get("error") == "structure_mismatch":
            messagebox.showerror(
                "Ошибка",
                "Структура файла конфига нарушена, или значительно обновлена в последнем обнавлении\nУдалите текущий файл конфига"
            )
        elif result.get("error") == "local_keys_mismatch":
            if messagebox.askyesno("внимане", "Вайл конфига был изменён, хотите внести возможные изменения?"):
                # Восстановление ключей
                pass
        elif result.get("error") == "missing_values":
            self._show_recovery_dialog(result["data"])
    
    def _show_recovery_dialog(self, data: dict):
        """Показывает диалог восстановления данных."""
        parent = tk.Toplevel(self.root)
        parent.title("Необходимо ввести данные")
        parent.geometry("600x180")
        
        try:
            icon_path = resource_path("static/icons/dirBook.ico")
            parent.iconbitmap(icon_path)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
        
        main_frame = LabelFrame(parent, text="Введите:", height=100)
        main_frame.pack(ipadx=5, ipady=5, fill=BOTH)
        
        count = 0
        entries = {}
        
        for key, values in data.items():
            for value in values:
                label = Label(main_frame, text=f"{value} для {key}")
                label.grid(row=count, column=0, sticky="w")
                entry = Entry(main_frame, width=71)
                entry.grid(row=count, column=1)
                entries[(key, value)] = entry
                count += 1
        
        def button_save_command():
            messagebox.showinfo("Недоступно", "В разработке")
        
        button_save = Button(parent, text="Сохранить", command=button_save_command)
        button_save.pack(anchor="se")
    
    def _create_gui(self):
        """Создает основной интерфейс."""
        size = self.app.get_window_size()
        
        main_frame = Frame(self.root)
        main_frame.place(x=0, y=0, height=size['y'], width=size['x'])
        
        # === Левое крыло ===
        left_frame = LabelFrame(main_frame, text="Варианты программ:")
        left_frame.place(x=0, y=0, height=size['y'] * 2 / 7, width=size['x'] / 5)
        
        left_down_frame = LabelFrame(main_frame, text="Информация:")
        left_down_frame.place(
            x=0, y=100, 
            height=size['y'] * 2 / 3 - size['y'] * 2 / 7 - 4,
            width=size['x'] / 5
        )
        
        # Информация
        label_time_now = Label(left_down_frame, text=f"Cейчас: {self.app.get_current_date()}")
        label_time_now.place(x=0, y=0)
        
        mod_info = self.app.get_file_modification_date()
        label_time_last = Label(
            left_down_frame,
            text=f"Файл изменён: {mod_info['date']}\n                              {mod_info['time']}"
        )
        label_time_last.place(x=0, y=15)
        
        # Чекбоксы подпрограмм
        Checkbutton(left_frame, text="Jp", variable=self.ubroutine_Jp_var, onvalue=1).grid(row=0, column=0, sticky="w")
        Checkbutton(left_frame, text="Cz", variable=self.ubroutine_Cz_var, onvalue=1).grid(row=1, column=0, sticky="w")
        Checkbutton(left_frame, text="Bam", variable=self.ubroutine_Bam_var, onvalue=1).grid(row=2, column=0, sticky="w")
        
        # === Правое крыло (вкладки настроек) ===
        right_frame = LabelFrame(main_frame, text="Настройки программ:")
        right_frame.place(
            x=size['x'] / 5 + 10, y=0,
            height=size['y'] * 2 / 3,
            width=size['x'] * 4 / 5 - 10
        )
        
        notebook = ttk.Notebook(right_frame)
        notebook.pack(expand=True, fill=BOTH)
        
        setings_jp = ttk.Frame(notebook)
        setings_cz = ttk.Frame(notebook)
        setings_bam = ttk.Frame(notebook)
        
        self._setup_jp_tab(setings_jp)
        self._setup_cz_tab(setings_cz)
        self._setup_bam_tab(setings_bam)
        
        setings_jp.pack(fill=BOTH, expand=True)
        setings_cz.pack(fill=BOTH, expand=True)
        setings_bam.pack(fill=BOTH, expand=True)
        
        notebook.add(setings_jp, text="Jp")
        notebook.add(setings_cz, text="Cz")
        notebook.add(setings_bam, text="Bam")
        
        # === Низ (прогресс) ===
        down_frame = LabelFrame(main_frame, text="Процесс")
        down_frame.place(
            x=0, y=size['y'] * 2 / 3,
            height=size['y'] * 1 / 3,
            width=size['x']
        )
        
        self.progressbar = Progressbar(down_frame, orient="horizontal", variable=self.brogressbar_value_var)
        self.progressbar.place(x=5, y=0, width=785)
        
        self.progresslabel = Label(down_frame, textvariable=self.progresslabel_var)
        self.progresslabel.place(x=0, y=25)
        
        Button(down_frame, text="Начать", command=self.start_button_command).place(
            x=size['x'] - 55, y=25, width=50
        )
    
    # === Вкладка ЖП ===
    def _setup_jp_tab(self, frame):
        settings = self.app.get_jp_settings()
        
        # Путь
        Label(frame, text="Путь:").grid(row=0, column=0)
        entry_puth = Entry(frame, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, settings["path"])
        
        def button_puth_command():
            path = updateInfoConfig(1)
            self.app.set_jp_path(path)
            entry_puth.delete(0, END)
            entry_puth.insert(0, path)
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getJPNameFile_input())
        
        Button(frame, text="Изменить", command=button_puth_command).grid(row=0, column=2, ipadx=5)
        
        # Имя
        Label(frame, text="Имя: ").grid(row=1, column=0)
        entry_name = Entry(frame, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, settings["name"])
        
        # Количество дней
        Label(frame, text="Количество дней отображения:").place(x=0, y=50)
        entry_date = Entry(frame, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)
        entry_date.insert(0, settings["list_date"])
        
        # Настройка колонок
        Label(frame, text="Настройка столбцов:").place(x=0, y=75)
        
        def label_column_command():
            self._show_jp_columns_dialog()
        
        Button(frame, text="Настроить", command=label_column_command).place(x=125, y=75)
        
        def button_save_command():
            self.app.set_jp_settings(
                entry_puth.get(),
                entry_date.get(),
                {}  # Колонки сохраняются через диалог
            )
            entry_puth.delete(0, END)
            entry_puth.insert(0, self.app.config.getJPPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getJPNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.app.config.getJPColumnName("Table of contents: List_date"))
        
        Button(frame, text="Сохранить", command=button_save_command).place(x=552, y=156)
    
    def _show_jp_columns_dialog(self):
        """Диалог настройки колонок ЖП."""
        parent = tk.Toplevel(self.root)
        parent.title("Изменение столбцов")
        parent.geometry("300x197")
        parent.wm_attributes("-topmost", True)
        
        try:
            icon_path = resource_path("static/icons/dirBook.ico")
            parent.iconbitmap(icon_path)
        except:
            pass
        
        settings = self.app.get_jp_settings()
        cols = settings["columns"]
        
        frame = Frame(parent)
        frame.pack(fill=BOTH)
        
        entries = {}
        fields = [
            ("Дата:", "date", "Table of contents: Date"),
            ("Дата выполнения::", "date_removed", "Table of contents: Date removed"),
            ("ДСЕ:", "dse", "Table of contents: DCE"),
            ("Вопрос принят ФИО:", "accepted_name", "Table of contents: Question accepted Full name"),
            ("Вопрос снят ФИО:", "removed_name", "Table of contents: Question removed Full name"),
            ("Номер:", "number", "Table of contents: Namber"),
            ("Наименование:", "name", "Table of contents: Name"),
            ("Перевод:", "translation", "Table of contents: Translation"),
        ]
        
        for i, (label_text, key, config_key) in enumerate(fields):
            Label(frame, text=label_text).grid(row=i, column=0, sticky="w")
            entry = Entry(frame, width=28)
            entry.grid(row=i, column=1)
            entry.delete(0, END)
            entry.insert(0, cols[key])
            entries[config_key] = entry
        
        def command_save_button():
            for config_key, entry in entries.items():
                self.app.config.setJPColumnName(config_key, entry.get())
            parent.destroy()
        
        def command_reset_button():
            defaults = self.app.reset_jp_columns()
            for config_key, entry in entries.items():
                entry.delete(0, END)
                entry.insert(0, defaults.get(config_key, ""))
        
        Button(frame, text="Сбросить всё", command=command_reset_button).grid(row=8, column=0, sticky="w")
        Button(frame, text="Сохранить", command=command_save_button).grid(row=8, column=1, sticky="e")
    
    # === Вкладка BAM ===
    def _setup_bam_tab(self, frame):
        settings = self.app.get_bam_settings()
        
        Label(frame, text="Путь:").grid(row=0, column=0)
        entry_puth = Entry(frame, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, settings["path"])
        
        def button_puth_command():
            path = updateInfoConfig(1)
            self.app.set_bam_path(path)
            entry_puth.delete(0, END)
            entry_puth.insert(0, path)
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getBAMNameFile_input())
        
        Button(frame, text="Изменить", command=button_puth_command).grid(row=0, column=2, ipadx=5)
        
        Label(frame, text="Имя: ").grid(row=1, column=0)
        entry_name = Entry(frame, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, settings["name"])
        
        Label(frame, text="Количество дней отображения:").place(x=0, y=50)
        entry_date = Entry(frame, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)
        entry_date.insert(0, settings["list_date"])
        
        Label(frame, text="Настройка столбцов:").place(x=0, y=75)
        
        def label_column_command():
            self._show_bam_columns_dialog()
        
        Button(frame, text="Настроить", command=label_column_command).place(x=125, y=75)
        
        def button_save_command():
            self.app.set_bam_settings(
                entry_puth.get(),
                entry_date.get(),
                {}
            )
            entry_puth.delete(0, END)
            entry_puth.insert(0, self.app.config.getBAMPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getBAMNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.app.config.getBAMColumnName("Table of contents: List_date"))
        
        Button(frame, text="Сохранить", command=button_save_command).place(x=552, y=156)
    
    def _show_bam_columns_dialog(self):
        """Диалог настройки колонок BAM."""
        parent = tk.Toplevel(self.root)
        parent.title("Изменение столбцов")
        parent.geometry("269x68")
        parent.wm_attributes("-topmost", True)
        
        settings = self.app.get_bam_settings()
        cols = settings["columns"]
        
        frame = Frame(parent)
        frame.pack(fill=BOTH)
        
        Label(frame, text="Дата:").grid(row=0, column=0, sticky="w")
        entry_Date = Entry(frame, width=28)
        entry_Date.grid(row=0, column=1)
        entry_Date.delete(0, END)
        entry_Date.insert(0, cols["date"])
        
        Label(frame, text="Назван. листов:").grid(row=1, column=0, sticky="w")
        entry_listes = Entry(frame, width=28)
        entry_listes.grid(row=1, column=1)
        entry_listes.delete(0, END)
        entry_listes.insert(0, str(cols["listes_excel"]).replace("[", "").replace("]", ""))
        
        def command_save_button():
            self.app.config.setBAMColumnName("Table of contents: Date", entry_Date.get())
            self.app.config.setBAMColumnName("Table of contents: listes_excel", entry_listes.get().split(","))
            parent.destroy()
        
        def command_reset_button():
            defaults = self.app.reset_bam_columns()
            entry_Date.delete(0, END)
            entry_Date.insert(0, defaults["Table of contents: Date"])
            entry_listes.delete(0, END)
            entry_listes.insert(0, str(defaults["Table of contents: listes_excel"]).replace("[", "").replace("]", ""))
        
        Button(frame, text="Сбросить всё", command=command_reset_button).grid(row=8, column=0, sticky="w")
        Button(frame, text="Сохранить", command=command_save_button).grid(row=8, column=1, sticky="e")
    
    # === Вкладка CZ ===
    def _setup_cz_tab(self, frame):
        settings = self.app.get_cz_settings()
        
        Label(frame, text="Путь:").grid(row=0, column=0)
        entry_puth = Entry(frame, width=85)
        entry_puth.grid(row=0, column=1)
        entry_puth.delete(0, END)
        entry_puth.insert(0, settings["path"])
        
        def button_puth_command():
            path = updateInfoConfig(1)
            self.app.set_cz_path(path)
            entry_puth.delete(0, END)
            entry_puth.insert(0, path)
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getCzNameFile_input())
        
        Button(frame, text="Изменить", command=button_puth_command).grid(row=0, column=2, ipadx=5)
        
        Label(frame, text="Имя: ").grid(row=1, column=0)
        entry_name = Entry(frame, width=85)
        entry_name.grid(row=1, column=1)
        entry_name.delete(0, END)
        entry_name.insert(0, settings["name"])
        
        Label(frame, text="Количество дней отображения:").place(x=0, y=50)
        entry_date = Entry(frame, width=3)
        entry_date.place(x=180, y=50)
        entry_date.delete(0, END)
        entry_date.insert(0, settings["list_date"])
        
        Label(frame, text="Настройка столбцов:").place(x=0, y=75)
        
        def label_column_command():
            self._show_cz_columns_dialog()
        
        Button(frame, text="Настроить", command=label_column_command).place(x=125, y=75)
        
        Label(frame, text="Настройка РЦ:").place(x=0, y=100)
        
        def label_rc_command():
            self._show_cz_rc_dialog()
        
        Button(frame, text="Настроить", command=label_rc_command).place(x=125, y=100)
        
        Label(frame, text="Дополнительные листы:").place(x=0, y=126)
        Checkbutton(frame, variable=self.dop_date_Cz_var).place(x=150, y=126)
        
        def button_save_command():
            self.app.set_cz_settings(
                entry_puth.get(),
                entry_date.get(),
                {},
                {}
            )
            entry_puth.delete(0, END)
            entry_puth.insert(0, self.app.config.getCzPathFile_input())
            entry_name.delete(0, END)
            entry_name.insert(0, self.app.config.getCzNameFile_input())
            entry_date.delete(0, END)
            entry_date.insert(0, self.app.config.getCzColumnName("Table of contents: List_date"))
        
        Button(frame, text="Сохранить", command=button_save_command).place(x=552, y=156)
    
    def _show_cz_columns_dialog(self):
        """Диалог настройки колонок CZ."""
        parent = tk.Toplevel(self.root)
        parent.title("Изменение столбцов")
        parent.geometry("270x173")
        parent.wm_attributes("-topmost", True)
        
        settings = self.app.get_cz_settings()
        cols = settings["columns"]
        
        frame = Frame(parent)
        frame.pack(fill=BOTH)
        
        fields = [
            ("Лист:", "list", "Table of contents: List"),
            ("Подписано:", "accepted", "Table of contents: Accepted"),
            ("ДСЕ:", "dse", "Table of contents: DCE"),
            ("РЦ:", "rc", "Table of contents: RC"),
            ("Закрыто:", "close", "Table of contents: Close"),
            ("Выполнено:", "done", "Table of contents: Done"),
            ("Наименование:", "date_writer", "Table of contents: Date writer"),
        ]
        
        entries = {}
        for i, (label_text, key, config_key) in enumerate(fields):
            Label(frame, text=label_text).grid(row=i, column=0, sticky="w")
            entry = Entry(frame, width=28)
            entry.grid(row=i, column=1)
            entry.delete(0, END)
            entry.insert(0, cols[key])
            entries[config_key] = entry
        
        def command_save_button():
            for config_key, entry in entries.items():
                self.app.config.setCzColumnName(config_key, entry.get())
            parent.destroy()
        
        def command_reset_button():
            defaults = self.app.reset_cz_columns()
            for config_key, entry in entries.items():
                entry.delete(0, END)
                entry.insert(0, defaults.get(config_key, ""))
        
        Button(frame, text="Сбросить всё", command=command_reset_button).grid(row=8, column=0, sticky="w")
        Button(frame, text="Сохранить", command=command_save_button).grid(row=8, column=1, sticky="e")
    
    def _show_cz_rc_dialog(self):
        """Диалог настройки РЦ CZ."""
        parent = tk.Toplevel(self.root)
        parent.title("Изменение столбцов")
        parent.geometry("326x110")
        parent.wm_attributes("-topmost", True)
        
        settings = self.app.get_cz_settings()
        rc = settings["rc_params"]
        
        frame = Frame(parent)
        frame.pack(fill=BOTH)
        
        Label(frame, text="Если нужно несколько значений пишите через запятую").place(x=0, y=0)
        
        fields = [
            ("ФОЦ:", "foc", "RC search parameters: Rc foc value"),
            ("ТОЦ:", "toc", "RC search parameters: Rc toc value"),
            ("ПОЦ:", "poc", "RC search parameters: Rc poc value"),
        ]
        
        entries = {}
        for i, (label_text, key, config_key) in enumerate(fields):
            Label(frame, text=label_text).grid(row=i+1, column=0, sticky="w")
            entry = Entry(frame, width=39)
            entry.grid(row=i+1, column=1)
            entry.delete(0, END)
            entry.insert(0, str(rc[key]).replace("[", "").replace("]", ""))
            entries[config_key] = entry
        
        def command_save_button():
            for config_key, entry in entries.items():
                self.app.config.setCzColumnName(config_key, entry.get().split(","))
            parent.destroy()
        
        def command_reset_button():
            defaults = self.app.reset_cz_rc_params()
            for config_key, entry in entries.items():
                entry.delete(0, END)
                entry.insert(0, str(defaults.get(config_key, "")).replace("[", "").replace("]", ""))
        
        Button(frame, text="Сбросить всё", command=command_reset_button).grid(row=4, column=0, sticky="w")
        Button(frame, text="Сохранить", command=command_save_button).grid(row=4, column=1, sticky="e")
    
    # === Обработчики команд ===
    
    def update_last_mouns(self, progress_barbar: bool = False):
        """Обновляет данные за прошлый месяц."""
        if progress_barbar:
            self.brogressbar_value_var.set(0)
            self.progressbar.update()
            self.progresslabel_var.set("Работает...")
        
        def progress_callback(value):
            self.brogressbar_value_var.set(self.brogressbar_value_var.get() + value)
            self.progressbar.update()
        
        self.app.update_last_mouns(progress_callback=progress_callback if progress_barbar else None)
        
        if progress_barbar:
            self.progresslabel_var.set("Программа завершена!")
            self.progresslabel.update()
    
    def start_button_command(self):
        """Обработчик кнопки 'Начать'."""
        # Проверка на обновление прошлого месяца
        if int(self.app.time_and_day_now[8:10]) <= 5:
            if messagebox.askyesno("Внимание", "Обновить данные за прошлый месяц?"):
                self.update_last_mouns()
                return
        
        self.brogressbar_value_var.set(0)
        self.progressbar.update()
        self.progresslabel_var.set("Работает...")
        
        def progress_callback(value):
            self.brogressbar_value_var.set(self.brogressbar_value_var.get() + value)
            self.progressbar.update()
        
        def completion_callback():
            self.progresslabel_var.set("Программа завершена!")
            self.progresslabel.update()
        
        self.app.start_processing(
            progress_callback=progress_callback,
            completion_callback=completion_callback
        )
    
    def change_value_progress_bar_var(self, value_pb):
        """Обновляет значение прогресс-бара."""
        self.brogressbar_value_var.set(self.brogressbar_value_var.get() + value_pb)
        self.progressbar.update()
    
    # === Debug mode ===
    
    def gui_debug_mode(self):
        """Открывает окно debug mode."""
        parent = tk.Toplevel(self.root)
        parent.title("Выбор параметров")
        parent.geometry("400x200")
        
        notebook = ttk.Notebook(parent)
        notebook.pack(expand=True, fill=BOTH)
        
        size_set = ttk.Frame(notebook)
        frame2 = ttk.Frame(notebook)
        
        self._setup_size_tab(size_set)
        
        size_set.pack(fill=BOTH, expand=True, padx=5, pady=5)
        frame2.pack(fill=BOTH, expand=True)
        
        notebook.add(size_set, text="Размер")
        notebook.add(frame2, text="...")
    
    def _setup_size_tab(self, size_set):
        """Вкладка настройки размеров окна."""
        size = self.app.get_window_size()
        strVarSizeX = StringVar(value=f"X: {size['x']}")
        strVarSizeY = StringVar(value=f"Y: {size['y']}")
        
        def button_x_command(sign: bool):
            new_x = self.app.modify_window_size("x", self.modification_var.get() if sign else -self.modification_var.get())
            strVarSizeX.set(f"X: {new_x}")
            self._refresh_window()
        
        def button_y_command(sign: bool):
            new_y = self.app.modify_window_size("y", self.modification_var.get() if sign else -self.modification_var.get())
            strVarSizeY.set(f"Y: {new_y}")
            self._refresh_window()
        
        def resset_command_button():
            new_size = self.app.reset_window_size()
            strVarSizeX.set(f"X: {new_size['x']}")
            strVarSizeY.set(f"Y: {new_size['y']}")
            self._refresh_window()
        
        Button(size_set, text="-", command=lambda: button_x_command(False)).grid(row=0, column=0)
        Label(size_set, textvariable=strVarSizeX).grid(row=0, column=1)
        Button(size_set, text="+", command=lambda: button_x_command(True)).grid(row=0, column=2)
        
        Button(size_set, text="-", command=lambda: button_y_command(False)).grid(row=1, column=0)
        Label(size_set, textvariable=strVarSizeY).grid(row=1, column=1)
        Button(size_set, text="+", command=lambda: button_y_command(True)).grid(row=1, column=2)
        
        Button(size_set, text="Reset", command=resset_command_button).grid(row=3, column=0)
        
        frame_right = Frame(size_set)
        frame_right.place(x=100, y=0)
        
        listmod = [1, 5, 10, 50, 100]
        for i, val in enumerate(listmod):
            Radiobutton(frame_right, value=val, text=val, variable=self.modification_var).grid(
                row=i, column=0, sticky="w"
            )
        
        return size_set
    
    def _refresh_window(self):
        """Обновляет размер окна и пересоздает GUI."""
        size = self.app.get_window_size()
        self.root.geometry(f"{size['x']}x{size['y']}")
        # Очистка и пересоздание
        for widget in self.root.winfo_children():
            widget.destroy()
        self._create_menu()
        self._create_gui()


if __name__ == "__main__":
    from report_conversion import ReportConversion
    
    app = ReportConversion()
    
    if app.should_open_config():
        app.open_config_dir()
        sys.exit()
    
    if app.is_no_gui_mode():
        app.start_processing()
    else:
        root = tk.Tk()
        gui = MainGUI(root, app)
        root.mainloop()