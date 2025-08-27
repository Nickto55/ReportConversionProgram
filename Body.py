# import webbrowser
import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox, BOTH

import pandas as pd

import Config
# import plyer
import JsonWork

class Main_gui:
    def __init__(self, root):
        """
        Создаёт главное окно приложения
        :param root: окно
        """
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

        """Основное меню"""
        main_menu.add_cascade(label="Settings", menu=settings_menu)
        self.root.config(menu=main_menu)

        """
        
        Переменные
        
        """
        self.modification = IntVar(value=1)
        self.availability_of_required_data = True

        """
        
        Вызовы функций
        
        """
        self.checking_data_inputs()

        self.create_gui()

    def recovery_data_inputs(self, dic_recovery:dict):
        count_recovery_data_inputs = 0
        def label_and_enntry(input_frame):
            nonlocal count_recovery_data_inputs

            label = Label(input_frame, text=f"{value} для {key}")
            label.grid(row=count_recovery_data_inputs,column=0,sticky="w")
            entryL = Entry(input_frame,width=71)
            entryL.grid(row=count_recovery_data_inputs, column=1)
            count_recovery_data_inputs+=1

            return input_frame

        def button_save_command():
            messagebox.showinfo("Недоступно", "В разработке")

        parent = tk.Toplevel(self.root)
        parent.title("Необходимо ввести данные")
        parent.geometry("600x180")

        main_faim = LabelFrame(parent, text="Введите:")
        main_faim.pack(ipadx=5, ipady=5, fill=BOTH)

        if not self.availability_of_required_data:
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
        recovery_data_inputs = {
            "Program:": [],
            "ЖП": [],
            "СЗ": []
        }
        data = self.config.getData()
        """
        Проверка наличия всех ключей, глубина = 2
        """
        lose_key_list = Config.configProgram.keys()-data.keys()
        if len(lose_key_list) != 0:
            messagebox.showerror("Ошибка", "Структура файла конфига нарушена, или значительно обновлена в последнем обнавлении\nУдалите текущий файл конфига")
        for local_key in data.keys():
            local_lose_key_list = Config.configProgram[local_key].keys() - data[local_key].keys()
            if len(local_lose_key_list) != 0:
                if messagebox.askyesno("внимане", "Вайл конфига был изменён, хотите внести возможные изменения?"):
                    for i in local_lose_key_list:
                        print("Измменены ключи:", i)
                        self.config.recoveryLoseKeyAndValue(local_key, i, Config.configProgram[local_key].get(i,""))

        """
        Проверка всех значений
        """

        for key_program in data.keys():
            for key_value in data[key_program].keys():
                value_config = data[key_program].get(key_value,"")
                if key_value in Config.keys_for_unnecessary_data[key_program]:
                    if value_config == "" or value_config == []:
                        if key_program == "Program:":
                            self.config.recoveryLoseKeyAndValue(key_program,key_value, Config.configProgram[key_program].get(key_value, ""))
                        else:
                            self.availability_of_required_data = False
                            recovery_data_inputs[key_program].append(key_value)
        self.recovery_data_inputs(recovery_data_inputs)

    def create_gui(self):
        pass

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

        size_set.pack(fill=BOTH, expand=True,padx=5,pady=5)
        frame2.pack(fill=BOTH, expand=True)

        # добавляем фреймы в качестве вкладок
        notebook.add(size_set, text="Размер")
        notebook.add(frame2, text="...")


        parent.mainloop()

    def create_size_se(self, size_set):
        strVarSizeX = StringVar(value=f"X: {self.config.getConfigSizeXProgram()}")
        strVarSizeY = StringVar(value=f"Y: {self.config.getConfigSizeYProgram()}")
        def button_x_command(sign:float):
            if sign:
                self.distance_x_root+=self.modification.get()
            else:
                self.distance_x_root-=self.modification.get()
            self.config.setConfigSizeXProgram(self.distance_x_root)
            strVarSizeX.set(f"X: {self.config.getConfigSizeXProgram()}")
            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
        def button_y_command(sign:float):
            if sign:
                self.distance_y_root+=self.modification.get()
            else:
                self.distance_y_root-=self.modification.get()
            self.config.setConfigSizeYProgram(self.distance_y_root)
            strVarSizeY.set(f"Y: {self.config.getConfigSizeYProgram()}")
            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")
        def resset_command_button():
            self.distance_x_root = int(Config.configProgram["Program:"].get("Size by X", ""))
            self.distance_y_root = int(Config.configProgram["Program:"].get("Size by Y", ""))
            strVarSizeX.set(f"X: {self.distance_x_root}")
            strVarSizeY.set(f"Y: {self.distance_y_root}")

            self.root.geometry(f"{self.distance_x_root}x{self.distance_y_root}")

        button_x_revers = Button(size_set, text="-", command=lambda : button_x_command(False))
        button_x_revers.grid(row=0,column=0)
        label_x = Label(size_set, textvariable=strVarSizeX)
        label_x.grid(row=0,column=1)
        button_x = Button(size_set, text="+", command=lambda : button_x_command(True))
        button_x.grid(row=0,column=2)

        button_y_revers = Button(size_set, text="-", command=lambda : button_y_command(False))
        button_y_revers.grid(row=1,column=0)
        label_y = Label(size_set, textvariable=strVarSizeY)
        label_y.grid(row=1,column=1)
        button_y = Button(size_set, text="+", command=lambda : button_y_command(True))
        button_y.grid(row=1,column=2)


        buton_reset = Button(size_set, text="Reset", command=resset_command_button)
        buton_reset.grid(row=3,column=0)
        frame_right = Frame(size_set)
        frame_right.place(x=100,y=0)


        listmod = [1,5,10,50,100]

        for i in range(len(listmod)):
            Radiobutton(frame_right, value=listmod[i], text=listmod[i], variable=self.modification).grid(row=i,column=0, sticky="w")

        return size_set





if __name__ == "__main__":
    root = tk.Tk()
    run = Main_gui(root)
    root.mainloop()

