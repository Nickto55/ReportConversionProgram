from datetime import timedelta, datetime as dt
from tkinter import messagebox
import pandas as pd

import scripts.excel_enter as excel_enter
from scripts.handlings.handling_json import JsonConfig
from scripts.handlings.handling_classes import SearchBam
from scripts.handlings.handling_log import logger


class BamMain:
    def __init__(self, mask_date: str = str(dt.now())):
        self.config = JsonConfig()
        # Получаем список листов/файлов из конфигурации
        self.listes_excel = self.config.getBAMColumnName("Table of contents: listes_excel", intOrlist=1)
        self.mask_date = mask_date
        self.error_masage_var = False

        # Заранее кешируем имена колонок из конфига, чтобы не дёргать его в циклах
        self.col_date = self.config.getBAMColumnName("Table of contents: Date")
        self.col_list_date = int(self.config.getBAMColumnName("Table of contents: List_date"))

    def get_allowed_dates(self) -> set:
        """Генерирует сет валидных дат для фильтрации (быстрее, чем поиск по списку)."""
        today = dt.strptime(self.mask_date[:10], '%Y-%m-%d')
        allowed_dates = set()
        for i in range(self.col_list_date):
            date_form = today - timedelta(days=i)
            allowed_dates.add(date_form.strftime("%Y.%m.%d"))
        return allowed_dates

    def result_creation_function(self, listes_excel):
        try:
            # Инициализируем поиск и получаем данные
            search = SearchBam(listes_excel)
            raw_data = search.get_dict_all_data()

            if not raw_data:
                return []

            # Превращаем словарь в DataFrame для удобной и быстрой работы
            df = pd.DataFrame.from_dict(raw_data, orient='index')

            # 1. Приведение дат к единому формату %Y.%m.%d
            # errors='coerce' превратит сломанные даты в NaT (Not a Time), программа не упадёт
            df['parsed_date'] = pd.to_datetime(df[self.col_date], errors='coerce')
            df['formatted_date'] = df['parsed_date'].dt.strftime("%Y.%m.%d")

            # 2. Фильтрация по периоду (дfloating окно дней из настроек)
            allowed_dates = self.get_allowed_dates()
            df_filtered = df[df['formatted_date'].isin(allowed_dates)].copy()

            if df_filtered.empty:
                if not self.error_masage_var:
                    messagebox.showinfo("Ошибка",
                                        "За указанный период нет данных. \nВ настройках программы БАМ увеличьте дни отображения")
                    self.error_masage_var = True
                return [["ПУСТО"]]

            # Сортируем по дате, чтобы группы шли по порядку
            df_filtered = df_filtered.sort_values(by='formatted_date')

            # Формируем заголовки (первая строчка)
            headers_row = list(df.columns)
            if 'parsed_date' in headers_row: headers_row.remove('parsed_date')
            if 'formatted_date' in headers_row: headers_row.remove('formatted_date')

            result = [headers_row + ["/::/"]]



            # col_up_name = df_filtered.columns[5]
            # col_seven_name = df_filtered.columns[7]

            col_seven_name = 'Наименование'
            col_up_name = 'УП'

            accumulated_dse = 0
            accumulated_up = 0

            for date_group, group in df_filtered.groupby('formatted_date'):
                result.append(["/../", "", "", "", "", "", "", "", ""])
                row_for_count_idx = len(result) - 1

                current_dse_count = 0
                current_up_sum = 0

                for _, row in group.iterrows():
                    row_dict = row.to_dict()
                    row_dict[self.col_date] = date_group

                    try:
                        val_up = row_dict.get(col_up_name, 0)
                        if pd.notna(val_up) and int(val_up) != 0:
                            current_up_sum += int(val_up)
                            current_dse_count += 1
                    except Exception as e:
                        print(e)

                    # Убираем технические колонки pandas перед выгрузкой в массив
                    del row_dict['parsed_date']
                    del row_dict['formatted_date']

                    # Превращаем в список, отсекая первый элемент (как list(row.values())[1:])
                    row_list = list(row_dict.values())[1:]
                    row_list.insert(0, "")  # Пустышка в начало
                    result.append(row_list)

                # Вычисляем разницу (текущие значения в группе)
                # Записываем суммы в строку разделителя `[row_for_count]`
                result[row_for_count_idx].insert(3, current_dse_count)
                result[row_for_count_idx].insert(6, current_up_sum)

                # Копируем 7-й элемент из первой строки данных этой группы
                if len(group) > 0:
                    first_row_data = list(group.iloc[0].to_dict().values())[1:]
                    result[row_for_count_idx].append(first_row_data[7] if len(first_row_data) > 7 else "")

            return result

        except Exception as e:
            logger.exception("result_creation_function критическая ошибка для %s: %s", listes_excel, e)
            return []

    def main(self):
        result = []
        for listes_excel in self.listes_excel:
            res = self.result_creation_function(listes_excel)
            if not res or res == [["ПУСТО"]]:
                continue

            res_0 = res[0]
            res.pop(0)
            res.reverse()
            res.insert(0, res_0)

            for i in res:
                if "№" in i:
                    i.remove("№")
                    i.insert(0, listes_excel)
                result.append(i)
            result.append([])
        return result


if __name__ == '__main__':
    run = BamMain()
    config = JsonConfig()
    for i in run.main():
        print(i)

    # excelPr = excel_enter.ExcelWriter(config.getJPPathFile_output(), min_prog="BAM")
    # excelPr.write_to_sheet(run.main(), "Бам по УП")
