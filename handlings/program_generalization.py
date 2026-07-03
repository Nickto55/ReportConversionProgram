"""Внимание, полоса отображения, и обрезка вставляемых формул, находятся в ExcelPrint,
при изменении 5 строчки(оглавления календаря), см. на совместимость"""

import calendar
import copy
from datetime import datetime as dt

import pandas as pd

from scripts.handlings.handling_json import JsonConfig
from scripts.handlings.handling_classes import SearchGe
from scripts.handlings.handling_log import logger, attempt_recover


class GeneProg:
    def __init__(self, mask_date: str = str(dt.now())):
        self.count_day_last_mounth: str = ""
        self.config = JsonConfig()
        self.search_in_sheet = SearchGe()

        self.row_static = []
        self.now_month = None

        self.data_jp = None
        self.data_cz = None
        self.data_bam = None

        self.foc_date_list = []
        self.toc_date_list = []
        self.poc_date_list = []

        self.data_ge = None

        self.date_day_now = None
        self.titel_mounth_now = None
        self.titel_mounth_last = None

        self.date_cz_2 = ""
        self.mask_date = mask_date

        self.get_datas_excel()

    def get_datas_excel(self, start_row=1, start_col=1):
        def _do():
            return self.search_in_sheet.sheet_name_list()

        def _reload_search_in_sheet():
            try:
                self.search_in_sheet = SearchGe()
                logger.info("Reinitialized SearchGe during recovery")
            except Exception:
                logger.exception("Failed to reinitialize SearchGe")

        try:
            self.data_jp, self.data_cz, self.data_bam, self.data_ge = attempt_recover(_do, recover_funcs=[
                _reload_search_in_sheet], attempts=2)
        except Exception:
            logger.exception("get_datas_excel failed; setting empty data structures")
            self.data_jp, self.data_cz, self.data_bam, self.data_ge = {}, {}, {}, {}

    def create_calendar(self):
        """Создание календаря, для таблицы"""

        result = []
        self.row_static = ["" for i in range(39)]
        result_row2 = ["" for i in range(8)]
        result_row2.append(r"\..../")
        for i in range(29):
            result_row2.append(r"\.../")
        for i in range(4): result.append(self.row_static.copy())
        date_str = self.mask_date[:10]
        current_year = int(date_str[:4])
        current_month = int(str(date_str[5:7]).replace("-", ""))

        if current_month == 1:
            last_month_num = 12
            last_year = current_year - 1
        else:
            last_month_num = current_month - 1
            last_year = current_year

        date_str_last = dt(last_year, last_month_num, 1)

        result[0][0] = date_str
        date_object = dt.strptime(date_str, '%Y-%m-%d')
        self.date_day_now = int(str(date_object)[8:10])
        titel_mounth_last = (date_str_last.strftime('%B'))
        titel_mounth_now = date_object.strftime('%B')
        self.titel_mounth_now = titel_mounth_now
        self.titel_mounth_last = titel_mounth_last
        count_day_now_mounth = list(calendar.monthrange(current_year, current_month))[-1]
        last_month = last_month_num
        self.now_month = current_month
        count_day_last_mounth = list(calendar.monthrange(last_year, last_month))[-1]
        self.count_day_last_mounth = count_day_last_mounth
        self.last_mouns_last_day = str(date_str_last)[:8] + str(count_day_last_mounth)

        result_row = []

        for i in range(7):
            result_row.append("")
        result_row.append(titel_mounth_last)
        result_row.append(titel_mounth_now)
        for i in range(29):
            result_row.append("")
        result.append(result_row)

        result_row = ["", "", "", "", ""]

        for i in range(count_day_last_mounth - 3 + 1, count_day_last_mounth + 1):
            result_row.append(i)

        for i in range(1, count_day_now_mounth + 1):
            result_row.append(i)
        result.append(result_row)

        return result

    def form_gen(self, result, row_idx, column_idx):
        try:
            if result[5][column_idx] <= self.date_day_now:
                if column_idx == 8:
                    num_form = 1
                else:
                    num_form = 2
                if num_form == 1: result[row_idx][column_idx] = r"\..../"
                if num_form == 2: result[row_idx][column_idx] = r"\.../"
        except Exception as e:
            print(e, f":\n    len(result[5]):{len(result[5])}, column_idx:{column_idx}")

    def chapter_gen(self, result: list, name_chapter: str, first_subsection: str, second_subsection: str,
                    thead_subsection: str = None, subsection_last: str = None):

        for i in range(3): result.append(self.row_static.copy())

        result[-1][0] = name_chapter

        result_row = ["", "", first_subsection]
        for i in range(36): result_row.append("")
        result.append(result_row)

        result_row = ["", "", second_subsection, "", ""]
        for i in range(34): result_row.append("")
        result.append(result_row)

        if thead_subsection is not None:
            result_row = ["", "", thead_subsection, "", ""]
            for i in range(34): result_row.append("")
            result.append(result_row)
            result_row = ["", "", subsection_last, "", ""]
            for i in range(34): result_row.append("")
            result.append(result_row)
        return result

    def _extract_date_from_row(self, key_row):
        """Извлекает дату из строки data_bam.
        Для обычных строк — из колонки 'Дата', для итоговых — из Unnamed: 11"""
        date_val = key_row.get("Дата", "")
        if not pd.isna(date_val) and date_val and str(date_val).strip():
            return str(date_val).strip()

        # Для итоговых строк ищем дату в Unnamed колонках
        for col_name in ["Unnamed: 11", "Unnamed: 12", "Unnamed: 13"]:
            val = key_row.get(col_name, "")
            if not pd.isna(val) and val and str(val).strip():
                val_str = str(val).strip()
                # Проверяем, что это похоже на дату (YYYY.MM.DD или YYYY-MM-DD)
                if len(val_str) >= 10 and val_str[4] in ".-" and val_str[7] in ".-":
                    return val_str
        return ""

    def _is_total_row(self, key_row):
        """Проверяет, является ли строка итоговой (число в Наименовании, nan в Дате)"""
        name = key_row.get("Наименование", "")
        date_val = key_row.get("Дата", "")
        return isinstance(name, (int, float)) and pd.isna(date_val)

    def _is_data_row(self, key_row):
        """Проверяет, является ли строка строкой с данными изделия"""
        name = key_row.get("Наименование", "")
        return (isinstance(name, str) and name.strip() and
                not pd.isna(name) and name.strip() not in ["nan", ""])

    def filling_data(self, result):
        for row_idx in range(len(result)):
            if row_idx < 8:
                continue
            for column_idx in range(len(result[row_idx])):
                if column_idx < 8:
                    continue
                if column_idx < len(result[5]):
                    if row_idx == 9:
                        self.jp_row_value_now(result, column_idx, row_idx)
                    if row_idx == 10:
                        self.jp_row_value_now(result, column_idx, row_idx, "Вып. Задач")
                    if row_idx == 14:
                        self.bam_row_value_now(result, row_idx, column_idx, 0, self.foc_date_list)
                    if row_idx == 15:
                        self.bam_row_value_now(result, row_idx, column_idx, 1, self.foc_date_list)

                    if row_idx == 19:
                        self.bam_row_value_now(result, row_idx, column_idx, 0, self.toc_date_list)
                    if row_idx == 20:
                        self.bam_row_value_now(result, row_idx, column_idx, 1, self.toc_date_list)

                    if row_idx == 24:
                        self.bam_row_value_now(result, row_idx, column_idx, 0, self.poc_date_list)
                    if row_idx == 25:
                        self.bam_row_value_now(result, row_idx, column_idx, 1, self.poc_date_list)

                    if row_idx == 29:
                        self.ldklds(result, row_idx, column_idx, self.egogo)
                    if row_idx == 30:
                        self.ldklds(result, row_idx, column_idx, self.foc)
                    if row_idx == 31:
                        self.ldklds(result, row_idx, column_idx, self.toc)
                    if row_idx == 32:
                        self.ldklds(result, row_idx, column_idx, self.poc)

                    # Обозначение места формул
                    if row_idx in [16, 17, 21, 22, 26, 27]:
                        self.form_gen(result, row_idx, column_idx)

        return result

    def ldklds(self, result, row_idx, column_idx, value):
        if column_idx < len(result[5]):
            if int(str(self.date_cz_2)[8:10]) == int(result[5][column_idx]) and int(
                    str(self.date_cz_2)[5:7]) == self.now_month:
                result[row_idx][column_idx] = int(value) if value != "" else ""

    def bam_row_value_now(self, result, row_idx, column_idx, depth, date_list):
        for dateList in date_list:
            # dateList: [наименование, УП, дата_строкой]
            if len(dateList) < 3:
                continue

            date_str = dateList[-1]  # последний элемент — дата
            if not date_str or not isinstance(date_str, str):
                continue

            try:
                # Формат даты: YYYY.MM.DD или YYYY-MM-DD
                day = int(date_str[8:10])
                month = int(date_str[5:7].replace(".", ""))

                if day == int(result[5][column_idx]) and month == self.now_month:
                    result[row_idx][column_idx] = dateList[depth]
                    break
            except (ValueError, IndexError) as e:
                print(f"ERROR: bam_row_value_now parsing date '{date_str}': {e}")
                continue

    def jp_row_value_now(self, result, column_idx, row_idx, column_name_data: str = "Задач"):
        cdsount = 0

        if column_idx < len(result[5]):
            if result[5][column_idx] == int(self.mask_date[8:10]):
                for r_idx, row in enumerate(self.data_jp, 1):
                    for value in self.data_jp[row]:
                        if r_idx == 2:
                            if int(self.mask_date[5:7].replace("-", "")) == self.now_month and cdsount == 1:
                                result[row_idx][column_idx] = self.data_jp[0].get(column_name_data, "")
                                break
                            else:
                                cdsount = 1

        if column_name_data == "Вып. Задач":
            for row_jp_num, row_jp in self.data_jp.items():
                if row_jp_num > 3 and pd.isna(row_jp.get('Unnamed: 0', '')) and not pd.isna(row_jp.get('Задач', '')):
                    if row_jp.get('Задач', '')[5:7] == str(dt.strptime(self.mask_date[:10], '%Y-%m-%d'))[5:7]:
                        result[10][int(row_jp.get('Задач', '')[8:10]) + 7] = row_jp.get('Переводов', '')

    def main(self):
        result = self.create_calendar()

        self.chapter_gen(result, "ЖП", "Всего ЖП", "Вып. Задач")
        self.chapter_gen(result, "ФОЦ", "Выполнено за день, ДСЕ", "Выполнено за день, УП")
        self.chapter_gen(result, "ТОЦ", "Выполнено за день, ДСЕ", "Выполнено за день, УП")
        self.chapter_gen(result, "ПОЦ", "Выполнено за день, ДСЕ", "Выполнено за день, УП")
        self.chapter_gen(result, "СЗ", "Итого", "ФОЦ", "ТОЦ", "ПОЦ")

        # === Определяем текущую секцию и собираем данные ===
        current_section = "ФОЦ"  # начальная секция по умолчанию

        last_data = '123'
        for key, key_row in self.data_bam.items():
            # Определяем секцию по заголовкам
            section_marker = key_row.get("ФОЦ", "")
            if isinstance(section_marker, str):
                if "ТОЦ" in section_marker:
                    current_section = "ТОЦ"
                    continue  # пропускаем строку заголовка
                elif "ПОЦ" in section_marker:
                    current_section = "ПОЦ"
                    continue  # пропускаем строку заголовка

            # Пропускаем пустые строки
            if not self._is_data_row(key_row) and not self._is_total_row(key_row):
                continue

            # Формируем date_list: [наименование, УП, дата]
            name = key_row.get("Наименование", "")

            up = key_row.get("УП", "")
            date_str = self._extract_date_from_row(key_row)

            # Пропускаем строки без даты
            if not date_str:
                date_str = last_data
            else:
                last_data = date_str
            if isinstance(name, str) or pd.isna(name):
                continue


            # Приводим типы
            if isinstance(name, (int, float)):
                name = int(name) if name == int(name) else name
            if isinstance(up, (int, float)):
                up = int(up) if up == int(up) else up

            date_list = [name, up, date_str]

            if current_section == "ФОЦ":
                self.foc_date_list.append(date_list)
            elif current_section == "ТОЦ":
                self.toc_date_list.append(date_list)
            elif current_section == "ПОЦ":
                self.poc_date_list.append(date_list)

        # === Обработка data_cz (СЗ) ===
        foc_search = ""
        for key_row in self.data_cz.values():
            for key_col in key_row.items():
                if "Итого" in key_col:
                    foc_search = "Итого"
                elif "ФОЦ" in key_col:
                    foc_search = "ФОЦ"
                elif "ТОЦ" in key_col:
                    foc_search = "ТОЦ"
                elif "ПОЦ" in key_col:
                    foc_search = "ПОЦ"
                break

            if key_col[0] == "Дата:":
                for key_s in key_row.keys():
                    if len(str(key_s)) > 9:
                        self.date_cz_2 = key_s
                        break
            if foc_search == "Итого":
                self.egogo = int(key_row.get(self.date_cz_2, "")) if not pd.isna(key_row.get(self.date_cz_2, "")) else 0
            if foc_search == "ФОЦ":
                self.foc = int(key_row.get(self.date_cz_2, "")) if not pd.isna(key_row.get(self.date_cz_2, "")) else 0
            if foc_search == "ТОЦ":
                self.toc = int(key_row.get(self.date_cz_2, "")) if not pd.isna(key_row.get(self.date_cz_2, "")) else 0
            if foc_search == "ПОЦ":
                poc_val = key_row.get(self.date_cz_2, "")
                self.poc = int(poc_val) if not pd.isna(poc_val) and poc_val != "" else 0

        self.filling_data(result)

        """Проверка какой сейчас месяц, для сохранения данных в конфиге и само сохранение"""

        result.append(["" for i in range(59)])
        result[-1].append("V2.01.color")

        if not self.data_ge is None:
            if int(str(list(self.data_ge.get(0).keys())[0])[5:7].replace("-", "")) == int(
                    self.mask_date[5:7].replace("-", "")):
                for r_idx, row in enumerate(self.data_ge):
                    for c_idx, value in enumerate(self.data_ge.get(r_idx, "")):
                        if len(result) > r_idx:
                            if len(result[r_idx]) > c_idx:
                                if c_idx < 3 or r_idx < 3:
                                    continue
                                if result[r_idx][c_idx] == "" and not pd.isna(
                                        self.data_ge[r_idx - 1].get(f"Unnamed: {c_idx}", "")):
                                    result[r_idx][c_idx] = self.data_ge[r_idx - 1].get(f"Unnamed: {c_idx}", "")

        self.config.set_ge_last_days(self.titel_mounth_now, self.month_data_last_3_day(result))

        last_day_in_mounth: dict = self.config.get_ge_last_days()

        if self.titel_mounth_last in last_day_in_mounth.keys():
            data: dict = last_day_in_mounth[self.titel_mounth_last]

            for row_idx in range(len(result)):
                for column_idx in range(len(result[row_idx])):
                    if 4 < column_idx < 8:
                        for key_data in data.keys():
                            if row_idx == 9 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ЖП", "").get("Всего ЖП", "")
                            if row_idx == 10 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ЖП", "").get("Выполнено ЖП", "")
                            if row_idx == 14 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ФОЦ", "").get("Выполнено за день, ДСЕ", "")
                            if row_idx == 15 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ФОЦ", "").get("Выполнено за день, УП", "")
                            if row_idx == 19 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ТОЦ", "").get("Выполнено за день, ДСЕ", "")
                            if row_idx == 20 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ТОЦ", "").get("Выполнено за день, УП", "")
                            if row_idx == 24 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ПОЦ", "").get("Выполнено за день, ДСЕ", "")
                            if row_idx == 25 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("ПОЦ", "").get("Выполнено за день, УП", "")
                            if row_idx == 29 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("СЗ", "").get("Итого", "")
                            if row_idx == 30 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("СЗ", "").get("ФОЦ", "")
                            if row_idx == 31 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("СЗ", "").get("ТОЦ", "")
                            if row_idx == 32 and result[5][column_idx] == int(key_data):
                                result[row_idx][column_idx] = data[key_data].get("СЗ", "").get("ПОЦ", "")

        return result

    def month_data_last_3_day(self, result):
        mounth_data1 = {
            "ЖП": {"Всего ЖП": "", "Выполнено ЖП": ""},
            "ФОЦ": {"Выполнено за день, ДСЕ": "", "Выполнено за день, УП": ""},
            "ТОЦ": {"Выполнено за день, ДСЕ": "", "Выполнено за день, УП": ""},
            "ПОЦ": {"Выполнено за день, ДСЕ": "", "Выполнено за день, УП": ""},
            "СЗ": {"Итого": "", "ФОЦ": "", "ТОЦ": "", "ПОЦ": ""}
        }

        # Берём последние 3 даты
        date_for_save_column = [result[5][-1], result[5][-2], result[5][-3]]
        mounth_data_day = {}

        for col_idx in range(len(result[5])):
            if result[5][col_idx] in date_for_save_column:
                # Используем deepcopy, чтобы избежать ссылок на вложенные словари
                current_data = copy.deepcopy(mounth_data1)

                # Обработка строк в текущем столбце
                for row_idx in [9, 10, 14, 15, 19, 20, 24, 25, 29, 30, 31, 32]:
                    if row_idx >= len(result):
                        continue
                    if col_idx < len(result[row_idx]):
                        value = result[row_idx][col_idx]

                        # ЖП
                        if row_idx == 9:
                            current_data["ЖП"]["Всего ЖП"] = value
                        elif row_idx == 10:
                            current_data["ЖП"]["Выполнено ЖП"] = value

                        # ФОЦ
                        elif row_idx == 14:
                            current_data["ФОЦ"]["Выполнено за день, ДСЕ"] = value
                        elif row_idx == 15:
                            current_data["ФОЦ"]["Выполнено за день, УП"] = value

                        # ТОЦ
                        elif row_idx == 19:
                            current_data["ТОЦ"]["Выполнено за день, ДСЕ"] = value
                        elif row_idx == 20:
                            current_data["ТОЦ"]["Выполнено за день, УП"] = value

                        # ПОЦ
                        elif row_idx == 24:
                            current_data["ПОЦ"]["Выполнено за день, ДСЕ"] = value
                        elif row_idx == 25:
                            current_data["ПОЦ"]["Выполнено за день, УП"] = value

                        # СЗ
                        elif row_idx == 29:
                            current_data["СЗ"]["Итого"] = value
                        elif row_idx == 30:
                            current_data["СЗ"]["ФОЦ"] = value
                        elif row_idx == 31:
                            current_data["СЗ"]["ТОЦ"] = value
                        elif row_idx == 32:
                            current_data["СЗ"]["ПОЦ"] = value

                # Сохраняем результат для текущей даты
                mounth_data_day[result[5][col_idx]] = current_data

        return mounth_data_day


if __name__ == "__main__":
    config = JsonConfig()

    ge_prog = GeneProg()
    jhjhhgjhjg=ge_prog.main()
    print('====================================================================================================')
    for i in jhjhhgjhjg:
        print(i)