import calendar
from datetime import datetime as dt

import pandas as pd

import ExcelPrint
from JsonWork import JsonConfig
from Search import SearchGe


class GeneProg:
    def __init__(self):
        self.config = JsonConfig()
        self.search_in_sheet = SearchGe()

        self.data_jp = None
        self.data_cz = None
        self.data_bam = None

        self.data_ge = None

        self.get_datas_excel()

    def get_datas_excel(self, start_row=1, start_col=1):
        self.data_jp, self.data_cz, self.data_bam, self.data_ge = self.search_in_sheet.sheet_name_list()

        # for r_idx, row in enumerate(self.data_jp.values, start=start_row):
        #     for c_idx, value in enumerate(row, start=start_col):
        #         print(value, end=" ")
        #     print()

    def main(self):
        result = []
        result_row1 = ["" for i in range(38)]
        result_row2 = ["" for i in range(8)]
        result_row2.append(r"\..../")
        for i in range(29):
            result_row2.append(r"\.../")
        for i in range(4): result.append(result_row1.copy())
        date_str = str(dt.now())[:10]
        date_str_last = str(dt.now())[:10]
        month_number = int(date_str_last[5:7]) - 1

        date_str_last = dt(2023, month_number, 1)

        result[0][0] = date_str
        date_object = dt.strptime(date_str, '%Y-%m-%d')
        titel_mounth_last = (date_str_last.strftime('%B'))
        titel_mounth_now = date_object.strftime('%B')
        count_day_now_mounth = list(calendar.monthrange(int(date_str[:4]), int(date_str[5:7])))[-1]
        last_month = int(date_str[5:7]) - 1
        now_month = int(date_str[5:7])
        if last_month == 0:
            last_month = 12
        count_day_last_mounth = list(calendar.monthrange(int(date_str[:4]), last_month))[-1]

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

        """Календарь окончен"""

        for i in range(3):
            result.append(result_row1.copy())

        result[-1][0] = "ЖП"

        result_row = ["", "", "Всего ЖП"]
        cdsount = 0

        for i in range(3, 38):
            if result[-4][i] == int(str(dt.now())[8:10]):
                for r_idx, row in enumerate(self.data_jp, 1):
                    for value in self.data_jp[row]:
                        if r_idx == 2:
                            if int(str(dt.now())[5:7])  == now_month and cdsount==1:
                                print("hjodsjhdhklllllllsfjlksdhldshl")
                                result_row.append(self.data_jp[0].get("Задач", ""))
                            else: cdsount=1
                            break


            else:
                result_row.append("")

        result.append(result_row)

        result_row = ["", "", "Выполнено ЖП", "", ""]
        # print(last_month, now_month)
        for i in range(33):
            result_row.append("")

        for key_row in self.data_jp.keys():
            varibel_break = False
            for c_idx in self.data_jp[key_row].keys():
                if key_row < 3:
                    continue
                if pd.isna(self.data_jp[key_row].get("Задач", "")):
                    break

                try:
                    int(self.data_jp[key_row].get("Переводов", ""))
                except:
                    continue

                if varibel_break:
                    result_row.append("")
                    break
                value = self.data_jp[key_row].get("Задач", "")
                if len(str(value)) > 7:
                    if int(str(value)[8:10]) in result[-5] and int(str(value)[5:7]) == now_month:
                        result_row[int(str(value)[8:10]) + 4 + 3] = (self.data_jp[key_row].get("Переводов", ""))
                        varibel_break = True
                        break
                    elif int(str(value)[8:10]) in result[-5] and int(str(value)[5:7]) == last_month and result[
                        -5].index(int(str(value)[8:10])) < 8:
                        result_row[int(str(value)[8:10]) - 29 + 5] = (self.data_jp[key_row].get("Переводов", ""))
                        varibel_break = True
                        break

        # print(result_row)
        result.append(result_row)

        for i in range(2): result.append(result_row2.copy())
        for i in range(1): result.append(result_row1.copy())
        result[-1][0] = "ФОЦ"
        count_titel = 0

        foc_search = "ФОЦ"
        foc_date_list = []
        toc_date_list = []
        poc_date_list = []

        result_row = ["", "", "Выполнено за день, ДСЕ", "", ""]

        for key_row in self.data_bam.values():
            for key_col in key_row.items():

                if "ТОЦ" in key_col:
                    foc_search = "ТОЦ"
                elif "ПОЦ" in key_col:
                    foc_search = "ПОЦ"
                break
            if foc_search == "ТОЦ":
                if pd.isna(key_row.get("Дата", "")) and not pd.isna(key_row.get("Фрезерные ЧПУ", "")):
                    toc_date_list.append([key_row.get("Фрезерные ЧПУ", ""), key_row.get("УП", ""), key_row.get("/::/", "")])
                elif pd.isna(key_row.get("Дата", "")):
                    continue

            if foc_search == "ПОЦ":
                if pd.isna(key_row.get("Дата", "")) and not pd.isna(key_row.get("Фрезерные ЧПУ", "")):
                    poc_date_list.append([key_row.get("Фрезерные ЧПУ", ""), key_row.get("УП", ""), key_row.get("/::/", "")])
                elif pd.isna(key_row.get("Дата", "")):
                    continue



            if foc_search == "ФОЦ":
                if pd.isna(key_row.get("Дата", "")) and not pd.isna(key_row.get("Фрезерные ЧПУ", "")):
                    foc_date_list.append([key_row.get("Фрезерные ЧПУ", ""), key_row.get("УП", ""), key_row.get("/::/", "")])
                elif pd.isna(key_row.get("Дата", "")):
                    continue

        for column in range(33):
            """ФОЦ"""
            dateList_cat = []
            for dateList in foc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))
                    if int(dateList[-1][8:10]) == int(result[-9][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[0])

                        break
                    if int(dateList[-1][8:10]) == int(result[-9][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[0])
                        break
                except:
                    pass
            if not int(result[-9][column + 5]) in dateList_cat:
                result_row.append("")

        result.append(result_row)
        result_row = ["", "", "Выполнено за день, УП", "", ""]
        for column in range(33):
            """ФОЦ"""
            dateList_cat = []
            for dateList in foc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))
                    if int(dateList[-1][8:10]) == int(result[-10][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[1])
                        break
                    if int(dateList[-1][8:10]) == int(result[-10][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[1])
                        break
                except:
                    pass
            if not int(result[-10][column + 5]) in dateList_cat:
                result_row.append("")

        result.append(result_row)

        for i in range(2): result.append(result_row2.copy())
        for i in range(1): result.append(result_row1.copy())
        result[-1][0] = "ТОЦ"

        result_row = ["", "", "Выполнено за день, ДСЕ", "", ""]

        for column in range(33):
            """ТОЦ"""
            dateList_cat = []
            for dateList in toc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))

                    if int(dateList[-1][8:10]) == int(result[-14][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[0])
                        break
                    if int(dateList[-1][8:10]) == int(result[-14][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[0])
                        break
                except:
                    pass
            if not int(result[-14][column + 5]) in dateList_cat:
                result_row.append("")

        result.append(result_row)
        result_row = ["", "", "Выполнено за день, УП", "", ""]
        for column in range(33):
            """ТОЦ"""
            dateList_cat = []
            for dateList in toc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))
                    if int(dateList[-1][8:10]) == int(result[-15][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[1])
                        break
                    if int(dateList[-1][8:10]) == int(result[-15][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[1])
                        break

                except:
                    pass
            if not int(result[-15][column + 5]) in dateList_cat:
                result_row.append("")

        result.append(result_row)

        for i in range(2): result.append(result_row2.copy())
        for i in range(1): result.append(result_row1.copy())
        result[-1][0] = "ПОЦ"

        result_row = ["", "", "Выполнено за день, ДСЕ", "", ""]

        for column in range(33):
            """ПОЦ"""
            dateList_cat = []
            for dateList in poc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))
                    if int(dateList[-1][8:10]) == int(result[-19][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[0])

                        break
                    if int(dateList[-1][8:10]) == int(result[-19][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[0])
                        break
                except:
                    pass
            if not int(result[-19][column + 5]) in dateList_cat:
                result_row.append("")

        result.append(result_row)
        result_row = ["", "", "Выполнено за день, УП", "", ""]
        for column in range(33):
            """ПОЦ"""
            dateList_cat = []
            for dateList in poc_date_list:
                try:
                    dateList_cat.append(int(dateList[-1][8:10]))
                    if int(dateList[-1][8:10]) == int(result[-20][column + 5]) and int(dateList[-1][5:7]) == now_month:
                        result_row.append(dateList[1])

                        break
                    if int(dateList[-1][8:10]) == int(result[-20][column + 5]) and int(
                            dateList[-1][5:7]) == last_month and len(result_row) < 10:
                        result_row.append(dateList[1])
                        break
                except:
                    pass
            if not int(result[-20][column + 5]) in dateList_cat:
                result_row.append("")
        result.append(result_row)

        for i in range(2): result.append(result_row2.copy())
        for i in range(1): result.append(result_row1.copy())
        result[-1][0] = "СЗ"

        foc_search = "Дата:"
        date_cz = ""

        egogo = 0
        foc = 0
        toc = 0
        poc = 0

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
            if foc_search == "Дата:":
                for key_s in key_row.keys():
                    if len(str(key_s)) > 9:
                        date_cz = key_s
                        break
            if foc_search == "Итого":
                egogo = int(key_row.get(date_cz, ""))
            if foc_search == "ФОЦ":
                foc = int(key_row.get(date_cz, ""))
            if foc_search == "ТОЦ":
                toc = int(key_row.get(date_cz, ""))
            if foc_search == "ПОЦ":
                poc = (key_row.get(date_cz, ""))

        result_row = ["", "", "Итого", "", ""]

        countjsa = 0
        for column in range(33):
            if int(str(date_cz)[8:10]) == int(result[-24][column + 5]) and int(str(date_cz)[5:7]) == now_month:
                if countjsa == 1:
                    result_row.append(int(egogo))
                else:
                    countjsa = 1
                    continue
                break
            if not int(result[-24][column + 5]) == int(str(date_cz)[8:10]):
                result_row.append("")

        result.append(result_row)
        result_row = ["", "", "ФОЦ", "", ""]
        countjsa = 0
        for column in range(33):
            if int(str(date_cz)[8:10]) == int(result[-25][column + 5]) and int(str(date_cz)[5:7]) == now_month:
                if countjsa == 1:
                    result_row.append(foc)
                else:
                    countjsa = 1
                    continue
                break
            if not int(result[-25][column + 5]) == int(str(date_cz)[8:10]):
                result_row.append("")

        countjsa = 0
        result.append(result_row)
        result_row = ["", "", "ТОЦ", "", ""]
        for column in range(36):
            if int(str(date_cz)[8:10]) == int(result[-26][column + 5]) and int(str(date_cz)[5:7]) == now_month:
                if countjsa == 1:
                    result_row.append(toc)
                else:
                    countjsa = 1
                    continue
                break
            result_row.append("")

        result.append(result_row)
        result_row = ["", "", "ПОЦ", "", ""]
        countjsa = 0
        for column in range(36):
            if int(str(date_cz)[8:10]) == int(result[-27][column + 5]) and int(str(date_cz)[5:7]) == now_month:
                if countjsa == 1:
                    result_row.append(poc)
                else:
                    countjsa = 1
                    continue
                break
            if not int(result[-27][column + 5]) == int(str(date_cz)[8:10]):
                result_row.append("")

        result.append(result_row)

        for i in range(2): result.append(result_row1.copy())
        result.append(["" for i in range(59)])
        result[-1].append("V2.01.color")

        if not self.data_ge is None:
            for r_idx, row in enumerate(self.data_ge):
                for c_idx, value in enumerate(self.data_ge.get(r_idx, "")):
                    if len(result) > r_idx:
                        if len(result[r_idx]) > c_idx:
                            if c_idx < 3 or r_idx < 3:
                                continue
                            if result[r_idx][c_idx] == "" and not pd.isna(
                                    self.data_ge[r_idx - 1].get(f"Unnamed: {c_idx}", "")):
                                result[r_idx][c_idx] = self.data_ge[r_idx - 1].get(f"Unnamed: {c_idx}", "")

        return result


if __name__ == "__main__":
    run = GeneProg()
    config = JsonConfig()

    # for i in run.main():
    #     print(i)
    excelPr = ExcelPrint.ExcelWriter(config.getJPPathFile_output(), min_prog="Ge")

    excelPr.write_to_sheet(run.main(), "Общая информация")
