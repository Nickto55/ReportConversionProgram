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

        self.data_jp: dict
        self.data_cz = None
        self.data_bam = None

        self.data_ge = None

        self.get_datas_excel()

    def get_datas_excel(self, start_row=1, start_col=1):
        self.data_jp, self.data_cz, self.data_bam = self.search_in_sheet.sheet_name_list()

        # for r_idx, row in enumerate(self.data_jp.values, start=start_row):
        #     for c_idx, value in enumerate(row, start=start_col):
        #         print(value, end=" ")
        #     print()

    def main(self):
        result = []
        result_row1 = ["" for i in range(38)]
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

        result_row = ["", "", "всего ЖП"]

        for i in range(3, 38):
            if result[-4][i] == int(str(dt.now())[8:10]):
                for r_idx, row in enumerate(self.data_jp, 1):
                    for value in self.data_jp[row]:
                        if r_idx == 2:
                            result_row.append(self.data_jp[0].get("Задач", ""))
                            break


            else:
                result_row.append("")




        result.append(result_row)


        result_row = ["", "", "выполнено ЖП","", ""]
        # print(last_month, now_month)
        for i in range(33):
            result_row.append("")



        for key_row in self.data_jp.keys():
            varibel_break = False
            for c_idx in self.data_jp[key_row].keys():
                if key_row<3:
                    continue
                if pd.isna(self.data_jp[key_row].get("Задач", "")):
                    break

                try:
                    int(self.data_jp[key_row].get("Переводов", ""))
                except:
                    continue


                if  varibel_break:
                    result_row.append("")
                    break
                value = self.data_jp[key_row].get("Задач", "")
                if len(str(value))>7:
                    if int(str(value)[8:10]) in result[-5]  and int(str(value)[5:7]) == now_month:
                        result_row[int(str(value)[8:10])+4+3]=(self.data_jp[key_row].get("Переводов", ""))
                        varibel_break =True
                        break
                    elif int(str(value)[8:10]) in result[-5] and int(str(value)[5:7]) == last_month and result[-5].index(int(str(value)[8:10])) < 8:
                        result_row[int(str(value)[8:10])-29+5]=(self.data_jp[key_row].get("Переводов", ""))
                        varibel_break =True
                        break


        # print(result_row)
        result.append(result_row)


        for i in range(3): result.append(result_row1.copy())
        result[-1][0] = "ФОЦ"
        count_titel = 0

        result_row = ["", "", "выполнено за день, ДСЕ"]
        for key_row in self.data_bam.values():
            for key_col in key_row.values():
                print(key_col, end=", ")
            print("\n")




        # for key_row in self.data_jp.keys():
        #     print(self.data_jp[key_row].get(5, ""))



        # for i in result:
        #     print(i)
        result.append(result_row)


        return  result


if __name__ == "__main__":
    run = GeneProg()
    config = JsonConfig()
    excelPr = ExcelPrint.ExcelWriter(config.getJPPathFile_output())

    excelPr.write_to_sheet(run.main(), "Test")
