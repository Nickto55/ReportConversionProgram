from datetime import timedelta, datetime as dt
from tkinter import messagebox

import ExcelPrint
from JsonWork import JsonConfig
from Search import SearchBam





def date_ref(date_n, len_date_value: int = 3, varibel: int = 0):
    if varibel:
        if len_date_value == 1:
            format_str = "%Y"
        elif len_date_value == 2:
            format_str = "%Y.%m"
        else:
            format_str = "%Y.%m.%d"
    else:
        if len_date_value == 1:
            format_str = "%Y"
        elif len_date_value == 2:
            format_str = "%Y-%m"
        else:
            format_str = "%Y-%m-%d"

    try:
        return dt.strftime(date_n, format_str)
    except:
        pass


class BamMain:
    def __init__(self,mask_date:str = str(dt.now())):
        self.config = JsonConfig()

        self.listes_excel = self.config.getBAMColumnName( "Table of contents: listes_excel", intOrlist=1)
        """Надо изменить на конфиг"""

        self.search = None
        self.listes_excel_last = ""
        self.mask_date = mask_date

        self.error_masage_var = False

    def get_dates(self, len_date: int, len_date_value: int = 3, year: bool = False):
        """
        Возвращает список из трех лет: позавчерашней, вчерашней и сегодняшней.

        """
        result_list = []
        today = dt.strptime(self.mask_date[:10], '%Y-%m-%d')

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
            result_list.append(dt.strftime(date_form, format_str))

        return result_list

    def headers_sort(self, headers):
        headers.sort()
        headers.remove("")
        result = []
        date_years = self.get_dates(int(self.config.getBAMColumnName("Table of contents: List_date")))
        for year_date in headers:
            if year_date in date_years:
                result.append(year_date)
        return result

    def headers_sort_full(self, headers):
        headers.sort()
        headers.remove("")
        return headers

    def result_creation_function(self, listes_excel):

        self.search = SearchBam(listes_excel)
        self.data = self.search.get_dict_all_data()

        self.headers = self.headers_sort(self.search.get_colum(self.config.getBAMColumnName("Table of contents: Date")))
        result = []
        result_row = []
        for row in self.data.values():
            for i in row.keys():
                result_row.append(i)
            break
        result_row.append("/::/")
        result.append(result_row)
        i_sort_last = ""
        try:
            i_sort_last = self.headers[-1]
        except:
            if not self.error_masage_var:
                messagebox.showinfo("Ошибка", "За указанный периуд нет данных. \nВ настройках программы БАМ увеличьте дни отображения")
                self.error_masage_var = True
                return ["ПУСТО"]


        row_for_count = 0
        row_count = 1
        count_dse = 0
        count_up = 0
        count_dse2 = 0
        count_up2 = 0

        for i_sort in self.headers:
            if i_sort_last != i_sort or self.listes_excel_last != listes_excel:
                if self.listes_excel_last != listes_excel:
                    self.listes_excel_last = listes_excel
                result.append(["/../","","","","","","","",""])
                row_count += 1
                row_for_count = row_count-1
            for row in self.data.values():
                try:
                    date_row = date_ref(row.get(self.config.getBAMColumnName("Table of contents: Date"), ""))
                except:
                    date_row = row.get(self.config.getBAMColumnName("Table of contents: Date"), "")

                if date_row == i_sort and date_row != "":
                    row_count += 1
                    row[self.config.getBAMColumnName("Table of contents: Date")] = date_ref(
                        row.get(self.config.getBAMColumnName("Table of contents: Date"), ""), varibel=1)
                    row = list(row.values())[1:]
                    if int(row[5]) != 0:
                        count_up += row[5]
                        count_dse +=1
                    row.insert(0, "")
                    result.append(row)
            i_sort_last = i_sort
            result[row_for_count].insert(3,  count_dse-count_dse2)
            result[row_for_count].insert(6,  count_up-count_up2)
            result[row_for_count].append(result[row_for_count+1][7])
            count_dse2 = count_dse

            count_up2 = count_up

        return result

    def main(self):
        """Главное тело"""
        result = []
        for listes_excel in self.listes_excel:
            res = self.result_creation_function(listes_excel)
            res_0 = res[0]
            res.pop(0)
            i:list
            res.reverse()
            res.insert(0,res_0)
            for i in res:
                if "№" in i:
                    i.remove("№")
                    i.insert(0,listes_excel)
                result.append(i)
            result.append([])
        return result


if __name__ == '__main__':
    run = BamMain()
    config = JsonConfig()
    excelPr = ExcelPrint.ExcelWriter(config.getJPPathFile_output(), min_prog="BAM")

    excelPr.write_to_sheet(run.main(), "Бам по УП")
