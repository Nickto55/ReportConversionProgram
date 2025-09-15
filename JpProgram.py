from datetime import date, timedelta

import pandas as pd
import plyer


import ExcelPrint
import JsonWork
import Search

import Body


def get_dates(len_date: int):
    """
    Возвращает список из трех дат: позавчерашней, вчерашней и сегодняшней.

    """
    result_list = []
    today = date.today()

    for i in range(0, len_date):
        date_form = today - timedelta(days=i)
        format_str = "%Y-%m-%d"
        result_list.append(date.strftime(date_form, format_str))

    return result_list


def send_notification(title, message, settime=15, file_path=""):
    plyer.notification.notify(
        title=title,
        message=message,
        app_name="Good Morning (JP)",
        timeout=settime
    )


class JpMain:
    def __init__(self, count_prog = None, root = None):
        self.config = JsonWork.JsonConfig()
        self.root = root

        if not pd.isna(root):
            self.pb = Body.Main_gui(root)
            if pd.isna(count_prog):
                self.pb.change_value_progress_bar_var(8)


        search = Search.SearchJP()
        self.data = search.get_dict_all_data()
        # self.first_t()
        # self.second_t()

    def first_t(self):
        result = []

        count_removed = 0
        count_conversion = 0

        pos_removed = 0
        pos_conversion = 0
        pos_removed_date = 0
        pos_add_date = 0
        pos_dse = 0

        # Определяем позиции колонок
        for i in (self.data.get(0, "")):
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Question removed Full name"):
                pos_removed = i
            if str(self.data.get(0, "").get(i, "")) in str(self.config.getJPColumnName("Table of contents: Translation")):
                pos_conversion = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName("Table of contents: Date removed"):
                pos_removed_date = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName("Table of contents: Date"):
                pos_add_date = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName("Table of contents: DCE"):
                pos_dse = i

        # Подсчитываем общие значения
        for i in self.data:
            data = self.data.get(i, "")
            n = data.get(pos_removed, "")
            if pd.isna(n):
                count_removed += 1
                n1 = data.get(pos_conversion, "")
                if n1 == "+":
                    count_conversion += 1

        # Получаем список дат
        date_list = get_dates(int(self.config.getJPColumnName("Table of contents: List_date")))

        # Словарь для хранения результатов по датам
        date_stats = {}

        # Проходим по данным и собираем статистику по датам
        for i in self.data:
            data = self.data.get(i, "")
            date_dse_value = data.get(pos_dse, "")
            removed_date_culum_value = data.get(pos_removed_date, "")

            if not pd.isna(date_dse_value) and not pd.isna(removed_date_culum_value):
                # Обрезаем до нужной длины (как в date_list)
                truncated_date = str(removed_date_culum_value)[:len(date_list[0])]

                if truncated_date in date_list:
                    if truncated_date not in date_stats:
                        date_stats[truncated_date] = {
                            "count_removed": 0,
                            "count_conversion": 0,
                            "count_date_add": 0,
                            "count": 0
                        }

                    date_stats[truncated_date]["count"] += 1

                    # Проверяем, если вопрос удален
                    if pd.isna(data.get(pos_removed, "")):
                        date_stats[truncated_date]["count_removed"] += 1
                        if data.get(pos_conversion, "") == "+":
                            date_stats[truncated_date]["count_conversion"] += 1
        for i in self.data:
            data = self.data.get(i, "")
            date_dse_value = data.get(pos_dse, "")
            date_add = data.get(pos_add_date, "")
            if not pd.isna(date_dse_value):
                truncated_date = str(date_add)[:len(date_list[0])]
                if truncated_date in date_list:
                    date_stats[truncated_date]["count_date_add"] += 1


        result.append(["Задач", "Переводов",""])
        result.append([count_removed, count_conversion])

        result.append(["", "", ""])
        result.append(["", "", ""])
        result.append(["Дата", "Вып. Задач", "Доб. Задач"])

        # Формируем результат
        for date in date_list:
            if date in date_stats:
                stats = date_stats[date]
                result.append([
                    date,
                    stats["count"],
                    stats["count_date_add"]
                ])
            else:
                result.append([date, 0,0])

        # # Выводим результат
        # for row in result:
        #     print(row)
        return result

    def second_t(self):
        result = [[self.config.getJPColumnName("Table of contents: Namber"),
                   self.config.getJPColumnName("Table of contents: Date"),
                   self.config.getJPColumnName("Table of contents: DCE"),
                   self.config.getJPColumnName("Table of contents: Name"),
                   self.config.getJPColumnName("Table of contents: Translation"),
                   self.config.getJPColumnName("Table of contents: Question accepted Full name")]
                  ]

        pos_removed = 0
        pos_number = 0
        pos_date = 0
        pos_dce = 0
        pos_name = 0
        pos_accepted_name = 0
        pos_Translation = 0

        # Определяем позиции колонок
        for i in (self.data.get(0, "")):
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Question removed Full name"):
                pos_removed = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Namber"):
                pos_number = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Date"):
                pos_date = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: DCE"):
                pos_dce = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Name"):
                pos_name = i
            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Translation"):
                pos_Translation = i

            if self.data.get(0, "").get(i, "") == self.config.getJPColumnName(
                    "Table of contents: Question accepted Full name"):
                pos_accepted_name = i

        if pos_number == 0 or pos_date == 0 or pos_dce == 0 or pos_name == 0 or pos_accepted_name == 0 or pos_removed == 0 or pos_Translation == 0:
            return None

        for i in self.data:
            data = self.data.get(i, "")
            if pd.isna(data.get(pos_removed, "")):
                result.append(
                    [
                        data.get(pos_number, ""),
                        str(data.get(pos_date, ""))[:len("10.07.2025")],
                        data.get(pos_dce, ""),
                        data.get(pos_name, ""),
                        data.get(pos_Translation, ""),
                        data.get(pos_accepted_name, "")
                    ]
                )
        # for i in result:
        #     print(i)
        # print(len(result))
        return result

    def main(self):
        result = []
        result1 = self.first_t()
        result2 = self.second_t()


        """
        2 версия
        """
        for row in range(len(result1)):
            res_row = []
            for column in range(max(len(result1[0]), len(result2[0])) + 1):
                if column == 0:
                    res_row.append("")
                    continue
                column -= 1
                if column < len(result1[0]):
                    try:
                        res_row.append(result1[row][column])
                    except:
                        pass
                else:
                    res_row.append("")
            result.append(res_row)

        result.append(["" for i in range(max(len(result1[0]), len(result2[0])))])
        result.append(["" for i in range(max(len(result1[0]), len(result2[0])))])

        for row in range(len(result2)):
            res_row = []
            for column in range(max(len(result1[0])-1, len(result2[0]))):
                # if column == 0:
                #     res_row.append("")
                #     continue
                # column -= 1
                if column < len(result2[0]):
                    res_row.append(result2[row][column])
                else:
                    res_row.append("")
            result.append(res_row)

        """
        1 версия
        # for row in range(min(len(result1), len(result2))):
        #     result_row = []
        #     for column in range(sum_count_column + 1):
        #         if column != 0 and column != sum_count_column - count_column2 - 1 and column != sum_count_column - count_column2:
        #             column -= 1
        #             if column < count_column1:
        #                 result_row.append(result1[row][column])
        #             column -= 2
        #             if count_column1 <= column <= count_column1 + count_column2:
        #                 result_row.append(result2[row][column - count_column1])
        #         else:
        #             result_row.append("")
        #     result.append(result_row)
        # 
        # if count_column1 > count_column2:
        #     for row in range(len(result2), len(result1)):
        #         result_row = []
        #         for column in range(sum_count_column + 1):
        #             if column != 0 and column > count_column1 + 1:
        #                 if column < count_column1:
        #                     result_row.append(result1[row][column])
        #             else:
        #                 result_row.append("")
        #         result.append(result_row)
        # 
        # if count_column1 < count_column2:
        #     for row in range(len(result1), len(result2)):
        #         result_row = []
        #         for column in range(sum_count_column + 3):
        #             if column > count_column1 + 2:
        #                 column -= 2
        #                 if count_column1 <= column <= count_column1 + count_column2:
        #                     result_row.append(result2[row][column - count_column1 - 1])
        #             else:
        #                 result_row.append("")
        #         result.append(result_row)
        """
        # for i in result:
        #     print(i)
        return result


if __name__ == "__main__":
    run = JpMain()
    config = JsonWork.JsonConfig()
    excelPr = ExcelPrint.ExcelWriter(config.getJPPathFile_output())

    excelPr.write_to_sheet(run.main(), "ЖП")
    # send_notification("Программа завершена", "Программа завершена, проверте файл",16)
