import sys
import os
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.handlings.handling_json import JsonConfig
from scripts.excel_enter import ExcelWriter
from scripts.handlings.handling_classes import SearchTrack


class TrackMain:
    def __init__(self):
        self.data = []

    def main(self):
        self.data = [
            [
                'РЦ'
                ,'Изделие'
                ,'ДСЕ'
                ,'Наименование'
                ,'Дата'
                ,'ФИО'
                ,'Комментарий'
                ,''
            ]
        ]
        for name_sheet in ["ФОЦ", "ПОЦ", "ТОЦ"]:
            app = SearchTrack(name_sheet)
            for number_row_get_data, get_data in app.get_dict_all_data().items():
                if name_sheet == "ФОЦ":
                    dse_key = "Наименование"
                    name_key = "Наименование.1"
                elif name_sheet == "ПОЦ":
                    dse_key = "Прутковые автоматы"
                    name_key = "Наименование"
                elif name_sheet == "ТОЦ":
                    dse_key = "Токарные ЧПУ"
                    name_key = "Наименование"
                else:
                    print(f"Произошла ошибка с наимен листа для отслеживания")
                    return None

                if not pd.isna(get_data.get('Дата', '')):
                    data_value = str(str(get_data.get('Дата', ''))[:10] )
                else:
                    data_value = " "
                self.data.append(
                    [
                        get_data.get('__sheet_origin__', '')
                        ,get_data.get('Изделие', '')
                        ,get_data.get(dse_key, '')
                        ,get_data.get(name_key, '')
                        ,data_value
                        ,get_data.get('Ф. И. О.', '')
                        ,get_data.get('Комментарий', '')
                        ,get_data.get('__color_D__', '')
                    ]
                )

        return self.data
        # return result_data


if __name__ == "__main__":
    run = TrackMain()
    config = JsonConfig()
    excelPr = ExcelWriter(config.getJPPathFile_output(), min_prog='Tr')
    excelPr.write_to_sheet(run.main(), "Отслеживание")
