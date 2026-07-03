import sys
import os

from typing import Optional, List, Dict, Any

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import scripts.handlings.handling_json as handling_json
from scripts.excel_reader import ExcelReader, MultiSheetReader



class SearchJP:
    """
    Класс для работы с JP-файлом.
    """
    
    def __init__(self):
        self.config = handling_json.JsonConfig()
        self.reader = ExcelReader(
            file_path=self.config.getJPPathFile_input()
        )
    
    def get_dict_all_data(self) -> Dict[int, Dict[str, Any]]:
        """Возвращает все данные в виде словаря."""
        return self.reader.get_dict_all_data()
    
    def load_excel(self) -> None:
        """Перезагружает Excel-файл."""
        self.reader.load_excel()


class SearchCz:
    """
    Класс для работы с CZ-файлом.
    """
    
    def __init__(self):
        self.config = handling_json.JsonConfig()
        self.reader = ExcelReader(
            file_path=self.config.getCzPathFile_input(),
            sheet_name=self.config.getCzColumnName("Table of contents: List")
        )
        self.reader.headers = self.reader.get_headers()
    
    def get_dict_all_data(self) -> Dict[int, Dict[str, Any]]:
        """Возвращает все данные в виде словаря."""
        return self.reader.get_dict_all_data()
    
    def get_headers(self) -> List[str]:
        """Возвращает заголовки колонок."""
        return self.reader.get_headers()
    
    def filter_and_save_columns(self, columns_to_save: List[str]) -> None:
        """Фильтрует и сохраняет указанные колонки."""
        self.reader.filter_and_save_columns(columns_to_save)
    
    def get_filtered_data(self) -> Optional[Dict[int, Dict[str, Any]]]:
        """Возвращает отфильтрованные данные."""
        return self.reader.get_filtered_data()
    
    def get_colum(self, get_colum: str, foc_mode: float = 0) -> List[str]:
        """
        Возвращает уникальные значения из указанной колонки.
        
        :param get_colum: имя колонки
        :param foc_mode: режим фильтрации (0 = выключен, !=0 = включен)
        """
        # Условие пропуска для FOC-режима:
        # пропускаем если (Done НЕ пустой ИЛИ Close НЕ пустой) ИЛИ DCE пустой
        def skip_condition(row_data: Dict[str, Any]) -> bool:
            done_col = self.config.getCzColumnName("Table of contents: Done")
            close_col = self.config.getCzColumnName("Table of contents: Close")
            dse_col = self.config.getCzColumnName("Table of contents: DCE")
            
            done_empty = self.reader.is_empty(row_data.get(done_col))
            close_empty = self.reader.is_empty(row_data.get(close_col))
            dse_empty = self.reader.is_empty(row_data.get(dse_col))
            
            # Оригинальная логика: пропускаем если НЕ (done_empty И close_empty И НЕ dse_empty)
            # Т.е. пропускаем если (done НЕ пуст ИЛИ close НЕ пуст) ИЛИ dse пуст
            return not (done_empty and close_empty and not dse_empty)
        
        return self.reader.get_column_values(
            get_column=get_colum,
            foc_mode=foc_mode != 0,
            skip_condition=skip_condition if foc_mode != 0 else None
        )


class SearchGe:
    """
    Класс для работы с GE-файлом (несколько листов).
    """
    
    def __init__(self):
        self.config = handling_json.JsonConfig()
        self.file_path = self.config.getJPColumnName("Path for output excel")
        self.reader = MultiSheetReader(self.file_path)
        
        self.data_jp: Optional[Dict[int, Dict[str, Any]]] = None
        self.data_cz: Optional[Dict[int, Dict[str, Any]]] = None
        self.data_bam: Optional[Dict[int, Dict[str, Any]]] = None
        self.data_ge: Optional[Dict[int, Dict[str, Any]]] = None
    
    def sheet_name_list(self) -> tuple:
        """
        Загружает все листы и возвращает данные в виде кортежа словарей.
        
        :return: (data_jp, data_cz, data_bam, data_ge)
        """
        sheet_name_list = self.config.getSheetNameList()
        self.reader.load_sheets(sheet_name_list)
        
        # Сопоставляем листы с атрибутами по порядку
        sheet_mapping = {
            0: 'data_jp',
            1: 'data_cz', 
            2: 'data_bam',
            3: 'data_ge'
        }
        
        result = []
        for i, sheet_name in enumerate(sheet_name_list):
            attr_name = sheet_mapping.get(i)
            if attr_name:
                sheet_dict = self.reader.get_sheet_as_dict(sheet_name)
                setattr(self, attr_name, sheet_dict)
                result.append(sheet_dict)
            else:
                result.append(None)
        
        # Дополняем None если листов меньше 4
        while len(result) < 4:
            result.append(None)
        
        return tuple(result)
    
    def load_excel(self, sheet_name: Optional[str] = None):
        """
        Загружает данные из Excel файла с указанного листа.
        (Метод оставлен для совместимости, но рекомендуется использовать sheet_name_list)
        """
        if sheet_name:
            reader = ExcelReader(self.file_path, sheet_name)
            return reader.return_data()
        return None
    

class SearchBam:
    """
    Класс для работы с BAM-файлом.
    """
    
    def __init__(self, sheet_name: str):
        self.config = handling_json.JsonConfig()
        self.reader = ExcelReader(
            file_path=self.config.getBAMPathFile_input(),
            sheet_name=sheet_name
        )
        # Автоматически фильтруем по колонке даты
        self.reader.filter_and_save_columns(
            self.config.getBAMColumnName("Table of contents: Date")
        )
    
    def return_data(self) -> Optional[Any]:
        """Возвращает загруженные данные."""
        return self.reader.return_data()
    
    def get_dict_all_data(self) -> Dict[int, Dict[str, Any]]:
        """Возвращает все данные в виде словаря."""
        return self.reader.get_dict_all_data()
    
    def get_colum(self, get_colum: str, foc_mode: float = 0) -> List[str]:
        """
        Возвращает уникальные значения из указанной колонки.
        
        :param get_colum: имя колонки
        :param foc_mode: режим фильтрации (0 = выключен, !=0 = включен)
        """
        # Условие пропуска для FOC-режима: пропускаем если колонка Date НЕ пустая
        def skip_condition(row_data: Dict[str, Any]) -> bool:
            date_col = self.config.getBAMColumnName("Table of contents: Date")
            return not self.reader.is_empty(row_data.get(date_col))

        
        return self.reader.get_column_values(
            get_column=get_colum,
            foc_mode=foc_mode != 0,
            skip_condition=skip_condition if foc_mode != 0 else None
        )

class SearchTrack:
    def __init__(self, sheet_name: str):
        self.config = handling_json.JsonConfig()
        
        # Получаем имя колонки D из конфига (или используем 'D' по умолчанию)
        # Предполагается, что в конфиге есть метод для получения имени колонки D
        if sheet_name == 'ФОЦ': column_d = 'Наименование'
        elif sheet_name == 'ПОЦ': column_d = 'Прутковые автоматы'
        elif sheet_name == 'ТОЦ': column_d = 'Токарные ЧПУ'


        self.reader = ExcelReader(
            file_path=self.config.getBAMPathFile_input(),
            sheet_name=sheet_name,
            color_filter_column=column_d,  # Фильтруем по цвету в колонке D
            track_sheet_origin=True        # Сохраняем имя листа
        )
        
        # Автоматически фильтруем по нужным колонкам
        # __sheet_origin__ добавляется автоматически при track_sheet_origin=True
        self.reader.filter_and_save_columns(
            self.config.getBAMColumnName("Table of contents: DCE")
        )
    
    def return_data(self) -> Optional[Any]:
        """Возвращает загруженные данные."""
        return self.reader.return_data()
    
    def get_dict_all_data(self) -> Dict[int, Dict[str, Any]]:
        """Возвращает все данные в виде словаря."""
        return self.reader.get_dict_all_data()
    
    def get_colum(self, get_colum: str, foc_mode: float = 0) -> List[str]:
        """
        Возвращает уникальные значения из указанной колонки.
        """
        # Условие пропуска для FOC-режима: пропускаем если колонка Date НЕ пустая
        def skip_condition(row_data: Dict[str, Any]) -> bool:
            date_col = self.config.getBAMColumnName("Table of contents: Date")
            return not self.reader.is_empty(row_data.get(date_col))
        
        return self.reader.get_column_values(
            get_column=get_colum,
            foc_mode=foc_mode != 0,
            skip_condition=skip_condition if foc_mode != 0 else None
        )
    
    def get_sheet_origin(self) -> Optional[str]:
        """Возвращает имя листа-источника."""
        return self.reader.sheet_origin
    

if __name__ == "__main__":
    app = SearchTrack("ФОЦ")
    ewhjrk =  app.get_dict_all_data()
    # for i,k in ewhjrk.items():
    #     print(i,k)