import sys

import pandas as pd  # type: ignore
import re
from typing import Any, Optional
from exeptions import NoneParameter
from datetime import datetime

title_table = pd.DataFrame()


def _search_cells(searching_row: str,
                  end_row: Optional[int] = None,
                  start_row: int = 0,
                  check_column: int = 0) -> Optional[int]:
    """Поиск ячеек на титульном листе"""
    for index, cell in enumerate(title_table.iloc[start_row:end_row,
                                 check_column]):
        if re.search(searching_row, str(cell)) is not None:
            return index
    return None


def _search_eq_and_spec(header_name: str, check_column: int) -> list:
    """Поиск оборудования или специалистов с помещением их в список"""
    eq_or_spec_list = []
    pre_header_row = _search_cells(header_name, check_column=check_column)
    if pre_header_row:
        header_row = pre_header_row + 1
    else:
        raise NoneParameter("На титульном листе удалены сведения о Зав.№ оборудования/ФИО специалистов")
    while not (pd.isnull(title_table.iloc[header_row, check_column]) and
               pd.isnull(title_table.iloc[header_row + 1, check_column])):
        specialist = title_table.iloc[header_row, check_column]
        if not pd.isnull(specialist):
            eq_or_spec_list.append(specialist)
        header_row += 1
    return eq_or_spec_list


def _check_none_parameter(*parameters: Any) -> None:
    """Проверка, что все обязательные поля заполнены"""
    for parameter in parameters:
        if not parameter:
            raise NoneParameter("На Титульном листе экспресс-отчета отсутствуют обязательные параметры")


def _normalize_parameters(vtd_obj_name: str,
                          vtd_obj_type: str,
                          pipeline_category: str,
                          pipeline_name: Optional[str],
                          contract_number_and_date: Optional[str],
                          pipeline_pressure: Optional[str]) -> Any:
    """Приведение параметров к нормальному виду"""

    def _remove_none_value(value: Optional[str]) -> Optional[str]:
        """Обнуление переменных в случае, если в поле лежит меньше двух символов"""
        if value:
            if not re.search(r'\w{2,}', value):
                value = None
        return value

    def _correcting_types(types: str) -> str:
        """Приведение типа трубопровода к шаблону"""
        if re.search(r'[пП]лощад|[цЦ]ех', types) and re.search(r'[шШ]лейф|[уУ]зел|[уУ]зла', types):
            types = "Технологические трубопроводы"
        elif re.search(r'[пП]лощад|[цЦ]ех', types):
            types = "Внутриплощадочные технологические трубопроводы"
        elif re.search(r'[шШ]лейф', types) and re.search(r'[уУ]зел|[уУ]зла', types):
            types = "Технологические трубопроводы узла подключения и подключающих шлейфов"
        elif re.search(r'[шШ]лейф', types):
            types = "Технологические трубопроводы подключающих шлейфов"
        elif re.search(r'[уУ]зел|[уУ]зла', types):
            types = "Технологические трубопроводы узла подключения"
        elif re.search(r'[мМ]агистрал', types):
            types = "Магистральный трубопровод"
        else:
            types = "Технологические трубопроводы"
        return types

    def _parse_vtd_obj_name(parse_name: str) -> Any:
        """Извлечение данных из наименования объекта"""
        match = re.search(r'КЦ-\d+',
                          parse_name)
        if match:
            kc_number = int(match[0].split('КЦ-')[1])
        else:
            kc_number = None
        match = re.search(r'КС-\d+',
                          parse_name)
        if match:
            ks_number = int(match[0].split('КС-')[1])
        else:
            ks_number = None
        match = re.search(r'КС\S*\s(?!КЦ|\w+ого|\w+ое)\S+',
                          parse_name)
        if match:
            ks_name = re.split(r'КС\S*\s', match[0])[1].replace('"', '')
        else:
            ks_name = None
        lpumg_search = re.search(r'\s\w+\sЛПУМГ', parse_name)
        if lpumg_search:
            lpumg_name = lpumg_search[0].split(' ЛПУМГ')[0].replace(' ', '')
        else:
            lpumg_name = None
        return kc_number, ks_number, ks_name, lpumg_name

    def _parse_num_and_date(number_and_date: Optional[str]) -> Any:
        """Извлечение даты и номера в отдельные переменные"""
        if number_and_date:
            if re.search(r'от', number_and_date):
                number_and_date_split = re.split(r'\sот\s', number_and_date)
                number = number_and_date_split[0]
                date = datetime.strptime(number_and_date_split[1], '%d.%m.%Y').date()
            else:
                number = number_and_date
                date = None
        else:
            number = None
            date = None
        return number, date

    pipeline_name = _remove_none_value(pipeline_name)
    contract_number, contract_date = _parse_num_and_date(
        _remove_none_value(contract_number_and_date))
    pipeline_category_list = pipeline_category.replace(" ", "").split(",")
    vtd_obj_type = _correcting_types(vtd_obj_type)
    if pipeline_pressure:
        pipeline_pressure = float(re.sub(r'[М|м|П|П|А|а]', '', pipeline_pressure).replace(',', '.'))
    parse_vtd_obj_parameters = _parse_vtd_obj_name(vtd_obj_name)
    normalize_return = (pipeline_category_list,
                        pipeline_name,
                        vtd_obj_type,
                        contract_number,
                        contract_date,
                        pipeline_pressure,
                        *parse_vtd_obj_parameters)

    return normalize_return


def title_table_parse(express_name) -> Any:
    global title_table
    try:
        title_table = pd.read_excel(
            express_name,
            sheet_name='Титульный лист',
            engine='openpyxl')
    except ValueError:
        raise ValueError("Титульный лист удален или переименован")
    try:
        # Наименование заказчика
        client_name = title_table.iloc[_search_cells(r'Наименование общества'), 1]
        # Наименование объекта контроля
        vtd_obj_name = title_table.iloc[_search_cells(r'Наименование об\wекта'), 1]
        # Наименование газопровода/трубопровода
        pipeline_name = title_table.iloc[_search_cells(r'Наименование \w+провода'), 1]
        # Вид объекта контроля
        vtd_obj_type = title_table.iloc[_search_cells(r'Вид об\wекта'), 1]
        # Даты начала и окончания работ
        vtd_start_date = title_table.iloc[_search_cells(r'Дата начала'), 1]
        vtd_end_date = title_table.iloc[_search_cells(r'Дата окончания'), 1]
        # Категория трубопровода
        pipeline_category = title_table.iloc[_search_cells(r'Категория', check_column=5), 6]
        # Номер и дата договора/письма
        contract_number_and_date_row = _search_cells(r'договор|письм\w', check_column=3)
        if contract_number_and_date_row:
            contract_number_and_date = title_table.iloc[contract_number_and_date_row + 1, 3]
        else:
            raise NoneParameter("На титульном листе удалены сведения о номере договора/письма")
        pipeline_pressure_row = _search_cells(r'[Дд]авление', check_column=1)
        if pipeline_pressure_row:
            pipeline_pressure = title_table.iloc[pipeline_pressure_row, 2]
        else:
            pipeline_pressure = None
        # Перечень заводских номеров оборудования
        equipment_numbers_list = list(map(int, _search_eq_and_spec(r'Зав', 2)))
        # Перечень специалистов
        specialists_list = _search_eq_and_spec(r'Ф\sИ\sО', 1)
        #
        _check_none_parameter(client_name,
                              vtd_obj_name,
                              vtd_obj_type,
                              vtd_start_date,
                              vtd_end_date,
                              pipeline_category,
                              equipment_numbers_list,
                              specialists_list)
    except ValueError:
        raise ValueError("На титульном листе удалены или переименованы необходимые строки")
    normalize_parameters = _normalize_parameters(vtd_obj_name,
                                                 vtd_obj_type,
                                                 pipeline_category,
                                                 pipeline_name,
                                                 contract_number_and_date,
                                                 pipeline_pressure)
    return_parameters = (client_name,
                         *normalize_parameters,
                         vtd_start_date.date(),
                         vtd_end_date.date(),
                         equipment_numbers_list,
                         specialists_list
                         )
    return return_parameters
