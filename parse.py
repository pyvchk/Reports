from title_table_parse import title_table_parse
import os
import re
from exeptions import NoSuchFileInFolder
from typing import Optional


def _check_express_file(parse_name: str = r'[Ээ]кспресс.+xlsx\Z'):

    for file_name in os.listdir():
        if re.search(parse_name, file_name):
            return file_name
    else:
        raise NoSuchFileInFolder("В папке с программой нет экспресс-отчета")


def express_parse(express_name: Optional[str]):
    if not express_name:
        express_name = _check_express_file()
    return title_table_parse(express_name)


def parse(dir_name: str = 'input/',
          express_name: Optional[str] = None):
    os.chdir(dir_name)
    return express_parse(express_name)


if __name__ == '__main__':
    parse()
