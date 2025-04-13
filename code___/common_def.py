from openpyxl.reader.excel import load_workbook
from openpyxl.worksheet import worksheet

from code___.classes import ObjectFromRaports


def open_sheet(filename: str, sheetname: str) -> worksheet:
    wb = load_workbook(filename, read_only=True, data_only=True, keep_links=False)

    return wb[sheetname]


def create_classes(dzo: str,
                   fieldname_list: list[str],
                   padname_list: list[str],
                   wellname_list: list[str]) -> list[ObjectFromRaports]:
    obj_list = []
    for i in range(len(wellname_list)):
        object_from_raport = ObjectFromRaports(dzo,
                                   fieldname=fieldname_list[i],
                                   padname=padname_list[i],
                                   wellname=str(wellname_list[i]))
        obj_list.append(object_from_raport)

    return obj_list