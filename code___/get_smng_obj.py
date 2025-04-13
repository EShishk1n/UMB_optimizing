from openpyxl.worksheet import worksheet
from code___.common_def import open_sheet, create_classes


def get_smng_obj():
    sheet_ranges = open_sheet(filename = '../daily_raports/smng.xlsx', sheetname = 'Север')
    dzo = 'ООО "ННК-Самаранефтегаз'
    wellname_list = get_smng_wellname(sheet_ranges)
    fieldname_list = get_smng_fieldname(sheet_ranges)
    padname_list = ['-'] * len(wellname_list)
    # smng_obj = create_classes(dzo, fieldname_list, padname_list, wellname_list)
    print(wellname_list)

    return 5

def get_smng_wellname(sheet_ranges: worksheet) -> list:
    wellname_list = []

    for i in range(8, 250, 8):
        cell_name = 'b' + str(i)
        value = sheet_ranges[cell_name].value
        proc = str(sheet_ranges['m' + str(i)].value) + str(sheet_ranges['m' + str(i + 7)].value)
        print(proc)
        if 'Демонтаж' in str(proc):
            print(value)
            continue
        if str(value).startswith('ИТОГО по проекту'):
            break
        wellname_list.append(value)

    return wellname_list

def get_smng_fieldname(sheet_ranges: worksheet) -> list:
    fieldname_list = []
    for i in range(13, 250, 8):
        cell_name = 'a' + str(i)
        value = sheet_ranges[cell_name].value
        fieldname_list.append(value)
    return fieldname_list


print(get_smng_obj())