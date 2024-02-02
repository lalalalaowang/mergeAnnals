import os
import xlrd
import xlwt


def merge():
    leaders = []
    staffs = []
    changes = []
    details_summary = []
    changes_summary = []
    all_summary = []
    retired_staffs = []

    file_path = os.getcwd()
    files = os.listdir(file_path)
    templates = [file for file in files if ".xls" in file]
    dirs = [file for file in files if "." not in file]
    for path in dirs:
        sub_file_path = os.path.join(file_path, path)
        sub_files = os.listdir(sub_file_path)
        for sub_file in sub_files:
            if "1、" in sub_file:
                data = xlrd.open_workbook(os.path.join(sub_file_path, sub_file))
                sheet = data.sheet_by_index(0)
                append_people(sheet, leaders)
                sheet2 = data.sheet_by_index(1)
                append_people(sheet2, staffs)

    print(leaders)
    print(staffs)


def append_people(sheet, people_list):
    for row in range(sheet.nrows):
        if sheet.row_values(row)[0]:
            if isinstance(sheet.row_values(row)[0], str):
                if sheet.row_values(row)[0].isdigit():
                    people_list.append(sheet.row_values(row))
            if isinstance(sheet.row_values(row)[0], (int, float)):
                people_list.append(sheet.row_values(row))


if __name__ == '__main__':
    merge()
