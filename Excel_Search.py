import os
import xlrd
from xlrd import XLRDError
import re


def search_files(folder_path):
    my_files_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                if os.path.basename(file)[0] != "~":
                    my_files_list.append(os.path.join(root, file))
    return my_files_list


def find_in_all_cells(sh, search_list):
    for search in search_list:
        for row in range(sh.nrows):
            for col in range(sh.ncols):
                if sh.cell_value(row, col) == str(search):
                    return True
                if sh.cell_value(row, col) == search:
                    return True
    return False


def find_regex(sheet, regex_list):
    status = False
    for regex in regex_list:
        re.compile(regex)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                if re.match(regex, str(sheet.cell_value(row, col))):
                    status = True
                    continue
    return status


def find_in_spec_cell(sheet, regex, row, col):
    re.compile(regex)
    if sheet.nrows > row and sheet.ncols > col:
        if re.match(regex, str(sheet.cell_value(row, col))):
            return True
    return False


def find_all(file_path, search_list):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_in_all_cells(sheet, search_list):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False


def find_all_regex(file_path, searched_list):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_regex(sheet, searched_list):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False


def find_one(file_path, search, row, col):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_in_spec_cell(sheet, search, row, col):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False


def save_to_file(folder, system, status):
    f = open("search_result.csv", "a")
    f.write(folder + "," + system + "," + status + "\n")
    f.close()


def search_in_all_cells():
    search_list = [r"^.*[rR].*[ -][Bb]", r"^[Bb]"]  # ['19062400033', '19070500145', '19070500148', '19070500161']
    start_path = r"G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2019"
    start_path2 = r"G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2020"
    start_path3 = r"G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2021"
    file_list = search_files(start_path) + search_files(start_path2) + search_files(start_path3)
    save_to_file("Folder", "System SN", "Status")
    for file in file_list:
        print("search in " + file)
        if find_all_regex(file, search_list):
            save_to_file(os.path.dirname(file), os.path.basename(file)[0:-5], "1")
        else:
            save_to_file(os.path.dirname(file), os.path.basename(file)[0:-5], "0")
        

def search_in_one_cell(search, row, col):
    start_path = r"G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2019"
    start_path2 = r"G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2020"

    file_list = search_files(start_path) + search_files(start_path2)
    save_to_file("Folder", "System SN", "Status")
    for file in file_list:
        print("search in " + file + " for " + str(search))
        if find_one(file, search, row, col):
            save_to_file(os.path.dirname(file), os.path.basename(file)[0:-5], "1")
        else:
            save_to_file(os.path.dirname(file), os.path.basename(file)[0:-5], "0")


search_in_all_cells()
# search_in_one_cell(r"^.*[ -]B", 1, 2)
input('Press ENTER to exit')
