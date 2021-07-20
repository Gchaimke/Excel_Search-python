import os
import xlrd
from xlrd import XLRDError


def search_files(folder_path):
    my_files_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                if os.path.basename(file)[0] != "~":
                    my_files_list.append(os.path.join(root, file))
    return my_files_list


def find_in_all_cells(sh, searched_value):
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            if sh.cell_value(row, col) == str(searched_value):
                return True
            if sh.cell_value(row, col) == searched_value:
                return True
    return False

def find_in_spec_cell(sh, searched_value, row, col):
    if sh.nrows > row and sh.ncols > col:
        if sh.cell_value(row, col) == str(searched_value):
            return True
        if sh.cell_value(row, col) == searched_value:
            return True
        if sh.cell_value(row, col) == "REV "+str(searched_value):
            return True
        if sh.cell_value(row, col) == "REV "+searched_value:
            return True
    return False


def find_all(file_path, searched_value):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_in_all_cells(sheet, searched_value):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False

def find_one(file_path, searched_value, row, col):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_in_spec_cell(sheet, searched_value, row, col):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False

def save_to_txt(system, rf):
    f = open("search_result.txt", "a")
    f.write(system + "," + rf + "\n")
    f.close()


def search_in_all_cells():
    a = ['59957'] #['19062400033', '19070500145', '19070500148', '19070500161', '19070500165']
    start_path = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2019"
    start_path2 = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2020"
    start_path3 = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2021"
    file_list = search_files(start_path) + search_files(start_path2) + search_files(start_path3)
    save_to_txt("System SN", "RF SN")
    for search in a:
        for file in file_list:
            print("search in " + file + " for " + str(search))
            if find_all(file, search):
                file_list.remove(file)
                save_to_txt(os.path.basename(file)[0:-5], str(search))
				
def search_in_one_cell(search, row, col):
    start_path = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2019"
    start_path2 = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2020"
    #start_path = "C:\\debug"
    #start_path2 = "C:\\debug2"
    file_list = search_files(start_path) + search_files(start_path2)
    save_to_txt("System SN", "RF SN")
    for file in file_list:
        print("search in " + file + " for " + str(search))
        if find_one(file, search, row, col):
            file_list.remove(file)
            save_to_txt(os.path.basename(file)[0:-5], str(search))


def debug():
    #a = [i for i in range(19081300155, 19081300160)] + 
    a = ['19062400033','19070500145']
    start_path = "C:\\search1"
    start_path2 = "C:\\search2"
    save_to_txt("System SN", "RF SN")
    file_list = search_files(start_path) + search_files(start_path2)
    for search in a:
        for file in file_list:
            print("search in " + file + " for " + str(search))
            if find_all(file, search):
                file_list.remove(file)
                save_to_txt(os.path.basename(file)[0:-5], str(search))


search_in_all_cells()
#search_in_one_cell("B", 1, 2)
#debug()
input('Press ENTER to exit')
# searchFiles()
