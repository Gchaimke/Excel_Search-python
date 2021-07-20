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


def find_cell(sh, searched_value):
    for row in range(sh.nrows):
        for col in range(sh.ncols):
            if sh.cell_value(row, col) == str(searched_value):
                return True
            if sh.cell_value(row, col) == searched_value:
                return True
    return False


def find_all(file_path, searched_value):
    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_name("Sheet1")
        if find_cell(sheet, searched_value):
            return True
    except XLRDError:
        print("An exception occurred in " + os.path.basename(file_path)[0:-5])
    return False


def save_to_txt(system, rf):
    f = open("search_result.txt", "a")
    f.write(system + "," + rf + "\n")
    f.close()


def start():
    #a = [i for i in range(19031500048, 19031500063)]
    #a = a + [i for i in range(19062400033, 19062400048)] + [i for i in range(19061300001, 19061300021)]
    #a = a + [i for i in range(19070500129, 19070500154)] + [i for i in range(19070500154, 19070500179)]
    #a = a + [i for i in range(19073000001, 19073000026)] + [i for i in range(19080500255, 19080500285)]
    #a = a + [i for i in range(19081300153, 19081300193)] + [i for i in range(19040900122, 19040900132)]
    a = ['31A65E9']
    start_path = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2019"
    start_path2 = "G:\Advantech\# Archive\@Clients\D-FEND\Production\Production 2020"
    file_list = search_files(start_path) + search_files(start_path2)
    save_to_txt("System SN", "RF SN")
    for search in a:
        for file in file_list:
            print("search in " + file + " for " + str(search))
            if find_all(file, search):
                file_list.remove(file)
                save_to_txt(os.path.basename(file)[0:-5], str(search))


def debug():
    #a = [i for i in range(19081300155, 19081300160)] + [i for i in range(19072600032, 19072600036)] + [i for i in range(19072600018, 19072600021)]
    a = ['31A65E9']
    start_path = "C:\\debug"
    start_path2 = "C:\\debug2"
    save_to_txt("System SN", "RF SN")
    file_list = search_files(start_path) + search_files(start_path2)
    for search in a:
        for file in file_list:
            print("search in " + file + " for " + str(search))
            if find_all(file, search):
                file_list.remove(file)
                save_to_txt(os.path.basename(file)[0:-5], str(search))


start()
#debug()
input('Press ENTER to exit')
# searchFiles()
