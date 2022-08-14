from time import sleep
from pywinauto import *

import xlrd
import os

parent_dir = os.path.abspath(os.pardir)
current_dir = os.getcwd()

t = 0.5
t1 = 2
t2 = 4

file_name = current_dir + "\\" + 'SRSmapping_Forerunner&Stingray.xlsx'
book = xlrd.open_workbook(file_name)
work_sheet_range = range(book.nsheets)
try:
    work_sheet = book.sheet_by_name("Stringray mapping V5")
except:
    print("work sheet not found is %s name" %file_name)

num_cols = work_sheet.ncols
#for j in range(2, num_cols):
for j in range(3, 5):
    prtname = work_sheet.cell_value(4,j)

    # install printer
    packpath = "C:\\ZD5-1-17-7408\\"
    packname = "PrnInst.exe"
    app = Application(backend="uia").start(packpath + packname)
    st_win = app.window(best_match="PrnInst-Welcome", top_level_only=False)
    # st_win.print_control_identifiers()

    st_win.child_window(title="Next >", control_type="Button").click()

    st_win = app.window(best_match="PrnInst-Options", top_level_only=False)
    st_win.child_window(title="Install Printer", control_type="Button").click_input()

    st_win = app.window(best_match="PrnInst-License Agreement", top_level_only=False)
    st_win.child_window(title="I accept the terms in the license agreement", control_type="RadioButton").click()
    st_win.child_window(title="Next >", control_type="Button").click()
    sleep(t2)

    st_win = app.window(best_match="PrnInst-Selecting the printer", top_level_only=False)

    # 通过键盘输入打印机名字
    st_win.type_keys(prtname)

    st_win = app.window(best_match="PrnInst-Selecting the printer", top_level_only=False)

    # selected the target pritner
    st_win.child_window(title=prtname, control_type="ListItem").click_input()
    st_win.child_window(title="Next >", control_type="Button").click()

    st_win = app.window(best_match="PrnInst-Options", top_level_only=False)

    st_win.child_window(title="D:\\pythondocument\\CSAUTO\\zebra_proj\\test\\NewDriver8\\checktext.txt", control_type="ListItem").click_input()
    st_win.child_window(title="Next >", control_type="Button").click()

    st_win = app.window(best_match="PrnInst-Additional Installations", top_level_only=False)
    #st_win.print_control_identifiers()
    st_win.child_window(title="Launch installation of Zebra Font Downloader Setup Wizard", control_type="CheckBox").click_input()
    st_win.child_window(title="Finish", control_type="Button").click()
    sleep(7)
