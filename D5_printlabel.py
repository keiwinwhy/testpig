import os
from time import sleep
from pywinauto import *

current_dir = os.getcwd()
t1 = 1

def action_printlabel(app2,dp2,prtname):
    # 右键-properties
    dp2.child_window(title=prtname).click_input(button='right')
    app2.Context.Printerproperties.click_input()
    sleep(t1)

    app4 = Application(backend="uia").connect(title_re=prtname + "*")
    dp4 = app4.window(best_match = prtname + " Properties",  top_level_only = False)
    # dp4.print_control_identifiers()

    sleep(t1)
    dp4.child_window(title="Print Test Page", control_type="Button").click_input()
    sleep(t1)

    # dailog 的处理, 先找到dailog所在的window
    app5 = Application(backend="uia").connect(title_re=prtname + "*")
    dp5 = app5.window(best_match=prtname + " Properties", top_level_only=False)
    sleep(t1)
    # dp5.print_control_identifiers()
    dp5.child_window(title="Close", auto_id="CommandButton_1", control_type="Button").click_input()

    dp4.child_window(title="Cancel", control_type="Button").click_input()
    sleep(5)

    return

