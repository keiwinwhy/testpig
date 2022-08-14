import os
current_dir = os.getcwd()

def get_value(printtype):
    if printtype == "ZPL":
        with open(current_dir + "\\" + "zpl.txt","r") as f:
            lines = f.readlines()
            p_data = lines[1]
            a = p_data.split("^")
            comb = a[9].split("~")
            p_dark = "file[" + comb[0] + "]"
            p_speed = "file[" + comb[1] + "]"
            w = lines[4].split("\n")[0]
            h = lines[5].split("\n")[0]
            p_width = "file[" + w + "]"
            p_height = "file[" + h + "]"
            f.close()

    elif printtype == "CPCL":
        with open(current_dir + "\\" + "output.txt","r",errors="ignore") as f:
            lines = f.readlines()
            p_width = "file[" + lines[1].split("\n")[0] + "]"
            p_dark = "file[" + lines[2].split("\n")[0] + "]"
            p_speed = "file[" + lines[3].split("\n")[0] + "]"
            h = lines[0]
            ph = h.split(" ")
            p_height = "file[" + ph[4] + "]"
            f.close()
    return p_width,p_dark,p_speed,p_height

#printtype = "ZPL"
#printtype = "CPCL"
#p_width,p_dark,p_speed,p_height = get_value(printtype)
#print(p_width,p_dark,p_speed,p_height)