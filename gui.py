from calendar import week
import tkinter as tk
import tkinter.font as tkFont
from pianoaule import generate_pngs, WORKING_DIR
from os import startfile, path, getcwd

single_file: bool
cleanup: list
week_offset: int

def print_configs():
    print(f"SingleFile: {single_file}|Cleanup={cleanup}|Offset={week_offset}")

def btn_generate_command():
    single_file = radio_value.get() == "single"
    week_offset = 0 if week_value.get() == "this" else 1
    cleanup = []
    if csv.get(): cleanup.append("csv")
    if xlsx.get(): cleanup.append("xlsx")
    # print(f"{single_file=} | {week_offset=} | {cleanup=}")

    btn_generate.config(text="Processando...") 
    btn_generate["state"] = tk.DISABLED
    generate_pngs(single_file=single_file, week_offset=week_offset, cleanup=cleanup)
    btn_generate.config(text="Genera!")
    btn_generate["state"] = tk.NORMAL
    return


def btn_opendir_command():
    try:
        startfile(WORKING_DIR)
    except:
        btn_opendir.config(text="Errore :(")


root = tk.Tk()
try:
    root.iconphoto(False, tk.PhotoImage(file=path.abspath(path.join( getcwd(),"..",  './logo.png'))))
except:
    pass

csv = tk.IntVar(None, 1)
xlsx = tk.IntVar(None, 1)
week_value = tk.StringVar(None, 'this')
radio_value = tk.StringVar(None, "single")


root.title("Piano Aule by VI")
width=600
height=200
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=False, height=False)


label_weekoffset=tk.Label(root)
ft = tkFont.Font(family='Times',size=10)
label_weekoffset["font"] = ft
label_weekoffset["fg"] = "#333333"
label_weekoffset["justify"] = "center"
label_weekoffset["text"] = "Quale settimana?"
label_weekoffset.place(x=20,y=10,width=215,height=30)

cbox_questaw=tk.Radiobutton(root, variable=week_value)
ft = tkFont.Font(family='Times',size=10)
cbox_questaw["font"] = ft
cbox_questaw["fg"] = "#333333"
cbox_questaw["justify"] = "center"
cbox_questaw["text"] = "Questa"
cbox_questaw.place(x=250,y=10,width=85,height=25)
cbox_questaw["value"] = "this"

cbox_nextw=tk.Radiobutton(root, variable=week_value)
ft = tkFont.Font(family='Times',size=10)
cbox_nextw["font"] = ft
cbox_nextw["fg"] = "#333333"
cbox_nextw["justify"] = "center"
cbox_nextw["text"] = "Prossima"
cbox_nextw.place(x=350,y=10,width=85,height=25)
cbox_nextw["value"] = "next"


btn_generate=tk.Button(root)
btn_generate["anchor"] = "center"
btn_generate["bg"] = "#efefef"
ft = tkFont.Font(family='Times',size=10)
btn_generate["font"] = ft
btn_generate["fg"] = "#000000"
btn_generate["justify"] = "center"
btn_generate["text"] = "Genera!"
btn_generate.place(x=470,y=80,width=117,height=57)
btn_generate["command"] = btn_generate_command

label_singlemulti=tk.Label(root)
label_singlemulti["borderwidth"] = "1px"
ft = tkFont.Font(family='Times',size=10)
label_singlemulti["font"] = ft
label_singlemulti["fg"] = "#333333"
label_singlemulti["justify"] = "center"
label_singlemulti["text"] = "Immagine singola o una per giorno"
label_singlemulti.place(x=20,y=120,width=215,height=30)

radio_single=tk.Radiobutton(root, variable=radio_value)
ft = tkFont.Font(family='Times',size=10)
radio_single["font"] = ft
radio_single["fg"] = "#333333"
radio_single["justify"] = "center"
radio_single["text"] = "Singolo"
radio_single.place(x=250,y=120,width=85,height=25)
radio_single["value"] = "single"

radio_multi=tk.Radiobutton(root, variable=radio_value)
ft = tkFont.Font(family='Times',size=10)
radio_multi["font"] = ft
radio_multi["fg"] = "#333333"
radio_multi["justify"] = "center"
radio_multi["text"] = "Multiplo"
radio_multi.place(x=350,y=120,width=85,height=25)
radio_multi["value"] = "multi"

btn_opendir=tk.Button(root)
btn_opendir["bg"] = "#efefef"
ft = tkFont.Font(family='Times',size=10)
btn_opendir["font"] = ft
btn_opendir["fg"] = "#000000"
btn_opendir["justify"] = "center"
btn_opendir["text"] = "Apri Cartella"
btn_opendir.place(x=470,y=150,width=115,height=30)
btn_opendir["command"] = btn_opendir_command

label_cleanup=tk.Label(root)
ft = tkFont.Font(family='Times',size=10)
label_cleanup["font"] = ft
label_cleanup["fg"] = "#333333"
label_cleanup["justify"] = "center"
label_cleanup["text"] = "Cancella questi file al termine"
label_cleanup.place(x=20,y=60,width=215,height=30)

cbox_csv=tk.Checkbutton(root, variable=csv)
ft = tkFont.Font(family='Times',size=10)
cbox_csv["font"] = ft
cbox_csv["fg"] = "#333333"
cbox_csv["justify"] = "center"
cbox_csv["text"] = "csv"
cbox_csv.place(x=250,y=60,width=85,height=25)
cbox_csv["offvalue"] = "0"
cbox_csv["onvalue"] = "1"


cbox_xlsx=tk.Checkbutton(root, variable=xlsx)
ft = tkFont.Font(family='Times',size=10)
cbox_xlsx["font"] = ft
cbox_xlsx["fg"] = "#333333"
cbox_xlsx["justify"] = "center"
cbox_xlsx["text"] = "xlsx"
cbox_xlsx.place(x=350,y=60,width=85,height=25)
cbox_xlsx["offvalue"] = "0"
cbox_xlsx["onvalue"] = "1"




    
root.mainloop()