import tkinter as tk
from tkinter import *
from tkinter import filedialog
from pathlib import Path
from tkinter import ttk
import docx2txt
import os

# ********************************#
# updated by costa_rica 8 Jan 2020#
# ********************************#

def end_app():
    Text_input_window.destroy()
    quit()

def click():
    StoreVarFile=open(TmpFlNm, "w+")
    StoreVarFile.write(entry_1.get())
    Text_input_window.destroy()
# make function where search_word is the parameter
# parameters for function srch_wd = the string searching for in the file,
# Dir_1 = the directory path to search for word files,
# Fl_Typ = could be file name but intended to use a wildcard + "." + extension,
# srch_lst= list where sentences will get stored
# fl_lst = list var where file paths are stored will be stored
# fl_list needs another function to parse out the file name from the path or there is a functino for just the file names
def function_1(srch_wd, Dir_1, Fl_Typ, srch_lst, fl_lst):
    for x in Dir_1.glob(Fl_Typ):
        try:
            paragraph_list = docx2txt.process(x).splitlines()
        except:
            paragraph_list = (f"searched for: {srch_wd} but file could not be searched due to formatting")
    # search each file in directory for search_word. If in file append only the sentence and file name to corresponding lists
        for i in paragraph_list:
            if i.find(srch_wd)>=0:
                # parse paragraph for the sentence the word is in
                if i.count(".")== 0 or i.count(".")==1:
                    srch_lst.append(i)
                    fl_lst.append(x)
                else:
                    while i.count(".") > 1:
                        ChrCnt = i.find(".")
                        SntncTmp = i[0: ChrCnt + 1]
                        if SntncTmp.count(srch_wd) > 0:
                            srch_lst.append(SntncTmp)
                            fl_lst.append(x)
                            i = i[ChrCnt + 1:].strip()
                        else:
                            i = i[ChrCnt + 1:].strip()

ApplctnNm="Featherlook"
Text_input_window= Tk()
Text_input_window.geometry('600x350+100+200')
Text_input_window.title(ApplctnNm +" 1.1")
label_1=Label (Text_input_window,text="Enter search word:", bg="black", fg="white")
label_1.grid(row=1, column=0, sticky=W)
entry_1=ttk.Entry(Text_input_window, width=40, background="white")
entry_1.grid(row=2, column=0, sticky=W)
TmpFlNm="StoreVar" + ApplctnNm + ".txt"


btn_1=ttk.Button(Text_input_window, text="SUBMIT", width=10, command=click)
btn_1.grid(row=3, column=1, sticky=W)
Text_input_window.mainloop()

search_list=[]
file_list=[]

root = Tk()
# .withdraw() method hides the window that is used to do the .askdirectory() method.
# we don't need it to be visible.
root.withdraw()
root.fileName = filedialog.askdirectory()
path = Path(root.fileName)
root.destroy()
root.mainloop()

# search word is saved in this text file becuase i don't wnat to use the .get() method
StoreVarFile_1=open(f"D:\Documents\App development/featherlook/" + TmpFlNm,"r")
search_word= StoreVarFile_1.readline()
StoreVarFile_1.close()
x=int()

# call Search_Dir_for_word function
function_1(search_word, path, "*.docx",search_list, file_list)

InfoWindow = Tk()

InfoWindow.title("Searched for:")
Close_Button = ttk.Button(InfoWindow,text="Close",command=InfoWindow.destroy)
Close_Button.grid(row=len(search_list)+3, column=2, sticky='w')
Cl_Heading_1 = ttk.Label(InfoWindow,text="Files:", font=("arial",10, "underline", "bold"))
Cl_Heading_1.grid(row=1, column=0, sticky='w')
Cl_Heading_2 = ttk.Label(InfoWindow,text="Paragraph word found in:", font=("arial",10, "underline", "bold"))
Cl_Heading_2.grid(row=1, column=4, sticky='w')

file_list_dict = {}
search_list_dict={}

for i in range(len(file_list)):
    file_list_dict[i] = ttk.Label(InfoWindow,text=file_list[i]).grid(row=2+i, column=0, sticky='w')
    for y in range(len(search_list)):
        search_list_dict[y] = ttk.Label(InfoWindow, text=search_list[y]).grid(row=2+y, column=4, sticky='w')

# *************** delete here
file_to_remove = os.path.join(f"D:\Documents\App development/featherlook/" + TmpFlNm)
os.remove(file_to_remove)

InfoWindow.mainloop()

