import tkinter as tk
from tkinter import *
from tkinter import filedialog
from pathlib import Path
from tkinter import ttk
import docx2txt
import os
import glob

# *********************************#
# updated by costa_rica 13 Jan 2020#
# *********************************#

def end_app():
    Text_input_window.destroy()
    quit()

def click():
    StoreVarFile=open(TmpFlNm, "w+")
    StoreVarFile.write(entry_1.get())
    Text_input_window.destroy()

# getting path from .exe
if getattr(sys, 'frozen', False):
    feath_name = os.path.basename(sys.executable)
    feath_dir = os.path.dirname(sys.executable)
    feath_dir_and_name = sys.executable.replace('\\', '/')
# getting absolute paths for .py:
elif __file__:
    feath_dir_and_name=__file__.replace('\\','/')
    feath_name=os.path.basename(__file__)
    feath_dir=os.path.dirname(__file__).replace('\\','/')

ApplctnNm="Featherlook"
Text_input_window= Tk()
Text_input_window.geometry('600x350+100+200')
Text_input_window.title(ApplctnNm +" 1.1")
search_word=StringVar()
label_1=Label (Text_input_window,text="Enter search word:", bg="black", fg="white")
label_1.grid(row=1, column=0, sticky=W)
entry_1=ttk.Entry(Text_input_window, width=40, background="white")
entry_1.grid(row=2, column=0, sticky=W)
TmpFlNm="StoreVar_" + ApplctnNm + ".txt"
btn_1=ttk.Button(Text_input_window, text="SUBMIT", width=10, command=click)
btn_1.grid(row=3, column=1, sticky=W)
Text_input_window.mainloop()


root = Tk()
# .withdraw() method hides the window that is used to do the .askdirectory() method.
# we don't need it to be visible.
root.withdraw()
root.fileName = filedialog.askopenfilename()
search_dir = os.path.dirname(root.fileName)
root.destroy()
root.mainloop()

# # not sure i need the line below:
# Store_1=feath_dir +"/"+ TmpFlNm

StoreVarFile_1=open(feath_dir.replace('\\','//') +"/"+ TmpFlNm,"r")
# StoreVarFile_1=open(feath_dir +"/"+ TmpFlNm,"r")
# print(StoreVarFile_1)
Store_1=feath_dir.replace('\\','//') +"/"+ TmpFlNm
search_word= StoreVarFile_1.readline()
StoreVarFile_1.close()

# variables and things needed for file and word parser
search_list=[]
file_list=[]
x = int()
y = int()
file_counter=int()
file_limit=2
paragraph_limit=3
para_sentence_limit=4


# ******put file word search and parser here*************
for a in glob.glob(search_dir+"/"+"*.docx"):
    # print('are we here?')
    para_counter=0
    file_counter+=1
    if a.count(search_word)>0:
        paragraph_list=docx2txt.process(a)
        # b is each paragraph in a file
        for b in paragraph_list.split('\n'):
            b_sentence_count = 0
            para_counter+=1
            # print(file_counter, para_counter, b_sentence_count)
            if b.count(search_word)>0:
                if b.count(".")<=1:
                    file_list.append(os.path.basename(a))
                    search_list.append(b)
                    x=len(file_list)
                    y=len(search_list)
                # establishes the paragraph has more than one sentence with elif below
                elif b.count(".")>1:
                    while b.count(".")>0:
                        # print("elif b.count(.)>1 is " + str(b.count(".")))
                        b_subsentence = b[0:b.find(".") +1]
                        b_sentence_count +=1
                        # establishes the search_word is not in b_subsentence
                        if b_subsentence.count(search_word) == 0:
                            # make b everything b was except b_subsentce
                            b = b[len(b_subsentence):]
                        # search_word is in b_subsentence
                        elif b_subsentence.count(search_word)>0:
                            search_list.append(b_subsentence)
                            file_list.append(os.path.basename(a))
                            x = len(file_list)
                            y = len(search_list)
                            b = b[len(b_subsentence):]
                            # print(file_list[x - 1], x, search_list[y - 1], para_counter)
                            # ***********LIMIT number of SENTENCES in paragraph***********
                            if b_sentence_count >= para_sentence_limit:
                                break
            # ***********LIMIT number of PARAGRAPHS in file***********
            elif para_counter>=paragraph_limit:
                # print('this filed_ 2')
                break
        # ***********LIMIT number of Files***********
        if file_counter>=file_limit:
            break



# print(y)
InfoWindow = Tk()
InfoWindow.title("Searched for:")
Close_Button = ttk.Button(InfoWindow,text="Close",command=InfoWindow.destroy)
Close_Button.grid(row=len(search_list)+3, column=2, sticky='w')
Cl_Heading_1 = ttk.Label(InfoWindow,text="Files:", font=("arial",10, "underline", "bold"))
Cl_Heading_1.grid(row=1, column=0, sticky='w')
Cl_Heading_2 = ttk.Label(InfoWindow,text="Sentence word found in:", font=("arial",10, "underline", "bold"))
Cl_Heading_2.grid(row=1, column=4, sticky='w')

# need to make this into a textbox to have scrollbar or at least it seems the .yview attribtue won't work on a Label
# scroll_1 = ttk.Scrollbar(InfoWindow, orient='vertical', command= Cl_Heading_2.yview)

file_list_dict = {}
search_list_dict={}

for i in range(len(file_list)):
    file_list_dict[i] = ttk.Label(InfoWindow,text=file_list[i]).grid(row=2+i, column=0, sticky='w')
    for y in range(len(search_list)):
        search_list_dict[y] = ttk.Label(InfoWindow, text=search_list[y]).grid(row=2+y, column=4, sticky='w')

# *************** delete temporary text file here
file_to_remove = feath_dir+'/' + TmpFlNm
os.remove(file_to_remove)

InfoWindow.mainloop()

