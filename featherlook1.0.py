import tkinter as tk
from tkinter import *
from tkinter import filedialog
from pathlib import Path
from tkinter import ttk
import docx2txt

def click():
    search_word = entry_1.get()
    print(search_word)
    Text_input_window.destroy()

def end_app():
    Text_input_window.destroy()
    quit()

Text_input_window= Tk()
Text_input_window.geometry('600x350+100+200')
Text_input_window.title("Featherlook 1.1")
label_1=Label (Text_input_window,text="Enter search word:", bg="black", fg="white")
label_1.grid(row=1, column=0, sticky=W)

entry_1=ttk.Entry(Text_input_window, width=40, background="white")
entry_1.grid(row=2, column=0, sticky=W)
search_word= entry_1.get()
btn_1=ttk.Button(Text_input_window, text="SUBMIT", width=10, command=click)
btn_1.grid(row=3, column=1, sticky=W)
Text_input_window.mainloop()


search_list=[]
file_list=[]

root = Tk()
root.fileName = filedialog.askdirectory()
path = Path(root.fileName)
root.destroy()
root.mainloop()


for x in path.glob('*.docx'):
    try:
        paragraph_list = docx2txt.process(x).splitlines()
    except:
        paragraph_list = (f"searched for: {search_word} but file could not be searched due to formatting")

    for i in paragraph_list:
        if i.find(search_word)>=0:
            search_list.append(i)
            file_list.append(x)

#print(search_list)
#print(file_list)

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
    file_list_dict[i] = ttk.Label(InfoWindow,text=file_list[i])
    file_list_dict[i].grid(row=2+i, column=0, sticky='w')
    for y in range(len(search_list)):
        search_list_dict[y] = tk.Label(InfoWindow, text=search_list[y])
        search_list_dict[y].grid(row=2+y, column=4, sticky='w')


InfoWindow.mainloop()

# counter=1
# a=0
# import openpyxl as xl
# wb = xl.load_workbook(r"D:\Documents\Nick test folder\Search output.xlsx")
# sheet = wb['Sheet1']
# while a <= len(search_list)-1:
#     #print(search_list[a])
#     sheet.cell(counter,1).value=search_list[a]
#     sheet.cell(counter, 2).value = str(file_list[a])
#     counter+=1
#     a +=1
# wb.save(r"D:\Documents\Nick test folder\Search output.xlsx")
