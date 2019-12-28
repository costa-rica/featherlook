import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from pathlib import Path
import docx2txt

root = Tk()
root.fileName = filedialog.askdirectory()
SearchWindow = Tk()
SearchWindow.title("Search Input")
SearchWindow.mainloop()
SearchInput= ttk.Label(SearchWindow,text = "Enter search word:")
SearchInput.grid(row=0,column = 0)
search_word= str(input("search for:"))
search_list=[]
file_list=[]
path = Path(root.fileName)
for x in path.glob('*.docx'):
    try:
        paragraph_list = docx2txt.process(x).splitlines()
    except:
        paragraph_list = (f"searched for: {search_word} but file could not be searched due to formatting")

    for i in paragraph_list:
        if i.find(search_word)>=0:
            search_list.append(i)
            file_list.append(x)

print(search_list)
print(file_list)
InfoWindow = tk.Label(root,text=search_list)
InfoWindow.pack()
root.mainloop()
counter=1
a=0
import openpyxl as xl
wb = xl.load_workbook(r"D:\Documents\Nick test folder\Search output.xlsx")
sheet = wb['Sheet1']
while a <= len(search_list)-1:
    #print(search_list[a])
    sheet.cell(counter,1).value=search_list[a]
    sheet.cell(counter, 2).value = str(file_list[a])
    counter+=1
    a +=1
wb.save(r"D:\Documents\Nick test folder\Search output.xlsx")