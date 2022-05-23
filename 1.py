#%%
import openpyxl
import datetime
import tkinter
from tkinter import ttk
from tkinter import messagebox
#%%
def save():
     a = str(description.get())
     b = float(income.get())
     c = float(outgo.get())
     d = str(MontYearselect.get())
     e = str(Categoryselect.get())
     book = openpyxl.load_workbook()
     sheet = book[d]
     sheet.append([datetime.datetime.now(),e,a,b,c])
     book.save()
     income.delete(0,'end')
     outgo.delete(0,'end')
     messagebox.showinfo('finished')
     
#%%

     