#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from pandas import ExcelWriter
import numpy as np
import pdfkit
import argparse
from PIL import Image,ImageTk
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog as fd

def __main__():
    window=tk.Tk()
    window.geometry("800x800")
    window.title("Automation")

    manager_name = tk.Label(text = "Desired Output/Target Value")
    manager_name.grid(column=0,row=1)
    departments_Entry = tk.Label(text = "Values to include in filtered output")
    departments_Entry.grid(column=0,row=2)
    message = tk.Label(text = "Every time you press 'Filter' an Excel and PDF will generate in correspondance to your specified inputs.")
    message.grid(column=0,row=3)

    manager_name_entry = tk.Entry()
    manager_name_entry.grid(column=1,row=1)
    departments_Entry = tk.Entry()
    departments_Entry.grid(column=1,row=2)
    
    def clear_text():
        manager_name_entry.delete(0, END)
        departments_Entry.delete(0, END)
        return messagebox.showinfo('New Files Generated','Excel and PDF files generated!')
    
    def converter(name):
        excel = name + '.xlsx'
        df = pd.concat(pd.read_excel(excel,sheet_name=None),ignore_index=True)
        df.to_html(name + '.html')
        options = {
        'margin-top': '0.25in',
        'margin-right': '0.25in',
        'margin-bottom': '0.25in',
        'margin-left': '0.25in',}
        pdfkit.from_file(name + '.html', name + '.pdf',options=options)

    
    def classify(name,*args):
        writer = ExcelWriter(name + '.xlsx')
        for x,y in grouped:
            if any([arg for arg in args if arg == x]):
                y.to_excel(writer,x,index=False) 
        writer.save()
        converter(name)
    
    def string_splitter():
        name = manager_name_entry.get()
        string_departments = departments_Entry.get()
        list_departments = string_departments.split(",")
        classify(name,*list_departments)
    
    
    def initiator():
        global filename 
        filename = fd.askopenfilename()
        
        excel_file = pd.read_excel(filename)
        
        global filtered
        initial = excel_file[excel_file['Last Name'] != 'Test']
        filtered = initial[initial['Login'] != 'Disabled']
        
        
        global grouped
        value = input("Enter the column you wish to operate on.")
        grouped = filtered.groupby(str(value))
    
        string_splitter()

    button=tk.Button(window,text="Filter",command=lambda:[initiator(), clear_text()])
    button.grid(column=1,row=3)
    destroy=tk.Button(window, text="Quit", command=window.destroy)
    destroy.grid(column=1,row=4)
    
    window.mainloop()
if __name__ == '__main__':
    __main__()


# In[ ]:




