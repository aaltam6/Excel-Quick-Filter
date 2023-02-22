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

    output_name = tk.Label(text = "Desired Output, This is what the file generated will be named.")
    output_name.grid(column=0,row=1)
    column_name = tk.Label(text = "Name of column to be filtered.")
    column_name.grid(column=0,row=2)
    filter_Entry = tk.Label(text = "Specify values to filter from column.")
    filter_Entry.grid(column=0,row=3)
    message = tk.Label(text = "Every time you press 'Filter' an Excel and PDF will generate in correspondance to your specified inputs.")
    message.grid(column=0,row=4)

    output_name_entry = tk.Entry()
    output_name_entry.grid(column=1,row=1)
    column_name_entry = tk.Entry()
    column_name_entry.grid(column=1,row=2)
    filter_Entry = tk.Entry()
    filter_Entry.grid(column=1,row=3)
    
    def clear_text():
        output_name_entry.delete(0, END)
        column_name_entry.delete(0, END)
        filter_Entry.delete(0, END)
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
        name = output_name_entry.get()
        string_departments = filter_Entry.get()
        list_departments = string_departments.split(",")
        classify(name,*list_departments)
    
    
    def initiator():
        global filename 
        filename = fd.askopenfilename()
        
        excel_file = pd.read_excel(filename)
        
        global grouped
        column = column_name_entry.get()
        grouped = excel_file.groupby(column)
    
        string_splitter()

    button=tk.Button(window,text="Filter",command=lambda:[initiator(), clear_text()])
    button.grid(column=1,row=4)
    destroy=tk.Button(window, text="Quit", command=window.destroy)
    destroy.grid(column=1,row=5)
    
    window.mainloop()
if __name__ == '__main__':
    __main__()
