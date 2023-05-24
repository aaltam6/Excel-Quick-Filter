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
    
    #GUI Output labeling
    output_name = tk.Label(text = "Desired Output, This is what the file generated will be named.")
    output_name.grid(column=0,row=1)
    column_name = tk.Label(text = "Name of column to be filtered.")
    column_name.grid(column=0,row=2)
    filter_Entry = tk.Label(text = "Specify values to filter from column.")
    filter_Entry.grid(column=0,row=3)
    message = tk.Label(text = "Every time you press 'Filter' an Excel and PDF will generate in correspondance to your specified inputs.")
    message.grid(column=0,row=4)
    
    #Input variable labels
    output_name_entry = tk.Entry()
    output_name_entry.grid(column=1,row=1)
    column_name_entry = tk.Entry()
    column_name_entry.grid(column=1,row=2)
    filter_Entry = tk.Entry()
    filter_Entry.grid(column=1,row=3)
    
    #This will clear text from input spaces and print a confirmation message.
    def clear_text():
        output_name_entry.delete(0, END)
        column_name_entry.delete(0, END)
        filter_Entry.delete(0, END)
        return messagebox.showinfo('New Files Generated','Excel and PDF files generated!')
    
    #Convert excel to dataframe to html, pdfkit does not like pandas or excel and html seems to port well.
    def converter(name):
        excel = name + '.xlsx'
        df = pd.concat(pd.read_excel(excel,sheet_name=None),ignore_index=True)
        #an html file is generated incidentally along with pdf and excel sheet.
        df.to_html(name + '.html')
        options = {
        'margin-top': '0.25in',
        'margin-right': '0.25in',
        'margin-bottom': '0.25in',
        'margin-left': '0.25in',}
        pdfkit.from_file(name + '.html', name + '.pdf',options=options)

    #Create excel sheet with 'writer' functionality, list comprehesnion to iterate over grouped and write to writer. 
    #List comprehension with two variables like that is kinda slick but might a hackjob, feels off to me for some reason.
    def classify(name,*args):
        writer = ExcelWriter(name + '.xlsx')
        for x,y in grouped:
            if any([arg for arg in args if arg == x]):
                y.to_excel(writer,x,index=False) 
        writer.save()
        converter(name)
    
    #Split specified departments by comma, could prob make this more robust by included other characters in the split functionality.
    def string_splitter():
        name = output_name_entry.get()
        string_departments = filter_Entry.get()
        list_departments = string_departments.split(",")
        classify(name,*list_departments)
    
    #initiate process, this code is kinda hack kinda slick, could make it really clean with a better GUI framework or something.
    def initiator():
        global filename 
        #file directory opens your file system and asks for your selection.
        filename = fd.askopenfilename()
        
        #Create dataframe of initial file selected.
        excel_file = pd.read_excel(filename)
        
        #group by specified column.
        global grouped
        column = column_name_entry.get()
        grouped = excel_file.groupby(column)
        
        #Start that bad boy up!
        string_splitter()

    #create buttons, lambda function was the best way I could find to get it to work in this sequential manner, plus its kinda cool.
    button=tk.Button(window,text="Filter",command=lambda:[initiator(), clear_text()])
    button.grid(column=1,row=4)
    destroy=tk.Button(window, text="Quit", command=window.destroy)
    destroy.grid(column=1,row=5)
    
    window.mainloop()
if __name__ == '__main__':
    __main__()
