
"""
@author: arn@ meisan-tech@outlook.com
"""
import glob
import os
import pandas as pd
from pandas import ExcelWriter
from tkinter import *
from tkinter import filedialog

root = Tk()
def browse():
    root.filename =  filedialog.askdirectory(initialdir = "/",title = "Select file folder")
    indir=root.filename
    concator(indir)

def concator(infile):
    writer=pd.ExcelWriter('concat_output.xlsx')
    os.chdir(infile)
    xlsx=glob.glob("*.xlsx")
    xls=glob.glob("*.xls")
    files=xlsx+xls
    for filename in files:
        excel_file = pd.ExcelFile(filename)
        sh_names=excel_file.sheet_names
        #sht_count=len(sh_names)
        print(filename)
        for sh_name in sh_names:
            print("Step done moving to next")
            df_excel = pd.read_excel(filename, sheet_name=sh_name)
            df_excel.to_excel(writer,sheet_name=sh_name)
            print("starting:-"+sh_name)
        print(filename+" is done")
    print("All files are combained to "+infile+"\concat_output.xlsx")
    writer.save()
l1=Label(root,text="Select the folder of files:")
b1=Button(root,text='Browse',command=browse)
l1.grid(row=0,column=1)
b1.grid(row=1,column=1)
root.mainloop()

"""
Make EXE by :
   pyinstaller --onefile <<scipt name>>
"""

    
    
    