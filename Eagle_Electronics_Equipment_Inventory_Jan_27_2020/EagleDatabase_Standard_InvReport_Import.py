#Inventory Report Backend
import os
from tkinter import*
import tkinter.messagebox
import EagleDatabase_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilename
import pandas as pd
import openpyxl
import csv
import time
import datetime

def ImportFile():
    window = Tk()
    window.title("Import Master DB Inventory File")
    width = 985
    height = 540
    window.config(bg="cadet blue")
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry("%dx%d+%d+%d" % (width, height, x, y))
    window.resizable(0, 0)
    TableMargin = Frame(window)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5", "column6", "column7", "column8", "column9", "column10" ),
                                height=25, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="Category", anchor=W)
    tree.heading("#2", text="Manufacturer", anchor=W)
    tree.heading("#3", text="Model", anchor=W)
    tree.heading("#4", text="ManfSN", anchor=W)
    tree.heading("#5", text="Description", anchor=W)            
    tree.heading("#6", text="AssetSN", anchor=W)
    tree.heading("#7", text="Date" ,anchor=W)
    tree.heading("#8", text="Location", anchor=W)
    tree.heading("#9", text="Condition", anchor=W)
    tree.heading("#10", text="Origin", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=80)            
    tree.column('#2', stretch=NO, minwidth=0, width=110)
    tree.column('#3', stretch=NO, minwidth=0, width=80)
    tree.column('#4', stretch=NO, minwidth=0, width=80)
    tree.column('#5', stretch=NO, minwidth=0, width=80)
    tree.column('#6', stretch=NO, minwidth=0, width=80)
    tree.column('#7', stretch=NO, minwidth=0, width=80)
    tree.column('#8', stretch=NO, minwidth=0, width=80)
    tree.column('#9', stretch=NO, minwidth=0, width=70)
    tree.column('#10', stretch=NO, minwidth=0, width=70)
    tree.pack()


    def Import_Inventory_File():
        name = askopenfilename(filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        if name:
            if name.endswith('.csv'):
                df = pd.read_csv(name, header = None,skiprows = {0})
                df.rename(columns = {0:'catg', 1:'manuf', 2:'model', 3:'main_SN', 4:'desc',
                                          5: 'asset_SN', 6:'datestamp', 7:'location', 8:'condition',
                                          9:'origin' },inplace = True)
                df['datestamp'] = pd.to_datetime(df['datestamp']).dt.strftime('%Y-%m-%d')
            else:
                df = pd.read_excel(name, header = None,skiprows = {0})
                df.rename(columns = {0:'catg', 1:'manuf', 2:'model', 3:'main_SN', 4:'desc',
                                          5: 'asset_SN', 6:'datestamp', 7:'location', 8:'condition',
                                          9:'origin' },inplace = True)
                df['datestamp'] = pd.to_datetime(df['datestamp']).dt.strftime("%Y-%m-%d")
                
            
            data = pd.DataFrame(df)
            if (data['main_SN'].duplicated().values.any() == True):
                tkinter.messagebox.showinfo("Add Error","Duplicate main_SN")
            else:
                if (data['asset_SN'].isnull().values.any() == True)| (data['main_SN'].isnull().values.any() == True):
                    tkinter.messagebox.showinfo("Import File Message","Manufacture and Asset SN can not be empty")
                else:
                    for each_rec in range(len(data)):
                        tree.insert("", tk.END, values=list(data.loc[each_rec]))
        ListBoxTotalImportEntries()
        con= sqlite3.connect("Eagle_Inventory.db")
        data.to_sql('Eagle_Inventory_temp',con, if_exists="replace", index=False)
        con.commit()
        con.close()
        

    def ListBoxTotalImportEntries():
        ImportTotalLBEntries.delete(0,END)
        Total_count = len(tree.get_children())
        ImportTotalLBEntries.insert(tk.END,Total_count)

                    
    def Submit_ImportToMasterDB():
        con= sqlite3.connect("Eagle_Inventory.db")
        cur=con.cursor()
        Imported_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory_temp ORDER BY `catg` ASC ;", con)
        ImportTotalLBEntries.delete(0,END)
        LengthDF = len(Imported_df)
        con.commit()
        con.close()
        
        if LengthDF == 0:
            tkinter.messagebox.showinfo("Import file","Please Select the Import File to Submit")
        else:
            iSubmit = tkinter.messagebox.askyesno("Entries Submit to Master DB", "Confirm if you want to Submit the Imported Entries to Master DB")
            if iSubmit >0:
                con= sqlite3.connect("Eagle_Inventory.db")
                cur=con.cursor()
                cur.execute("DELETE FROM Eagle_Inventory WHERE EXISTS (SELECT * FROM Eagle_Inventory_temp WHERE Eagle_Inventory.main_SN = Eagle_Inventory_temp.main_SN and Eagle_Inventory.model = Eagle_Inventory_temp.model)")                        
                cur.execute("INSERT INTO Eagle_Inventory (catg, manuf, model, main_SN, desc, asset_SN,\
                                datestamp, location, condition, origin) SELECT catg, manuf, model, main_SN, desc, asset_SN, datestamp, location, condition, origin FROM Eagle_Inventory_temp")
                cur.execute("DELETE FROM Eagle_Inventory_temp")
                time.sleep(2)
                con.commit()
                con.close()          
                tree.delete(*tree.get_children())
                tkinter.messagebox.showinfo("Submitted to Inventory Database(DB)","You have Submitted a Record to Master Inventory Database(DB)")
            return
        

    def Exit():
        window.destroy()

    btnImport = Button(window, text="Import Master DB File", font=('aerial', 9, 'bold'), bg= 'orange', height =1, width=18, bd=4, command = Import_Inventory_File)
    btnImport.place(x=840,y=420)

    btnSubmit = Button(window, text="Submit To Master DB", font=('aerial', 9, 'bold'), bg= 'orange', height =1, width=18, bd=4, command = Submit_ImportToMasterDB)
    btnSubmit.place(x=840,y=460)

    btnExit = Button(window, text="Exit", font=('aerial', 9, 'bold'), bg= 'orange', height =1, width=8, bd=4, command = Exit)
    btnExit.place(x=915,y=508)

    ImportTotalLBEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable= IntVar(), width = 6)
    ImportTotalLBEntries.place(x=870,y=80)
    L1Import = Label(window, text = "Total Import Entries", font=("arial", 10,'bold'),bg = "yellow").place(x=840,y=55)

