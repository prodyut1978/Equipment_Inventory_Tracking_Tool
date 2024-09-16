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

def Generate_Inv_Report():
    conn= sqlite3.connect("Eagle_Inventory.db")    
    InventoryCount_DF = pd.read_sql_query("select * from Eagle_Inventory ;", conn)
    Inv_data = pd.DataFrame(InventoryCount_DF)
    time.sleep(1)
    Inv_CountReport1   = Inv_data.groupby(['catg','model'], as_index=False).main_SN.count()
    Inv_CountReport1   = pd.DataFrame(Inv_CountReport1)
    Inv_CountReport1.rename(columns={'catg':'Category', 'model':'Model_Name', 'main_SN':'Total_Count'},inplace = True)
    Inv_CountReport1.to_sql('Eagle_Inventory_Report_1',conn, if_exists="replace", index=False)
    time.sleep(1)
    Inv_CountReport2   = Inv_data.groupby(['catg', 'location', 'model'], as_index=False).main_SN.count()
    Inv_CountReport2   = pd.DataFrame(Inv_CountReport2)
    Inv_CountReport2.rename(columns={'catg':'Category','location':'Location', 'model':'Model_Name', 'main_SN':'Total_Count'},inplace = True)
    Inv_CountReport2.to_sql('Eagle_Inventory_Report_2',conn, if_exists="replace", index=False)
    time.sleep(1)
    Inv_CountReport3   = Inv_data.groupby(['catg'], as_index=False).main_SN.count()
    Inv_CountReport3   = pd.DataFrame(Inv_CountReport3)
    Inv_CountReport3.rename(columns={'catg':'Category', 'main_SN':'Total_Count'},inplace = True)
    Inv_CountReport3.to_sql('Eagle_Inventory_Report_3',conn, if_exists="replace", index=False)            
    conn.commit()
    conn.close()
    tkinter.messagebox.showinfo("Generate Inventory Report","Inventory Report Generated Please Press View Report")
