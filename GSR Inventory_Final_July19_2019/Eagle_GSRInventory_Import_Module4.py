#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_GSRInventory_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
import pandas as pd
import openpyxl
import csv
import time
import datetime
Default_Date_today   = datetime.date.today()

class GSRInventoryImport:    
    def __init__(self,root):
        self.root =root
        self.root.title ("Eagle GSR Inventory Import Wizard")
        self.root.geometry("1350x650+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)
        TableMargin = Frame(self.root, bd = 2, padx= 10, pady= 8, relief = RIDGE)
        TableMargin.pack(side=TOP)
        TableMargin.pack(side=LEFT)
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                                 "column6", "column7", "column8", "column9", "column10",
                                                 "column11", "column12", "column13", "column14", "column15",
                                                 "column16", "column17" ), height=26, show='headings')
        scrollbary.config(command=tree.yview)
        scrollbary.pack(side=RIGHT, fill=Y)
        scrollbarx.config(command=tree.xview)
        scrollbarx.pack(side=BOTTOM, fill=X)
        tree.heading("#1", text="CaseSrNo", anchor=W)
        tree.heading("#2", text="DeviceType", anchor=W)
        tree.heading("#3", text="ProjectID", anchor=W)
        tree.heading("#4", text="CpuSerialNumber", anchor=W)
        tree.heading("#5", text="BootVersion", anchor=W)            
        tree.heading("#6", text="ApplicationVersion", anchor=W)
        tree.heading("#7", text="FlashSerialNumber" ,anchor=W)
        tree.heading("#8", text="FlashCapacityGB", anchor=W)
        tree.heading("#9", text="LastTimeSeenInDTMDt", anchor=W)
        tree.heading("#10", text="LastTimeLineViewedDt", anchor=W)
        tree.heading("#11", text="LastTimeReapedDt" ,anchor=W)
        tree.heading("#12", text="LastTimeTestedDt", anchor=W)
        tree.heading("#13", text="FirstTimeScriptedDt", anchor=W)
        tree.heading("#14", text="InitialScript", anchor=W)
        tree.heading("#15", text="LastTimeScriptedDt", anchor=W)        
        tree.heading("#16", text="CurrentScript", anchor=W)
        tree.heading("#17", text="DuplicatedEntries", anchor=W)        
        tree.column('#1', stretch=NO, minwidth=0, width=60)            
        tree.column('#2', stretch=NO, minwidth=0, width=70)
        tree.column('#3', stretch=NO, minwidth=0, width=55)
        tree.column('#4', stretch=NO, minwidth=0, width=80)
        tree.column('#5', stretch=NO, minwidth=0, width=60)
        tree.column('#6', stretch=NO, minwidth=0, width=60)
        tree.column('#7', stretch=NO, minwidth=0, width=80)
        tree.column('#8', stretch=NO, minwidth=0, width=80)
        tree.column('#9', stretch=NO, minwidth=0, width=100)
        tree.column('#10', stretch=NO, minwidth=0, width=100)
        tree.column('#11', stretch=NO, minwidth=0, width=100)
        tree.column('#12', stretch=NO, minwidth=0, width=100)
        tree.column('#13', stretch=NO, minwidth=0, width=100)
        tree.column('#14', stretch=NO, minwidth=0, width=80)
        tree.column('#15', stretch=NO, minwidth=0, width=100)
        tree.column('#16', stretch=NO, minwidth=0, width=80)
        tree.column('#17', stretch=NO, minwidth=0, width=80)
        tree.pack()
        
### All Functions defining 
        
        self.df = None
        Bad_Timestamp_CSV = '1900/1/00 00:00'
        Bad_Timestamp_Excel = datetime.time(0, 0)
        Fixed_Timestamp   = '1900/1/01 00:00'

        def iExit():
            iExit= tkinter.messagebox.askyesno("Eagle GSR Inventory Management System", "Confirm if you want to exit")
            if iExit >0:
                self.root.destroy()
                return

        def ResetCount():
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)

        def ClearView():
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            tree.delete(*tree.get_children())

        def ClearMasterDB():
            iDelete = tkinter.messagebox.askyesno("Delete Master DB", "Confirm if you want to Clear Master GSR Inventory DB and Start Again")
            if iDelete >0:
                ClearView()
                conn = sqlite3.connect("Eagle_GSRInventory.db")
                cur = conn.cursor()
                cur.execute("DELETE FROM Eagle_GSRInventory_TEMP")
                cur.execute("DELETE FROM Eagle_GSRInventory_ANALYZED_TEMP")
                cur.execute("DELETE FROM Eagle_GSRInventory")
                conn.commit()
                conn.close()
                return

        def TotalEntries():
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            TotalEntries = len(data)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)              
            conn.commit()
            conn.close()
            
        def DeleteSelectedImportData():
            iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
            if iDelete >0:
                self.txtTotalEntries.delete(0,END)
                self.txtValidEntries.delete(0,END)
                self.txtDuplicatedEntries.delete(0,END)
                conn = sqlite3.connect("Eagle_GSRInventory.db")
                cur = conn.cursor()                
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_GSRInventory_TEMP WHERE CaseSrNo =? AND DeviceType=? AND \
                                LastTimeSeenInDTMDt=? AND LastTimeLineViewedDt=? AND LastTimeReapedDt=? ",\
                                (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#9'),\
                                 tree.set(selected_item, '#10'),tree.set(selected_item, '#11'),)) 
                    conn.commit()
                    tree.delete(selected_item)
                conn.commit()
                conn.close()
                TotalEntries()
                return

        def AnalyzeImport():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)
            self.cur=conn.cursor()                
            data.to_sql('Eagle_GSRInventory_ANALYZED_TEMP',conn, if_exists="replace", index=False)            
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalDF = pd.DataFrame(Complete_df)
            TotalEntries = len(TotalDF)
            ValidEntries = len(data)
            DuplicatedEntries = len(TotalDF)-len(data)            
            self.txtTotalEntries.insert(tk.END,TotalEntries)
            self.txtValidEntries.insert(tk.END,ValidEntries)
            self.txtDuplicatedEntries.insert(tk.END,DuplicatedEntries)            
            tkinter.messagebox.showinfo("Analyze Complete","Invalid and Duplicated Entries are Removed")
            conn.commit()
            conn.close()

        def ViewDuplicateEntries():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data = data.loc[data.DuplicatedEntries == True, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)                     
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            Duplicate_count = len(tree.get_children())
            self.txtDuplicatedEntries.insert(tk.END,Duplicate_count)
            conn.commit()
            conn.close()

        def ViewTotalImport():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalEntries = len(data)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)              
            conn.commit()
            conn.close()

        def ViewAnalyzeValidEntries():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_ANALYZED_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            ValidEntries = len(data)
            self.txtValidEntries.insert(tk.END,ValidEntries)
            conn.commit()
            conn.close()

        def UpdateDuplicateMasterDB():
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            data.to_sql('Eagle_GSRInventory',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()
            
            
        def SubmitToMasterDB():
            iSubmit = tkinter.messagebox.askyesno("Valid Entries Submit to Master DB", "Confirm if you want to Submit the Analyzed Valid Entries to Master DB")
            if iSubmit >0:
                tree.delete(*tree.get_children())
                conn = sqlite3.connect("Eagle_GSRInventory.db")
                cur=conn.cursor()
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_ANALYZED_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data.to_sql('Eagle_GSRInventory',conn, if_exists="append", index=False)
                cur.execute("DELETE FROM Eagle_GSRInventory_TEMP")
                cur.execute("DELETE FROM Eagle_GSRInventory_ANALYZED_TEMP")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Submit Complete","All Valid Import Entries are Submitted to Master DB")
                UpdateDuplicateMasterDB()
                return
                                                       
        def ImportGSRInventoryFile():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            fileList = askopenfilenames(filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, header = None, skiprows = {0})
                    else:
                        df = pd.read_excel(filename, header = None, skiprows = {0})
                    dfList.append(df)
                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns = {0:'CaseSrNo', 1:'DeviceType', 2:'ProjectID', 3:'CpuSerialNumber', 4:'BootVersion',
                                          5: 'ApplicationVersion', 6:'FlashSerialNumber', 7:'FlashCapacityGB', 8:'LastTimeSeenInDTMDt',
                                          9:'LastTimeLineViewedDt', 10:'LastTimeReapedDt',11:'LastTimeTestedDt',
                                          12:'FirstTimeScriptedDt',13:'InitialScript',14:'LastTimeScriptedDt',
                                          15:'CurrentScript'},inplace = True)

                self.df = pd.DataFrame(concatDf)
                self.df["InitialScript"].fillna("Unknown", inplace = True)
                self.df["CurrentScript"].fillna("Unknown", inplace = True)
                self.df["DeviceType"].fillna("Unknown", inplace = True)
                self.df["ProjectID"].fillna("Unknown", inplace = True)
                self.df["CpuSerialNumber"].fillna("Unknown", inplace = True)
                self.df["BootVersion"].fillna("Unknown", inplace = True)
                self.df["ApplicationVersion"].fillna("Unknown", inplace = True)
                self.df["FlashSerialNumber"].fillna("Unknown", inplace = True)
                self.df["FlashCapacityGB"].fillna("Unknown", inplace = True)
                self.df["ProjectID"].fillna("Unknown", inplace = True)
                self.df["LastTimeSeenInDTMDt"].fillna(Fixed_Timestamp, inplace = True)
                self.df["LastTimeLineViewedDt"].fillna(Fixed_Timestamp, inplace = True)
                self.df["LastTimeReapedDt"].fillna(Fixed_Timestamp, inplace = True)
                self.df["LastTimeTestedDt"].fillna(Fixed_Timestamp, inplace = True)
                self.df["FirstTimeScriptedDt"].fillna(Fixed_Timestamp, inplace = True)
                self.df["LastTimeScriptedDt"].fillna(Fixed_Timestamp, inplace = True)

                def trans_TimeFixCSV(x):
                    if x == Bad_Timestamp_CSV:
                        return Fixed_Timestamp
                    else:
                        return x

                def trans_TimeFixExcel(y):
                    if y == Bad_Timestamp_Excel:
                        return Fixed_Timestamp
                    else:
                        return y

                self.df['LastTimeLineViewedDt']  = self.df['LastTimeLineViewedDt'].apply(trans_TimeFixCSV)
                self.df['LastTimeSeenInDTMDt']   = self.df['LastTimeSeenInDTMDt'].apply(trans_TimeFixCSV)
                self.df['LastTimeReapedDt']      = self.df['LastTimeReapedDt'].apply(trans_TimeFixCSV)
                self.df['LastTimeTestedDt']      = self.df['LastTimeTestedDt'].apply(trans_TimeFixCSV)
                self.df['FirstTimeScriptedDt']   = self.df['FirstTimeScriptedDt'].apply(trans_TimeFixCSV)
                self.df['LastTimeScriptedDt']    = self.df['LastTimeScriptedDt'].apply(trans_TimeFixCSV)
                self.df['LastTimeLineViewedDt']  = self.df['LastTimeLineViewedDt'].apply(trans_TimeFixExcel)
                self.df['LastTimeSeenInDTMDt']   = self.df['LastTimeSeenInDTMDt'].apply(trans_TimeFixExcel)
                self.df['LastTimeReapedDt']      = self.df['LastTimeReapedDt'].apply(trans_TimeFixExcel)
                self.df['LastTimeTestedDt']      = self.df['LastTimeTestedDt'].apply(trans_TimeFixExcel)
                self.df['FirstTimeScriptedDt']   = self.df['FirstTimeScriptedDt'].apply(trans_TimeFixExcel)
                self.df['LastTimeScriptedDt']    = self.df['LastTimeScriptedDt'].apply(trans_TimeFixExcel)
                self.df['LastTimeLineViewedDt'] = pd.to_datetime(self.df['LastTimeLineViewedDt']).dt.strftime('%Y-%m-%d')
                self.df['LastTimeSeenInDTMDt'] = pd.to_datetime(self.df['LastTimeSeenInDTMDt']).dt.strftime('%Y-%m-%d')                    
                self.df['LastTimeReapedDt'] = pd.to_datetime(self.df['LastTimeReapedDt']).dt.strftime('%Y-%m-%d')
                self.df['LastTimeTestedDt'] = pd.to_datetime(self.df['LastTimeTestedDt']).dt.strftime('%Y-%m-%d')
                self.df['FirstTimeScriptedDt'] = pd.to_datetime(self.df['FirstTimeScriptedDt']).dt.strftime('%Y-%m-%d')
                self.df['LastTimeScriptedDt'] = pd.to_datetime(self.df['LastTimeScriptedDt']).dt.strftime('%Y-%m-%d')                                    
                data = pd.DataFrame(self.df)
                data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
                for each_rec in range(len(data)):
                    tree.insert("", tk.END, values=list(data.loc[each_rec]))            
                con= sqlite3.connect("Eagle_GSRInventory.db")
                self.cur=con.cursor()                
                data.to_sql('Eagle_GSRInventory_TEMP',con, if_exists="replace", index=False)
                TotalEntries = len(data)       
                self.txtTotalEntries.insert(tk.END,TotalEntries)  
                con.commit()
                con.close()

        def ExportValidEntries():
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_ANALYZED_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortLastTimeSeenInDTMDt = pd.DataFrame(Complete_df)
            data_SortLastTimeSeenInDTMDt = data_SortLastTimeSeenInDTMDt.sort_values(by =['LastTimeSeenInDTMDt'])
            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortLastTimeSeenInDTMDt.to_excel(file,sheet_name='SortByLastTimeSeen',index=False)
                    file.close
                    tkinter.messagebox.showinfo("Inventory Export","Inventory Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def ExportMasterDB():
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortLastTimeSeenInDTMDt = pd.DataFrame(Complete_df)
            data_SortLastTimeSeenInDTMDt = data_SortLastTimeSeenInDTMDt.sort_values(by =['LastTimeSeenInDTMDt'])
            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortLastTimeSeenInDTMDt.to_excel(file,sheet_name='SortByLastTimeSeen',index=False)
                    file.close
                    tkinter.messagebox.showinfo("Inventory Export","Inventory Report Saved as Excel")                                        
            conn.commit()
            conn.close()

            
        
### Entry Wizard
        self.txtTotalEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
        self.txtTotalEntries.place(x=1110,y=6)

        self.txtValidEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
        self.txtValidEntries.place(x=167,y=6)

        self.txtDuplicatedEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
        self.txtDuplicatedEntries.place(x=744,y=6)

### Button Wizard  
        btnImport = Button(self.root, text="Import GSR Inventory Files", font=('aerial', 9, 'bold'), height =1, width=22, bd=4, command = ImportGSRInventoryFile)
        btnImport.place(x=2,y=620)
        btnAnalyzeImport = Button(self.root, text="Analyze Imported Inventory Files ", font=('aerial', 9, 'bold'), height =1, width=26, bd=4, command = AnalyzeImport)
        btnAnalyzeImport.place(x=172,y=620)
        btnAnalyzeSubmit = Button(self.root, text="Submit Analyzed Valid Entries To Master DB", font=('aerial', 9, 'bold'), height =1, width=35, bd=4, command = SubmitToMasterDB)
        btnAnalyzeSubmit.place(x=370,y=620)
        btnExportMasterDBValidEntries = Button(self.root, text="Export Master DB", font=('aerial', 9, 'bold'), height =1, width=16, bd=4, command = ExportMasterDB)
        btnExportMasterDBValidEntries.place(x=631,y=620)

        btnExit = Button(self.root, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
        btnExit.place(x=1267,y=620)
        btnClearView = Button(self.root, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
        btnClearView.place(x=1181,y=620)
        btnClearDB = Button(self.root, text="Clear Master DB", font=('aerial', 9, 'bold'), height =1, width=13, bd=4, command = ClearMasterDB)
        btnClearDB.place(x=1073,y=620)
        btnDelete = Button(self.root, text="Delete Selected Import Entries", font=('aerial', 9, 'bold'), height =1, width=24, bd=4, command = DeleteSelectedImportData)
        btnDelete.place(x=888,y=620)

        btnAnalyzedValidView = Button(self.root, text="View Analyzed Valid Entries", font=('aerial', 9, 'bold'), height =1, width=22, bd=1, command = ViewAnalyzeValidEntries)
        btnAnalyzedValidView.place(x=2,y=6)
        btnExportAnalyzedValidView = Button(self.root, text="Export Analyzed Valid Entries", font=('aerial', 9, 'bold'), height =1, width=24, bd=1, command = ExportValidEntries)
        btnExportAnalyzedValidView.place(x=252,y=6)
        btnViewDuplicateEntries = Button(self.root, text="View Duplicate Entries", font=('aerial', 9, 'bold'), height =1, width=19, bd=1, command = ViewDuplicateEntries)
        btnViewDuplicateEntries.place(x=600,y=6)
        btnViewTotalImport = Button(self.root, text="View Total Import", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
        btnViewTotalImport.place(x=995,y=6)

        btnResetTotal = Button(self.root, text="Reset Count", font=('aerial', 9, 'bold'), height =1, width=10, bd=1, command = ResetCount)
        btnResetTotal.place(x=1267,y=6)

   
if __name__ == '__main__':
    root = Tk()
    application  = GSRInventoryImport(root)
    root.mainloop()
