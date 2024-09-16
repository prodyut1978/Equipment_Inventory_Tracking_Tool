#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_GSRInventory_BackEnd
import Eagle_GSRMergedInventory_BackEnd
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

class MergedGSRInventoryRepairQuery:    
    def __init__(self,root):
        
##  ----------------- Define Variables-------------
        
        CaseSrNo    = IntVar()
        DeviceType  = IntVar()
        DeviceType1 = IntVar()
        DeviceType2 = IntVar()
        ProjectID   = StringVar()
        List_Year   = {2015,2016,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025}
        List_Month  = {1,2,3,4,5,6,7,8,9,10,11,12}
        List_Day    = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31}

##  ----------------- Define Event Treeview-------------
        def InventoryRec(event):
            for nm in tree.selection():
                sd = tree.item(nm, 'values')

##  ----------------- Define TreeView Window-------------
        self.root =root
        self.root.title ("Eagle Merged GSR Inventory Query Wizard")
        self.root.geometry("1350x665+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)        
        TableMargin = Frame(self.root,  bd = 2, padx= 10, pady= 10, relief = RIDGE)
        TableMargin.place(x=2, y=120, anchor="nw", width=1350, height=510)
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                                 "column6", "column7", "column8", "column9", "column10",
                                                 "column11", "column12", "column13", "column14", "column15",
                                                 "column16", "column17"), height=35, show='headings')
        scrollbary.config(command=tree.yview)
        scrollbary.pack(side=RIGHT, fill=Y)
        scrollbarx.config(command=tree.xview)
        scrollbarx.pack(side=BOTTOM, fill=X)
        tree.heading("#1", text="CaseSrNo", anchor=W)
        tree.heading("#2", text="DeviceType", anchor=W)        
        tree.heading("#3", text="ProjectID", anchor=W)
        tree.heading("#4", text="FlashCapacityGB", anchor=W)
        tree.heading("#5", text="LastTimeSeenInDTMDt", anchor=W)
        tree.heading("#6", text="LastTimeLineViewedDt", anchor=W)
        tree.heading("#7", text="LastTimeReapedDt", anchor=W)
        tree.heading("#8", text="FlagsRepair" ,anchor=W)        
        tree.heading("#9", text="WorkOrderNo", anchor=W)
        tree.heading("#10", text="PartNo", anchor=W)
        tree.heading("#11", text="TechnicianInput", anchor=W)
        tree.heading("#12", text="CrewReported", anchor=W)
        tree.heading("#13", text="DateRepaired", anchor=W)
        tree.heading("#14", text="FlagsDeployment", anchor=W)
        tree.heading("#15", text="StartTimeUTC", anchor=W)
        tree.heading("#16", text="EndTimeUTC", anchor=W)
        tree.heading("#17", text="JobName", anchor=W)
        
        tree.column('#1', stretch=NO, minwidth=0, width=70)            
        tree.column('#2', stretch=NO, minwidth=0, width=80)
        tree.column('#3', stretch=NO, minwidth=0, width=60)
        tree.column('#4', stretch=NO, minwidth=0, width=80)
        tree.column('#5', stretch=NO, minwidth=0, width=140)
        tree.column('#6', stretch=NO, minwidth=0, width=140)
        tree.column('#7', stretch=NO, minwidth=0, width=110)
        tree.column('#8', stretch=NO, minwidth=0, width=80)
        tree.column('#9', stretch=NO, minwidth=0, width=80)
        tree.column('#10', stretch=NO, minwidth=0, width=68)
        tree.column('#11', stretch=NO, minwidth=0, width=80)
        tree.column('#12', stretch=NO, minwidth=0, width=80)
        tree.column('#13', stretch=NO, minwidth=0, width=90)
        tree.column('#14', stretch=NO, minwidth=0, width=100)
        tree.column('#15', stretch=NO, minwidth=0, width=80)
        tree.column('#16', stretch=NO, minwidth=0, width=80)
        tree.column('#17', stretch=NO, minwidth=0, width=90)
        tree.pack()
        tree.bind('<<TreeviewSelect>>',InventoryRec)

### All Functions defining 

        def iExitGSRMerged():
            iExit= tkinter.messagebox.askyesno("Eagle Merged GSR Inventory Query Wizard", "Confirm if you want to exit")
            if iExit >0:
                self.root.destroy()
                return

        def GSRMergedInvBeforeSearch_A():
            DeviceType1 = (self.txtBeforeDeviceType.get())
            try:
                self.txtAfterYear.delete(0,END)
                self.txtAfterMonth.delete(0,END)
                self.txtAfterDay.delete(0,END)
                self.txtAfterDeviceType.delete(0,END)
                self.txtAfterDeviceTypeName.delete(0,END)
                self.txtCaseSrNo.delete(0,END)
                self.txtDeviceType.delete(0,END)
                self.txtProjectID.delete(0,END)
                self.txtNumberofSearch.delete(0,END)
                tree.delete(*tree.get_children())
                BeforedateSearch_A = datetime.date(int(self.txtBeforeYear.get()),int(self.txtBeforeMonth.get()), int(self.txtBeforeDay.get()))
                conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data['LastTimeSeenInDTMDt'] = pd.to_datetime(data['LastTimeSeenInDTMDt']).dt.strftime('%Y-%m-%d')
                BeforedateSearch_A = pd.to_datetime(BeforedateSearch_A).strftime('%Y-%m-%d')
                data = data[data['LastTimeSeenInDTMDt'] <= BeforedateSearch_A]
                data = data.reset_index(drop=True)
                if (DeviceType1 != ""):
                    data = data[data['DeviceType'] == DeviceType1]
                    data = data.reset_index(drop=True)
                    TotalSearchEntries = len(data)
                    self.txtNumberofSearch.insert(tk.END,TotalSearchEntries)
                    for each_rec in range(len(data)):
                        tree.insert("", tk.END, values=list(data.loc[each_rec]))
                    conn.commit()
                    conn.close()
                else:
                    TotalSearchEntries = len(data)
                    self.txtNumberofSearch.insert(tk.END,TotalSearchEntries)
                    for each_rec in range(len(data)):
                        tree.insert("", tk.END, values=list(data.loc[each_rec]))
                    conn.commit()
                    conn.close()
                                                        
            except:
                tkinter.messagebox.showerror("Search Input Error", "Please Input Valid Date to Search")


        def GSRMergedInvAfterSearch_B():
            DeviceType2 = (self.txtAfterDeviceType.get())
            try:
                self.txtBeforeYear.delete(0,END)
                self.txtBeforeMonth.delete(0,END)
                self.txtBeforeDay.delete(0,END)
                self.txtBeforeDeviceType.delete(0,END)
                self.txtBeforeDeviceTypeName.delete(0,END)
                self.txtCaseSrNo.delete(0,END)
                self.txtDeviceType.delete(0,END)
                self.txtProjectID.delete(0,END)
                self.txtNumberofSearch.delete(0,END)
                tree.delete(*tree.get_children())
                AfterdateSearch_B = datetime.date(int(self.txtAfterYear.get()),int(self.txtAfterMonth.get()), int(self.txtAfterDay.get()))
                conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data['LastTimeSeenInDTMDt'] = pd.to_datetime(data['LastTimeSeenInDTMDt']).dt.strftime('%Y-%m-%d')
                AfterdateSearch_B = pd.to_datetime(AfterdateSearch_B).strftime('%Y-%m-%d')
                data = data[data['LastTimeSeenInDTMDt'] >= AfterdateSearch_B]
                data = data.reset_index(drop=True)
                if (DeviceType2 != ""):
                    data = data[data['DeviceType'] == DeviceType2]
                    data = data.reset_index(drop=True)
                    TotalSearchEntries = len(data)
                    self.txtNumberofSearch.insert(tk.END,TotalSearchEntries)
                    for each_rec in range(len(data)):
                        tree.insert("", tk.END, values=list(data.loc[each_rec]))
                    conn.commit()
                    conn.close()
                else:
                    TotalSearchEntries = len(data)
                    self.txtNumberofSearch.insert(tk.END,TotalSearchEntries)
                    for each_rec in range(len(data)):
                        tree.insert("", tk.END, values=list(data.loc[each_rec]))
                    conn.commit()
                    conn.close()
                
            except:
                tkinter.messagebox.showerror("Search Input Error", "Please Input Valid Date to Search")

        def GSRMergedInvSearch_C():
            if (self.txtCaseSrNo.get() != "") | (self.txtDeviceType.get() != "")| (self.txtProjectID.get() != ""):
                tree.delete(*tree.get_children())
                self.txtNumberofSearch.delete(0,END)
                self.txtAfterYear.delete(0,END)
                self.txtAfterMonth.delete(0,END)
                self.txtAfterDay.delete(0,END)
                self.txtBeforeYear.delete(0,END)
                self.txtBeforeMonth.delete(0,END)
                self.txtBeforeDay.delete(0,END)
                self.txtBeforeDeviceType.delete(0,END)
                self.txtAfterDeviceType.delete(0,END)
                self.txtBeforeDeviceTypeName.delete(0,END)
                self.txtAfterDeviceTypeName.delete(0,END)
                conn= sqlite3.connect("Eagle_GSRMergedInventory.db")
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM `Eagle_GSRMergedInventory_MASTER` WHERE `CaseSrNo`= ? COLLATE NOCASE OR DeviceType = ? COLLATE NOCASE OR ProjectID = ? ",\
                               (self.txtCaseSrNo.get(), self.txtDeviceType.get(), self.txtProjectID.get(),))                
                fetch = cursor.fetchall()
                TotalSearchEntries = len(fetch)
                self.txtNumberofSearch.insert(tk.END,TotalSearchEntries)
                for data in fetch:
                    tree.insert('', 'end', values=(data))
                cursor.close()
                conn.close()
            else:
                tkinter.messagebox.showinfo("Search Error","Please Select CaseSrNo or Device Type or Project ID to Search")
                

        def ExportGSRMergeSearch():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            search_DF = pd.DataFrame(dfList)
            search_DF.rename(columns = {0:'CaseSrNo', 1:'DeviceType', 2:'ProjectID', 3:'FlashCapacityGB', 4:'LastTimeSeenInDTMDt',
                                          5: 'LastTimeLineViewedDt', 6:'LastTimeReapedDt', 7:'Flags', 8:'WorkOrderNo',
                                          9:'PartNo', 10:'TechnicianInput', 11:'CrewReported', 12:'DateRepaired'},inplace = True)
            data_SortByCaseSrNo = search_DF.sort_values(by =['CaseSrNo'])
            data_SortLastTimeSeenInDTMDt = search_DF.sort_values(by =['LastTimeSeenInDTMDt'])
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortLastTimeSeenInDTMDt.to_excel(file,sheet_name='SortByLastTimeSeen',index=False)
                    file.close
                    tkinter.messagebox.showinfo("Search Results Export"," Search Query Report Saved as Excel")
            tree.delete(*tree.get_children())
                        
        def GSRMergedInvSummary_D():
            self.txtDateFrom.delete(0,END)
            self.txtTotal.delete(0,END)
            self.txtDuplicated.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            DateFromDF = data.sort_values(by =['LastTimeSeenInDTMDt'])
            DateFrom   = (DateFromDF['LastTimeSeenInDTMDt'].min())
            TotalEntries = len(data)
            DuplicatedDF = data.loc[data.DuplicatedEntries == True, 'CaseSrNo': 'JobName']
            DuplicatedDF = DuplicatedDF.reset_index(drop=True)
            DuplicatedEntries = len(DuplicatedDF)            
            self.txtDateFrom.insert(tk.END,DateFrom)
            self.txtTotal.insert(tk.END,TotalEntries)
            self.txtDuplicated.insert(tk.END,DuplicatedEntries)
            conn.commit()
            conn.close()

        def ClearGSRMergedInvSummary_D():
            self.txtDateFrom.delete(0,END)
            self.txtTotal.delete(0,END)
            self.txtDuplicated.delete(0,END)

        def ViewGSRMergedDuplicateEntries_D():
            tree.delete(*tree.get_children())
            self.txtDuplicated.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            data = data.loc[data.DuplicatedEntries == True, 'CaseSrNo': 'JobName']
            data = data.reset_index(drop=True)                     
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            DuplicatedEntries = len(tree.get_children())
            self.txtDuplicated.insert(tk.END,DuplicatedEntries)
            conn.commit()
            conn.close()
                
        
        def callbackFuncCaseSrNo(event):
            CaseSrNo = (self.txtCaseSrNo.get())
            print('Selected CaseSrNo:'+ CaseSrNo)


        def callbackFuncDeviceType(event):
            DeviceType = (self.txtDeviceType.get())                
            print('Selected DeviceType:'+ DeviceType)

        def callbackFuncDeviceType1(event):
            DeviceType1 = (self.txtBeforeDeviceType.get())
            if DeviceType1 == '279':
                Dev_Name = 'GSRx3'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '273':
                Dev_Name = 'SDRx'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '270':
                Dev_Name = 'SDR'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '266':
                Dev_Name = 'GSIx'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '264':
                Dev_Name = 'GSRx1'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '263':
                Dev_Name = 'GSRx4'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '257':
                Dev_Name = 'GSR4'
                self.txtBeforeDeviceTypeName.delete(0,END)
            elif DeviceType1 == '256':
                Dev_Name = 'GSR1'
                self.txtBeforeDeviceTypeName.delete(0,END)
            else:
                Dev_Name = 'Unknown'
                self.txtBeforeDeviceTypeName.delete(0,END)
                
            self.txtBeforeDeviceTypeName.insert(tk.END,Dev_Name)
            print('Selected DeviceType:'+ DeviceType1)

        def callbackFuncDeviceType2(event):
            DeviceType2 = (self.txtAfterDeviceType.get())
            if DeviceType2 == '279':
                Dev_Name = 'GSRx3'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '273':
                Dev_Name = 'SDRx'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '270':
                Dev_Name = 'SDR'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '266':
                Dev_Name = 'GSIx'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '264':
                Dev_Name = 'GSRx1'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '263':
                Dev_Name = 'GSRx4'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '257':
                Dev_Name = 'GSR4'
                self.txtAfterDeviceTypeName.delete(0,END)
            elif DeviceType2 == '256':
                Dev_Name = 'GSR1'
                self.txtAfterDeviceTypeName.delete(0,END)
            else:
                Dev_Name = 'Unknown'
                self.txtAfterDeviceTypeName.delete(0,END)

            self.txtAfterDeviceTypeName.insert(tk.END,Dev_Name)
            print('Selected DeviceType:'+ DeviceType2)

        def callbackFuncDeviceType3(event):
            DeviceType3 = (self.txtDeviceType.get())
            if DeviceType3 == '279':
                Dev_Name = 'GSRx3'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '273':
                Dev_Name = 'SDRx'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '270':
                Dev_Name = 'SDR'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '266':
                Dev_Name = 'GSIx'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '264':
                Dev_Name = 'GSRx1'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '263':
                Dev_Name = 'GSRx4'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '257':
                Dev_Name = 'GSR4'
                self.DeviceTypeNameQuick.delete(0,END)
            elif DeviceType3 == '256':
                Dev_Name = 'GSR1'
                self.DeviceTypeNameQuick.delete(0,END)
            else:
                Dev_Name = 'Unknown'
                self.DeviceTypeNameQuick.delete(0,END)
                
            self.DeviceTypeNameQuick.insert(tk.END,Dev_Name)
            print('Selected DeviceType:'+ DeviceType3)



        def callbackFuncProjectID(event):
            ProjectID = (self.txtProjectID.get())
            print('Selected ProjectID:'+ ProjectID)

        def callbackBeforeYear(event):
            BeforeYear = (self.txtBeforeYear.get())

        def callbackBeforeMonth(event):
            BeforeMonth = (self.txtBeforeMonth.get())

        def callbackBeforeDay(event):
            BeforeDay = (self.txtBeforeDay.get())

        def callbackAfterYear(event):
            AfterYear = (self.txtAfterYear.get())

        def callbackAfterMonth(event):
            AfterMonth = (self.txtAfterMonth.get())

        def callbackAfterDay(event):
            AfterDay = (self.txtAfterDay.get())       

        def ClearMergedGSRView():
            self.txtBeforeYear.delete(0,END)
            self.txtBeforeMonth.delete(0,END)
            self.txtBeforeDay.delete(0,END)
            self.txtAfterYear.delete(0,END)
            self.txtAfterMonth.delete(0,END)
            self.txtAfterDay.delete(0,END)
            self.txtCaseSrNo.delete(0,END)
            self.txtDeviceType.delete(0,END)
            self.txtProjectID.delete(0,END)
            self.txtDateFrom.delete(0,END)
            self.txtTotal.delete(0,END)
            self.txtDuplicated.delete(0,END)
            self.txtNumberofSearch.delete(0,END)            
            self.txtBeforeDeviceType.delete(0,END)
            self.txtAfterDeviceType.delete(0,END)
            self.txtBeforeDeviceTypeName.delete(0,END)
            self.txtAfterDeviceTypeName.delete(0,END)
            tree.delete(*tree.get_children())

        def ResetSearchA():
            tree.delete(*tree.get_children())
            self.txtBeforeYear.delete(0,END)
            self.txtBeforeMonth.delete(0,END)
            self.txtBeforeDay.delete(0,END)
            self.txtBeforeDeviceType.delete(0,END)
            self.txtAfterYear.delete(0,END)
            self.txtAfterMonth.delete(0,END)
            self.txtAfterDay.delete(0,END)
            self.txtAfterDeviceType.delete(0,END)
            self.txtCaseSrNo.delete(0,END)
            self.txtDeviceType.delete(0,END)
            self.txtProjectID.delete(0,END)
            self.txtNumberofSearch.delete(0,END)
            self.txtBeforeDeviceTypeName.delete(0,END)
            self.txtAfterDeviceTypeName.delete(0,END)

        def ResetSearchB():
            tree.delete(*tree.get_children())
            self.txtAfterYear.delete(0,END)
            self.txtAfterMonth.delete(0,END)
            self.txtAfterDay.delete(0,END)
            self.txtAfterDeviceType.delete(0,END)
            self.txtBeforeYear.delete(0,END)
            self.txtBeforeMonth.delete(0,END)
            self.txtBeforeDay.delete(0,END)
            self.txtBeforeDeviceType.delete(0,END)
            self.txtCaseSrNo.delete(0,END)
            self.txtDeviceType.delete(0,END)
            self.txtProjectID.delete(0,END)
            self.txtNumberofSearch.delete(0,END)
            self.txtBeforeDeviceTypeName.delete(0,END)
            self.txtAfterDeviceTypeName.delete(0,END)

        def ResetSearchC():
            tree.delete(*tree.get_children())
            self.txtCaseSrNo.delete(0,END)
            self.txtDeviceType.delete(0,END)
            self.txtProjectID.delete(0,END)
            self.txtBeforeYear.delete(0,END)
            self.txtBeforeMonth.delete(0,END)
            self.txtBeforeDay.delete(0,END)
            self.txtBeforeDeviceType.delete(0,END)
            self.txtAfterYear.delete(0,END)
            self.txtAfterMonth.delete(0,END)
            self.txtAfterDay.delete(0,END)
            self.txtAfterDeviceType.delete(0,END)
            self.txtNumberofSearch.delete(0,END)
            self.txtBeforeDeviceTypeName.delete(0,END)
            self.txtAfterDeviceTypeName.delete(0,END)
            
        def ExportMasterMergedDB():
            conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
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
                    tkinter.messagebox.showinfo("Master Merged Inventory Export"," Master Merged Inventory Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def UpdateMasterMergedDB():
            iUpdate = tkinter.messagebox.askyesno("Update Master Merged DB", "Confirm if you want to Update Master Merged DB Removing Duplicated CaseSrNo with Older Dates")
            if iUpdate >0:
                tree.delete(*tree.get_children())
                self.txtDateFrom.delete(0,END)
                self.txtTotal.delete(0,END)
                self.txtDuplicated.delete(0,END)  
                conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
                data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'JobName']
                data = data.reset_index(drop=True)
                self.cur=conn.cursor()
                data.to_sql('Eagle_GSRMergedInventory_MASTER',conn, if_exists="replace", index=False)
                for each_rec in range(len(data)):
                    tree.insert("", tk.END, values=list(data.loc[each_rec]))
                tkinter.messagebox.showinfo("Eagle GSR Merged Inventory Update Complete","Old Duplicated Entries are Removed and Replaced by Newer One")
                conn.commit()
                conn.close()
                GSRMergedInvSummary_D()

        def ViewMasterMergedDB():
            ClearMergedGSRView
            conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRMergedInventory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            conn.commit()
            conn.close()
            GSRMergedInvSummary_D()

        def DeleteMasterMergedSelectedData():
            iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
            if iDelete >0:
                self.txtDateFrom.delete(0,END)
                self.txtTotal.delete(0,END)
                self.txtDuplicated.delete(0,END)
                conn = sqlite3.connect("Eagle_GSRMergedInventory.db")
                cur = conn.cursor()
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_GSRMergedInventory_MASTER WHERE CaseSrNo =? AND DeviceType=? AND \
                                LastTimeSeenInDTMDt=? AND LastTimeLineViewedDt=? AND LastTimeReapedDt=? ",\
                                (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#5'),\
                                 tree.set(selected_item, '#6'),tree.set(selected_item, '#7'),))                                 
                    conn.commit()
                    tree.delete(selected_item)
                conn.close()
                GSRMergedInvSummary_D()
                return



### Labeling Windows 

        InvL1 = Label(self.root, text = "A: Search Before Date:", font=("arial", 10,'bold'),bg = "green").place(x=6,y=10)
        InvL2 = Label(self.root, text = "1: Year:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=6,y=38)
        InvL3 = Label(self.root, text = "2: Month:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=6,y=65)
        InvL4 = Label(self.root, text = "3: Day:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=6,y=92)
        InvL18 = Label(self.root, text = "4: DeviceType:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=144,y=38)

        InvL5 = Label(self.root, text = "B: Search After Date:", font=("arial", 10,'bold'),bg = "green").place(x=342,y=10)
        InvL6 = Label(self.root, text = "1: Year:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=342,y=38)
        InvL7 = Label(self.root, text = "2: Month:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=342,y=65)
        InvL8 = Label(self.root, text = "3: Day:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=342,y=92)
        InvL19 = Label(self.root, text = "4: DeviceType:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=474,y=38)

        InvL9 = Label(self.root, text = "C: Quick Search:", font=("arial", 10,'bold'),bg = "green").place(x=684,y=10)
        InvL10 = Label(self.root, text = "1: DeviceType:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=684,y=38)
        InvL11 = Label(self.root, text = "2: CaseSrNo:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=684,y=65)
        InvL12 = Label(self.root, text = "3: Project ID:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=684,y=92)

        InvL13 = Label(self.root, text = "D: GSR Merged Master Inventory Summary:", font=("arial", 10,'bold'),bg = "green").place(x=1038,y=10)
        InvL14 = Label(self.root, text = "1: Date From:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=1038,y=38)
        InvL15 = Label(self.root, text = "2: Total:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=1038,y=65)
        InvL16 = Label(self.root, text = "3: Duplicated:", font=("arial", 10,'bold'),bg = "cadet blue").place(x=1038,y=92)

        InvL17 = Label(self.root, text = "Number of Search Items:", font=("arial", 10,'bold'),bg = "green", bd=4).place(x=500,y=632)


### Entry Wizard
        self.txtBeforeYear  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtBeforeYear.place(x=80,y=38)
        self.txtBeforeMonth  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtBeforeMonth.place(x=80,y=65)
        self.txtBeforeDay  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtBeforeDay.place(x=80,y=92)
        self.txtBeforeDeviceType  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=DeviceType1, width = 4)
        self.txtBeforeDeviceType.place(x=252,y=38)
        
        DeviceTypeNameA = StringVar(self.root)
        self.txtBeforeDeviceTypeName  = Entry(self.root, font=('aerial', 10, 'bold'), bd=2,bg = 'cadet blue', textvariable= DeviceTypeNameA, width = 7)
        self.txtBeforeDeviceTypeName.place(x=253,y=65)

        self.txtAfterYear  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtAfterYear.place(x=410,y=38)
        self.txtAfterMonth  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtAfterMonth.place(x=410,y=65)
        self.txtAfterDay  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 4)
        self.txtAfterDay.place(x=410,y=92)
        self.txtAfterDeviceType  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=DeviceType2, width = 4)
        self.txtAfterDeviceType.place(x=580,y=38)

        DeviceTypeNameB = StringVar(self.root)
        self.txtAfterDeviceTypeName  = Entry(self.root, font=('aerial', 10, 'bold'), bd=2, bg = 'cadet blue', textvariable= DeviceTypeNameB, width = 7)
        self.txtAfterDeviceTypeName.place(x=582,y=65)

        
        self.txtDeviceType  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=DeviceType, width = 8)
        self.txtDeviceType.place(x=790,y=38)
        self.txtCaseSrNo  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=CaseSrNo, width = 8)
        self.txtCaseSrNo.place(x=790,y=65)        
        self.txtProjectID  = ttk.Combobox(self.root, font=('aerial', 12, 'bold'),textvariable=ProjectID, width = 8)
        self.txtProjectID.place(x=790,y=92)
        
        DeviceTypeNameC = StringVar(self.root)
        self.DeviceTypeNameQuick  = Entry(self.root, font=('aerial', 10, 'bold'), bd=2, bg = 'cadet blue', textvariable= DeviceTypeNameC, width = 7)
        self.DeviceTypeNameQuick.place(x=910,y=38)
        
        self.txtCaseSrNo['values']     = sorted(list(set(Eagle_GSRMergedInventory_BackEnd.Combo_input_CaseSrNo())))
        self.txtDeviceType['values']   = sorted(list(set(Eagle_GSRMergedInventory_BackEnd.Combo_input_DeviceType())))
        self.txtBeforeDeviceType['values']   = sorted(list(set(Eagle_GSRMergedInventory_BackEnd.Combo_input_DeviceType())))
        self.txtAfterDeviceType['values']   = sorted(list(set(Eagle_GSRMergedInventory_BackEnd.Combo_input_DeviceType())))
        self.txtProjectID['values']    = sorted(list(set(Eagle_GSRInventory_BackEnd.Combo_input_ProjectID())))

        self.txtBeforeYear['values']  = sorted(list(set(List_Year)))
        self.txtBeforeMonth['values'] = sorted(list(set(List_Month)))
        self.txtBeforeDay['values']   = sorted(list(set(List_Day)))

        self.txtAfterYear['values']  = sorted(list(set(List_Year)))
        self.txtAfterMonth['values'] = sorted(list(set(List_Month)))
        self.txtAfterDay['values']   = sorted(list(set(List_Day)))
        
        self.txtCaseSrNo.bind('<<ComboboxSelected>>',callbackFuncCaseSrNo)
        self.txtDeviceType.bind('<<ComboboxSelected>>',callbackFuncDeviceType3)
        self.txtBeforeDeviceType.bind('<<ComboboxSelected>>',callbackFuncDeviceType1)
        self.txtAfterDeviceType.bind('<<ComboboxSelected>>',callbackFuncDeviceType2)        
        self.txtProjectID.bind('<<ComboboxSelected>>',callbackFuncProjectID)

        self.txtBeforeYear.bind('<<ComboboxSelected>>',callbackBeforeYear)
        self.txtBeforeMonth.bind('<<ComboboxSelected>>',callbackBeforeMonth)
        self.txtBeforeDay.bind('<<ComboboxSelected>>',callbackBeforeDay)

        self.txtAfterYear.bind('<<ComboboxSelected>>',callbackAfterYear)
        self.txtAfterMonth.bind('<<ComboboxSelected>>',callbackAfterMonth)
        self.txtAfterDay.bind('<<ComboboxSelected>>',callbackAfterDay)
        
        self.txtDateFrom  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 10)
        self.txtDateFrom.place(x=1138,y=38)
        self.txtTotal  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
        self.txtTotal.place(x=1138,y=65)
        self.txtDuplicated  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
        self.txtDuplicated.place(x=1138,y=92)

        self.txtNumberofSearch  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
        self.txtNumberofSearch.place(x=670,y=632)

### Button Wizard  
        btnSearchBeforeDate = Button(self.root, text="Search Entries A", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command =GSRMergedInvBeforeSearch_A)
        btnSearchBeforeDate.place(x=144,y=65)
        btnClearSearchBeforeDate = Button(self.root, text="Reset Search A", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command =ResetSearchA)
        btnClearSearchBeforeDate.place(x=144,y=92)

        btnSearchAfterDate = Button(self.root, text="Search Entries B", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = GSRMergedInvAfterSearch_B)
        btnSearchAfterDate.place(x=474,y=65)
        btnClearSearchAfterDate = Button(self.root, text="Reset Search B", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command =ResetSearchB)
        btnClearSearchAfterDate.place(x=474,y=92)

        btnSearchGen = Button(self.root, text="Search Entries C", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = GSRMergedInvSearch_C)
        btnSearchGen.place(x=890,y=65)
        btnClearSearchGen = Button(self.root, text="Reset Search C", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ResetSearchC)
        btnClearSearchGen.place(x=890,y=92)

        btnGSRMergedInvSummary = Button(self.root, text="View Summary", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = GSRMergedInvSummary_D)
        btnGSRMergedInvSummary.place(x=1240,y=38)
        btnClearGSRMergedInvSummary = Button(self.root, text="Clear Summary", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ClearGSRMergedInvSummary_D)
        btnClearGSRMergedInvSummary.place(x=1240,y=65)
        btnViewGSRMergedDuplicateEntries = Button(self.root, text="View Duplicate", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ViewGSRMergedDuplicateEntries_D)
        btnViewGSRMergedDuplicateEntries.place(x=1240,y=92)

        btnViewMasterMergedDB = Button(self.root, text="View Merged DB", font=('aerial', 9, 'bold'), height =1, width=14, bd=4, command = ViewMasterMergedDB)
        btnViewMasterMergedDB.place(x=2,y=632)
        btnUpdateMasterMergedDB = Button(self.root, text="Update Merged DB", font=('aerial', 9, 'bold'), height =1, width=15, bd=4, command = UpdateMasterMergedDB)
        btnUpdateMasterMergedDB.place(x=119,y=632)
        btnExportMasterMergedDB = Button(self.root, text="Export Merged DB", font=('aerial', 9, 'bold'), height =1, width=15, bd=4, command = ExportMasterMergedDB)
        btnExportMasterMergedDB.place(x=243,y=632)
        
        btnClearMergedGSRView = Button(self.root, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearMergedGSRView)
        btnClearMergedGSRView.place(x=1175,y=632)
        btnMergedGSRExit = Button(self.root, text="Exit Query", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExitGSRMerged)
        btnMergedGSRExit.place(x=1264,y=632)
        btnDeleteMasterMergedSelectedData = Button(self.root, text="Delete Selected", font=('aerial', 9, 'bold'), height =1, width=13, bd=4, command = DeleteMasterMergedSelectedData)
        btnDeleteMasterMergedSelectedData.place(x=1067,y=632)
        btnExportGSRMergeSearch = Button(self.root, text="Export Search Query", font=('aerial', 9, 'bold'), height =1, width=18, bd=4 ,command = ExportGSRMergeSearch)
        btnExportGSRMergeSearch.place(x=750,y=632)


if __name__ == '__main__':
    root = Tk()
    application  = MergedGSRInventoryRepairQuery(root)
    root.mainloop()
