#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_GSRInventory_BackEnd
import Eagle_GSRRepairInventory_BackEnd
import Eagle_GSRMergedInventory_BackEnd
import Eagle_GSRDeploymentHistory_BackEnd
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

class GSRRepairMaster_Merge_GSRInvMaster:    
    def __init__(self,root):
        self.root =root
        self.root.title ("Eagle GSR Repair Inventory Merge With GSR Inventory")
        self.root.geometry("1340x672+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)        
        TableMargin = Frame(self.root,  bd = 2, padx= 2, pady= 10, relief = RIDGE)
        TableMargin.place(x=2, y=35, anchor="nw", width=1335, height=602)
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
        SEARCHCaseSN   = StringVar()
        SEARCHWorkOrder = StringVar()
        Fixed_Timestamp   = '1900/1/01 00:00'
        
##### All Functions defining

        def UpdateGSRInvMasterDB():
            conn = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['LastTimeSeenInDTMDt', 'LastTimeReapedDt']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)
            self.cur=conn.cursor()
            data.to_sql('Eagle_GSRInventory_MERGED_TEMP',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()

        def UpdateGSRDeploymentMasterDB():
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['StartTimeUTC']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)
            self.cur=conn.cursor()
            data.to_sql('Eagle_GSRDeploymentHistory_MERGED_TEMP',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()
           
        def iExit():
            iExit= tkinter.messagebox.askyesno("Eagle GSR Inventory Management System", "Confirm if you want to exit")
            if iExit >0:
                self.root.destroy()
                return

        def ClearView():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            

        def ExportMergedReport():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            search_DF = pd.DataFrame(dfList)
            search_DF.rename(columns = {0:'CaseSrNo', 1:'DeviceType', 2:'ProjectID', 3:'FlashCapacityGB', 4:'LastTimeSeenInDTMDt',
                                        5: 'LastTimeLineViewedDt', 6:'LastTimeReapedDt', 7:'Flags', 8:'WorkOrderNo', 9:'PartNo',
                                        10:'TechnicianInput', 11:'CrewReported', 12:'DateRepaired'},inplace = True)
            data_SortByCaseSrNo = pd.DataFrame(search_DF)            
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortLastTimeSeenInDTMDt = pd.DataFrame(search_DF)            
            data_SortLastTimeSeenInDTMDt = data_SortLastTimeSeenInDTMDt.sort_values(by =['LastTimeSeenInDTMDt'])

            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortLastTimeSeenInDTMDt.to_excel(file,sheet_name='SortByLastTimeSeenInDTMDt',index=False)
                    file.close
                    tkinter.messagebox.showinfo("ListBox Entries Export"," ListBox Entries Report Saved as Excel")
            tree.delete(*tree.get_children())

            
        def MergeMasterGSRInvGSRRepairGSRDeployDB():
            UpdateGSRInvMasterDB()
            UpdateGSRDeploymentMasterDB()
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)

            connRepair = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df_Repair = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory ORDER BY `CaseSrNo` ASC ;", connRepair)
            data_Repair = pd.DataFrame(Complete_df_Repair) 
            data_Repair ['DuplicatedEntries']=data_Repair.sort_values(by =['WorkOrderNo']).duplicated(['CaseSrNo', 'DeviceType'],keep='last')            
            data_Repair = data_Repair.loc[data_Repair.DuplicatedEntries == False, 'WorkOrderNo': 'DuplicatedEntries']            
            data_Repair = data_Repair.reset_index(drop=True)            
            data_Repair = data_Repair.loc[:,['WorkOrderNo','CaseSrNo','PartNo','DeviceType', 'TechnicianInput','CrewReported','DateRepaired']]

            connMaster = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df_Master = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_MERGED_TEMP ORDER BY `CaseSrNo` ASC ;", connMaster)
            data_Master = pd.DataFrame(Complete_df_Master)                   
            data_Master = data_Master.loc[:,['CaseSrNo','DeviceType','ProjectID','FlashCapacityGB',
                                             'LastTimeSeenInDTMDt','LastTimeLineViewedDt','LastTimeReapedDt']]

            connDeploy = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df_Deploy = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_MERGED_TEMP ORDER BY `CaseSrNo` ASC ;", connDeploy)
            data_Deploy = pd.DataFrame(Complete_df_Deploy)                   
            data_Deploy = data_Deploy.loc[:,['CaseSrNo','DeviceType','StartTimeUTC','EndTimeUTC','JobName']]


            Merge_Repair_Master = pd.merge(data_Master, data_Repair, on =['CaseSrNo','DeviceType'] ,how ='outer',
                       left_index = False, right_index = False, sort = True, indicator = True )

            def trans_CaseSrNo_MISSING(x):
                if x   == 'both':
                    return 'Repaired'
                elif x == 'right_only':
                    return 'CaseSrNo Missing in GSRMasterInventory from GSR Repair'
                elif x == 'left_only':
                    return 'No Repair Report'
                else:
                    return x

            Merge_Repair_Master['FlagsRepair']               = Merge_Repair_Master['_merge'].apply(trans_CaseSrNo_MISSING)
            Merge_Repair_Master["LastTimeSeenInDTMDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master['LastTimeSeenInDTMDt'] = pd.to_datetime(Merge_Repair_Master['LastTimeSeenInDTMDt']).dt.strftime('%Y-%m-%d')

            Merge_Repair_Master["LastTimeLineViewedDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master['LastTimeLineViewedDt'] = pd.to_datetime(Merge_Repair_Master['LastTimeLineViewedDt']).dt.strftime('%Y-%m-%d')

            Merge_Repair_Master["LastTimeReapedDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master['LastTimeReapedDt'] = pd.to_datetime(Merge_Repair_Master['LastTimeReapedDt']).dt.strftime('%Y-%m-%d')
                        
            Merge_Repair_Master = Merge_Repair_Master.loc[:,
                    ['CaseSrNo','DeviceType','ProjectID','FlashCapacityGB','LastTimeSeenInDTMDt','LastTimeLineViewedDt','LastTimeReapedDt','FlagsRepair',
                     'WorkOrderNo','PartNo','TechnicianInput','CrewReported','DateRepaired']]
            Merge_Repair_Master = Merge_Repair_Master.reset_index(drop=True)
            Merge_Repair_Master = pd.DataFrame(Merge_Repair_Master)



            Merge_Repair_Master_Deploy = pd.merge(Merge_Repair_Master, data_Deploy, on =['CaseSrNo','DeviceType'] ,how ='outer',
                       left_index = False, right_index = False, sort = True, indicator = True )

            def trans_CaseSrNo_MISSING_Deploy(y):
                if y   == 'both':
                    return 'Matched'
                elif y == 'right_only':
                    return 'CaseSrNo Missing in GSRMasterInventory From GSRDeployment'
                elif y == 'left_only':
                    return 'No Deployment Info'
                else:
                    return y


            Merge_Repair_Master_Deploy['FlagsDeployment'] = Merge_Repair_Master_Deploy['_merge'].apply(trans_CaseSrNo_MISSING_Deploy)
            Merge_Repair_Master_Deploy["LastTimeSeenInDTMDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master_Deploy['LastTimeSeenInDTMDt'] = pd.to_datetime(Merge_Repair_Master_Deploy['LastTimeSeenInDTMDt']).dt.strftime('%Y-%m-%d')

            Merge_Repair_Master_Deploy["LastTimeLineViewedDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master_Deploy['LastTimeLineViewedDt'] = pd.to_datetime(Merge_Repair_Master_Deploy['LastTimeLineViewedDt']).dt.strftime('%Y-%m-%d')

            Merge_Repair_Master_Deploy["LastTimeReapedDt"].fillna(Fixed_Timestamp, inplace = True)
            Merge_Repair_Master_Deploy['LastTimeReapedDt'] = pd.to_datetime(Merge_Repair_Master_Deploy['LastTimeReapedDt']).dt.strftime('%Y-%m-%d')



            Merge_Repair_Master_Deploy = Merge_Repair_Master_Deploy.loc[:,
                    ['CaseSrNo','DeviceType','ProjectID','FlashCapacityGB','LastTimeSeenInDTMDt','LastTimeLineViewedDt','LastTimeReapedDt','FlagsRepair',
                     'WorkOrderNo','PartNo','TechnicianInput','CrewReported','DateRepaired', 'FlagsDeployment', 'StartTimeUTC','EndTimeUTC','JobName']]
            Merge_Repair_Master_Deploy = Merge_Repair_Master_Deploy.reset_index(drop=True)
            Merge_Repair_Master_Deploy = pd.DataFrame(Merge_Repair_Master_Deploy)

            
            for each_rec in range(len(Merge_Repair_Master_Deploy)):
                tree.insert("", tk.END, values=list(Merge_Repair_Master_Deploy.loc[each_rec]))

            tkinter.messagebox.showinfo("Merge Complete","Merge Complete and Duplicated Entries are Removed")

            TotalEntries = len(Merge_Repair_Master_Deploy)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)

            connMergedMaster = sqlite3.connect("Eagle_GSRMergedInventory.db")
            Merge_Repair_Master_Deploy.to_sql('Eagle_GSRMergedInventory_MASTER',connMergedMaster, if_exists="append", index=False)
            
            connRepair.commit()
            connRepair.close()
            
            connMaster.commit()
            connMaster.close()
            
            connDeploy.commit()
            connDeploy.close()
            
            connMergedMaster.commit()
            connMergedMaster.close()

        def ViewRepairMasterDB():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            connRepair = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df_Repair = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory ORDER BY `CaseSrNo` ASC ;", connRepair)
            data_Repair = pd.DataFrame(Complete_df_Repair)
            data_Repair ['DuplicatedEntries']=data_Repair.sort_values(by =['WorkOrderNo']).duplicated(['CaseSrNo', 'DeviceType'],keep='last')            
            data_Repair = data_Repair.loc[data_Repair.DuplicatedEntries == False, 'WorkOrderNo': 'DuplicatedEntries']
            data_Repair = data_Repair.reset_index(drop=True)            
            data_Repair = data_Repair.loc[:,['CaseSrNo', 'DeviceType', 'WorkOrderNo', 'PartNo', 'TechnicianInput','CrewReported','DateRepaired']]

            data_Repair["ProjectID"]           = data_Repair.shape[0]*[" "]
            data_Repair["FlashCapacityGB"]     = data_Repair.shape[0]*[" "]
            data_Repair["LastTimeSeenInDTMDt"] = data_Repair.shape[0]*[" "]
            data_Repair["LastTimeLineViewedDt"] = data_Repair.shape[0]*[" "]
            data_Repair["LastTimeReapedDt"] = data_Repair.shape[0]*[" "]
            data_Repair["FlagsRepair"] = data_Repair.shape[0]*[" "]            
            data_Repair = data_Repair.loc[:,['CaseSrNo', 'DeviceType','ProjectID','FlashCapacityGB','LastTimeSeenInDTMDt','LastTimeLineViewedDt',
                                             'LastTimeReapedDt','FlagsRepair','WorkOrderNo', 'PartNo', 'TechnicianInput','CrewReported','DateRepaired']]            
            data_Repair = data_Repair.reset_index(drop=True) 
            for each_rec in range(len(data_Repair)):
                tree.insert("", tk.END, values=list(data_Repair.loc[each_rec]))
            TotalEntries = len(data_Repair)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)                       
            connRepair.commit()
            connRepair.close()

        def ViewGSRInvMasterDB():
            UpdateGSRInvMasterDB()
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            connMaster = sqlite3.connect("Eagle_GSRInventory.db")
            Complete_df_Master = pd.read_sql_query("SELECT * FROM Eagle_GSRInventory_MERGED_TEMP ORDER BY `CaseSrNo` ASC ;", connMaster)
            data_Master = pd.DataFrame(Complete_df_Master)                   
            data_Master = data_Master.loc[:,['CaseSrNo','DeviceType','ProjectID','FlashCapacityGB',
                                             'LastTimeSeenInDTMDt','LastTimeLineViewedDt','LastTimeReapedDt']]

            for each_rec in range(len(data_Master)):
                tree.insert("", tk.END, values=list(data_Master.loc[each_rec]))
            TotalEntries = len(data_Master)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)  
            connMaster.commit()
            connMaster.close()

        def ViewGSRDeploymentMasterDB():
            UpdateGSRDeploymentMasterDB()
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            connMaster = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df_Master = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_MERGED_TEMP ORDER BY `CaseSrNo` ASC ;", connMaster)
            data_Master = pd.DataFrame(Complete_df_Master)                   
            data_Master = data_Master.loc[:,['CaseSrNo','DeviceType','StartTimeUTC','EndTimeUTC',
                                             'JobName']]

            data_Master["ProjectID"]           = data_Master.shape[0]*[" "]
            data_Master["FlashCapacityGB"]     = data_Master.shape[0]*[" "]
            data_Master["LastTimeSeenInDTMDt"] = data_Master.shape[0]*[" "]
            data_Master["LastTimeLineViewedDt"] = data_Master.shape[0]*[" "]
            data_Master["LastTimeReapedDt"] =   data_Master.shape[0]*[" "]
            data_Master["FlagsRepair"] =       data_Master.shape[0]*[" "]
            data_Master["WorkOrderNo"]           = data_Master.shape[0]*[" "]
            data_Master["PartNo"]     = data_Master.shape[0]*[" "]
            data_Master["TechnicianInput"] = data_Master.shape[0]*[" "]
            data_Master["CrewReported"] = data_Master.shape[0]*[" "]
            data_Master["DateRepaired"] =   data_Master.shape[0]*[" "]
            data_Master["FlagsDeployment"] =       data_Master.shape[0]*[" "]

            data_Master = data_Master.loc[:,['CaseSrNo', 'DeviceType','ProjectID','FlashCapacityGB','LastTimeSeenInDTMDt','LastTimeLineViewedDt',
                                             'LastTimeReapedDt','FlagsRepair','WorkOrderNo', 'PartNo', 'TechnicianInput','CrewReported','DateRepaired',
                                             'FlagsDeployment','StartTimeUTC','EndTimeUTC','JobName']]            
            data_Master = data_Master.reset_index(drop=True)


            for each_rec in range(len(data_Master)):
                tree.insert("", tk.END, values=list(data_Master.loc[each_rec]))
            TotalEntries = len(data_Master)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)  
            connMaster.commit()
            connMaster.close()
            

## Label
        InvL1 = Label(self.root, text = "Total Entries:", font=("arial", 12,'bold'),bg = "green", bd=4).place(x=900,y=640)
        InvL2 = Label(self.root, text = "<<<< Populated From Master GSR Inventory >>>>", font=("arial", 12,'bold'),bg = "green", width = 67, bd=4).place(x=2,y=2)
        InvL3 = Label(self.root, text = "<< Populated From Master GSRRepair & GSR Deployment >>", font=("arial", 12,'bold'),bg = "green", width = 50, bd=4).place(x=830,y=2)
        InvL4 = Label(self.root, text = "<< Merged Flags >>", font=("arial", 12,'bold'),bg = "purple", bd=4).place(x=683,y=2)

##### Entry Wizard
        self.txtTotalEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 6, bd =4)
        self.txtTotalEntries.place(x=1010,y=640)


##### Button Wizard  
        btnViewMasterGSRInvDB = Button(self.root, text="View MasterGSRInventory", font=('aerial', 9, 'bold'),
                                 height =1, width=21, bd=4, command = ViewGSRInvMasterDB)
        btnViewMasterGSRInvDB.place(x=2,y=640)

        btnViewMasterRepairDB = Button(self.root, text="View MasterRepairInventory", font=('aerial', 9, 'bold'),
                                 height =1, width=23, bd=4, command = ViewRepairMasterDB)
        btnViewMasterRepairDB.place(x=165,y=640)

        btnViewMasterDeploymentDB = Button(self.root, text="View MasterDeployInventory", font=('aerial', 9, 'bold'),
                                 height =1, width=23, bd=4, command = ViewGSRDeploymentMasterDB)
        btnViewMasterDeploymentDB.place(x=342,y=640)


        btnExit = Button(self.root, text="Exit Wizard", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
        btnExit.place(x=1252,y=640)

        btnClearView = Button(self.root, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
        btnClearView.place(x=1078,y=640)

        btnMergeGSRInv_RepairReport = Button(self.root, text="Merge All Databases", font=('aerial', 9, 'bold'),
                                      height =1, width=28, bd=4, bg= 'ghost white', command = MergeMasterGSRInvGSRRepairGSRDeployDB)
        btnMergeGSRInv_RepairReport.place(x=520,y=640)

        btnExportMergedReport = Button(self.root, text="Export Merged Report", font=('aerial', 9, 'bold'),
                                  height =1, width=18, bd=4, bg= 'ghost white', command = ExportMergedReport)
        btnExportMergedReport.place(x=732,y=640)


if __name__ == '__main__':
    root = Tk()
    application  = GSRRepairMaster_Merge_GSRInvMaster(root)
    root.mainloop()
