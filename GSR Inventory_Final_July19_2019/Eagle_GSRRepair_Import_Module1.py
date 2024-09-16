#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_GSRRepairInventory_BackEnd
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

class GSRRepairImport:    
    def __init__(self,root):
        self.root =root
        self.root.title ("Eagle GSR Repair Report Import Wizard")
        self.root.geometry("1200x650+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)
        TableMargin = Frame(self.root, bd = 2, padx= 10, pady= 10, relief = RIDGE)
        TableMargin.pack(side=TOP)
        TableMargin.pack(side=LEFT)
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                                 "column6", "column7", "column8", "column9", "column10", "column11","column12"), height=25, show='headings')
        scrollbary.config(command=tree.yview)
        scrollbary.pack(side=RIGHT, fill=Y)
        scrollbarx.config(command=tree.xview)
        scrollbarx.pack(side=BOTTOM, fill=X)
        tree.heading("#1", text="WorkOrderNo", anchor=W)
        tree.heading("#2", text="CaseSrNo", anchor=W)
        tree.heading("#3", text="PartNo", anchor=W)
        tree.heading("#4", text="DeviceType", anchor=W)
        tree.heading("#5", text="TechnicianInput", anchor=W)
        tree.heading("#6", text="CrewReported", anchor=W)
        tree.heading("#7", text="WarrantyStatus", anchor=W)            
        tree.heading("#8", text="Chargeable", anchor=W)
        tree.heading("#9", text="PricePer" ,anchor=W)
        tree.heading("#10", text="DiscountApplied", anchor=W)
        tree.heading("#11", text="SubTotal", anchor=W)
        tree.heading("#12", text="DateRepaired", anchor=W)
        
        tree.column('#1', stretch=NO, minwidth=0, width=100)            
        tree.column('#2', stretch=NO, minwidth=0, width=80)
        tree.column('#3', stretch=NO, minwidth=0, width=100)
        tree.column('#4', stretch=NO, minwidth=0, width=80)
        tree.column('#5', stretch=NO, minwidth=0, width=120)
        tree.column('#6', stretch=NO, minwidth=0, width=120)
        tree.column('#7', stretch=NO, minwidth=0, width=100)
        tree.column('#8', stretch=NO, minwidth=0, width=80)
        tree.column('#9', stretch=NO, minwidth=0, width=80)
        tree.column('#10', stretch=NO, minwidth=0, width=110)
        tree.column('#11', stretch=NO, minwidth=0, width=80)
        tree.column('#12', stretch=NO, minwidth=0, width=100)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=('aerial', 8), foreground="black")
        style.configure("Treeview", foreground='black')
        style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
        
        tree.pack()
        SEARCHCaseSN   = StringVar()
        SEARCHWorkOrder = StringVar()
        
        GSRX_3_279_1 = '450-01480-03'
        GSRX_3_279_2 = '450-01480-13'

        GSRX_4_263_1 = '450-01480-14'

        GSRX_1_264_1 = '450-01480-21'
        GSRX_1_264_2 = '450-01480-11'

        GSR_1_256_1 = '450-00800-01'
        GSR_1_256_2 = '450-00800-01-CAP'
        GSR_1_256_3 = '450-00800-11'

        GSR_4_257_1 = '450-00800-21'
        GSR_4_257_2 = '450-00800-04'

      
##### All Functions defining 

        def iExit():
            iExit= tkinter.messagebox.askyesno("Eagle GSR Inventory Management System", "Confirm if you want to exit")
            if iExit >0:
                self.root.destroy()
                return

        def ClearView():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            self.txtSearchCaseSNMasterDB.delete(0,END)
            self.txtSearchWorkorderMasterDB.delete(0,END)

        def CaseSNSearch():
            if SEARCHCaseSN.get() != "":
                tree.delete(*tree.get_children())
                self.txtTotalEntries.delete(0,END)
                conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM `Eagle_GSRRepairInventory` WHERE `CaseSrNo`= ? ",\
                               (self.txtSearchCaseSNMasterDB.get(),))                          
                                                            
                fetch = cursor.fetchall()
                for data in fetch:
                    tree.insert('', 'end', values=(data))
                cursor.close()
                conn.close()

        def WorkorderSearch():
            if SEARCHWorkOrder.get() != "":
                tree.delete(*tree.get_children())
                self.txtTotalEntries.delete(0,END)
                conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM `Eagle_GSRRepairInventory` WHERE `WorkOrderNo`= ? ",\
                               (self.txtSearchWorkorderMasterDB.get(),))                          
                                                            
                fetch = cursor.fetchall()
                for data in fetch:
                    tree.insert('', 'end', values=(data))
                cursor.close()
                conn.close()

        def ClearSearchCaseSN():
            self.txtSearchCaseSNMasterDB.delete(0,END)
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)

        def ClearSearchWorkOrder():
            self.txtSearchWorkorderMasterDB.delete(0,END)
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            
        def ViewMasterDB():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory ORDER BY `WorkOrderNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalEntries = len(data)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)              
            conn.commit()
            conn.close()

        def DeleteSelectedData():
            iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
            if iDelete >0:
                self.txtTotalEntries.delete(0,END)
                conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
                cur = conn.cursor()
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_GSRRepairInventory_TEMP WHERE WorkOrderNo =? AND CaseSrNo=? AND \
                                PartNo =?  AND DeviceType =? AND TechnicianInput =? AND CrewReported =? ",\
                                (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#3'),\
                                 tree.set(selected_item, '#4'), tree.set(selected_item, '#5'), tree.set(selected_item, '#6'),))
                    cur.execute("DELETE FROM Eagle_GSRRepairInventory WHERE WorkOrderNo =? AND CaseSrNo=? AND \
                                 PartNo =?  AND DeviceType =? AND TechnicianInput =? AND CrewReported =? ",\
                                (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#3'),\
                                 tree.set(selected_item, '#4'), tree.set(selected_item, '#5'), tree.set(selected_item, '#6'),))
                    conn.commit()
                    tree.delete(selected_item)
                conn.close()
                Total_count = len(tree.get_children())
                self.txtTotalEntries.insert(tk.END,Total_count)
                return

        def ViewImport():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory_TEMP ORDER BY `WorkOrderNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalEntries = len(data)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)              
            conn.commit()
            conn.close()

        def SubmitToMasterDB():
            iSubmit = tkinter.messagebox.askyesno("Entries Submit to Repair Master DB", "Confirm if you want to Submit the Imported Entries to Master Repair DB")
            if iSubmit >0:
                tree.delete(*tree.get_children())
                conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
                cur=conn.cursor()
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data.to_sql('Eagle_GSRRepairInventory',conn, if_exists="append", index=False)
                cur.execute("DELETE FROM Eagle_GSRRepairInventory_TEMP")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Submit Complete","All Import Entries are Submitted to Master DB")
                self.txtTotalEntries.delete(0,END)
                return
                                                       
        def ImportGSRRepairFile():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            fileList = askopenfilenames(filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, header = None, skiprows = {0},skipfooter=3, dtype={0: int, 1: int})
                    else:
                        df = pd.read_excel(filename, header = None, skiprows = {0},skipfooter=3, dtype={0: int, 1: int})
                    dfList.append(df)
                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns = {0:'WorkOrderNo', 1:'CaseSrNo', 2:'PartNo', 3:'TechnicianInput', 4:'CrewReported', 5:'WarrantyStatus',
                                          6: 'Chargeable', 7:'PricePer', 8:'DiscountApplied', 9:'SubTotal', 10:'DateRepaired'},inplace = True)
                self.df = pd.DataFrame(concatDf)

                def trans_AssignDeviceType(x):
                    if x == GSRX_3_279_1:
                        return 279
                    elif x == GSRX_3_279_2:
                        return 279
                    
                    elif x == GSRX_4_263_1:
                        return 263

                    elif x == GSRX_1_264_1:
                        return 264
                    elif x == GSRX_1_264_2:
                        return 264

                    elif x == GSR_1_256_1:
                        return 256
                    elif x == GSR_1_256_2:
                        return 256
                    elif x == GSR_1_256_3:
                        return 256

                    elif x == GSR_4_257_1:
                        return 257
                    elif x == GSR_4_257_2:
                        return 257
                    
                    else:
                        return 'unknown'


                self.df['DeviceType']  = self.df['PartNo'].apply(trans_AssignDeviceType)
                self.df['DateRepaired'] = pd.to_datetime(self.df['DateRepaired']).dt.strftime('%Y-%m-%d')
        
                self.df = self.df.loc[:,['WorkOrderNo','CaseSrNo','PartNo','DeviceType','TechnicianInput','CrewReported','WarrantyStatus',
                                         'Chargeable','PricePer','DiscountApplied','SubTotal','DateRepaired']]                
                data = pd.DataFrame(self.df)
                data.rename(columns = {0:'WorkOrderNo', 1:'CaseSrNo', 2:'PartNo', 3:'DeviceType', 4:'TechnicianInput', 5:'CrewReported', 6:'WarrantyStatus',
                                          7: 'Chargeable', 8:'PricePer', 9:'DiscountApplied', 10:'SubTotal', 11:'DateRepaired'},inplace = True)                
                data = data[pd.to_numeric(data.WorkOrderNo,errors='coerce').notnull()]                
                data = data.reset_index(drop=True)
                for each_rec in range(len(data)):
                    tree.insert("", tk.END, values=list(data.loc[each_rec]))            
                con= sqlite3.connect("Eagle_GSRRepairInventory.db")
                self.cur=con.cursor()                
                data.to_sql('Eagle_GSRRepairInventory_TEMP',con, if_exists="replace", index=False)
                TotalEntries = len(data)       
                self.txtTotalEntries.insert(tk.END,TotalEntries)  
                con.commit()
                con.close()

        def SaveImportedEntries():
            conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortWorkOrderNo = pd.DataFrame(Complete_df)
            data_SortWorkOrderNo = data_SortWorkOrderNo.sort_values(by =['WorkOrderNo'])
            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortWorkOrderNo.to_excel(file,sheet_name='SortByWorkOrderNo',index=False)
                    file.close
                    tkinter.messagebox.showinfo("Inventory Repair Report Export","Inventory Repair Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def ExportMasterDB():
            conn = sqlite3.connect("Eagle_GSRRepairInventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRRepairInventory ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortWorkOrderNo = pd.DataFrame(Complete_df)
            data_SortWorkOrderNo = data_SortWorkOrderNo.sort_values(by =['WorkOrderNo'])
            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortWorkOrderNo.to_excel(file,sheet_name='SortByWorkOrderNo',index=False)
                    file.close
                    tkinter.messagebox.showinfo("Inventory Repair Master DB Export","Inventory Repair Master DB Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def ExportListBoxView():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            search_DF = pd.DataFrame(dfList)
            search_DF.rename(columns = {0:'WorkOrderNo', 1:'CaseSrNo', 2:'PartNo', 3:'TechnicianInput', 4:'CrewReported',
                                          5: 'WarrantyStatus', 6:'Chargeable', 7:'PricePer', 8:'DiscountApplied',
                                          9:'SubTotal'},inplace = True)
            data_SortByCaseSrNo = pd.DataFrame(search_DF)            
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortWorkOrderNo = pd.DataFrame(search_DF)            
            data_SortWorkOrderNo = data_SortWorkOrderNo.sort_values(by =['WorkOrderNo'])

            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortWorkOrderNo.to_excel(file,sheet_name='SortByWorkOrderNo',index=False)
                    file.close
                    tkinter.messagebox.showinfo("ListBox Entries Export"," ListBox Entries Report Saved as Excel")
            tree.delete(*tree.get_children())
        
##### Entry Wizard
        self.txtTotalEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 6, bd =4)
        self.txtTotalEntries.place(x=775,y=613)

        self.txtSearchCaseSNMasterDB  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=SEARCHCaseSN, width = 9)
        self.txtSearchCaseSNMasterDB.place(x=150,y=14)

        self.txtSearchWorkorderMasterDB  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=SEARCHWorkOrder, width = 9)
        self.txtSearchWorkorderMasterDB.place(x=500,y=14)

## Label
        InvL1 = Label(self.root, text = "Total Entries:", font=("arial", 12,'bold'),bg = "green", bd=4).place(x=660,y=613)
        InvL2 = Label(self.root, text = "Search Master DB:", font=("arial", 12,'bold'),bg = "green", bd=1).place(x=2,y=14)
        


### Button Wizard  
        btnImport = Button(self.root, text="Import GSR Repair File", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = ImportGSRRepairFile)
        btnImport.place(x=2,y=613)

        btnSubmitMasterDB = Button(self.root, text="Submit to Repair MasterDB", font=('aerial', 9, 'bold'), height =1, width=22, bd=4, command = SubmitToMasterDB)
        btnSubmitMasterDB.place(x=145,y=613)

        btnViewMasterDB = Button(self.root, text="View Repair MasterDB", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = ViewMasterDB)
        btnViewMasterDB.place(x=316,y=613)

        btnExportMasterDB = Button(self.root, text="Export Repair MasterDB", font=('aerial', 9, 'bold'), height =1, width=20, bd=4, command = ExportMasterDB)
        btnExportMasterDB.place(x=458,y=613)

        btnSaveImportEntries = Button(self.root, text="Save Import Files", font=('aerial', 9, 'bold'), height =1, width=14, bd=4, command = SaveImportedEntries)
        btnSaveImportEntries.place(x=875,y=613)

        btnViewImport = Button(self.root, text="View Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ViewImport)
        btnViewImport.place(x=990,y=613)

        btnExit = Button(self.root, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
        btnExit.place(x=1110,y=613)

        btnSearchMasterCaseSN = Button(self.root, text="Search By CaseSN", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = CaseSNSearch)
        btnSearchMasterCaseSN.place(x=240,y=14)

        btnSearchClearCaseSN = Button(self.root, text="Reset", font=('aerial', 9, 'bold'), height =1, width=6, bd=1, command = ClearSearchCaseSN)
        btnSearchClearCaseSN.place(x=362,y=14)

        btnSearchMasterWorkOrder = Button(self.root, text="Search By Workorder", font=('aerial', 9, 'bold'), height =1, width=18, bd=1, command = WorkorderSearch)
        btnSearchMasterWorkOrder.place(x=590,y=14)

        btnSearchClear = Button(self.root, text="Reset", font=('aerial', 9, 'bold'), height =1, width=6, bd=1, command = ClearSearchWorkOrder)
        btnSearchClear.place(x=726,y=14)

        btnExportSelectedListbox = Button(self.root, text="Export Search Entries ", font=('aerial', 9, 'bold'), height =1, width=20, bd=1, command = ExportListBoxView)
        btnExportSelectedListbox.place(x=833,y=14)
        
        btnDeleteSelected = Button(self.root, text="Delete Selected", font=('aerial', 9, 'bold'), height =1, width=13, bd=1, command = DeleteSelectedData)
        btnDeleteSelected.place(x=1015,y=14)

        btnClearView = Button(self.root, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=1, command = ClearView)
        btnClearView.place(x=1118,y=14)


if __name__ == '__main__':
    root = Tk()
    application  = GSRRepairImport(root)
    root.mainloop()
