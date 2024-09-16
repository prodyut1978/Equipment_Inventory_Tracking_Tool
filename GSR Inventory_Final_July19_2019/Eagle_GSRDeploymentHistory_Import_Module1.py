#Front End
import os
from tkinter import*
import tkinter.messagebox
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

class GSRDeploymentHistoryImport:    
    def __init__(self,root):
        self.root =root
        self.root.title ("Eagle GSR Deployment History Import Wizard")
        self.root.geometry("1350x650+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)
        TableMargin = Frame(self.root, bd = 2, padx= 10, pady= 8, relief = RIDGE)
        TableMargin.pack(side=TOP)
        TableMargin.pack(side=LEFT)
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                                 "column6", "column7", "column8"), height=26, show='headings')
        scrollbary.config(command=tree.yview)
        scrollbary.pack(side=RIGHT, fill=Y)
        scrollbarx.config(command=tree.xview)
        scrollbarx.pack(side=BOTTOM, fill=X)
        tree.heading("#1", text="CaseSrNo", anchor=W)
        tree.heading("#2", text="DeviceType", anchor=W)
        tree.heading("#3", text="Line", anchor=W)
        tree.heading("#4", text="OccupiedStations", anchor=W)
        tree.heading("#5", text="StartTimeUTC", anchor=W)            
        tree.heading("#6", text="EndTimeUTC", anchor=W)
        tree.heading("#7", text="JobName" ,anchor=W)
        tree.heading("#8", text="DuplicatedEntries" ,anchor=W)             
        tree.column('#1', stretch=NO, minwidth=0, width=80)            
        tree.column('#2', stretch=NO, minwidth=0, width=80)
        tree.column('#3', stretch=NO, minwidth=0, width=80)
        tree.column('#4', stretch=NO, minwidth=0, width=130)
        tree.column('#5', stretch=NO, minwidth=0, width=180)
        tree.column('#6', stretch=NO, minwidth=0, width=180)
        tree.column('#7', stretch=NO, minwidth=0, width=450)
        tree.column('#8', stretch=NO, minwidth=0, width=120)        
        tree.pack()
        self.df = None
        Bad_Timestamp_Excel = '-'
        Fixed_Timestamp   = '1900-01-01 00:00:00Z'
        
### All Functions defining
        def Treepopup_do_popup(event):
            try:
                Treepopup.tk_popup(event.x_root, event.y_root, 0)
            finally:
                Treepopup.grab_release()
                
        def ImportGSRDeploymentHistoryFiles():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)
            fileList = askopenfilenames(filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, header = None, skiprows = {0})
                        filename_w_ext = os.path.basename(filename)
                        job_name, file_extension = os.path.splitext(filename_w_ext)
                        df["JobName"]     = df.shape[0]*[job_name]
                    else:
                        df = pd.read_excel(filename, header = None, skiprows = {0})
                        filename_w_ext = os.path.basename(filename)
                        job_name, file_extension = os.path.splitext(filename_w_ext)
                        df["JobName"]     = df.shape[0]*[job_name]
                    dfList.append(df)

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns = {0:'CaseSrNo', 1:'DeviceType', 2:'Line', 3:'OccupiedStations', 4:'TotalStationsOccupied',
                              5: 'StationInterval', 6:'DeploymentDirection', 7:'Station1', 8:'Station2',
                              9:'Station3', 10:'Station4',11:'GSRStation',
                              12:'StartTimeUTC',13:'EndTimeUTC', 14:'JobName'},inplace = True)
                
                self.df = pd.DataFrame(concatDf)                
                self.df["StartTimeUTC"].fillna(Fixed_Timestamp, inplace = True)
                self.df["EndTimeUTC"].fillna(Fixed_Timestamp, inplace = True)

                def trans_TimeFixExcel(y):
                    if y == Bad_Timestamp_Excel:
                        return Fixed_Timestamp
                    else:
                        return y

                self.df['StartTimeUTC']  = self.df['StartTimeUTC'].apply(trans_TimeFixExcel)
                self.df['EndTimeUTC']    = self.df['EndTimeUTC'].apply(trans_TimeFixExcel)            
                self.df['StartTimeUTC'] = pd.to_datetime(self.df['StartTimeUTC']).dt.strftime('%Y-%m-%d')
                self.df['EndTimeUTC']   = pd.to_datetime(self.df['EndTimeUTC']).dt.strftime('%Y-%m-%d')                
                data = pd.DataFrame(self.df)
                dataUnknownDeviceType = pd.DataFrame(self.df)
                data = data.loc[:,['CaseSrNo','DeviceType','Line','OccupiedStations',
                                   'StartTimeUTC','EndTimeUTC','JobName']]                
                data ['DuplicatedEntries']=data.sort_values(by =['StartTimeUTC','EndTimeUTC']).duplicated(['CaseSrNo','DeviceType'],keep='last')
                data = data[(data.DeviceType == 'SDRx')|(data.DeviceType == 'SDR')|
                            (data.DeviceType == 'GSR-4')|(data.DeviceType == 'GSR-3')|(data.DeviceType == 'GSR-1')|
                            (data.DeviceType == 'GSRx-1')|(data.DeviceType == 'GSRx-3')|(data.DeviceType == 'GSRx-4')]
                data = data.reset_index(drop=True)

                def trans_AssignDeviceType(x):

                    if x == 'SDRx':
                        return 273

                    elif x == 'SDR':
                        return 270
                    
                    elif x == 'GSR-4':
                        return 257

                    elif x == 'GSR-3':
                        return 279

                    elif x == 'GSR-1':
                        return 256

                    elif x == 'GSRx-1':
                        return 264

                    elif x == 'GSRx-4':
                        return 263

                    elif x == 'GSRx-3':
                        return 279
                    
                    else:
                        return x

                data['DeviceType']  = data['DeviceType'].apply(trans_AssignDeviceType)
                data = data.reset_index(drop=True)
                for each_rec in range(len(data)):
                    tree.insert("", tk.END, values=list(data.loc[each_rec]))            

                con= sqlite3.connect("Eagle_GSRDeploymentHistory.db")
                self.cur=con.cursor()                
                data.to_sql('Eagle_GSRDeploymentHistory_TEMP',con, if_exists="replace", index=False)


                dataUnknownDeviceType = dataUnknownDeviceType.loc[:,['CaseSrNo','DeviceType','Line','OccupiedStations',
                                   'StartTimeUTC','EndTimeUTC','JobName']]

                def trans_FindBadDeviceType(x):

                    if x == 'SDRx':
                        return 'OK'

                    elif x == 'SDR':
                        return 'OK'
                    
                    elif x == 'GSR-4':
                        return 'OK'

                    elif x == 'GSR-3':
                        return 'OK'

                    elif x == 'GSR-1':
                        return 'OK'

                    elif x == 'GSRx-1':
                        return 'OK'

                    elif x == 'GSRx-4':
                        return 'OK'

                    elif x == 'GSRx-3':
                        return 'OK'
                    
                    else:
                        return x
                    
                dataUnknownDeviceType['DeviceType']  = dataUnknownDeviceType['DeviceType'].apply(trans_FindBadDeviceType)
                dataUnknownDeviceType = dataUnknownDeviceType.query("DeviceType not in ['OK']")
                dataUnknownDeviceType = dataUnknownDeviceType.reset_index(drop=True)               
                dataUnknownDeviceType ['DuplicatedEntries']=dataUnknownDeviceType.sort_values(by =['StartTimeUTC']).duplicated(['CaseSrNo','DeviceType'],keep='last')
                dataUnknownDeviceType = dataUnknownDeviceType.reset_index(drop=True)
                dataUnknownDeviceType.to_sql('Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_MASTER',con, if_exists="append", index=False)
                dataUnknownDeviceType.to_sql('Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_TEMP',con, if_exists="replace", index=False)
                TotalInValidDeviceEntries = len(dataUnknownDeviceType)                
                TotalEntries = len(data)+len(dataUnknownDeviceType)       
                self.txtTotalEntries.insert(tk.END,TotalEntries)
                self.txtInValidDeviceEntries.insert(tk.END,TotalInValidDeviceEntries)
                con.commit()
                con.close()

        def AnalyzeGSRDeploymentHistoryImport():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)
            self.cur=conn.cursor()                
            data.to_sql('Eagle_GSRDeploymentHistory_ANALYZED',conn, if_exists="replace", index=False)            
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
            

        def UpdateDuplicateGSRDeploymentHistory_MASTER():
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['StartTimeUTC']).duplicated(['CaseSrNo','DeviceType'],keep='last')
            data = data.loc[data.DuplicatedEntries == False, 'CaseSrNo': 'DuplicatedEntries']
            data = data.reset_index(drop=True)
            data.to_sql('Eagle_GSRDeploymentHistory_MASTER',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()


        def SubmitAnalyzeGSRDeploymentHistoryValidToMasterDB():
            iSubmit = tkinter.messagebox.askyesno("Valid Entries Submit to Master DB", "Confirm if you want to Submit the Analyzed Valid Entries to Master DB")
            if iSubmit >0:
                tree.delete(*tree.get_children())
                conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
                cur=conn.cursor()
                Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_ANALYZED ORDER BY `CaseSrNo` ASC ;", conn)
                data = pd.DataFrame(Complete_df)
                data.to_sql('Eagle_GSRDeploymentHistory_MASTER',conn, if_exists="append", index=False)
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_TEMP")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_ANALYZED")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_TEMP")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Submit Complete","All Valid Import Entries are Submitted to Master DB")
                UpdateDuplicateGSRDeploymentHistory_MASTER()
                return
                                                       

        def ExportAnalyzedValidEntries():
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_ANALYZED ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])

            data_SortStartTimeUTC = pd.DataFrame(Complete_df)
            data_SortStartTimeUTC = data_SortStartTimeUTC.sort_values(by =['StartTimeUTC'])
            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortStartTimeUTC.to_excel(file,sheet_name='data_SortStartTimeUTC',index=False)
                    file.close
                    tkinter.messagebox.showinfo("GSRDeploymentHistory Export","GSRDeploymentHistory Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def ExportGSRDeploymentHistoryMasterDB():
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_MASTER ORDER BY `CaseSrNo` ASC ;", conn)
            data_SortByCaseSrNo = pd.DataFrame(Complete_df)
            data_SortByCaseSrNo = data_SortByCaseSrNo.sort_values(by =['CaseSrNo'])
            data_SortStartTimeUTC = pd.DataFrame(Complete_df)
            data_SortStartTimeUTC = data_SortStartTimeUTC.sort_values(by =['StartTimeUTC'])            
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename) as file:
                        data_SortByCaseSrNo.to_excel(file,sheet_name='SortByCaseSrNo',index=False)
                        data_SortStartTimeUTC.to_excel(file,sheet_name='data_SortStartTimeUTC',index=False)
                    file.close
                    tkinter.messagebox.showinfo("GSRDeploymentHistory Export","GSRDeploymentHistory Report Saved as Excel")                                        
            conn.commit()
            conn.close()

        def ViewAnalyzeValidEntries():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_ANALYZED ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            ValidEntries = len(data)
            self.txtValidEntries.insert(tk.END,ValidEntries)
            conn.commit()
            conn.close()


        def iExit():
            iExit= tkinter.messagebox.askyesno("Eagle GSR Inventory Management System", "Confirm if you want to exit")
            if iExit >0:
                self.root.destroy()
                return

        def ResetCount():
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)

        def ClearView():
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)
            tree.delete(*tree.get_children())

        def ClearMasterDB():
            iDelete = tkinter.messagebox.askyesno("Delete GSRDeploymentHistory Master DB", "Confirm if you want to Clear Master GSRDeploymentHistory DB and Start Again")
            if iDelete >0:
                ClearView()
                conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
                cur = conn.cursor()
                tree.delete(*tree.get_children())
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_TEMP")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_ANALYZED")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_MASTER")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_MASTER")
                cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_TEMP")
                conn.commit()
                conn.close()
                return

        def TotalEntries():
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
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
                conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
                cur = conn.cursor()                
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_GSRDeploymentHistory_TEMP WHERE CaseSrNo =? AND DeviceType=? AND \
                                StartTimeUTC =? AND EndTimeUTC =? AND JobName =? ",\
                                (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#5'),\
                                 tree.set(selected_item, '#6'),tree.set(selected_item, '#7'),)) 
                    conn.commit()
                    tree.delete(selected_item)
                conn.commit()
                conn.close()
                TotalEntries()
                return


        def ViewDuplicateEntries():
            tree.delete(*tree.get_children())
            self.txtValidEntries.delete(0,END)
            self.txtTotalEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
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
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalEntries = len(data)       
            self.txtTotalEntries.insert(tk.END,TotalEntries)              
            conn.commit()
            conn.close()

        def ViewInvalidDeviceEntries():
            tree.delete(*tree.get_children())
            self.txtTotalEntries.delete(0,END)
            self.txtValidEntries.delete(0,END)
            self.txtDuplicatedEntries.delete(0,END)
            self.txtInValidDeviceEntries.delete(0,END)            
            conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_TEMP ORDER BY `CaseSrNo` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            TotalInValidDeviceEntries = len(data)       
            self.txtInValidDeviceEntries.insert(tk.END,TotalInValidDeviceEntries)              
            conn.commit()
            conn.close()

        def UpdateSlectedJobName():
            iUpdateSlectedJobName = tkinter.messagebox.askyesno("Update JobName in Database", "Confirm if you want to Update JobName")
            if iUpdateSlectedJobName >0:
                conn = sqlite3.connect("Eagle_GSRDeploymentHistory.db")
                cur = conn.cursor()
                application_window = self.root
                Job_update = simpledialog.askstring("Input Updated JobName", "What is your updated JobName?",
                                parent=application_window)
                if Job_update is not None:
                    for selected_item in tree.selection():
                        cur.execute("UPDATE Eagle_GSRDeploymentHistory_TEMP SET JobName =? WHERE CaseSrNo =? AND DeviceType=? AND \
                                    StartTimeUTC =? AND EndTimeUTC =? ",\
                                    (Job_update, tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#5'),\
                                     tree.set(selected_item, '#6'),))                        
                    conn.commit()
                    conn.close()
                else:
                    tkinter.messagebox.showinfo("Update Error","Please Input Updated JobName") 
                ViewTotalImport()                
                return

                
### Entry Wizard
        self.txtValidEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
        self.txtValidEntries.place(x=167,y=6)

        self.txtInValidDeviceEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 5)
        self.txtInValidDeviceEntries.place(x=650,y=6)

        self.txtDuplicatedEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
        self.txtDuplicatedEntries.place(x=914,y=6)

        self.txtTotalEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
        self.txtTotalEntries.place(x=1150,y=6)

### Button Wizard  
        btnImport = Button(self.root, text="Import GSRDeploymentHistory Files", font=('aerial', 9, 'bold'), height =1, width=29, bd=4,
                    command = ImportGSRDeploymentHistoryFiles)
        btnImport.place(x=2,y=620)
        
        btnAnalyzeImport = Button(self.root, text="Analyze Imported Files ", font=('aerial', 9, 'bold'), height =1, width=19, bd=4,
                           command = AnalyzeGSRDeploymentHistoryImport)        
        btnAnalyzeImport.place(x=222,y=620)
        
        btnAnalyzeSubmit = Button(self.root, text="Submit Analyzed Valid Entries To MasterDB", font=('aerial', 9, 'bold'), height =1, width=35, bd=4,
                           command = SubmitAnalyzeGSRDeploymentHistoryValidToMasterDB)
        btnAnalyzeSubmit.place(x=372,y=620)
        
        btnExportMasterDBValidEntries = Button(self.root, text="Export Master DB", font=('aerial', 9, 'bold'), height =1, width=16, bd=4,
                                        command = ExportGSRDeploymentHistoryMasterDB)
        btnExportMasterDBValidEntries.place(x=632,y=620)

        btnExportAnalyzedValidView = Button(self.root, text="Export Analyzed Valid Entries", font=('aerial', 9, 'bold'), height =1, width=24, bd=1,
                                     command = ExportAnalyzedValidEntries)
        btnExportAnalyzedValidView.place(x=252,y=6)

        btnAnalyzedValidView = Button(self.root, text="View Analyzed Valid Entries", font=('aerial', 9, 'bold'), height =1, width=22, bd=1,
                               command = ViewAnalyzeValidEntries)
        btnAnalyzedValidView.place(x=2,y=6)

        btnViewInvalidDeviceEntries = Button(self.root, text="View Invalid Device Entries", font=('aerial', 9, 'bold'), height =1, width=22, bd=1,
                                             command = ViewInvalidDeviceEntries)
        btnViewInvalidDeviceEntries.place(x=485,y=6)

        btnDelete = Button(self.root, text="Delete Selected Import Entries", font=('aerial', 9, 'bold'), height =1, width=24, bd=4, command = DeleteSelectedImportData)
        btnDelete.place(x=888,y=620)
        btnViewDuplicateEntries = Button(self.root, text="View Duplicate Entries", font=('aerial', 9, 'bold'), height =1, width=19, bd=1, command = ViewDuplicateEntries)
        btnViewDuplicateEntries.place(x=770,y=6)
        btnViewTotalImport = Button(self.root, text="View Total Import", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
        btnViewTotalImport.place(x=1035,y=6)
        btnResetTotal = Button(self.root, text="Reset Count", font=('aerial', 9, 'bold'), height =1, width=10, bd=1, command = ResetCount)
        btnResetTotal.place(x=1267,y=6)
        btnClearView = Button(self.root, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
        btnClearView.place(x=1181,y=620)
        btnClearMasterDB = Button(self.root, text="Clear Master DB", font=('aerial', 9, 'bold'), height =1, width=13, bd=4, command = ClearMasterDB)
        btnClearMasterDB.place(x=1073,y=620)
        btnExit = Button(self.root, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
        btnExit.place(x=1267,y=620)

        tree.bind("<Button-3>", Treepopup_do_popup)
        Treepopup = Menu(tree, tearoff=0)
        Treepopup.add_command(label="Delete Selected Entries", command = DeleteSelectedImportData)
        Treepopup.add_command(label="Update Selected JobName", command = UpdateSlectedJobName)    
        Treepopup.add_separator()
        Treepopup.add_command(label="Exit", command = iExit)


   
if __name__ == '__main__':
    root = Tk()
    application  = GSRDeploymentHistoryImport(root)
    root.mainloop()
