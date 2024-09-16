#Front End
import os
from tkinter import*
import tkinter.messagebox
import EagleDatabase_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog
import pandas as pd
import openpyxl
import csv
import time
import datetime
import EagleDatabase_Inv_Report_Viewer
import EagleDatabase_Inv_Report_Generator
import EagleDatabase_Standard_InvReport_Import
import EagleDatabase_Scanned_InvReport_Import
import EagleDatabase_TransmittalOut_1
import EagleDatabase_Transmittal_In_Import


Default_Date_today   = datetime.date.today()

class Inventory:
    
    def __init__(self,root):
        self.root =root
        self.root.title ("Eagle Electronics Equipment Inventory ")
        self.root.geometry("1350x718+0+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)
        

##  ----------------- Define Variables-------------
        
        catg      = StringVar()
        manuf     = StringVar()
        model     = StringVar()
        main_SN   = StringVar()
        desc      = StringVar()
        asset_SN  = StringVar()
        datestamp = StringVar(self.root, value=Default_Date_today)
        location  = StringVar()
        condition = StringVar()
        origin    = StringVar()
        SEARCH    = StringVar()
        TOTALE    = StringVar()
        SEARCHM   = StringVar()
        TOTALLB    = StringVar()
        

##        #----------------- Function-------------        

        def iExit():
            iExit= tkinter.messagebox.askyesno("Eagle Inventory Management System", "Confirm if you want to exit")
            if iExit >0:
                global root
                root.destroy()
                return

        def ClearData():
            self.txtCatg.delete(0,END)
            self.txtManuf.delete(0,END)
            self.txtModel.delete(0,END)
            self.txtMain_SN.delete(0,END)
            self.txtDesc.delete(0,END)            
            self.txtAsset_SN.delete(0,END)
            self.txtDatestamp.delete(0,END)
            self.txtLocation.delete(0,END)
            self.txtCondition.delete(0,END)
            self.txtOrigin.delete(0,END)            
            self.txtKeySearch.delete(0,END)

        def Reset_TotalDBCount():
            self.txtTotalEntries.delete(0,END)

        def Reset_ListboxCount():
            self.txtTotalLBEntries.delete(0,END)

        def ListBoxTotalEntries():
            self.txtTotalLBEntries.delete(0,END)
            Total_count = len(tree.get_children())
            self.txtTotalLBEntries.insert(tk.END,Total_count)

        def AddData():
            if(len(main_SN.get())!=0) & (len(asset_SN.get())!=0):
                try:
                    EagleDatabase_BackEnd.addInvRec(catg.get(), manuf.get(), model.get(), main_SN.get(),desc.get(), asset_SN.get(), datestamp.get(), location.get(), condition.get(), origin.get())
                    tree.delete(*tree.get_children())
                    tree.insert("", tk.END,values=(catg.get(), manuf.get(), model.get(), main_SN.get(),desc.get(), asset_SN.get(), datestamp.get(), location.get(), condition.get(), origin.get()))
                except:
                    tkinter.messagebox.showinfo("Add Error","Duplicate Manuf SN")
            else:
                    tkinter.messagebox.showinfo("Add Error","Manufacture and Asset SN can not be empty")
            self.txtCatg['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Categ())))
            self.txtManuf['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Manuf())))
            self.txtModel['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Model())))
            self.txtMain_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Main_SN())))
            self.txtDesc['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Desc())))
            self.txtAsset_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Asset_SN())))
            self.txtDatestamp['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Datestamp())))
            self.txtLocation['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Location())))
            self.txtCondition['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Condition())))
            self.txtOrigin['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Origin())))
            ListBoxTotalEntries()
            UpdateNewEntryToMASTER()
            TotalInvCount()
                    
        def DisplayData():
               tree.delete(*tree.get_children())
               for row in EagleDatabase_BackEnd.viewData():
                   tree.insert("", tk.END, values=row)                   
               self.txtCatg['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Categ())))
               self.txtManuf['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Manuf())))
               self.txtModel['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Model())))
               self.txtMain_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Main_SN())))
               self.txtDesc['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Desc())))
               self.txtAsset_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Asset_SN())))
               self.txtDatestamp['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Datestamp())))
               self.txtLocation['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Location())))
               self.txtCondition['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Condition())))
               self.txtOrigin['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Origin())))               
               TotalInvCount()
               ListBoxTotalEntries()

        def searchDatabase():
            tree.delete(*tree.get_children())
            for row in EagleDatabase_BackEnd.searchData(catg.get(), manuf.get(), model.get(), main_SN.get(),desc.get(), asset_SN.get(), datestamp.get(), location.get(), condition.get(), origin.get()):
                tree.insert("", tk.END, values=row)
            ListBoxTotalEntries()

        def InventoryRec(event):
            for nm in tree.selection():
                sd = tree.item(nm, 'values')
                self.txtCatg.delete(0,END)
                self.txtCatg.insert(tk.END,sd[0])                
                self.txtManuf.delete(0,END)
                self.txtManuf.insert(tk.END,sd[1])
                self.txtModel.delete(0,END)
                self.txtModel.insert(tk.END,sd[2])
                self.txtMain_SN.delete(0,END)
                self.txtMain_SN.insert(tk.END,sd[3])
                self.txtDesc.delete(0,END)
                self.txtDesc.insert(tk.END,sd[4])                
                self.txtAsset_SN.delete(0,END)
                self.txtAsset_SN.insert(tk.END,sd[5])
                self.txtDatestamp.delete(0,END)
                self.txtDatestamp.insert(tk.END,sd[6])
                self.txtLocation.delete(0,END)
                self.txtLocation.insert(tk.END,sd[7])
                self.txtCondition.delete(0,END)
                self.txtCondition.insert(tk.END,sd[8])
                self.txtOrigin.delete(0,END)
                self.txtOrigin.insert(tk.END,sd[9])

        def DeleteData():
            SelectionTree = tree.selection()
            if len(SelectionTree)>0:
                iDelete = tkinter.messagebox.askyesno("Delete Entry From Master Database", "Confirm if you want to Delete From Master DataBase")
                if iDelete >0:
                    conn = sqlite3.connect("Eagle_Inventory.db")
                    cur = conn.cursor()
                    if(len(main_SN.get())!=0):
                        for selected_item in tree.selection():
                            cur.execute("DELETE FROM Eagle_Inventory WHERE model =? AND main_SN = ? ", (tree.set(selected_item, '#3'), tree.set(selected_item, '#4'),))
                            conn.commit()
                            tree.delete(selected_item)
                        conn.commit()
                        conn.close()
                    ClearData()
                    DisplayData()
                    TotalInvCount()
                    ListBoxTotalEntries()
                    return
            else:
                tkinter.messagebox.showinfo("Delete Error","Please Select Entries To Delete From Master DB")

        def DeleteSelectedFromLB():
            SelectionTree = tree.selection()
            if len(SelectionTree)>0:
                if(len(main_SN.get())!=0):
                    for selected_item in tree.selection():
                        tree.delete(selected_item)
                    ListBoxTotalEntries()
                    TotalInvCount()
            else:
                tkinter.messagebox.showinfo("Delete Error","Please Select Entries To Delete From List Box")
                

        def UpdateSlectedLocation():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            if len(ListBox_DF)>0:
                SelectionTree = tree.selection()
                if len(SelectionTree)>0:
                    iUpdateSlectedLocation = tkinter.messagebox.askyesno("Update Location in Database", "Confirm if you want to Update Location")
                    if iUpdateSlectedLocation >0:
                        conn = sqlite3.connect("Eagle_Inventory.db")
                        cur = conn.cursor()
                        application_window = self.root
                        location_update = simpledialog.askstring("Input Updated Location", "What is your updated location?",
                                        parent=application_window)
                        if location_update is not None:
                            for selected_item in tree.selection():
                                cur.execute("UPDATE Eagle_Inventory SET location =? WHERE model = ? AND main_SN =?", (location_update, tree.set(selected_item, '#3'), tree.set(selected_item, '#4')))
                                conn.commit()                        
                            conn.commit()
                            conn.close()
                        else:
                            tkinter.messagebox.showinfo("Update Error","Please Input Updated Location") 
                        ClearData()
                        DisplayData()
                        TotalInvCount()
                        ListBoxTotalEntries()
                        return
                else:
                    tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Location")
            else:
                tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Location")
                

        def UpdateSlectedDescription():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            if len(ListBox_DF)>0:
                SelectionTree = tree.selection()
                if len(SelectionTree)>0:
                    iUpdateSlectedDesc = tkinter.messagebox.askyesno("Update Description in Database", "Confirm if you want to Update Description")
                    if iUpdateSlectedDesc >0:
                        conn = sqlite3.connect("Eagle_Inventory.db")
                        cur = conn.cursor()
                        application_window = self.root
                        desc_update = simpledialog.askstring("Input Updated Description", "What is your updated description?",
                                        parent=application_window)
                        if desc_update is not None:
                            for selected_item in tree.selection():
                                cur.execute("UPDATE Eagle_Inventory SET desc =? WHERE model = ? AND main_SN =?", (desc_update, tree.set(selected_item, '#3'), tree.set(selected_item, '#4')))
                                conn.commit()                        
                            conn.commit()
                            conn.close()
                        else:
                            tkinter.messagebox.showinfo("Update Error","Please Input Updated Description") 
                        ClearData()
                        DisplayData()
                        TotalInvCount()
                        return
                else:
                    tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Description")
            else:
                tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Description")

        def UpdateSlectedCondition():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            if len(ListBox_DF)>0:
                SelectionTree = tree.selection()
                if len(SelectionTree)>0:
                    iUpdateSlectedCond = tkinter.messagebox.askyesno("Update Condition in Database", "Confirm if you want to Update Condition")
                    if iUpdateSlectedCond >0:
                        conn = sqlite3.connect("Eagle_Inventory.db")
                        cur = conn.cursor()
                        application_window = self.root
                        cond_update = simpledialog.askstring("Input Updated Condition", "What is your updated condition?",
                                        parent=application_window)
                        if cond_update is not None:
                            for selected_item in tree.selection():
                                cur.execute("UPDATE Eagle_Inventory SET condition =? WHERE model = ? AND main_SN =?", (cond_update, tree.set(selected_item, '#3'), tree.set(selected_item, '#4')))
                                conn.commit()                        
                            conn.commit()
                            conn.close()
                        else:
                            tkinter.messagebox.showinfo("Update Error","Please Input Updated Condition") 
                        ClearData()
                        DisplayData()
                        TotalInvCount()
                    return
                else:
                    tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Condition")            
            else:
                tkinter.messagebox.showinfo("Update Error","Please Select Entries To Update Condition")

        def TotalInvCount():
            self.txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("Eagle_Inventory.db")
            Inv_df = pd.read_sql_query("select * from Eagle_Inventory ;", conn)
            Inv_count_data = pd.DataFrame(Inv_df)
            Total_count = Inv_count_data['main_SN'].count()
            self.txtTotalEntries.insert(tk.END,Total_count)
            conn.commit()
            conn.close()

        def ExportCompleteDB():
            conn = sqlite3.connect("Eagle_Inventory.db")
            Complete_df = pd.read_sql_query("select * from Eagle_Inventory ;", conn)
            Export_Database = pd.DataFrame(Complete_df)
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,\
                       defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.csv'):
                    with open(filename, 'w') as file:
                        Export_Database.to_csv(file,index=None)
                    file.close
                    tkinter.messagebox.showinfo("DB Export","DB Saved as CSV")
                else:
                    with pd.ExcelWriter(filename) as file:
                        Export_Database.to_excel(file,sheet_name='InventoryDB',index=False)
                    file.close
                    tkinter.messagebox.showinfo("DB Export","DB Saved as Excel")                    
                        
            conn.commit()
            conn.close()


        def ExportCompleteListBoxEntries():
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            ListBox_DF.rename(columns = {0:'Category', 1:'Manuf', 2:'Model', 3:'ManfSN', 4:'Desc',
                                          5: 'Asset_SN', 6:'Date', 7:'Location', 8:'Condition',9:'Origin'},inplace = True)
                        
            Export_ListBox = pd.DataFrame(ListBox_DF)
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,
                       defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if filename:
                if filename.endswith('.xlsx'):
                    with pd.ExcelWriter(filename, engine='xlsxwriter') as file:
                        Export_ListBox.to_excel(file,sheet_name='Transmittal Front Page', index=False, startrow=0, header=True)       
                    file.close
                    tkinter.messagebox.showinfo("List Box Export","All Entries in ListBox Saved as Excel")                    
                        
        def update():
            if(len(main_SN.get())!=0) & (len(asset_SN.get())!=0):
                conn = sqlite3.connect("Eagle_Inventory.db")
                cur = conn.cursor()
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_Inventory WHERE main_SN =?", (tree.set(selected_item, '#4'),))
                    conn.commit()
                    tree.delete(selected_item)
                    conn.close()

            if(len(main_SN.get())!=0) & (len(asset_SN.get())!=0):
                try:
                    EagleDatabase_BackEnd.addInvRec(catg.get(), manuf.get(), model.get(), main_SN.get(), desc.get(), asset_SN.get(), datestamp.get(), location.get(), condition.get(), origin.get())
                    tree.delete(*tree.get_children())
                    tree.insert("", tk.END,values=(catg.get(), manuf.get(), model.get(), main_SN.get(), desc.get(), asset_SN.get(), datestamp.get(), location.get(), condition.get(), origin.get()))
                except:
                    tkinter.messagebox.showinfo("Add Error","Duplicate Asset SN")
            else:
                    tkinter.messagebox.showinfo("Add Error","Manufacture & Asset SN Can Not Be Empty")
                    
        def ClearListBoxView():
               tree.delete(*tree.get_children())
               Reset_TotalDBCount()
               Reset_ListboxCount()

        def InventoryReportGenerator():
            window = Tk()
            window.title("Inventory Report Genarator")
            window.config(bg="purple")
            width = 500
            height = 200
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = (screen_width/2) - (width/2)
            y = (screen_height/2) - (height/2)
            window.geometry("%dx%d+%d+%d" % (width, height, x, y))
            window.grid_rowconfigure(1, weight=1)
            window.grid_columnconfigure(0, weight=1)
            window.resizable(0, 0)

            def WExit():
                window.destroy()

            L1 = Label(window, text = "A: Generate Inventory Report Count By Category, Model and Location", font=("arial", 10,'bold'),bg = "cadet blue").place(x=10,y=20)
            btnViewInvA = Button(window, text="Generate Inventory Report", font=('aerial', 10, 'bold'), height =1, width=21, bd=4, command =EagleDatabase_Inv_Report_Generator.Generate_Inv_Report)
            btnViewInvA.place(x=10,y=50)
            L2 = Label(window, text = "B: View Gerated Report", font=("arial", 10,'bold'),bg = "cadet blue").place(x=10,y=100)
            btnViewInvB = Button(window, text="View Report", font=('aerial', 10, 'bold'), height =1, width=10, bd=4, command =EagleDatabase_Inv_Report_Viewer.View_Inv_Report)
            btnViewInvB.place(x=10,y=130)

            btnEXIT= Button(window, text="Exit Viewer", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, bg="red", command =WExit)
            btnEXIT.place(x=410,y=160)

        def KeySearch():
            if SEARCH.get() != "":
                tree.delete(*tree.get_children())
                conn = sqlite3.connect("Eagle_Inventory.db")
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM `Eagle_Inventory` WHERE `catg` LIKE ? OR `manuf` LIKE ? OR\
                               `model` LIKE ? OR `main_SN` LIKE ? OR `asset_SN` LIKE ? OR `location` LIKE ? OR \
                               `condition` LIKE ? OR `origin` LIKE ? OR `desc` LIKE ? ",\
                               ('%'+str(SEARCH.get())+'%', '%'+str(SEARCH.get())+'%', '%'+str(SEARCH.get())+'%',\
                                '%'+ str(SEARCH.get())+'%' , '%'+ str(SEARCH.get())+'%' , '%'+ str(SEARCH.get())+'%' ,\
                                '%'+ str(SEARCH.get())+'%' , '%'+ str(SEARCH.get())+'%' , '%'+ str(SEARCH.get())+'%'))                               
                                                            
                fetch = cursor.fetchall()
                for data in fetch:
                    tree.insert('', 'end', values=(data))
                cursor.close()
                conn.close()
            ListBoxTotalEntries()


        def ExportListBoxView():
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,\
                       defaultextension='.csv', filetypes = (("CSV file",".csv"),("Text file",".txt")))
            if filename:
                    if filename.endswith('.csv'):
                        with open(filename, 'w') as file:
                            file.write('Category' + ',' + 'Manufacturer' + ',' + 'Model' + ',' + 'Manuf SN' + ',' +  'Description' + ',' + 'Asset SN' +\
                                       ',' + 'Date' + ',' + 'Location' + ',' + 'Condition' + ',' + 'Origin' + '\n')
                            for item in tree.selection():
                                list_item = (tree.item(item, 'values'))
                                x1= list_item[0]
                                x2= list_item[1]
                                x3= list_item[2]
                                x4= list_item[3]
                                x5= list_item[4]
                                x6= list_item[5]
                                x7= list_item[6]
                                x8= list_item[7]
                                x9= list_item[8]
                                x10= list_item[9]                                
                                file.write( x1 + ',' + x2 + ',' + x3 + ',' + x4 + ',' + x5 + ',' + x6 + ',' + x7 + ',' + x8 +  ',' + x9 + ',' + x10 + '\n')
                        file.close
                        tkinter.messagebox.showinfo("Save file","File Saved as CSV")

                    else:
                        with open(filename, 'w') as file:
                            file.write('Category' + ',' + 'Manufacturer' + ',' + 'Model' + ',' + 'Manuf SN' + ',' +  'Description' + ',' + 'Asset SN' +\
                                       ',' + 'Date' + ',' + 'Location' + ',' + 'Condition' + ',' + 'Origin' + '\n')
                            for item in tree.selection():
                                list_item = (tree.item(item, 'values'))
                                x1= list_item[0]
                                x2= list_item[1]
                                x3= list_item[2]
                                x4= list_item[3]
                                x5= list_item[4]
                                x6= list_item[5]
                                x7= list_item[6]
                                x8= list_item[7]
                                x9= list_item[8]
                                x10= list_item[9]
                                file.write( x1 + ',' + x2 + ',' + x3 + ',' + x4 + ',' + x5 + ',' + x6 + ',' + x7 + ',' + x8 +  ',' + x9 + ',' + x10 + '\n')                                
                        file.close
                        tkinter.messagebox.showinfo("Save file","File Saved as TEXT")


        def AddListForTransmittalOut():
            cur_id = tree.focus()
            selvalue = tree.item(cur_id)['values']
            Length_Selected  =  (len(selvalue))
            if Length_Selected != 0:
                for item in tree.selection():
                    list_item = (tree.item(item, 'values'))                
                    con= sqlite3.connect("Eagle_Inventory.db")
                    cur=con.cursor()
                    cur.execute("INSERT INTO Eagle_Inventory_transmittalOut1 VALUES (?,?,?,?,?,?,?,?,?,?)",(list_item))
                    con.commit()
                    con.close()
                tkinter.messagebox.showinfo("Add List Message","Selected List of Entries Added To Transmittal Out Database")
            else:
                tkinter.messagebox.showinfo("Add List Message","Please Select List of Entries To Generate Transmittal Out Database")
                

        def GenerateTransmittalOut1():
            iSubmit = tkinter.messagebox.askyesno("Generate Transmittal Out",
                                        "Confirm if you want to Generate Transmittal Out")
            if iSubmit >0:
                EagleDatabase_TransmittalOut_1.GenerateTransmittalOutFirstOption()
            return

        def UpdateImportToMASTER():
            conn = sqlite3.connect("Eagle_Inventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory ORDER BY `catg` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['datestamp']).duplicated(['main_SN','model'],keep='last')
            data = data.loc[data.DuplicatedEntries == False, 'catg': 'origin']
            data = data.reset_index(drop=True)
            data.to_sql('Eagle_Inventory',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()
            DisplayData()

        def UpdateNewEntryToMASTER():
            conn = sqlite3.connect("Eagle_Inventory.db")
            Complete_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory ORDER BY `catg` ASC ;", conn)
            data = pd.DataFrame(Complete_df)
            data ['DuplicatedEntries']=data.sort_values(by =['datestamp']).duplicated(['main_SN','model'],keep='last')
            data = data.loc[data.DuplicatedEntries == False, 'catg': 'origin']
            data = data.reset_index(drop=True)
            data.to_sql('Eagle_Inventory',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()

        def ClearAllTempDataBase():
            try:
                conn = sqlite3.connect("Eagle_Inventory.db")
                cur=conn.cursor()
                cur.execute("DELETE FROM Eagle_Inventory_transmittalOut1")
                cur.execute("DELETE FROM Eagle_Inventory_transmittalOut2")
                cur.execute("DELETE FROM Eagle_Inventory_Merged_temp")
                cur.execute("DELETE FROM Eagle_Inventory_Scan_TEMP")
                cur.execute("DELETE FROM Eagle_Inventory_temp")
                cur.execute("DELETE FROM Eagle_Inventory_Report_1")
                cur.execute("DELETE FROM Eagle_Inventory_Report_2")
                cur.execute("DELETE FROM Eagle_Inventory_Report_3")
                cur.execute("DELETE FROM Eagle_Inventory_transmittalOutFrontPage")
                
                cur.execute("DROP TABLE Eagle_Inventory_transmittalOut1")
                cur.execute("DROP TABLE Eagle_Inventory_transmittalOut2")
                cur.execute("DROP TABLE Eagle_Inventory_Merged_temp")
                cur.execute("DROP TABLE Eagle_Inventory_Scan_TEMP")
                cur.execute("DROP TABLE Eagle_Inventory_temp")
                cur.execute("DROP TABLE Eagle_Inventory_Report_1")
                cur.execute("DROP TABLE Eagle_Inventory_Report_2")
                cur.execute("DROP TABLE Eagle_Inventory_Report_3")
                cur.execute("DROP TABLE Eagle_Inventory_transmittalOutFrontPage")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Clear DB","All Temp Database are Cleared")
            except:
                tkinter.messagebox.showinfo("Clear DB","All Temp Database are Already Cleared")

        def GenerateAllTempDataBase():
            EagleDatabase_BackEnd.inventoryData()
            tkinter.messagebox.showinfo("Temp DB Message","All Temp Database are Generated")

        def ImportMasterDBFile():
            EagleDatabase_Standard_InvReport_Import.ImportFile()

        def ImportTransmittalInToMasterDB():
            EagleDatabase_Transmittal_In_Import.ImportTransmittalInFile()

        def ImportBatchScannedFileTransmittalOut2():
            iSubmit = tkinter.messagebox.askyesno("Import Scanned File And Generate Transmittal Out",
                                        "Confirm if you want to Import Scanned File And Generate Transmittal Out")
            if iSubmit >0:
                EagleDatabase_Scanned_InvReport_Import.ImportScannedFile()
            return
            
            
                                        
        #----------------- Frames-------------
        menu = Menu(self.root)
        self.root.config(menu=menu)
        filemenu = Menu(menu, tearoff=0)
        menu.add_cascade(label="File", menu=filemenu)
        filemenu.add_command(label="Import MasterDB File", command=EagleDatabase_Standard_InvReport_Import.ImportFile)
        filemenu.add_command(label="Import Transmittal In File", command=EagleDatabase_Transmittal_In_Import.ImportTransmittalInFile)
        filemenu.add_command(label="Export Complete DB", command=ExportCompleteDB)
        filemenu.add_command(label="Export All Listbox Entries", command=ExportCompleteListBoxEntries)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=iExit)

        Reportmenu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Report", menu=Reportmenu)
        Reportmenu.add_command(label="Generate Inventory Report", command=InventoryReportGenerator)

        Advancedmenu = Menu(menu, tearoff=0)
        menu.add_cascade(label="Advanced", menu=Advancedmenu)
        Advancedmenu.add_command(label="Clear All Temp Database", command=ClearAllTempDataBase)
        Advancedmenu.add_command(label="Generate All Temp Database", command=GenerateAllTempDataBase)
        
        
        TitFrame = Frame(self.root, bd = 2, padx= 5, pady= 4, bg = "#006dcc", relief = RIDGE)
        TitFrame.pack(side = TOP)

        self.lblTit = Label(TitFrame, font=('aerial', 12, 'bold'), text="Eagle Equipment Management System", bg="yellow")
        self.lblTit.grid()

        L1 = Label(self.root, text = "A: Inventory Info", font=("arial", 11,'bold'),bg = "cadet blue").place(x=10,y=14)
        L2 = Label(self.root, text = "B: List Box (LB) Details", font=("arial", 11,'bold'),bg = "cadet blue").place(x=521,y=80)
        
        

        DataFrameLEFT = LabelFrame(self.root, bd = 2, width = 520, height = 700, padx= 8, pady= 10,relief = RIDGE,
                                   bg = "Ghost White",font=('aerial', 15, 'bold'))
        DataFrameLEFT.place(x=10,y=40)

                #----------------- Tree View Frames-------------
        TableMargin = Frame(self.root)
        TableMargin.pack(side=TOP)
        TableMargin.pack(side=RIGHT)
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5", "column6", "column7", "column8",  "column9", "column10"),
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
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=('aerial', 8), foreground="black")
        style.configure("Treeview", foreground='black')
        style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')

        Treepopup = Menu(tree, tearoff=0)
        Treepopup.add_command(label="Add Entries For TransmittalOut", command=AddListForTransmittalOut)
        Treepopup.add_separator()
        Treepopup.add_command(label="Delete Selected Entries", command=DeleteData)        
        Treepopup.add_command(label="Update Selected Location", command=UpdateSlectedLocation)
        Treepopup.add_command(label="Update Selected Description", command=UpdateSlectedDescription)
        Treepopup.add_command(label="Update Selected Condition", command=UpdateSlectedCondition)
        Treepopup.add_separator()
        Treepopup.add_command(label="Exit", command=iExit)

        def Treepopup_do_popup(event):
            # display the popup menu
            try:
                Treepopup.tk_popup(event.x_root, event.y_root, 0)
            finally:
                # make sure to release the grab (Tk 8.0a1 only)
                Treepopup.grab_release()

        tree.bind("<Button-3>", Treepopup_do_popup)
        tree.pack()
    

        #----------------- Labels and Entry Wizard------------

        self.lblCatg = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "1. Category:", padx =10, pady= 10, bg = "Ghost White")
        self.lblCatg.grid(row =0, column = 0, sticky =W)
        self.txtCatg = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = catg, width = 40)
        self.txtCatg.grid(row =0, column = 1)
        self.txtCatg['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Categ())))

        self.lblManuf = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "2. Manufacturer:", padx =10, pady= 10, bg = "Ghost White")
        self.lblManuf.grid(row =1, column = 0, sticky =W)
        self.txtManuf = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = manuf, width = 40)
        self.txtManuf.grid(row =1, column = 1)
        self.txtManuf['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Manuf())))

        self.lblModel = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "3. Model Name:", padx =10, pady= 10, bg = "Ghost White")
        self.lblModel.grid(row =2, column = 0, sticky =W)
        self.txtModel = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = model, width = 40)
        self.txtModel.grid(row =2, column = 1)
        self.txtModel['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Model())))
        
        self.lblMain_SN = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "4. Manf. SN:", padx =10, pady= 10, bg = "Ghost White")
        self.lblMain_SN.grid(row =3, column = 0, sticky =W)
        self.txtMain_SN = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = main_SN, width = 40)
        self.txtMain_SN.grid(row =3, column = 1)
        self.txtMain_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Main_SN())))
        
        self.lblDesc = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "5. Description:", padx =10, pady= 10, bg = "Ghost White")
        self.lblDesc.grid(row =4, column = 0, sticky =W)
        self.txtDesc = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = desc, width = 40)
        self.txtDesc.grid(row =4, column = 1)
        self.txtDesc['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Desc())))

        self.lblAsset_SN = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "6. Asset SN:", padx =10, pady= 10, bg = "Ghost White")
        self.lblAsset_SN.grid(row =5, column = 0, sticky =W)
        self.txtAsset_SN = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = asset_SN, width = 40)
        self.txtAsset_SN.grid(row =5, column = 1)
        self.txtAsset_SN['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Asset_SN())))

        self.lblDatestamp = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "7. Date (yyyy-mm-dd):", padx =10 , pady= 10, bg = "Ghost White")
        self.lblDatestamp.grid(row =6, column = 0, sticky =W)
        self.txtDatestamp = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = datestamp, width = 40)
        self.txtDatestamp.grid(row =6, column = 1)
        self.txtDatestamp['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Datestamp())))

        self.lblLocation = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "8. Location:", padx =10, pady= 10, bg = "Ghost White")
        self.lblLocation.grid(row =7, column = 0, sticky =W)
        self.txtLocation = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = location, width = 40)
        self.txtLocation.grid(row =7, column = 1)
        self.txtLocation['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Location())))

        self.lblCondition = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "9. Condition:", padx =10, pady= 10, bg = "Ghost White")
        self.lblCondition.grid(row =8, column = 0, sticky =W)
        self.txtCondition = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = condition, width = 40)
        self.txtCondition.grid(row =8, column = 1)
        self.txtCondition['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Condition())))

        self.lblOrigin = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "10. Origin:", padx =10, pady= 10, bg = "Ghost White")
        self.lblOrigin.grid(row =9, column = 0, sticky =W)
        self.txtOrigin = ttk.Combobox(DataFrameLEFT, font=('aerial', 10, 'bold'), textvariable = origin, width = 40)
        self.txtOrigin.grid(row =9, column = 1)
        self.txtOrigin['values'] = sorted(list(set(EagleDatabase_BackEnd.Combo_input_Origin())))

        self.txtKeySearch  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=SEARCH, width = 20)
        self.txtKeySearch.place(x=710,y=80)

        #----------------- Tree View Select Event------------
        
        tree.bind('<<TreeviewSelect>>',InventoryRec)

        #----------------- Button Widget------------

        btnAddData = Button(self.root, text="Add Entry", font=('aerial', 10, 'bold'), height =1, width=10, bd=4,command = AddData )
        btnAddData.place(x=10,y=467)
        btnClearData = Button(self.root, text="Clear Entry", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearData)
        btnClearData.place(x=403,y=8)        
        btnSearchData = Button(self.root, text="Search Entry", font=('aerial', 10, 'bold'), height =1, width=10, bd=4, command = searchDatabase)
        btnSearchData.place(x=395,y=467)
        btnUpdateData = Button(self.root, text="Modify Entry", font=('aerial', 10, 'bold'), height =1, width=10, bd=4, command = update)
        btnUpdateData.place(x=110,y=467)
        btnDeleteData = Button(self.root, text="Delete Entry", font=('aerial', 10, 'bold'), height =1, width=10, bd=4, command = DeleteData)
        btnDeleteData.place(x=295,y=467)
        btnKeySearchListbox = Button(self.root, text="Search by Keyword", font=('aerial', 9, 'bold'), height =1, width=16, bd=2, command = KeySearch)
        btnKeySearchListbox.place(x=710,y=52)

        L4 = Label(self.root, text = "C: Import Master DB File", font=("arial", 11,'bold'),bg = "orange").place(x=10,y=530)

        btnImportMasterDBFile = Button(self.root, text="Import MasterDB File", font=('aerial', 10, 'bold'), height =1, width=17, bd=4, command = ImportMasterDBFile)
        btnImportMasterDBFile.place(x=10,y=560)
                
        L5 = Label(self.root, text = "D: Update Master DB File", font=("arial", 11,'bold'),bg = "orange").place(x=10,y=615)
        btnUpdatetoDB = Button(self.root, text="Update Import To DB", font=('aerial', 10, 'bold'), height =1, width=17, bd=4, command =UpdateImportToMASTER)
        btnUpdatetoDB.place(x=10,y=645)

        L6 = Label(self.root, text = "E: View Master DB File", font=("arial", 11,'bold'),bg = "orange").place(x=330,y=530)
        btnViewMasterDB = Button(self.root, text="Populate Master DB", font=('aerial', 10, 'bold'), height =1, width=17, bd=4, command = DisplayData)
        btnViewMasterDB.place(x=330,y=560)
        self.txtTotalEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=TOTALE, width = 9, bd=1)
        self.txtTotalEntries.place(x=360,y=595)

        L8 = Label(self.root, text = "Quantity LB", font=("arial", 11,'bold'),bg = "cadet blue").place(x=1258,y=57)
        self.txtTotalLBEntries  = Entry(self.root, font=('aerial', 12, 'bold'),textvariable=TOTALLB, width = 9)
        self.txtTotalLBEntries.place(x=1262,y=78)

        btnClearListbox = Button(self.root, text="Clear Listbox", font=('aerial', 10, 'bold'), height =1, width=11, bd=2, command = ClearListBoxView)
        btnClearListbox.place(x=518,y=653)        
        btnExportListBox = Button(self.root, text="Export Selected", font=('aerial', 10, 'bold'), height =1, width=13, bd=2, command = ExportListBoxView)
        btnExportListBox.place(x=617,y=653)
        btnDeleteSelectedFromLB = Button(self.root, text="Delete From LB", font=('aerial', 10, 'bold'), height =1, width=13, bd=2, command = DeleteSelectedFromLB)
        btnDeleteSelectedFromLB.place(x=733,y=653)

        L9 = Label(self.root, text = "Import Scanned File & Update Master DB & TransmittalOut Wizard :", font=("arial", 8,'bold'),bg = "cadet blue").place(x=870,y=653)

        btnImportScannedFile = Button(self.root, text="Import - UpdateDB \n TransmittalOut Wizard", font=('aerial', 10, 'bold'), height =2, width=18, bd=2, command = ImportBatchScannedFileTransmittalOut2)
        btnImportScannedFile.place(x=1000,y=675)

        btnExitData = Button(self.root, text="Exit", font=('aerial', 10, 'bold'), height =1, width=6, bd=2, command = iExit)
        btnExitData.place(x=1290,y=653)

        L10 = Label(self.root, text = "Add List Entries and Generate Transmittal From List:", font=("arial", 8,'bold'),bg = "cadet blue").place(x=935,y=52)
        btnAddListTransmittalOut = Button(self.root, text="Add List", font=('aerial', 10, 'bold'), height =1, width=8, bd=1, command = AddListForTransmittalOut)
        btnAddListTransmittalOut.place(x=940,y=75)

        L11 = Label(self.root, text = "> & > ", font=("arial", 10,'bold'), bg= 'cadet blue').place(x=1015,y=75)

        btnAddListToGenerateTransmittalOut = Button(self.root, text="Generate Transmittal Out", font=('aerial', 10, 'bold'), height =1, width=20, bd=1, command = GenerateTransmittalOut1)
        btnAddListToGenerateTransmittalOut.place(x=1060,y=75)

if __name__ == '__main__':
    root = Tk()
    application  = Inventory (root)
    root.mainloop()

