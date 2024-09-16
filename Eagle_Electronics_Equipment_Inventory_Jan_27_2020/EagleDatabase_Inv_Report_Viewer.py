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

def View_Inv_Report():
    tkinter.messagebox.showinfo("View Inventory Report A","Please Make Sure You View Report After Generating Inventory Report")
    SEARCHMODEL = StringVar()
    SEARCHCATG = StringVar()
    window = Tk()
    window.title("Inventory Report Viewer")
    window.config(bg="ghost white")
    width = 1250
    height = 600
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry("%dx%d+%d+%d" % (width, height, x, y))
    window.grid_rowconfigure(1, weight=1)
    window.grid_columnconfigure(0, weight=1)
    window.resizable(0, 0)
    TableMargin = Frame(window)
    TableMargin1 = Frame(window)
    TableMargin2 = Frame(window)
    TableMargin.place(x=10, y =40)
    TableMargin1.place(x=340, y =40)
    TableMargin2.place(x=770, y =40)

    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    scrollbarx1 = Scrollbar(TableMargin1, orient=HORIZONTAL)
    scrollbary1 = Scrollbar(TableMargin1, orient=VERTICAL)
    scrollbarx2 = Scrollbar(TableMargin2, orient=HORIZONTAL)
    scrollbary2 = Scrollbar(TableMargin2, orient=VERTICAL)

    tree = ttk.Treeview(TableMargin, column=("column1", "column2"),selectmode="extended",
                                height=20, show='headings')
    tree1 = ttk.Treeview(TableMargin1, column=("column1", "column2", "column3"),selectmode="extended",
                                height=20, show='headings')
    tree2 = ttk.Treeview(TableMargin2, column=("column1", "column2", "column3",  "column4" ),selectmode="extended",
                                    height=20, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbary1.config(command=tree1.yview)
    scrollbary1.pack(side=RIGHT, fill=Y)
    scrollbary2.config(command=tree2.yview)
    scrollbary2.pack(side=RIGHT, fill=Y)

    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    scrollbarx1.config(command=tree1.xview)
    scrollbarx1.pack(side=BOTTOM, fill=X)
    scrollbarx2.config(command=tree2.xview)
    scrollbarx2.pack(side=BOTTOM, fill=X)

    tree.heading("#1", text="Category", anchor=W)
    tree.heading("#2", text="Count By Category", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=150)            
    tree.column('#2', stretch=NO, minwidth=0, width=150)

    tree1.heading("#1", text="Category", anchor=W)
    tree1.heading("#2", text="Model", anchor=W)
    tree1.heading("#3", text="Count By Model", anchor=W)
    tree1.column('#1', stretch=NO, minwidth=0, width=150)            
    tree1.column('#2', stretch=NO, minwidth=0, width=100)
    tree1.column('#3', stretch=NO, minwidth=0, width=150)

    tree2.heading("#1", text="Category", anchor=W)
    tree2.heading("#2", text="Location", anchor=W)
    tree2.heading("#3", text="Model Name", anchor=W)
    tree2.heading("#4", text="Count By Model", anchor=W)
    tree2.column('#1', stretch=NO, minwidth=0, width=150)            
    tree2.column('#2', stretch=NO, minwidth=0, width=80)
    tree2.column('#3', stretch=NO, minwidth=0, width=100)
    tree2.column('#4', stretch=NO, minwidth=0, width=100)

    tree.pack()
    tree1.pack()
    tree2.pack()

    # Defining Functions ######
    def InvExit():
        window.destroy()

    def ResetSearch():
        tree.delete(*tree.get_children())
        tree1.delete(*tree1.get_children())
        tree2.delete(*tree.get_children())
        ComboCatglL2.delete(0,END)
        ComboModelL2.delete(0,END)
        
    def Combo_Inv_Category():
        conn= sqlite3.connect("Eagle_Inventory.db")
        CategoryCount_DF = pd.read_sql_query("SELECT Category FROM Eagle_Inventory_Report_3 ;", conn)
        CategoryCount_DF = pd.DataFrame(CategoryCount_DF)
        data = []
        for each_rec in CategoryCount_DF.Category:
            data.append((each_rec))
        conn.close()
        return data

    def Combo_Inv_Model():
        conn= sqlite3.connect("Eagle_Inventory.db")
        ModelCount_DF = pd.read_sql_query("SELECT Model_Name FROM Eagle_Inventory_Report_2 ;", conn)
        ModelCount_DF = pd.DataFrame(ModelCount_DF)
        data = []
        for each_rec in ModelCount_DF.Model_Name:
            data.append((each_rec))
        conn.close()
        return data

    def ExportCompleteINV():
        conn = sqlite3.connect("Eagle_Inventory.db")
        Complete_df2 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_2 ORDER BY `Model_Name` ASC ;", conn)
        Complete_df3 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_3 ORDER BY `Category` ASC ;", conn)
        Complete_df1 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_1 ORDER BY `Category` ASC ;", conn)        
        Export_Database2 = pd.DataFrame(Complete_df2)
        Export_Database1 = pd.DataFrame(Complete_df1)
        Export_Database3 = pd.DataFrame(Complete_df3)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
        if filename:
            if filename.endswith('.xlsx'):
                with pd.ExcelWriter(filename) as file:
                    Export_Database3.to_excel(file,sheet_name='Inventory Report1',index=False)
                    Export_Database1.to_excel(file,sheet_name='Inventory Report2',index=False)
                    Export_Database2.to_excel(file,sheet_name='Inventory Report3',index=False)
                file.close
                tkinter.messagebox.showinfo("Inventory Export","Inventory Report Saved as Excel")                    
                    
        conn.commit()
        conn.close()

    def ClearInvView():
        tree.delete(*tree.get_children())
        tree1.delete(*tree1.get_children())
        tree2.delete(*tree2.get_children())

    def callbackFunc(event):
        print("Selected Category")
        CategoryName = (ComboCatglL2.get())
        print (CategoryName)

    def callbackFunc1(event):
        print("Selected Model")
        ModelName = (ComboModelL2.get())
        print (ModelName)
        
    def ViewCompleteInv():
        tree.delete(*tree.get_children())
        tree1.delete(*tree1.get_children())
        tree2.delete(*tree2.get_children())
        conn = sqlite3.connect("Eagle_Inventory.db")    
        InventoryCount_DF1 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_1 ORDER BY `Category` ASC ;", conn)
        InventoryCount_DF2 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_2 ORDER BY `Model_Name` ASC ;", conn)
        InventoryCount_DF3 = pd.read_sql_query("SELECT * FROM Eagle_Inventory_Report_3 ORDER BY `Category` ASC ;", conn)
            
        Inv_data1 = pd.DataFrame(InventoryCount_DF1)
        Inv_data2 = pd.DataFrame(InventoryCount_DF2)
        Inv_data3 = pd.DataFrame(InventoryCount_DF3)
        
        for each_rec in range(len(Inv_data1)):
            tree1.insert("", tk.END, values=list(Inv_data1.loc[each_rec]))
        for each_rec in range(len(Inv_data2)):
            tree2.insert("", tk.END, values=list(Inv_data2.loc[each_rec]))
        for each_rec in range(len(Inv_data3)):
            tree.insert("", tk.END, values=list(Inv_data3.loc[each_rec]))

        conn.commit()
        conn.close()

    def InvModel_search():
        if ComboModelL2.get() != "":
            tree.delete(*tree.get_children())
            tree1.delete(*tree1.get_children())
            tree2.delete(*tree2.get_children())
            conn= sqlite3.connect("Eagle_Inventory.db")
            cursor1 = conn.cursor()
            cursor2 = conn.cursor()
            cursor1.execute("SELECT * FROM `Eagle_Inventory_Report_1` WHERE `Model_Name` LIKE ? ", ('%'+ str(ComboModelL2.get()) +'%', ))        
            cursor2.execute("SELECT * FROM `Eagle_Inventory_Report_2` WHERE `Model_Name` LIKE ? ", ('%'+ str(ComboModelL2.get()) +'%', ))
            fetch1 = cursor1.fetchall()
            fetch2 = cursor2.fetchall()
            for data in fetch1:
                tree1.insert('', 'end', values=(data))
            for data in fetch2:
                tree2.insert('', 'end', values=(data))
            cursor1.close()
            cursor2.close()
            conn.close()
        else:
                tkinter.messagebox.showinfo("Search Error","Please Select Model Name and Search")

    def InvCatg_search():
        if ComboCatglL2.get() != "":
            tree.delete(*tree.get_children())
            tree1.delete(*tree1.get_children())
            tree2.delete(*tree2.get_children())
            conn= sqlite3.connect("Eagle_Inventory.db")
            cursor1 = conn.cursor()
            cursor2 = conn.cursor()
            cursor3 = conn.cursor()        
            cursor1.execute("SELECT * FROM `Eagle_Inventory_Report_1` WHERE `Category` LIKE ? ", ('%'+ str(ComboCatglL2.get()) +'%', ))
            cursor2.execute("SELECT * FROM `Eagle_Inventory_Report_2` WHERE `Category` LIKE ? ", ('%'+ str(ComboCatglL2.get()) +'%', ))
            cursor3.execute("SELECT * FROM `Eagle_Inventory_Report_3` WHERE `Category` LIKE ? ", ('%'+ str(ComboCatglL2.get()) +'%', ))        
            fetch1 = cursor1.fetchall()
            fetch2 = cursor2.fetchall()
            fetch3 = cursor3.fetchall()
            for data3 in fetch3:
                tree.insert('', 'end', values=(data3))        
            for data1 in fetch1:
                tree1.insert('', 'end', values=(data1))
            for data2 in fetch2:
                tree2.insert('', 'end', values=(data2))
            cursor1.close()
            cursor2.close()
            cursor3.close()
            conn.close()
        else:
            tkinter.messagebox.showinfo("Search Error","Please Select Model Name and Search")
                        
    def ExportSelectedInvReportA():
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,\
               defaultextension='.csv', filetypes = (("csv files",".csv"),("text files",".txt")))
        if filename:
            if filename.endswith('.csv'):
                with open(filename, 'w') as file:
                    file.write('Category' +  ',' + 'Total Count By Model' + '\n')
                    for item in tree.selection():
                        list_item = (tree.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        file.write( x1 + ',' + x2 + '\n')
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as CSV")

            else:
                with open(filename, 'w') as file:
                    file.write('Category' +  ',' + 'Total Count By Model' + '\n')
                    for item in tree.selection():
                        list_item = (tree.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        file.write( x1 + ',' + x2 + '\n')                                
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as TEXT")

    def ExportSelectedInvReportB():
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,\
               defaultextension='.csv', filetypes = (("csv files",".csv"),("text files",".txt")))
        if filename:
            if filename.endswith('.csv'):
                with open(filename, 'w') as file:
                    file.write('Category' + ',' + 'Model Name' + ',' + 'Total Count By Model' + '\n')
                    for item in tree1.selection():
                        list_item = (tree1.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        x3= list_item[2]
                        file.write( x1 + ',' + x2 + ',' + x3 + '\n')
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as CSV")

            else:
                with open(filename, 'w') as file:
                    file.write('Category' + ',' + 'Model Name' + ',' + 'Total Count By Model' + '\n')
                    for item in tree1.selection():
                        list_item = (tree1.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        x3= list_item[2]                        
                        file.write( x1 + ',' + x2 + ',' + x3 + '\n')                                
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as TEXT")

    def ExportSelectedInvReportC():
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file" ,\
               defaultextension='.csv', filetypes = (("csv files",".csv"),("text files",".txt")))
        if filename:
            if filename.endswith('.csv'):
                with open(filename, 'w') as file:
                    file.write('Category' + ',' + 'Model Name' + ',' + 'Total Count By Model' + '\n')
                    for item in tree2.selection():
                        list_item = (tree2.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        x3= list_item[2]
                        x4= list_item[3]
                        file.write( x1 + ',' + x2 + ',' + x3 + ',' + x4 + '\n')
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as CSV")

            else:
                with open(filename, 'w') as file:
                    file.write('Category' + ',' + 'Model Name' + ',' + 'Total Count By Model' + '\n')
                    for item in tree2.selection():
                        list_item = (tree2.item(item, 'values'))
                        x1= list_item[0]
                        x2= list_item[1]
                        x3= list_item[2]
                        x4= list_item[3]
                        file.write( x1 + ',' + x2 + ',' + x3 + ',' + x4 + '\n')                                
                file.close
                tkinter.messagebox.showinfo("Inventory file Export","Inventory File Saved as TEXT")



    ##### Labeling #######
    InvL1 = Label(window, text = "1: Inventory Report A", font=("arial", 10,'bold')).place(x=10,y=10)
    InvL2 = Label(window, text = "2: Inventory Report B", font=("arial", 10,'bold')).place(x=340,y=10)
    InvL3 = Label(window, text = "3: Inventory Report C", font=("arial", 10,'bold')).place(x=770,y=10)
    ModelL2 = Label(window, text = "4: Search By Model", font=("arial", 10,'bold'),bg = "green").place(x=600,y=520)
    ModelL2 = Label(window, text = "5: Search By Category", font=("arial", 10,'bold'),bg = "green").place(x=1100,y=520)

    menu = Menu(window)
    window.config(menu=menu)
    filemenu = Menu(menu, tearoff=0)
    menu.add_cascade(label="File", menu=filemenu)
    filemenu.add_command(label="Export Complete Report", command = ExportCompleteINV)
    filemenu.add_separator()
    filemenu.add_command(label="Exit", command=InvExit)

    ComboModelL2 = ttk.Combobox(window, font=('aerial', 10, 'bold'), textvariable = SEARCHMODEL, width = 28)
    ComboModelL2.place(x=600,y=545)
    ComboModelL2['values'] = list(set(Combo_Inv_Model()))
    ComboModelL2.bind('<<ComboboxSelected>>',callbackFunc1)
        
    ComboCatglL2 = ttk.Combobox(window, font=('aerial', 10, 'bold'), textvariable = SEARCHCATG, width = 32)
    ComboCatglL2.place(x=1000,y=545)
    ComboCatglL2['values'] = list(set(Combo_Inv_Category()))
    ComboCatglL2.bind('<<ComboboxSelected>>',callbackFunc)

    btnReset_Inv = Button(window, text="Reset All", font=('aerial', 11, 'bold'), height =1, width=9, bd=1, command = ResetSearch)
    btnReset_Inv.place(x=1155,y=6)

    btnExport_InvA = Button(window, text="Export Selected Report A", font=('aerial', 9, 'bold'), height =1,width=22, bd=1, command=ExportSelectedInvReportA)
    btnExport_InvA.place(x=10,y=490)

    btnExport_InvB = Button(window, text="Export Selected Report B", font=('aerial', 9, 'bold'), height =1,width=22, bd=1, command =ExportSelectedInvReportB)
    btnExport_InvB.place(x=348,y=490)

    btnExport_InvC = Button(window, text="Export Selected Report C", font=('aerial', 9, 'bold'), height =1,width=22, bd=1, command =ExportSelectedInvReportC)
    btnExport_InvC.place(x=770,y=490)


    btnComplete_Inv = Button(window, text="View Generated Report", font=('aerial', 11, 'bold'), height =2, width=19, bd=4, command = ViewCompleteInv)
    btnComplete_Inv.place(x=2,y=540)

    btnExport_Complete_Report = Button(window, text="Export All Report", font=('aerial', 11, 'bold'), height =2, width=19, bd=4, command = ExportCompleteINV)
    btnExport_Complete_Report.place(x=200,y=540)



    btnPopulate_Catg = Button(window, text="Populate Search", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = InvCatg_search)
    btnPopulate_Catg.place(x=1140,y=570)

    btnPopulate_Model = Button(window, text="Populate Search", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = InvModel_search)
    btnPopulate_Model.place(x=600,y=570)



