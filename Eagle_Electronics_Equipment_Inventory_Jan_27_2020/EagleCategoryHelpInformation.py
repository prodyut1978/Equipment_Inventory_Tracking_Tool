from tkinter import*
import tkinter.ttk as ttk
import tkinter as tk
import pandas as pd
import tkinter.messagebox

def EquipmentCatgHelp():
    EquipmentName_RecEquipList           = ["GSR 1", "GSX 1", "GSX 3","GSX 4","HAWK FSU","Lipo","BX10","BN18","BN25"]
    EquipmentName_BackupMediaList        = ["Startech Drive", "Startech Case of 3", "HDD Backup Case"]
    EquipmentName_AdminList              = ["Invoice/PO", "HSE Equipment", "Office Supplies","Laptop","Aircard"]
    EquipmentName_FleetList              = ["F150", "F250", "F350","F450","F550",
                                            "UTV","ATV","Snowmobile","Trailer","Van",
                                            "ATS60","UniVibe","Y2400","HV4"]
    EquipmentName_AccessoriesList        = ["Line Viewer (batts, charger,dongle)", "LHR", "R1", "Orientation Tool",
                                            "Garmin","Handheld Radio","Handheld Chargers","Tiger Nav",
                                            "Tiger Nav Lite","Heli Picker","Heli Bags","Hand Auger",
                                            "Ice Auger","Planting Poles","Seisnet Key","Testif-I Key"]
    EquipmentName_SourceElectronicsList  = ["Boombox", "Shotpro", "GSI/GSIx", "SDR/SDRx","VibPro", "Force2","Force3",
                                            "UE","UE2","UE3","Blaster Batteries"]
    EquipmentName_RecEquipList           = pd.DataFrame({'RecEquip': EquipmentName_RecEquipList})
    EquipmentName_BackupMediaList        = pd.DataFrame({'BackupMedia': EquipmentName_BackupMediaList})
    EquipmentName_AdminList              = pd.DataFrame({'Admin': EquipmentName_AdminList})
    EquipmentName_FleetList              = pd.DataFrame({'Fleet': EquipmentName_FleetList})
    EquipmentName_AccessoriesList        = pd.DataFrame({'Accessories': EquipmentName_AccessoriesList})
    EquipmentName_SourceElectronicsList  = pd.DataFrame({'SourceElectronics': EquipmentName_SourceElectronicsList})

    window = Tk()
    window.title("Eagle Electronics Inventory Category and Item Viewer")
    window.config(bg="ghost white")
    width = 520
    height = 800
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry("%dx%d+%d+%d" % (width, height, x, y))
    window.grid_rowconfigure(1, weight=1)
    window.grid_columnconfigure(0, weight=1)
    window.resizable(0, 0)
    TableMargin1 = Frame(window)
    TableMargin2 = Frame(window)
    TableMargin3 = Frame(window)
    TableMargin4 = Frame(window)
    TableMargin5 = Frame(window)
    TableMargin6 = Frame(window)
    TableMargin1.place(x=2, y =2)
    TableMargin2.place(x=2, y =228)
    TableMargin3.place(x=2, y =330)
    TableMargin4.place(x=2, y =470)
    TableMargin5.place(x=260, y =2)
    TableMargin6.place(x=260, y =368)

    tree1 = ttk.Treeview(TableMargin1, column=("column1"),selectmode="extended",
                                height=10, show='headings')
    tree2 = ttk.Treeview(TableMargin2, column=("column1"),selectmode="extended",
                                height=4, show='headings')
    tree3 = ttk.Treeview(TableMargin3, column=("column1"),selectmode="extended",
                                height=6, show='headings')
    tree4 = ttk.Treeview(TableMargin4, column=("column1"),selectmode="extended",
                                height=14, show='headings')


    tree5 = ttk.Treeview(TableMargin5, column=("column1"),selectmode="extended",
                                height=17, show='headings')
    tree6 = ttk.Treeview(TableMargin6, column=("column1"),selectmode="extended",
                                height=12, show='headings')
    tree1.heading("#1", text="Category : Recording Equipment", anchor=W)
    tree2.heading("#1", text="Category : Backup Media", anchor=W)
    tree3.heading("#1", text="Category : Admin", anchor=W)
    tree4.heading("#1", text="Category : Fleet", anchor=W)
    tree5.heading("#1", text="Category : Accessories", anchor=W)
    tree6.heading("#1", text="Category : Source Electronics", anchor=W)
    tree1.column('#1', stretch=NO, minwidth=0, width=200)
    tree2.column('#1', stretch=NO, minwidth=0, width=200)
    tree3.column('#1', stretch=NO, minwidth=0, width=200)
    tree4.column('#1', stretch=NO, minwidth=0, width=200)
    tree5.column('#1', stretch=NO, minwidth=0, width=200)
    tree6.column('#1', stretch=NO, minwidth=0, width=200)

    for each_rec in range(len(EquipmentName_RecEquipList)):
        tree1.insert("", tk.END, values=list(EquipmentName_RecEquipList.loc[each_rec]))
    for each_rec in range(len(EquipmentName_BackupMediaList)):
        tree2.insert("", tk.END, values=list(EquipmentName_BackupMediaList.loc[each_rec]))
    for each_rec in range(len(EquipmentName_AdminList)):
        tree3.insert("", tk.END, values=list(EquipmentName_AdminList.loc[each_rec]))
    for each_rec in range(len(EquipmentName_FleetList)):
        tree4.insert("", tk.END, values=list(EquipmentName_FleetList.loc[each_rec]))
    for each_rec in range(len(EquipmentName_AccessoriesList)):
        tree5.insert("", tk.END, values=list(EquipmentName_AccessoriesList.loc[each_rec]))
    for each_rec in range(len(EquipmentName_SourceElectronicsList)):
        tree6.insert("", tk.END, values=list(EquipmentName_SourceElectronicsList.loc[each_rec]))

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 9,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree1.pack()
    tree2.pack()
    tree3.pack()
    tree4.pack()
    tree5.pack()
    tree6.pack()

