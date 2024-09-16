#Inventory Scanned Import
import os
from tkinter import*
import tkinter.messagebox
import EagleDatabase_BackEnd
import EagleCategoryHelpInformation
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilename
import pandas as pd
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import csv
import time
import datetime
def GenerateTransmittalOutFirstOption():
    conn = sqlite3.connect("Eagle_Inventory.db")
    TransmittalOut_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory_transmittalOut1;", conn)
    data = pd.DataFrame(TransmittalOut_df)
    data = data.sort_values(by =['main_SN'])
    data = data.reset_index(drop=True)
    data = pd.DataFrame(data)
    LengthDF = len(data)

    ## Making Transmittal Summary
    Inv_CountReport2   = data.groupby(['catg'], as_index=False).main_SN.count()
    Inv_CountReport2   = pd.DataFrame(Inv_CountReport2)
    Inv_CountReport2.rename(columns={'catg':'Category', 'main_SN':'Quantity'},inplace = True)
    Inv_CountReport2["Item"]     = Inv_CountReport2.shape[0]*[""]
    Inv_CountReport2["Owner"]     = Inv_CountReport2.shape[0]*[""]
    Inv_CountReport2["ManfSN"]     = Inv_CountReport2.shape[0]*["See Sheet Detail List"]
    Inv_CountReport2["UnitWeight"]  = Inv_CountReport2.shape[0]*[""]
    Inv_CountReport2["TotalWeight"]= Inv_CountReport2.shape[0]*[""]
    Inv_CountReport2["Comments"]          = Inv_CountReport2.shape[0]*[""]
    Inv_CountReport2    = Inv_CountReport2.loc[:,['Category','Item', 'ManfSN', 'Owner', 'Quantity', 'UnitWeight', 'TotalWeight', 'Comments']]
    Inv_CountReport2    = pd.DataFrame(Inv_CountReport2)
    Inv_CountReport2    = Inv_CountReport2.reset_index(drop=True)
    Inv_CountReport2.to_sql('Eagle_Inventory_transmittalOutFrontPage',conn, if_exists="replace", index=False)
    conn.commit()
    conn.close()

    if LengthDF == 0:
        tkinter.messagebox.showinfo("Transmittal Out Message",
                "Transmittal Out database is empty")
    else:        
    ## Tree View
        window = Tk()
        window.title("Generated Transmittal Output View")
        window.config(bg="ghost white")
        width = 1260
        height = 880
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        window.geometry("%dx%d+%d+%d" % (width, height, x, y))
        window.resizable(0, 0)                    
        TableMargin = Frame(window, bd = 1, pady= 8)
        TableMargin1 = Frame(window, bd = 1, pady= 8)

        TableMargin_label = Label(window, text = "Equipment List Details:", font=("arial", 12,'bold'),bg = "ghost white").place(x=10,y=90)    
        TableMargin.place(x=10, y =118)

        TableMargin1_label = Label(window, text = "Generated Transmittal Out Summary:", font=("arial", 12,'bold'),bg = "ghost white").place(x=10,y=560)

        
        TableMargin1.place(x=10, y =588)
        
        scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
        tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4",
                                                 "column5", "column6", "column7", "column8",
                                                 "column9", "column10"), height=18, show='headings')
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
        tree.column('#1', stretch=NO, minwidth=0, width=87)            
        tree.column('#2', stretch=NO, minwidth=0, width=90)
        tree.column('#3', stretch=NO, minwidth=0, width=90)
        tree.column('#4', stretch=NO, minwidth=0, width=80)
        tree.column('#5', stretch=NO, minwidth=0, width=70)
        tree.column('#6', stretch=NO, minwidth=0, width=80)
        tree.column('#7', stretch=NO, minwidth=0, width=80)
        tree.column('#8', stretch=NO, minwidth=0, width=80)
        tree.column('#9', stretch=NO, minwidth=0, width=60)
        tree.column('#10', stretch=NO, minwidth=0, width=60) 
        tree.pack()

        scrollbarx = Scrollbar(TableMargin1, orient=HORIZONTAL)
        scrollbary = Scrollbar(TableMargin1, orient=VERTICAL)
        tree1 = ttk.Treeview(TableMargin1, column=("column1", "column2", "column3", "column4",
                                                 "column5", "column6", "column7", "column8"), height=8, show='headings')
        scrollbary.config(command=tree1.yview)
        scrollbary.pack(side=RIGHT, fill=Y)
        scrollbarx.config(command=tree1.xview)
        scrollbarx.pack(side=BOTTOM, fill=X)
        tree1.heading("#1", text="Category", anchor=W)
        tree1.heading("#2", text="Item", anchor=W)
        tree1.heading("#3", text="Serial#/Unit#", anchor=W)
        tree1.heading("#4", text="Owner", anchor=W)
        tree1.heading("#5", text="Quantity", anchor=W)            
        tree1.heading("#6", text="UnitWeight(If Reqd)", anchor=W)
        tree1.heading("#7", text="Total Weight (lbs)" ,anchor=W)
        tree1.heading("#8", text="Comments", anchor=W)              
        tree1.column('#1', stretch=NO, minwidth=0, width=100)            
        tree1.column('#2', stretch=NO, minwidth=0, width=116)
        tree1.column('#3', stretch=NO, minwidth=0, width=90)
        tree1.column('#4', stretch=NO, minwidth=0, width=100)
        tree1.column('#5', stretch=NO, minwidth=0, width=70)
        tree1.column('#6', stretch=NO, minwidth=0, width=114)
        tree1.column('#7', stretch=NO, minwidth=0, width=110)
        tree1.column('#8', stretch=NO, minwidth=0, width=80)    
        tree1.pack()

        def TransmittalRec(event):
            for nm in tree.selection():
                sd = tree.item(nm, 'values')

        def TransmittalRec1(event):
            for nm in tree1.selection():
                sd = tree1.item(nm, 'values')
                EquipmentType.delete(0,END)
                EquipmentType.insert(tk.END,sd[0])                
                EquipmentName.delete(0,END)
                EquipmentName.insert(tk.END,sd[1])                
                ItemSN.delete(0,END)                
                ItemSN.insert(tk.END,sd[2])                
                OwnerName.delete(0,END)                
                OwnerName.insert(tk.END,sd[3])
                ItemQuantity.delete(0,END)
                ItemQuantity.insert(tk.END,sd[4])                
                UnitWeight.delete(0,END)
                UnitWeight.insert(tk.END,sd[5])
                TotalWeight.delete(0,END)
                TotalWeight.insert(tk.END,sd[6])
                Comments.delete(0,END)
                Comments.insert(tk.END,sd[7])


    ##----------------- Tree View Select Event------------
            
        tree.bind('<<TreeviewSelect>>',TransmittalRec)
        tree1.bind('<<TreeviewSelect>>',TransmittalRec1)
                
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        for each_rec in range(len(Inv_CountReport2)):
            tree1.insert("", tk.END, values=list(Inv_CountReport2.loc[each_rec]))
       

    ## Label and Entry
        TitFrame = Frame(window, bd = 2, padx= 2, pady= 2, bg = "#006dcc", relief = RIDGE)
        TitFrame.place(x=890, y =2)    
        lblTit = Label(TitFrame, font=('aerial', 10, 'bold'),
                            text="Eagle Canada Equipment Transmittal Out \n 6806 Railway Street SE \n Calgary, AB T2H 3A8\n Ph: (403) 263-7770",
                            bg="#006dcc")
        lblTit.grid()

        L1 = Label(window, text = "A: Count in Transmittal :", font=("arial", 10,'bold'),bg = "ghost white").place(x=6,y=8)
        TotalTransmittalEntry  = Entry(window, font=('aerial', 12, 'bold'),textvariable = IntVar(), width = 11, bd=2)
        TotalTransmittalEntry.place(x=180,y=8)
        TotalTransmittalEntry.delete(0,END)
        TotalTransmittalEntry.insert(tk.END,LengthDF)

        Default_Date_today   = datetime.date.today()
        L2 = Label(window, text = "B: Transmittal Out Date: ** ", font=("arial", 10,'bold'),bg = "ghost white").place(x=6,y=48)
        TransmittalDate  = Entry(window, font=('aerial', 12, 'bold'),textvariable = StringVar(window,value=Default_Date_today), width = 11, bd=2)
        TransmittalDate.place(x=180,y=48)

        L3 = Label(window, text = "C: Transmittal Number :", font=("arial", 10,'bold'),bg = "ghost white").place(x=390,y=8)
        TransmittalBatch  = Entry(window, font=('aerial', 12, 'bold'),textvariable = StringVar(), width = 15, bd=2)
        TransmittalBatch.place(x=559,y=8)

        Sending_Reason = ["For Crew - Production", "Transfer To Calgary Shop", "Crew To Crew Transfer", "For Repair - Bad Equipment",]
        L4 = Label(window, text = "D: Transmittal Reason :", font=("arial", 10,'bold'),bg = "ghost white").place(x=390,y=48)
        ReasonSending = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 31, values= Sending_Reason)
        ReasonSending.current(0)
        ReasonSending.place(x=559,y=48)

        L5 = Label(window, text = "H: Eagle TDG Form Completed For Batteries? :", font=("arial", 10,'bold'),bg = "ghost white").place(x=250,y=830)
        Answer_List = ["No, Not Required", "Yes"]                                        
        TDG  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 20, values= Answer_List)
        TDG.current(0)
        TDG.place(x=560,y=830)

        L6 = Label(window, text = "E: Job/Program Information Entries:", font=("arial", 12,'bold'),bg = "ghost white").place(x=840,y=100)

        L7 = Label(window, text = "1: Program Name :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=130)    
        ProjectName = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        ProjectName.place(x=1010,y=130)

        L8 = Label(window, text = "2: Program Number :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=170)
        ProgramNumber  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        ProgramNumber.place(x=1010,y=170)

        Crew_List = ["Suncor Energy", "Cenovus Energy", "CNRL", "Explor", "IGC", "LXL Consulting", "North American Helium", "RPS", "Synterra","TGS",
                     "Eagle Office Calgary", "Eagle Shop Calgary", "Crew 101", "Crew 102", "Crew 103", "Crew 104", "Crew 105", "Crew 106", "Crew 107", "Crew 108"]
        L9 = Label(window, text = "3: Crew :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=210)    
        Crew_Name = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= Crew_List)
        Crew_Name.place(x=1010,y=210)

        CrewManager_List = [ "Polachek, Kris", "Tofsrud, Don", "Dewit, Mark", "Pilkey, John", "Harris, Tim", "Jackson, Terry",
                            "Renaud, Corey R ", "Sheppard, Greg ", "Taylor, Lenard R ", "Bowman, Doug ",
                            "Graychick, Scott P ", "McFarlane, Michael M ", "Trotter Doug ", "Croken, Terry "]

        L10 = Label(window, text = "4: Crew Manager Name :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=250)
        CrewManager  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= CrewManager_List)
        CrewManager.place(x=1010,y=250)

        L11 = Label(window, text = "F: Transmittal Sender/Receiver Information Entries:", font=("arial", 12,'bold'),bg = "ghost white").place(x=840,y=300)

        From_List = ["Eagle Office Calgary", "Eagle Shop Calgary", "Crew 101", "Crew 102", "Crew 103", "Crew 104", "Crew 105", "Crew 106", "Crew 107", "Crew 108"]
        L12 = Label(window, text = "1: From Location :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=330)    
        FromAddress  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= From_List)
        FromAddress.place(x=1010,y=330)

        To_List = ["Eagle Office Calgary", "Eagle Shop Calgary", "Crew 101", "Crew 102", "Crew 103", "Crew 104", "Crew 105", "Crew 106", "Crew 107", "Crew 108",
                   "Geospace Technology Calgary", "Geo-Check Calgary", "Mitcham Calgary", "Gobal Calgary", "Dawson USA", "Inova Geophysical Calgary"]

        L12 = Label(window, text = "2: To Location : ** ", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=370)    
        ToAddress  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= To_List)
        ToAddress.place(x=1010,y=370)

        L13 = Label(window, text = "3: Other To/From :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=410)    
        OtherToFrom  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        OtherToFrom.place(x=1010,y=410)

        L14 = Label(window, text = "4: Transported By :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=450)
        VehicalNumber  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        VehicalNumber.place(x=1010,y=450)

        L15 = Label(window, text = "5: Receiver Name :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=490)
        ReceiverName  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        ReceiverName.place(x=1010,y=490)

        L16 = Label(window, text = "6: Shipper Name :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=530)
        ShipperName  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        ShipperName.place(x=1010,y=530)

        L17 = Label(window, text = "7: PO Number :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=570)
        PONumber  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        PONumber.place(x=1010,y=570)

        L18 = Label(window, text = "G: Generated Transmittal Equipment Information :", font=("arial", 12,'bold'),bg = "ghost white").place(x=840,y=620)
                          
        L19 = Label(window, text = "1: Equipment Category :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=650)
        EquipmentType  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)
        EquipmentType.place(x=1010,y=650)

        EquipmentName_RecEquipList           = ["GSR 1", "GSX 1", "GSX 3","GSX 4","HAWK FSU","Lipo","BX10","BN18","BN25"]
        EquipmentName_BackupMediaList        = ["Startech Drive", "Startech Case of 3", "HDD Backup Case"]
        EquipmentName_AccessoriesList        = ["Line Viewer (batts, charger,dongle)", "LHR", "R1", "Orientation Tool",
                                                "Garmin","Handheld Radio","Handheld Chargers","Tiger Nav",
                                                "Tiger Nav Lite","Heli Picker","Heli Bags","Hand Auger",
                                                "Ice Auger","Planting Poles","Seisnet Key","Testif-I Key"]
        EquipmentName_SourceElectronicsList  = ["Boombox", "Shotpro", "GSI/GSIx", "SDR/SDRx","VibPro", "Force2","Force3",
                                                "UE","UE2","UE3","Blaster Batteries"]
        EquipmentName_AdminList              = ["Invoice/PO", "HSE Equipment", "Office Supplies","Laptop","Aircard"]
        EquipmentName_FleetList              = ["F150", "F250", "F350","F450","F550",
                                                "UTV","ATV","Snowmobile","Trailer","Van",
                                                "ATS60","UniVibe","Y2400","HV4"]
        EquipmentName_List =  [ "GSR 1", "GSX 1", "GSX 3","GSX 4","HAWK FSU","Lipo","BX10","BN18","BN25",
                                "Startech Drive", "Startech Case of 3", "HDD Backup Case",
                                "Line Viewer (batts, charger,dongle)", "LHR", "R1", "Orientation Tool",
                                "Garmin","Handheld Radio","Handheld Chargers","Tiger Nav",
                                "Tiger Nav Lite","Heli Picker","Heli Bags","Hand Auger",
                                "Ice Auger","Planting Poles","Seisnet Key","Testif-I Key",
                                "Boombox", "Shotpro", "GSI/GSIx", "SDR/SDRx","VibPro", "Force2","Force3",
                                "UE","UE2","UE3","Blaster Batteries","Invoice/PO", "HSE Equipment",
                                "Office Supplies","Laptop","Aircard","F150", "F250", "F350","F450","F550",
                                "UTV","ATV","Snowmobile","Trailer","Van",
                                "ATS60","UniVibe","Y2400","HV4"]
        def itemhelp():
            EagleCategoryHelpInformation.EquipmentCatgHelp()
        
        L21 = Label(window, text = "2: Item Name :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=690)
        btnItemHelp = Button(window, text="?", font=('aerial', 8, 'bold'), height =1, width=2, bd=1, command = itemhelp)
        btnItemHelp.place(x=940,y=693)
        EquipmentName  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= EquipmentName_List)
        EquipmentName.place(x=1010,y=690)

        Owner_List = ["Eagle Office Calgary", "Geospace Technology Calgary", "Geo-Check Calgary",
                      "Mitcham Calgary", "Gobal Calgary", "Dawson USA", "Inova Geophysical Calgary"]
        
        L22 = Label(window, text = "3: Owner :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=730)
        OwnerName  = ttk.Combobox(window, font=('aerial', 10, 'bold'),  width = 30, values= Owner_List)
        OwnerName.place(x=1010,y=730)

        L23 = Label(window, text = "Item Quantity :", font=("arial", 10,'bold'),bg = "ghost white").place(x=350,y=560)
        ItemQuantity  = Entry(window, font=('aerial', 10, 'bold'), textvariable = IntVar(), width = 10, bd=2)
        ItemQuantity.place(x=450,y=560)

        L24 = Label(window, text = "Serial #/Unit # :", font=("arial", 10,'bold'),bg = "ghost white").place(x=540,y=560)
        ItemSN  = Entry(window, font=('aerial', 10, 'bold'), textvariable = StringVar(), width = 22, bd=2)
        ItemSN.place(x=645,y=560)

        L25 = Label(window, text = "4: UnitWeight (If Reqd) :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=760)
        UnitWeight  = Entry(window, font=('aerial', 10, 'bold'),textvariable = IntVar(), width = 8, bd=2)
        UnitWeight.place(x=1010,y=760)
        L26 = Label(window, text = "5: Total (lbs):", font=("arial", 10,'bold'),bg = "ghost white").place(x=1095,y=760)
        TotalWeight  = Entry(window, font=('aerial', 10, 'bold'),textvariable = IntVar(), width = 8, bd=2)
        TotalWeight.place(x=1185,y=760)

        L27 = Label(window, text = "6: Comments :", font=("arial", 10,'bold'),bg = "ghost white").place(x=840,y=790)
        Comments  = Entry(window, font=('aerial', 10, 'bold'),textvariable = StringVar(), width = 33, bd=2)                    
        Comments.place(x=1010,y=790)

        
    ## Define functions
                        
        def ExportTransmittal():
            

            Total_Count_Export = TotalTransmittalEntry.get()
            Total_Count = ("Transmittal Quantity: ")
            
            TransmittalDate_Export = TransmittalDate.get()
            Transmittal_Date = ("Transmittal Date: ")
            
            TransmittalNumber_Export = TransmittalBatch.get()
            TransmittalNumber = ("Transmittal No : ")

            ProjectName_Export = ProjectName.get()
            Project_Name = ("Job/Program Name : ")

            ProjectLocation_Export = FromAddress.get()
            Project_Location = ("From Location : ")

            Crew_Name_Export = Crew_Name.get()
            CrewName = ("Crew : ")

            ProgramNumber_Export = ProgramNumber.get()
            Program_Number = ("Job/Program Number : ")

            CrewManager_Export = CrewManager.get()
            Crew_Manager = ("Crew Manager Name : ")
            
            EquipmentType_Export = EquipmentType.get()
            Equipment_Type = ("Equipment Type : ")

            EquipmentName_Export = EquipmentName.get()
            Equipment_Name = ("Equipment Name : ")

            ReceiverName_Export = ReceiverName.get()
            Receiver_Name = ("Receiver Name : ")

            VehicalNumber_Export = VehicalNumber.get()
            Vehical_Number = ("Transported by : ")

            ShipperName_Export = ShipperName.get()
            Shipper_Name = ("Shipper Name : ")
            
            Comments_Export = Comments.get()
            CommentsText = ("Comments: ")

            ReceiverAddress_Export = ToAddress.get()
            Receiver_Address = ("Receiver Address : ")

            SendingReason_Export = ReasonSending.get()

            Owner_Export        = OwnerName.get()
            Weight_Export       = TotalWeight.get()
            TDG_Export          = TDG.get()
            PO_Number_Export    = PONumber.get()
            OtherToFrom_Export  = OtherToFrom.get()

            if(len(ReceiverAddress_Export)!=0) & (len(TransmittalDate_Export)!=0):
                UpdateTransmittalOutToMasterDB()
                dfList =[] 
                for child in tree.get_children():
                    df = tree.item(child)["values"]
                    dfList.append(df)
                Transmittal_DF = pd.DataFrame(dfList)
                Transmittal_DF.rename(columns = {0:'Category', 1:'Manuf', 2:'Model', 3:'ManfSN', 4:'Desc',
                                              5: 'Asset_SN', 6:'Date', 7:'Location', 8:'Condition',9:'Origin'},inplace = True)
                Transmittal_DF_SortByCaseSrNo = Transmittal_DF.sort_values(by =['ManfSN'])

                dfListFront =[] 
                for child in tree1.get_children():
                    df1 = tree1.item(child)["values"]
                    dfListFront.append(df1)
                Transmittal_Front_DF = pd.DataFrame()
                Transmittal_Front_DF_Data = pd.DataFrame(dfListFront)
                Transmittal_Front_DF_Data.rename(columns = {0:'Category', 1:'Item', 2:'Serial # / Unit #', 3:'Owner', 4:'Quantity',
                                              5: 'Weight (if Reqd)', 6:'Total Weight (lbs)', 7:'Comment'},inplace = True)
                TransmittalSummary = (['',''],
                                    [Total_Count, Total_Count_Export],
                                    [Transmittal_Date,   TransmittalDate_Export],
                                    [TransmittalNumber,  TransmittalNumber_Export],['',''],
                                    ['',''],
                                    [Project_Name,  ProjectName_Export],
                                    [Project_Location,  ProjectLocation_Export],
                                    [CrewName,  Crew_Name_Export],
                                    [Program_Number,  ProgramNumber_Export],
                                    [Crew_Manager,  CrewManager_Export])

                ReceiverSummary = (['',''],
                                   [Receiver_Name, ReceiverName_Export],
                                   [Shipper_Name,  ShipperName_Export],
                                   [Vehical_Number,  VehicalNumber_Export],['',''],
                                   ['',''])

                ReceiverAddress_Split = ReceiverAddress_Export.split(",")                            
                
                Transmittal_Front_A = (['Date:',TransmittalDate_Export],
                               ['Shipper:', ShipperName_Export],
                               ['Receiver:',  ReceiverName_Export],
                               ['PO Number:', PO_Number_Export ])

                Transmittal_Front_B = (['From Location:',ProjectLocation_Export],
                                   ['To Location:', ReceiverAddress_Split[0]],
                                   ['Other To/From:',  OtherToFrom_Export],
                                   ['Transported By:',  VehicalNumber_Export])

                Transmittal_Front_C = (['Crew:',Crew_Name_Export],
                                   ['Program Name:', ProjectName_Export],
                                   ['Program Number:',  ProgramNumber_Export],
                                   ['Crew Manager:',  CrewManager_Export])
               
                filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name To Export Transmittal" ,
                           defaultextension='.xlsx', filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
                if filename:
                    if filename.endswith('.xlsx'):
                        with pd.ExcelWriter(filename, engine='xlsxwriter') as file:
                            Transmittal_Front_DF.to_excel(file,sheet_name='Transmittal Front Page', index=False, startrow=7, header=False)                    
                            Transmittal_DF_SortByCaseSrNo.to_excel(file,sheet_name='Detail List Equipment',index=False, startrow=14, header=False)
                            workbook_ListBadGSR  = file.book                    
                            workbook_ListBadGSR.formats[0].set_align('center')                    
                            worksheet_Front      = file.sheets['Transmittal Front Page']
                            worksheet_ListBadGSR = file.sheets['Detail List Equipment']

                            worksheet_Front.set_margins(0.3, 0.4, 1.6, 1.1)                                    
                            worksheet_Front.set_landscape()
                            worksheet_Front.print_area('A1:O27')
                            worksheet_Front.print_across()
                            worksheet_Front.fit_to_pages(1, 1)                                    
                            worksheet_Front.set_paper(9)
                            worksheet_Front.set_start_page(1)
                            worksheet_Front.hide_gridlines(0)
                            worksheet_Front.set_page_view()
                            
                            worksheet_ListBadGSR.set_margins(0.3, 0.1, 1.6, 1.1)                                    
                            worksheet_ListBadGSR.set_portrait()
                            worksheet_ListBadGSR.print_area('A1:J44')
                            worksheet_ListBadGSR.print_across()
                            worksheet_ListBadGSR.fit_to_pages(1, 0)                                    
                            worksheet_ListBadGSR.set_paper(9)
                            worksheet_ListBadGSR.set_start_page(1)
                            worksheet_ListBadGSR.hide_gridlines(0)
                            worksheet_ListBadGSR.set_page_view()

                            headerFront = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' +  'Ph: (403) 263-7770' +  '&R&U&24&"cambria, bold"Transmittal'
                            header1 = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' +  'Ph: (403) 263-7770' +  '&R&U&18&"cambria, bold"Transmittal' +'\n' +'Equipment List'
                            worksheet_Front.set_header(headerFront,{'image_left':'eagle logo.jpg'})
                            worksheet_ListBadGSR.set_header(header1,{'image_left':'eagle logo.jpg'})
                            footerFront = ('&LRevised Date : &D')+ ('&CEFXX-XX-XX') + ('&RReceiver Name (Print): &Y-------------------------------------------------------------' + '\n' + '&RReceiver Signature: &Y-------------------------------------------------------------')
                            footer1 = ('&LDate : &D') + ('&RReceiver Name (Print): &Y-----------------------------------------' + '\n' + '&RReceiver Signature: &Y-----------------------------------------')
                            worksheet_ListBadGSR.set_footer(footer1)
                            worksheet_Front.set_footer(footerFront)
                            
                            cell_format_1 = workbook_ListBadGSR.add_format({'bold': True, 'text_wrap': True, 'align': 'top', 'valign': 'top', 'border': 0})
                            cell_format_1.set_underline(1)
                            cell_format_2 = workbook_ListBadGSR.add_format({'bold': False, 'text_wrap': True, 'align': 'top', 'valign': 'top', 'border': 0})
                            cell_format_3 = workbook_ListBadGSR.add_format({'bold': False, 'text_wrap': True, 'align': 'top', 'valign': 'top', 'border': 0})
                            cell_format_3.set_font_size(15)
                            cell_format_4 = workbook_ListBadGSR.add_format({'bold': True, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'border': 0})
                            cell_format_4.set_underline(1)
                            cell_format_5 = workbook_ListBadGSR.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'center', 'border': 0})
                            cell_format_5.set_underline(1)
                            cell_format_6 = workbook_ListBadGSR.add_format({'bold': False, 'text_wrap': True, 'align': 'center', 'valign': 'top', 'border': 0})
                            cell_format_6.set_font_size(12)
                            cell_format_7 = workbook_ListBadGSR.add_format({'bold': True, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'border': 0})
                            cell_format_7.set_font_size(13)


                            cell_format_8 = workbook_ListBadGSR.add_format({'bold': False, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'border': 0})
                            cell_format_8.set_font_size(11)

                            worksheet_Front.merge_range('A1:B1', "")
                            worksheet_Front.merge_range('A2:B2', "")
                            worksheet_Front.merge_range('A3:B3', "")
                            worksheet_Front.merge_range('A4:B4', "")
                            worksheet_Front.merge_range('C1:E1', "")
                            worksheet_Front.merge_range('C2:E2', "")
                            worksheet_Front.merge_range('C3:E3', "")
                            worksheet_Front.merge_range('C4:E4', "")
                            worksheet_Front.merge_range('F1:G1', "")
                            worksheet_Front.merge_range('F2:G2', "")
                            worksheet_Front.merge_range('F3:G3', "")
                            worksheet_Front.merge_range('F4:G4', "")
                            worksheet_Front.merge_range('H1:J1', "")
                            worksheet_Front.merge_range('H2:J2', "")
                            worksheet_Front.merge_range('H3:J3', "")
                            worksheet_Front.merge_range('H4:J4', "")
                            worksheet_Front.merge_range('K1:L1', "")
                            worksheet_Front.merge_range('K2:L2', "")
                            worksheet_Front.merge_range('K3:L3', "")
                            worksheet_Front.merge_range('K4:L4', "")
                            worksheet_Front.merge_range('M1:O1', "")
                            worksheet_Front.merge_range('M2:O2', "")
                            worksheet_Front.merge_range('M3:O3', "")
                            worksheet_Front.merge_range('M4:O4', "")
                            worksheet_Front.merge_range('A26:H27', 'Transmittal Reason : ' +  SendingReason_Export, cell_format_3)

                            worksheet_ListBadGSR.merge_range('A2:B2', "")
                            worksheet_ListBadGSR.merge_range('A3:B3', "")
                            worksheet_ListBadGSR.merge_range('A4:B4', "")
                            worksheet_ListBadGSR.merge_range('A5:J5', "")                                    
                            worksheet_ListBadGSR.merge_range('A7:B7', "")
                            worksheet_ListBadGSR.merge_range('A8:B8', "")
                            worksheet_ListBadGSR.merge_range('A9:B9', "")
                            worksheet_ListBadGSR.merge_range('A10:B10', "")
                            worksheet_ListBadGSR.merge_range('A11:B11', "")                                  
                            worksheet_ListBadGSR.merge_range('C2:E2', "")
                            worksheet_ListBadGSR.merge_range('C3:E3', "")
                            worksheet_ListBadGSR.merge_range('C4:E4', "")
                            worksheet_ListBadGSR.merge_range('C7:E7', "")
                            worksheet_ListBadGSR.merge_range('C8:E8', "")
                            worksheet_ListBadGSR.merge_range('C9:E9', "")
                            worksheet_ListBadGSR.merge_range('C10:E10', "")
                            worksheet_ListBadGSR.merge_range('C11:E11', "")
                            worksheet_ListBadGSR.merge_range('F2:G2', "")
                            worksheet_ListBadGSR.merge_range('F3:G3', "")
                            worksheet_ListBadGSR.merge_range('F4:G4', "")                                    
                            worksheet_ListBadGSR.merge_range('H2:J2', "")
                            worksheet_ListBadGSR.merge_range('H3:J3', "")
                            worksheet_ListBadGSR.merge_range('H4:J4', "")
                            worksheet_ListBadGSR.merge_range('F7:J7', "")
                            worksheet_ListBadGSR.merge_range('F8:J8', "")
                            worksheet_ListBadGSR.merge_range('F9:J9', "")
                            worksheet_ListBadGSR.merge_range('F10:J10', "")
                                                                
                            row_TransmittalSummary = 0
                            col_TransmittalSummary = 0
                            for item, values in (TransmittalSummary):
                                worksheet_ListBadGSR.write(row_TransmittalSummary, col_TransmittalSummary,     item)
                                worksheet_ListBadGSR.write(row_TransmittalSummary, col_TransmittalSummary+2, values)
                                row_TransmittalSummary += 1
                            
                            row_TransmittalReceiver = 0
                            col_TransmittalReceiver = 5
                            for item, values in (ReceiverSummary):
                                worksheet_ListBadGSR.write(row_TransmittalReceiver, col_TransmittalReceiver,     item)
                                worksheet_ListBadGSR.write(row_TransmittalReceiver, col_TransmittalReceiver+2, values)
                                row_TransmittalReceiver += 1

                            row_ReceiverAddress = 6
                            col_ReceiverAddress = 5
                            for item in ReceiverAddress_Split:
                                worksheet_ListBadGSR.write(row_ReceiverAddress, col_ReceiverAddress, item)
                                row_ReceiverAddress += 1

                            row_TransmittalFrontA = 0
                            col_TransmittalFrontA = 0
                            for item, values in (Transmittal_Front_A):
                                worksheet_Front.write(row_TransmittalFrontA, col_TransmittalFrontA,     item, cell_format_8)
                                worksheet_Front.write(row_TransmittalFrontA, col_TransmittalFrontA + 2, values, cell_format_8)
                                row_TransmittalFrontA += 1

                            row_TransmittalFrontB = 0
                            col_TransmittalFrontB = 5
                            for item, values in (Transmittal_Front_B):
                                worksheet_Front.write(row_TransmittalFrontB, col_TransmittalFrontB,     item, cell_format_8)
                                worksheet_Front.write(row_TransmittalFrontB, col_TransmittalFrontB + 2, values, cell_format_8)
                                row_TransmittalFrontB += 1

                            row_TransmittalFrontC = 0
                            col_TransmittalFrontC = 10
                            for item, values in (Transmittal_Front_C):
                                worksheet_Front.write(row_TransmittalFrontC, col_TransmittalFrontC,     item, cell_format_8)
                                worksheet_Front.write(row_TransmittalFrontC, col_TransmittalFrontC + 2, values, cell_format_8)
                                row_TransmittalFrontC += 1

                            cell_format_Centre = workbook_ListBadGSR.add_format()
                            cell_format_Left = workbook_ListBadGSR.add_format()
                            cell_format_Centre.set_align('center')
                            cell_format_Left.set_align('left')
                            worksheet_ListBadGSR.set_column('A:A',12, cell_format_Left)
                            worksheet_ListBadGSR.set_column('B:B', 8, cell_format_Left)
                            worksheet_ListBadGSR.set_column('C:C', 8, cell_format_Left)
                            worksheet_ListBadGSR.set_column('D:D', 8, cell_format_Left)
                            worksheet_ListBadGSR.set_column('E:E', 11, cell_format_Left)
                            worksheet_ListBadGSR.set_column('F:F', 10, cell_format_Left)
                            worksheet_ListBadGSR.set_column('G:G', 7, cell_format_Left)
                            worksheet_ListBadGSR.set_column('H:H', 8, cell_format_Left)
                            worksheet_ListBadGSR.set_column('I:I', 9, cell_format_Left)
                            header_format_ListBadGSR = workbook_ListBadGSR.add_format({
                                            'bold': True,
                                            'text_wrap': True,
                                            'valign': 'top',
                                            'fg_color': '#808080',
                                            'border': 2})
                            header_format_ListBadGSR.set_align('center')
                            worksheet_ListBadGSR.merge_range('A1:E1', "Transmittal Summary:", cell_format_1)
                            worksheet_ListBadGSR.merge_range('A6:E6', "Transmittal Out Information:", cell_format_1)
                            worksheet_ListBadGSR.merge_range('F1:J1', "Receiving Information:", cell_format_1)
                            worksheet_ListBadGSR.merge_range('F6:J6', "Receiving Location:", cell_format_1)                                    
                            worksheet_ListBadGSR.merge_range('F11:J11', "Reason For Transmittal:", cell_format_1)                                    
                            worksheet_Front.merge_range('A5:G6', " Eagle TDG Form Completed For Batteries ? :   (YES / NO, Not Required) ", cell_format_4)
                            worksheet_Front.merge_range('H5:O6', TDG_Export , cell_format_4)

                            worksheet_Front.merge_range('A7:B7', "Category", cell_format_5)                                        
                            worksheet_Front.merge_range('C7:D7', "Item", cell_format_5)                    
                            worksheet_Front.merge_range('E7:F7', "Serial #/Unit #", cell_format_5)                    
                            worksheet_Front.merge_range('G7:H7', "Owner", cell_format_5)                    
                            worksheet_Front.write('I7', "Quantity", cell_format_5)                    
                            worksheet_Front.merge_range('J7:K7', "Unit Weight", cell_format_5)                    
                            worksheet_Front.merge_range('L7:M7', "Total Weight (lbs)", cell_format_5)                    
                            worksheet_Front.merge_range('N7:O7', "Comments", cell_format_5)
                            worksheet_Front.merge_range('I26:K27', " Total Overall Weight (lbs) : ", cell_format_7)
                            worksheet_Front.merge_range('L26:O27', '=SUM(L8:L25)', cell_format_7)
                            
                            Numberof_Row = 18
                            StartRow = 7
                            for i in range(Numberof_Row):
                                worksheet_Front.merge_range(i+StartRow, 0, i+StartRow, 1, '')
                                worksheet_Front.merge_range(i+StartRow, 2, i+StartRow, 3, '')
                                worksheet_Front.merge_range(i+StartRow, 4, i+StartRow, 5, '')
                                worksheet_Front.merge_range(i+StartRow, 6, i+StartRow, 7, '')
                                worksheet_Front.merge_range(i+StartRow, 9, i+StartRow, 10, '')
                                worksheet_Front.merge_range(i+StartRow, 11, i+StartRow, 12, '')
                                worksheet_Front.merge_range(i+StartRow, 13, i+StartRow, 14, '')
                                                        
                            Transmittal_Front_DF_Data['Category'].to_excel(file,sheet_name='Transmittal Front Page', index=False, startrow=7, startcol=0, header=False, merge_cells=False)
                            Transmittal_Front_DF_Data['Item'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=2, header=False)
                            Transmittal_Front_DF_Data['Serial # / Unit #'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=4, header=False)
                            Transmittal_Front_DF_Data['Owner'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=6, header=False)
                            Transmittal_Front_DF_Data['Quantity'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=8, header=False)
                            Transmittal_Front_DF_Data['Weight (if Reqd)'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=9, header=False)
                            Transmittal_Front_DF_Data['Total Weight (lbs)'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=11, header=False)
                            Transmittal_Front_DF_Data['Comment'].to_excel(file,sheet_name='Transmittal Front Page',index=False, startrow=7, startcol=13, header=False)                                                
                            worksheet_ListBadGSR.merge_range('A12:B13', CommentsText, cell_format_1)
                            worksheet_ListBadGSR.merge_range('C12:E13', '' , cell_format_2)
                            worksheet_ListBadGSR.merge_range('F12:J13', SendingReason_Export, cell_format_3)

                            for col_num, value in enumerate(Transmittal_DF_SortByCaseSrNo.columns.values):
                                worksheet_ListBadGSR.write(13, col_num, value, header_format_ListBadGSR)
                        file.close
                        tkinter.messagebox.showinfo("Transmittal Export"," Transmittal Out Report Saved as Excel")
                tree.delete(*tree.get_children())
                tree1.delete(*tree1.get_children())                
                ClearDB()
                iExit()
            else:
                tkinter.messagebox.showinfo("Update Master DB Error Message","You Must Need to Provide Updated Receiving Location and Transmittal Out Date")

        def UpdateTransmittalOutToMasterDB():
            UpdateLocationDateToMasterDB()
            con= sqlite3.connect("Eagle_Inventory.db")
            cur=con.cursor()
            cur.execute("DELETE FROM Eagle_Inventory WHERE EXISTS (SELECT * FROM Eagle_Inventory_transmittalOut1 WHERE Eagle_Inventory.main_SN = Eagle_Inventory_transmittalOut1.main_SN)")
            cur.execute("INSERT INTO Eagle_Inventory (catg, manuf, model, main_SN, desc, asset_SN,\
                                        datestamp, location, condition, origin) SELECT catg, manuf, model, main_SN, desc, asset_SN, datestamp, location, condition, origin FROM Eagle_Inventory_transmittalOut1")        
            con.commit()
            cur.close()
            con.close()

        def UpdateLocationDateToMasterDB():
            ReceiverAddress_Export = ToAddress.get()
            TransmittalDate_Export = TransmittalDate.get()
            conn = sqlite3.connect("Eagle_Inventory.db")
            cur = conn.cursor()        
            cur.execute("UPDATE Eagle_Inventory_transmittalOut1 SET location =? , datestamp =? ", (ReceiverAddress_Export, TransmittalDate_Export,))
            TransmittalOut_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory_transmittalOut1;", conn)
            data = pd.DataFrame(TransmittalOut_df)
            data = data.sort_values(by =['main_SN'])
            data = data.reset_index(drop=True)
            data = pd.DataFrame(data)
            tree.delete(*tree.get_children())
            for each_rec in range(len(data)):
                tree.insert("", tk.END, values=list(data.loc[each_rec]))
            conn.commit()
            conn.close()
            
        def ReconnectDBAfterDelete():
            conn = sqlite3.connect("Eagle_Inventory.db")
            TransmittalOut_df = pd.read_sql_query("SELECT * FROM Eagle_Inventory_transmittalOut1;", conn)
            data = pd.DataFrame(TransmittalOut_df)
            data = data.sort_values(by =['main_SN'])
            data = data.reset_index(drop=True)
            data = pd.DataFrame(data)
            LengthDF = len(data)

            ## Making Transmittal Summary
            Inv_CountReport2   = data.groupby(['catg'], as_index=False).main_SN.count()
            Inv_CountReport2   = pd.DataFrame(Inv_CountReport2)
            Inv_CountReport2.rename(columns={'catg':'Category', 'main_SN':'Quantity'},inplace = True)
            Inv_CountReport2["Item"]  = Inv_CountReport2.shape[0]*[""]
            Inv_CountReport2["Owner"]  = Inv_CountReport2.shape[0]*[""]
            
            Inv_CountReport2["ManfSN"]     = Inv_CountReport2.shape[0]*["See Sheet Detail List"]
            Inv_CountReport2["UnitWeight"]  = Inv_CountReport2.shape[0]*[""]
            Inv_CountReport2["TotalWeight"]= Inv_CountReport2.shape[0]*[""]
            Inv_CountReport2["Comments"]          = Inv_CountReport2.shape[0]*[""]
            Inv_CountReport2    = Inv_CountReport2.loc[:,['Category','Item', 'ManfSN', 'Owner', 'Quantity', 'UnitWeight', 'TotalWeight', 'Comments']]
            Inv_CountReport2    = pd.DataFrame(Inv_CountReport2)
            Inv_CountReport2    = Inv_CountReport2.reset_index(drop=True)
            Inv_CountReport2.to_sql('Eagle_Inventory_transmittalOutFrontPage',conn, if_exists="replace", index=False)
            conn.commit()
            conn.close()
            for each_rec in range(len(Inv_CountReport2)):
                tree1.insert("", tk.END, values=list(Inv_CountReport2.loc[each_rec]))
        
        def DeleteSelectData():
            iDelete = tkinter.messagebox.askyesno("Delete Entry From Transmittal", "Confirm if you want to Delete")
            if iDelete >0:
                conn = sqlite3.connect("Eagle_Inventory.db")
                cur = conn.cursor()            
                TotalTransmittalEntry.delete(0,END)
                for selected_item in tree.selection():
                    cur.execute("DELETE FROM Eagle_Inventory_transmittalOut1 WHERE main_SN =? ", (tree.set(selected_item, '#4'),))
                    conn.commit()
                    tree.delete(selected_item)
                Total_count = len(tree.get_children())
                TotalTransmittalEntry.insert(tk.END,Total_count)
                tree1.delete(*tree1.get_children())
                conn.commit()
                conn.close()
                ReconnectDBAfterDelete()
            return

        def iExit():
            window.destroy()

        def ClearDB():
            conn = sqlite3.connect("Eagle_Inventory.db")
            cur = conn.cursor()
            cur.execute("DELETE FROM Eagle_Inventory_transmittalOutFrontPage")
            cur.execute("DELETE FROM Eagle_Inventory_transmittalOut1")
            conn.commit()
            conn.close()
                

        def UpdateTransmittal():
            CategoryUpdate      = EquipmentType.get()
            ItemUpdate          = EquipmentName.get()
            ItemSerialNo        = ItemSN.get()
            OwnerUpdate         = OwnerName.get()
            QuantityUpdate      = ItemQuantity.get()
            Unit_WeightUpdate   = UnitWeight.get()
            Total_WeightUpdate  = TotalWeight.get()
            CommentsUpdate  = Comments.get()                
            con= sqlite3.connect("Eagle_Inventory.db")
            cur=con.cursor()
            sqlite_update_query = """Update Eagle_Inventory_transmittalOutFrontPage set Item = ?, ManfSN = ?, Owner = ?, Quantity =?, UnitWeight = ?, TotalWeight = ?, Comments = ? where Category = ?"""
            columnValues = (ItemUpdate, ItemSerialNo, OwnerUpdate, QuantityUpdate, Unit_WeightUpdate, Total_WeightUpdate, CommentsUpdate, CategoryUpdate )
            cur.execute(sqlite_update_query, columnValues)
            tree1.delete(*tree1.get_children())
            TransmittalFront_DF = pd.read_sql_query("SELECT * FROM Eagle_Inventory_transmittalOutFrontPage;", con)
            data = pd.DataFrame(TransmittalFront_DF)        
            data = data.reset_index(drop=True)
            data = pd.DataFrame(data)
            for each_rec in range(len(data)):
                tree1.insert("", tk.END, values=list(data.loc[each_rec]))       
            EquipmentType.delete(0,END)
            EquipmentName.delete(0,END)
            OwnerName.delete(0,END)                                                
            ItemQuantity.delete(0,END)
            UnitWeight.delete(0,END)
            TotalWeight.delete(0,END)
            Comments.delete(0,END)
            ItemSN.delete(0,END)
            con.commit()
            cur.close()
            con.close()


        ## Command Button
        btnExitData = Button(window, text="Exit", font=('aerial', 10, 'bold'), height =1, width=6, bd=4, command = iExit)
        btnExitData.place(x=1193,y=847)

        btnUpdateTransmittalInfo = Button(window, text="Update Transmittal", font=('aerial', 10, 'bold'), height =1, width=16, bd=4, command = UpdateTransmittal)
        btnUpdateTransmittalInfo.place(x=1115,y=813)

        btnDeleteSelected = Button(window, text="Delete Selected", font=('aerial', 10, 'bold'), height =1, width=14, bd=2, command = DeleteSelectData)
        btnDeleteSelected.place(x=200,y=85)

        btnExportTransmittalOut = Button(window, text="Export Transmittal Out\n As Excel", font=('aerial', 10, 'bold'), height =2, width=19, bd=4, command = ExportTransmittal)
        btnExportTransmittalOut.place(x=6,y=820)





























