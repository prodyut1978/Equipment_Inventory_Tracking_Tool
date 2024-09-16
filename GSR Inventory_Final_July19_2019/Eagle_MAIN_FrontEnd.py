#Front End
import os
from tkinter import*
import tkinter.messagebox

import Eagle_GSRInventory_BackEnd
import Eagle_GSRRepairInventory_BackEnd
import Eagle_GSRMergedInventory_BackEnd
import Eagle_GSRDeploymentHistory_BackEnd

import Eagle_GSRInventory_Import_Module4
import Eagle_GSRInventory_Search_Module2

import Eagle_GSRRepair_Import_Module1
import Eagle_GSRRepair_Search_Module1

import Eagle_GSRDeploymentHistory_Import_Module1

import Eagle_GSRMerge_Module2
import Eagle_GSRMerge_QueryModule1

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

class GSRInventoryFrontEnd:    
    def __init__(self,root):
        
##  ----------------- Define Window-------------
        self.root =root
        self.root.title ("Eagle GSR Inventory Main")
        self.root.geometry("1180x480+10+0")
        self.root.config(bg="ghost white")
        self.root.resizable(0, 0)
        TitFrame = Frame(self.root, bd = 1, padx= 5, pady= 4, bg = "ghost white", relief = RIDGE)
        TitFrame.pack(side = TOP)
        self.lblTit = Label(TitFrame, bd= 4, font=('aerial', 14, 'bold'), bg = "ghost white", text="EAGLE GSR INVENTORY MANAGEMENT SYSTEM")
        self.lblTit.grid()

        def OpenGSRInventoryImportWizard():
            from Eagle_GSRInventory_Import_Module4 import GSRInventoryImport
            if __name__ == '__main__':
                root = Tk()
                application  = GSRInventoryImport(root)
                root.mainloop()
        def OpenGSRRepairImportWizard():
            from Eagle_GSRRepair_Import_Module1 import GSRRepairImport
            if __name__ == '__main__':
                root = Tk()
                application  = GSRRepairImport(root)
                root.mainloop()

        def OpenGSRDeploymentImportWizard():
            from Eagle_GSRDeploymentHistory_Import_Module1 import GSRDeploymentHistoryImport
            if __name__ == '__main__':
                root = Tk()
                application  = GSRDeploymentHistoryImport(root)
                root.mainloop()

        def MergeGSRInv_GSRRepairWizard():
            from Eagle_GSRMerge_Module2 import GSRRepairMaster_Merge_GSRInvMaster
            if __name__ == '__main__':
                root = Tk()
                application  = GSRRepairMaster_Merge_GSRInvMaster(root)
                root.mainloop()

        def MergedGSRInv_GSRRepairQueryWizard():
            from Eagle_GSRMerge_QueryModule1 import MergedGSRInventoryRepairQuery
            if __name__ == '__main__':
                root = Tk()
                application  = MergedGSRInventoryRepairQuery(root)
                root.mainloop()

        def GSRInventorySearchWizard():
            from Eagle_GSRInventory_Search_Module2 import GSRInventorySearch
            if __name__ == '__main__':
                root = Tk()
                application  = GSRInventorySearch(root)
                root.mainloop()

        def GSRRepairSearchWizard():
            from Eagle_GSRRepair_Search_Module1 import GSRRepairSearch
            if __name__ == '__main__':
                root = Tk()
                application  = GSRRepairSearch(root)
                root.mainloop()
    
##  ----------------- Define Labels-------------
        L1 = Label(self.root, text = "A: Import Modules GSRInventory & GSRRepair & GSRDeployment Files:", font=("arial", 12,'bold'), bg= "green").place(x=10,y=55)
        L2 = Label(self.root, text = "1:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=30,y=88)
        L3 = Label(self.root, text = "2:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=30,y=134)
        L4 = Label(self.root, text = "3:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=30,y=180)
        
        btnImportGSRInv = Button(self.root, text="GSR Inventory Reports Import Wizard", font=('aerial', 11, 'bold'), height =1, width=32, bd=4,
                                 command = OpenGSRInventoryImportWizard)
        btnImportGSRInv.place(x=55,y=84)

        btnImportGSRRepair = Button(self.root, text="GSR Repaired Reports Import Wizard", font=('aerial', 11, 'bold'), height =1, width=32, bd=4,
                                 command = OpenGSRRepairImportWizard)
        btnImportGSRRepair.place(x=55,y=130)

        btnImportGSRDeployment = Button(self.root, text="GSR Deployment Reports Import Wizard", font=('aerial', 11, 'bold'), height =1, width=32, bd=4,
                                 command = OpenGSRDeploymentImportWizard)
        btnImportGSRDeployment.place(x=55,y=176)


        L5 = Label(self.root, text = "B: Merge Master GSR Inventory & GSR Repair DB & GSRDeployment History:", font=("arial", 12,'bold'), bg= "green").place(x=10,y=250)
        L6 = Label(self.root, text = "1:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=30,y=283)

        btnMergeGSRInvGSRRepair = Button(self.root, text=" Merge Master Database A1 & A2 & A3", font=('aerial', 11, 'bold'), height =2, width=32, bd=4,
                                         command = MergeGSRInv_GSRRepairWizard)
        btnMergeGSRInvGSRRepair.place(x=55,y=279)


        L7 = Label(self.root, text = "C: Query in Merged GSR Inventory & GSR Repair DB & GSRDeployment History:", font=("arial", 12,'bold'), bg= "green").place(x=10,y=350)
        L8 = Label(self.root, text = "1:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=30,y=383)

        btnQueryGSRInvGSRRepairMerged = Button(self.root, text=" Query In Merged Databases From B1", font=('aerial', 11, 'bold'), height =2, width=32, bd=4,
                                         command = MergedGSRInv_GSRRepairQueryWizard)
        btnQueryGSRInvGSRRepairMerged.place(x=55,y=379)

        L9 = Label(self.root, text = "D: Search Modules For GSR Inventory Files OR GSR Repair Files:", font=("arial", 12,'bold'), bg= "green").place(x=650,y=55)
        L10 = Label(self.root, text = "1:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=670,y=88)
        L11 = Label(self.root, text = "2:", font=("arial", 12,'bold'), bg= "cadet blue", bd=4).place(x=670,y=134)
        
        btnSearchGSRInv = Button(self.root, text="GSR Inventory Search Wizard", font=('aerial', 11, 'bold'), height =1, width=24, bd=4,
                                 command = GSRInventorySearchWizard)
        btnSearchGSRInv.place(x=695,y=84)

        btnSearchGSRRepair = Button(self.root, text="GSR Repair Search Wizard", font=('aerial', 11, 'bold'), height =1, width=24, bd=4,
                                    command = GSRRepairSearchWizard)
        btnSearchGSRRepair.place(x=695,y=130)



       


if __name__ == '__main__':
    root = Tk()
    application  = GSRInventoryFrontEnd(root)
    root.mainloop()
