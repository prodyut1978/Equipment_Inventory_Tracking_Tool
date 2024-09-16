import os
from tkinter import*
import tkinter.messagebox
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog
from tkinter import filedialog
import pandas as pd
import openpyxl
import csv
import time
import datetime
import win32com.client

def Generate_ReceiverAddress():
    window = Tk()
    window.title("Input Entry For Receiver Location")
    window.config(bg="ghost white")
    width = 550
    height = 400
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry("%dx%d+%d+%d" % (width, height, x, y))
    window.resizable(0, 0)
    window.grid()
    TitFrame = Frame(window, bd = 2, padx= 5, pady= 4, relief = RIDGE)
    TitFrame.pack(side = TOP)
    InputHeader = Label(TitFrame, font=('aerial', 12, 'bold'), text="Entry Widget For Receiver Location File")
    InputHeader.grid()

    DataFrameLEFT = LabelFrame(window, bd = 1, width = 490, height = 400, padx= 6, pady= 10,relief = RIDGE,
                                       bg = "Ghost White",font=('aerial', 15, 'bold'))
    DataFrameLEFT.place(x=0,y=80)

    def ClearAll():
        lblNameEntries.delete(0,END)
        lblStreetEntries.delete(0,END)
        lblCityEntries.delete(0,END)                                                
        lblCityEntries.delete(0,END)
        lblProvinceEntries.delete(0,END)
        lblPostalCodeEntries.delete(0,END)
        lblPhoneNumberEntries.delete(0,END)

    def iExit():
        window.destroy()
        

    def SubmitReceiverAddressToDB():
        OrganizationName = lblNameEntries.get()
        StreetName       = lblStreetEntries.get()
        CityName         = lblCityEntries.get()
        ProvinceName     = lblProvinceEntries.get()
        PostalCode       = lblPostalCodeEntries.get()
        PhoneNumber      = 'Ph: ' + lblPhoneNumberEntries.get()
        City_Province_PostalCode = CityName + ' ' + ' ' + ProvinceName + ' ' + ' ' + PostalCode
        DBEntry =  [OrganizationName + ',' + StreetName + ',' + City_Province_PostalCode + ',' + PhoneNumber]
        
        con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
        cur=con.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS Eagle_ReceiverLocation (ReceiverAddress text)")
        cur.execute("INSERT INTO Eagle_ReceiverLocation VALUES (?)",(DBEntry))
        con.commit()
        con.close()
        ClearAll()
        
        

    lblName = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "1. Receiving Organization Name:", padx =10, pady= 10, bg = "Ghost White")
    lblName.grid(row =0, column = 0, sticky =W)
    lblNameEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblNameEntries.grid(row =0, column = 1)

    lblStreet = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "2. Street Address:", padx =10, pady= 10, bg = "Ghost White")
    lblStreet.grid(row =1, column = 0, sticky =W)
    lblStreetEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblStreetEntries.grid(row =1, column = 1)

    lblCity = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "3. City:", padx =10, pady= 10, bg = "Ghost White")
    lblCity.grid(row =2, column = 0, sticky =W)
    lblCityEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblCityEntries.grid(row =2, column = 1)

    lblProvince = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "4. Province:", padx =10, pady= 10, bg = "Ghost White")
    lblProvince.grid(row =3, column = 0, sticky =W)
    lblProvinceEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblProvinceEntries.grid(row =3, column = 1)

    lblPostalCode = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "5. Postal Code:", padx =10, pady= 10, bg = "Ghost White")
    lblPostalCode.grid(row =4, column = 0, sticky =W)
    lblPostalCodeEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblPostalCodeEntries.grid(row =4, column = 1)

    lblPhoneNumber = Label(DataFrameLEFT, font=('aerial', 10, 'bold'), text = "6. Phone Number:", padx =10, pady= 10, bg = "Ghost White")
    lblPhoneNumber.grid(row =5, column = 0, sticky =W)
    lblPhoneNumberEntries  = Entry(DataFrameLEFT, font=('aerial', 12, 'bold'),textvariable= StringVar(), width = 30, bd=2)
    lblPhoneNumberEntries.grid(row =5, column = 1)


    btnExit = Button(window, text="Exit", font=('aerial', 10, 'bold'), height =1, width=12, bd=2, command =iExit)
    btnExit.place(x=6,y=350)


    btnClear = Button(window, text="Clear Entries", font=('aerial', 10, 'bold'), height =1, width=12, bd=2, command =ClearAll)
    btnClear.place(x=156,y=350)


    btnSubmit = Button(window, text="Submit Entries", font=('aerial', 10, 'bold'), height =1, width=12, bd=2, command =SubmitReceiverAddressToDB)
    btnSubmit.place(x=415,y=350)
