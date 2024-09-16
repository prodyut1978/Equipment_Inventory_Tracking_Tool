import sqlite3
#backend

def BadGSRinventoryData():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_BadGSRInventoryDatabase (BatchNumber text, JobName text, CrewNumber text,\
                 Location text, Date text, Unit_SN integer, DeviceType text, Opened text, FaultFound text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_BadGSRInventoryDatabase_ACCUMULATED_DUPLICATED(BatchNumber text, JobName text, CrewNumber text,\
                 Location text, Date text, Unit_SN integer, DeviceType text, Opened text, FaultFound text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_BadGSRInventoryDatabase_TEMP_DUPLICATED(BatchNumber text, JobName text, CrewNumber text,\
                 Location text, Date text, Unit_SN integer, DeviceType text, Opened text, FaultFound text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_BadGSRInventoryDatabase_TEMP_IMPORT(BatchNumber text, JobName text, CrewNumber text,\
                 Location text, Date text, Unit_SN integer, DeviceType text, Opened text, FaultFound text, DuplicatedEntries text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_BadGSRInventoryDatabase_TRANSMITTAL_OUT(BatchNumber text, JobName text, CrewNumber text,\
                 Location text, Date text, Unit_SN integer, DeviceType text, Opened text, FaultFound text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_ReceiverLocation (ReceiverAddress text)")
    con.commit()
    con.close()


def addInvRec(BatchNumber, JobName, CrewNumber, Location, Date, Unit_SN, DeviceType, Opened, FaultFound):
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()    
    cur.execute("INSERT INTO Eagle_BadGSRInventoryDatabase VALUES (?,?,?,?,?,?,?,?,?)",(BatchNumber, JobName, CrewNumber, Location, Date, Unit_SN, DeviceType, Opened, FaultFound))
    con.commit()
    con.close()


def viewData():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT * FROM Eagle_BadGSRInventoryDatabase ORDER BY `BatchNumber` ASC")
    rows=cur.fetchall()
    con.close()
    return rows


def searchData(BatchNumber = "", JobName = "", CrewNumber = "", Location = "", Date = "", Unit_SN = "", DeviceType = "", Opened = "", FaultFound = ""):
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT * FROM Eagle_BadGSRInventoryDatabase WHERE BatchNumber = ? COLLATE NOCASE OR JobName = ? COLLATE NOCASE OR CrewNumber = ? COLLATE NOCASE OR Location = ? COLLATE NOCASE OR \
                Date = ? COLLATE NOCASE OR Unit_SN = ? COLLATE NOCASE OR DeviceType = ? OR Opened = ? COLLATE NOCASE OR FaultFound = ? COLLATE NOCASE",\
                (BatchNumber, JobName, CrewNumber, Location, Date, Unit_SN, DeviceType, Opened, FaultFound))
    rows=cur.fetchall()
    con.close()
    return rows

def Combo_input_BatchNumber():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT BatchNumber FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_JobName():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT JobName FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ProjectName():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT JobName FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_CrewNumber():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT CrewNumber FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ProjectCrewNumber():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT CrewNumber FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Location():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT Location FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ProjectLocation():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT Location FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Date():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT Date FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_Unit_SN():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT Unit_SN FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_DeviceType():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT DeviceType FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Opened():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT Opened FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    return data

def Combo_input_FaultFound():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT FaultFound FROM Eagle_BadGSRInventoryDatabase")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ReceiverLocation():
    con= sqlite3.connect("Eagle_BadGSRInventoryDatabase.db")
    cur=con.cursor()
    cur.execute("SELECT ReceiverAddress FROM Eagle_ReceiverLocation")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data




BadGSRinventoryData()
