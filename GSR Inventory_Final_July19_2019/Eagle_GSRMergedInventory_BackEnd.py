import sqlite3
#backend

def GSRMergedInventoryData():
    con= sqlite3.connect("Eagle_GSRMergedInventory.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRMergedInventory_MASTER(CaseSrNo integer NOT NULL, DeviceType text, ProjectID text,\
                FlashCapacityGB text, LastTimeSeenInDTMDt text, LastTimeLineViewedDt text, LastTimeReapedDt text, FlagsRepair text,\
                WorkOrderNo text, PartNo text,TechnicianInput text, CrewReported text,DateRepaired text, FlagsDeployment text,\
                StartTimeUTC text, EndTimeUTC text , JobName text)")

    con.commit()
    con.close()

def Combo_input_CaseSrNo():
    con= sqlite3.connect("Eagle_GSRMergedInventory.db")
    cur=con.cursor()
    cur.execute("SELECT CaseSrNo FROM Eagle_GSRMergedInventory_MASTER")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_DeviceType():
    con= sqlite3.connect("Eagle_GSRMergedInventory.db")
    cur=con.cursor()
    cur.execute("SELECT DeviceType FROM Eagle_GSRMergedInventory_MASTER")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ProjectID():
    con= sqlite3.connect("Eagle_GSRMergedInventory.db")
    cur=con.cursor()
    cur.execute("SELECT ProjectID FROM Eagle_GSRMergedInventory_MASTER")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data



GSRMergedInventoryData()
