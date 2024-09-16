import sqlite3
#backend

def GSRinventoryData():
    con= sqlite3.connect("Eagle_GSRInventory.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRInventory (CaseSrNo integer NOT NULL, DeviceType text, ProjectID text,\
                CpuSerialNumber text, BootVersion text, ApplicationVersion text ,\
                FlashSerialNumber text, FlashCapacityGB text,\
                LastTimeSeenInDTMDt text, LastTimeLineViewedDt text, LastTimeReapedDt text,\
                LastTimeTestedDt text, FirstTimeScriptedDt text, InitialScript text, LastTimeScriptedDt text, CurrentScript text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRInventory_TEMP (CaseSrNo integer NOT NULL, DeviceType text, ProjectID text,\
                CpuSerialNumber text, BootVersion text, ApplicationVersion text ,\
                FlashSerialNumber text, FlashCapacityGB text,\
                LastTimeSeenInDTMDt text, LastTimeLineViewedDt text, LastTimeReapedDt text,\
                LastTimeTestedDt text, FirstTimeScriptedDt text, InitialScript text, \
                LastTimeScriptedDt text, CurrentScript text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRInventory_ANALYZED_TEMP (CaseSrNo integer NOT NULL, DeviceType text, ProjectID text,\
                CpuSerialNumber text, BootVersion text, ApplicationVersion text ,\
                FlashSerialNumber text, FlashCapacityGB text,\
                LastTimeSeenInDTMDt text, LastTimeLineViewedDt text, LastTimeReapedDt text,\
                LastTimeTestedDt text, FirstTimeScriptedDt text, InitialScript text, \
                LastTimeScriptedDt text, CurrentScript text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRInventory_MERGED_TEMP (CaseSrNo integer NOT NULL, DeviceType text, ProjectID text,\
                CpuSerialNumber text, BootVersion text, ApplicationVersion text ,\
                FlashSerialNumber text, FlashCapacityGB text,\
                LastTimeSeenInDTMDt text, LastTimeLineViewedDt text, LastTimeReapedDt text,\
                LastTimeTestedDt text, FirstTimeScriptedDt text, InitialScript text, \
                LastTimeScriptedDt text, CurrentScript text, DuplicatedEntries text)")
    
    con.commit()
    con.close()


def viewGSRInventoryMaster():
    con= sqlite3.connect("Eagle_GSRInventory.db")
    cur=con.cursor()
    cur.execute("SELECT * FROM Eagle_GSRInventory ORDER BY `CaseSrNo` ASC")
    rows=cur.fetchall()
    con.close()
    return rows

def Combo_input_CaseSrNo():
    con= sqlite3.connect("Eagle_GSRInventory.db")
    cur=con.cursor()
    cur.execute("SELECT CaseSrNo FROM Eagle_GSRInventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_DeviceType():
    con= sqlite3.connect("Eagle_GSRInventory.db")
    cur=con.cursor()
    cur.execute("SELECT DeviceType FROM Eagle_GSRInventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_ProjectID():
    con= sqlite3.connect("Eagle_GSRInventory.db")
    cur=con.cursor()
    cur.execute("SELECT ProjectID FROM Eagle_GSRInventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data




GSRinventoryData()
