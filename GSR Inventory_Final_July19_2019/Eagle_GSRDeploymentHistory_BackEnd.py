import sqlite3
#backend

def GSRDeploymentHistoryData():
    con= sqlite3.connect("Eagle_GSRDeploymentHistory.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_MASTER (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_TEMP (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_ANALYZED (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_MERGED_TEMP (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_MASTER (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRDeploymentHistory_UNKNOWN_DEVICE_TYPE_TEMP (CaseSrNo integer NOT NULL, DeviceType text, Line text,\
                OccupiedStations text, StartTimeUTC text, EndTimeUTC text , JobName text, DuplicatedEntries text)")
    
    con.commit()
    con.close()



def Combo_input_CaseSrNo():
    con= sqlite3.connect("Eagle_GSRDeploymentHistory.db")
    cur=con.cursor()
    cur.execute("SELECT CaseSrNo FROM Eagle_GSRDeploymentHistory_MASTER")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_DeviceType():
    con= sqlite3.connect("Eagle_GSRDeploymentHistory.db")
    cur=con.cursor()
    cur.execute("SELECT DeviceType FROM Eagle_GSRDeploymentHistory_MASTER")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data






GSRDeploymentHistoryData()
