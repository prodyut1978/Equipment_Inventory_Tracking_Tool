import sqlite3
#backend

def inventoryData():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory (catg text, manuf text, model text, main_SN text NOT NULL, desc text, asset_SN text,\
                datestamp text, location text, condition text, origin text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_temp (catg text, manuf text, model text, main_SN text NOT NULL, desc text, asset_SN text,\
                              datestamp text, location text, condition text, origin text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_Merged_temp (catg text, manuf text, model text, main_SN text NOT NULL, desc text, asset_SN text,\
                              datestamp text, location text, condition text, origin text, Status text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_transmittalOut1 (catg text, manuf text, model text, main_SN text NOT NULL, desc text, asset_SN text,\
                              datestamp text, location text, condition text, origin text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_transmittalOut2 (catg text, manuf text, model text, main_SN text NOT NULL, desc text, asset_SN text,\
                              datestamp text, location text, condition text, origin text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_transmittalOutFrontPage (Category text, Item text, ManfSN text, Owner text, Quantity integer, UnitWeight real,\
                              TotalWeight real, Comments text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_Scan_TEMP (asset_SN text, main_SN text, location text, datestamp text)")    
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_Report_1 (Category text, Model_Name text, Total_Count text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_Report_2 (Category text, Location text, Model_Name text, Total_Count text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_Inventory_Report_3 (Category text, Total_Count text)")
    con.commit()
    con.close()


def addInvRec(catg, manuf, model, main_SN, desc, asset_SN, datestamp, location, condition, origin):
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()    
    cur.execute("INSERT INTO Eagle_Inventory VALUES (?,?,?,?,?,?,?,?,?,?)",(catg, manuf, model, main_SN, desc, asset_SN, datestamp, location, condition, origin))
    con.commit()
    con.close()


def viewData():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT * FROM Eagle_Inventory ORDER BY `catg` ASC")
    rows=cur.fetchall()
    con.close()
    return rows


def searchData(catg = "", manuf = "", model = "", main_SN = "", desc = "", asset_SN = "", datestamp = "", location = "", condition = "", origin = ""):
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT * FROM Eagle_Inventory WHERE catg = ? COLLATE NOCASE OR manuf = ? COLLATE NOCASE OR model = ? COLLATE NOCASE OR main_SN = ? COLLATE NOCASE OR \
                desc = ? COLLATE NOCASE OR asset_SN = ? COLLATE NOCASE OR datestamp = ? OR location = ? COLLATE NOCASE OR condition = ? COLLATE NOCASE OR origin = ? COLLATE NOCASE",\
                (catg, manuf, model, main_SN, desc, asset_SN, datestamp, location, condition, origin))
    rows=cur.fetchall()
    con.close()
    return rows

def Combo_input_Categ():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT catg FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_Manuf():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT manuf FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_Model():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT model FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Main_SN():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT main_SN FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Desc():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT desc FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_Asset_SN():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT asset_SN FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data

def Combo_input_Datestamp():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT datestamp FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Location():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT location FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    return data

def Combo_input_Condition():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT condition FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data


def Combo_input_Origin():
    con= sqlite3.connect("Eagle_Inventory.db")
    cur=con.cursor()
    cur.execute("SELECT origin FROM Eagle_Inventory")
    data = []
    for row in cur.fetchall():
        data.append(row[0])
    con.close()
    return data



inventoryData()
