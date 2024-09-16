import sqlite3
#backend

def GSRRepairInventoryData():
    con= sqlite3.connect("Eagle_GSRRepairInventory.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRRepairInventory (WorkOrderNo integer NOT NULL, CaseSrNo integer NOT NULL, PartNo text,\
                DeviceType text, TechnicianInput text, CrewReported text, WarrantyStatus text, Chargeable text , PricePer text, DiscountApplied text, SubTotal text, DateRepaired text)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_GSRRepairInventory_TEMP (WorkOrderNo integer NOT NULL, CaseSrNo integer NOT NULL, PartNo text,\
                DeviceType text, TechnicianInput text, CrewReported text, WarrantyStatus text, Chargeable text , PricePer text, DiscountApplied text, SubTotal text, DateRepaired text)")

    con.commit()
    con.close()



GSRRepairInventoryData()
