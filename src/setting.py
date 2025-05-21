import os
import eel
import sqlite3

@eel.expose
def init():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()
    sql = """CREATE TABLE IF NOT EXISTS excel(id INTEGER PRIMARY KEY AUTOINCREMENT, name, item_head1, item_head2, item_head3, item_head4, item_head5, item_head6, item_head7, item_head8, item_head9, item_head10, item_head11, item_head12, item_head13, item_head14, item_head15)"""
    cursor.execute(sql)

    sql = """CREATE TABLE IF NOT EXISTS option(id INTEGER PRIMARY KEY AUTOINCREMENT, selected_id INTEGER)"""
    cursor.execute(sql)

    sql = """CREATE TABLE IF NOT EXISTS settings(id INTEGER PRIMARY KEY AUTOINCREMENT, excel_type, setting1, setting2, setting3, setting4, setting5, setting6, setting7, setting8, setting9, setting10, setting11, setting12, setting13, setting14, setting15)"""
    cursor.execute(sql)

    conn.commit()

@eel.expose
def getCurrentId():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    cursor.execute("SELECT selected_id FROM option")
    rows = cursor.fetchall()

    for row in rows:
        for i in row:
            id = i

    conn.close()
    
    return id

@eel.expose
def getIndexList():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    index_num = []
    index = []

    cursor.execute("SELECT id FROM excel")
    rows = cursor.fetchall()

    for row in rows:
        index_num.append(row)

    cursor.execute("SELECT name FROM excel")
    rows = cursor.fetchall()

    for row in rows:
        index.append(row)

    conn.close()

    return (index_num, index)

@eel.expose
def getSelectedItem():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()
    
    cursor.execute("SELECT selected_id FROM option")
    rows = cursor.fetchall()
    for row in rows:
        for i in row:
            id = str(i)
            
    cursor.execute("SELECT * FROM excel WHERE id=?", (id))
    rows = cursor.fetchall()
    if rows:
        rows = rows[0][1:]

    cursor.execute("SELECT * FROM settings WHERE id=?", (id))
    setting_rows = cursor.fetchall()
    if setting_rows:
        setting_rows = setting_rows[0][1:]

    conn.close()

    return (rows,setting_rows)

@eel.expose
def addItemHeads(add_inputs, add_setting_inputs):
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    cursor.execute("INSERT OR REPLACE INTO excel (name, item_head1, item_head2, item_head3, item_head4, item_head5, item_head6, item_head7, item_head8, item_head9, item_head10, item_head11, item_head12, item_head13, item_head14, item_head15) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (add_inputs))
    conn.commit()

    cursor.execute("INSERT OR REPLACE INTO settings(excel_type, setting1, setting2, setting3, setting4, setting5, setting6, setting7, setting8, setting9, setting10, setting11, setting12, setting13, setting14, setting15) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (add_setting_inputs))
    conn.commit()

    conn.close()

@eel.expose
def editItemHeads(inputs, setting_inputs):
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()
    
    cursor.execute("SELECT selected_id FROM option")
    rows = cursor.fetchall()
    for row in rows:
        for i in row:
            id = i

    edit_inputs = []
    edit_inputs.append(id)
    edit_setting_inputs = []
    edit_setting_inputs.append(id)

    for i in inputs:
        edit_inputs.append(i)
    
    for si in setting_inputs:
        edit_setting_inputs.append(si)

    cursor.execute("INSERT OR REPLACE INTO excel (id, name, item_head1, item_head2, item_head3, item_head4, item_head5, item_head6, item_head7, item_head8, item_head9, item_head10, item_head11, item_head12, item_head13, item_head14, item_head15) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (edit_inputs))
    conn.commit()
    cursor.execute("INSERT OR REPLACE INTO settings (id, excel_type, setting1, setting2, setting3, setting4, setting5, setting6, setting7, setting8, setting9, setting10, setting11, setting12, setting13, setting14, setting15) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (edit_setting_inputs))
    conn.commit()

    conn.close()

@eel.expose
def changeIndex(id):
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    cursor.execute("INSERT OR REPLACE INTO option (id,selected_id) VALUES (1,?)", (id))
    conn.commit()

    conn.close()


# 初期設定

init()
dbname = "excel.db"
connection = sqlite3.connect(dbname)
cursor = connection.cursor()

cursor.execute("SELECT * FROM option")

if cursor.fetchone() == None:
    changeIndex("1")

# 削除
# dbname = ('excel.db')
# conn = sqlite3.connect(dbname, isolation_level=None)
# cursor = conn.cursor()
# sql = """DROP TABLE if exists excel"""

# # #命令を実行
# conn.execute(sql)
# conn.commit()#コミットする

# conn.close()
