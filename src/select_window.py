import wx
import os
import eel
import math
import sqlite3
import datetime
import pandas as pd

class FileDropTarget(wx.FileDropTarget):
    """ Drag & Drop Class """
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        self.window.write(filenames)

        return 0


class App(wx.Frame):
    """ GUI """
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, size=(500, 500), style=wx.DEFAULT_FRAME_STYLE)

        p = wx.Panel(self, wx.ID_ANY)

        label = wx.StaticText(p, wx.ID_ANY, 'DnD file', style=wx.SIMPLE_BORDER | wx.TE_CENTER)
        label.SetBackgroundColour("#e0ffff")

        dt1 = FileDropTarget(self)
        
        label.SetDropTarget(dt1)

        self.text_entry = wx.TextCtrl(p, wx.ID_ANY)

        layout = wx.BoxSizer(wx.VERTICAL)
        layout.Add(label, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        layout.Add(self.text_entry, flag=wx.EXPAND | wx.ALL, border=10)
        p.SetSizer(layout)
        
        self.Show()
    
    def write(self,pathnames):
        write_excel(pathnames)
        self.Destroy()

@eel.expose
def openWindow():
    app = wx.App()
    App(None, -1, 'エクセル変換')
    app.MainLoop()


def write_excel(pathnames):
    item_names = getCurrentItem()
    item_names = [a for a in item_names if a != '']
    item_settings = getCurrentSettings()

    sub_columns = {}
    data = {}
    max = 50
    end = 42

    for n in item_names:
        sub_columns[n] = []

    if item_settings[0] == "0" or item_settings[0] == 0: #0:表, 1:複数表
        if type(pathnames) == str:
            input_file = pd.ExcelFile(pathnames)
            sheet_names = input_file.sheet_names
            
            column_list = []
            for sheet_name in sheet_names:
                main_columns = {}
                df = pd.read_excel(p,index_col=None,header=None,sheet_name=sheet_name)
                df = df.iloc[1:,:end].dropna(how="all")
                x = 0

                for item_name in item_names:
                    x += 1
                    empty_list = [0]*len(df)

                    if item_name != "":
                        if item_name in df.columns:
                            for m in df[item_name]:
                                if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                    if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                        column_list.append(m.strftime('%Y/%m/%d'))
                                    else:
                                        column_list.append(m)
                                else:
                                    if type(m) is int or type(m) is float:
                                        if math.isnan(m):
                                            column_list.append(0)
                                        else: 
                                            if item_settings[x] == "1":
                                                column_list.append(round(m))
                                            elif item_settings[x] == "2":
                                                column_list.append(math.floor(m))
                                            elif item_settings[x] == "3":
                                                column_list.append(math.ceil(m))
                                            elif item_settings[x] == "4":
                                                if type(m) is float:
                                                    int_value = int(m)
                                                    float_value = m - int_value
                                                    float_value = float('{:.2f}'.format(float_value))
                                                    sexa_value = round(float_value*0.6,2)
                                                    total_value = int_value + sexa_value
                                                    total_value = '{:.2f}'.format(total_value)
                                                    column_list.append(total_value)
                                                else:
                                                    column_list.append(m)
                                            elif item_settings[x] == "5":
                                                if type(m) is float:
                                                    int_value = int(m)
                                                    float_value = m - int_value
                                                    float_value = float('{:.2f}'.format(float_value))
                                                    decimal_value = round(float_value*10/6,2)
                                                    total_value = int_value + decimal_value
                                                    total_value = '{:.2f}'.format(total_value)
                                                    column_list.append(total_value)
                                                else:
                                                    column_list.append(m)
                                            else:
                                                column_list.append(m)

                                    else: 
                                        column_list.append(0)
                            main_columns[item_name] = column_list
                                
                        else:
                            main_columns[item_name] = empty_list
                                            

        else:
            main_columns = {}
            for item_name in item_names:
                main_columns[item_name] = []
            for p in pathnames:
                input_file = pd.ExcelFile(p)
                sheet_names = input_file.sheet_names
                column_list = []
                for sheet_name in sheet_names:
                    df = pd.read_excel(p,index_col=None,header=None,sheet_name=sheet_name)
                    filtered_df = df.iloc[:,:end]
                    dfg = filtered_df.groupby((df.isnull().all(axis=1)).cumsum())
                    index_head = []
                    for index, g in dfg:
                        g = g.dropna(how="all")
                        if len(g) > 0:
                            g = g.values
                            if set(item_names) <= set(g[0]):
                                index_head = g[0]
                                d = pd.DataFrame(g[1:,:], columns=index_head)
                            elif len(index_head) > 0:
                                d = pd.DataFrame(g, columns=index_head)
                            else:
                                d = []

                            index_num = 0
                            divisions = {}

                            if len(d) > 0:
                                for i,r in d.items():
                                    if type(i) is str:
                                        divisions[index_num] = r
                                    index_num += 1
                                divisions[len(d)] = ""

                                num = 0
                                prev = 0
                                for i in divisions:
                                    if i >= 1:
                                        if num == 0:
                                            column_item = d.iloc[:,0:i]
                                        else:
                                            column_item = d.iloc[:,prev:i]
                                        num += 1
                                        prev = i

                                        x = 0
                                        for item_name in item_names:
                                            x += 1
                                            empty_list = [0]*len(column_item)
                                            
                                            column_list = []
                                            if item_name in d.columns:
                                                if item_name in column_item.columns:
                                                    if column_item.shape[1] <= 1:
                                                        for m in column_item[item_name]:
                                                            if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                    column_list.append(m.strftime('%Y/%m/%d'))
                                                                else:
                                                                    column_list.append(m)
                                                            else:
                                                                if type(m) is int or type(m) is float:
                                                                    if math.isnan(m):
                                                                        column_list.append(0)
                                                                    else: 
                                                                        if item_settings[x] == "1":
                                                                            column_list.append(round(m))
                                                                        elif item_settings[x] == "2":
                                                                            column_list.append(math.floor(m))
                                                                        elif item_settings[x] == "3":
                                                                            column_list.append(math.ceil(m))
                                                                        elif item_settings[x] == "4":
                                                                            if type(m) is float:
                                                                                int_value = int(m)
                                                                                float_value = m - int_value
                                                                                print(float_value)
                                                                                float_value = float('{:.2f}'.format(float_value))
                                                                                sexa_value = round(float_value*0.6,2)
                                                                                print(sexa_value)
                                                                                total_value = int_value + sexa_value
                                                                                total_value = '{:.2f}'.format(total_value)
                                                                                column_list.append(total_value)
                                                                            else:
                                                                                column_list.append(m)
                                                                        elif item_settings[x] == "5":
                                                                            print("type5")
                                                                            if type(m) is float:
                                                                                int_value = int(m)
                                                                                float_value = m - int_value
                                                                                float_value = float('{:.2f}'.format(float_value))
                                                                                decimal_value = round(float_value*10/6,2)
                                                                                total_value = int_value + decimal_value
                                                                                total_value = '{:.2f}'.format(total_value)
                                                                                column_list.append(total_value)
                                                                            else:
                                                                                column_list.append(m)
                                                                        else:
                                                                            column_list.append(m)

                                                                else: 
                                                                    column_list.append(0)
                                                        main_columns[item_name] += column_list
                                                    else: 
                                                        for m in column_item.iloc[:,-1]:
                                                            if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                    column_list.append(m.strftime('%Y/%m/%d'))
                                                                else:
                                                                    column_list.append(m)
                                                            else:
                                                                if type(m) is int or type(m) is float:
                                                                    if math.isnan(m):
                                                                        column_list.append(0)  
                                                                    else: 
                                                                        if item_settings[x] == "1":
                                                                            column_list.append(round(m))
                                                                        elif item_settings[x] == "2":
                                                                            column_list.append(math.floor(m))
                                                                        elif item_settings[x] == "3":
                                                                            column_list.append(math.ceil(m))
                                                                        elif item_settings[x] == "4":
                                                                            if type(m) is float:
                                                                                int_value = int(m)
                                                                                float_value = m - int_value
                                                                                float_value = float('{:.2f}'.format(float_value))
                                                                                decimal_value = round(float_value*0.6,2)
                                                                                total_value = int_value + decimal_value
                                                                                total_value = '{:.2f}'.format(total_value)
                                                                                column_list.append(total_value)
                                                                            else:
                                                                                column_list.append(m)
                                                                        elif item_settings[x] == "5":
                                                                            if type(m) is float:
                                                                                int_value = int(m)
                                                                                float_value = m - int_value
                                                                                float_value = float('{:.2f}'.format(float_value))
                                                                                sexa_value = round(float_value*10/6,2)
                                                                                total_value = int_value + sexa_value
                                                                                total_value = '{:.2f}'.format(total_value)
                                                                                column_list.append(total_value)
                                                                            else:
                                                                                column_list.append(m)
                                                                        else:
                                                                            column_list.append(m)
                                                                else: 
                                                                    column_list.append(0)                                 
                                                        main_columns[item_name] += column_list

                                            else:
                                                main_columns[item_name] += empty_list
                                                
        data = main_columns
    
    else:
        if type(pathnames) == str:
            input_file = pd.ExcelFile(pathnames)
            sheet_names = input_file.sheet_names
            
            for sheet_name in sheet_names:
                main_columns = {}
                for i in range(max):
                    main_columns[i] = {}
                df = pd.read_excel(p,index_col=None,header=None,sheet_name=sheet_name)
                filtered_df = df.iloc[1:,:end]
                dfg = filtered_df.groupby((df.isnull().all(axis=1)).cumsum())

                for index, g in dfg:
                    g = g.dropna(how="all")
                    if len(g) > 0:
                        g = g.values
                        d = pd.DataFrame(g[1:,:], columns=g[0])
                        index_num = 0
                        divisions = {}
                        for i,r in d.items():
                            if type(i) is str:
                                divisions[index_num] = r
                            index_num += 1
                        divisions[len(d)] = ""
                        
                        num = 0
                        prev = 0
                        for i in divisions:
                            if i >= 1:
                                if num == 0:
                                    column_item = d.iloc[:,0:i]
                                else:
                                    column_item = d.iloc[:,prev:i]
                                num += 1
                                prev = i

                                x = 0
                                for item_name in item_names:
                                    x += 1
                                    empty_list = [0]*len(column_item)
                                    
                                    column_list = []
                                    if item_name in d.columns:
                                        if item_name in column_item.columns:
                                            if column_item.shape[1] <= 1:
                                                for m in column_item[item_name]:
                                                    if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                        if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                            column_list.append(m.strftime('%Y/%m/%d'))
                                                        else:
                                                            column_list.append(m)
                                                    else:
                                                        if type(m) is int or type(m) is float:
                                                            if math.isnan(m):
                                                                column_list.append(0)
                                                            else: 
                                                                if item_settings[x] == "1":
                                                                    column_list.append(round(m))
                                                                elif item_settings[x] == "2":
                                                                    column_list.append(math.floor(m))
                                                                elif item_settings[x] == "3":
                                                                    column_list.append(math.ceil(m))
                                                                elif item_settings[x] == "4":
                                                                    if type(m) is float:
                                                                        int_value = int(m)
                                                                        float_value = m - int_value
                                                                        float_value = float('{:.2f}'.format(float_value))
                                                                        sexa_value = round(float_value*0.6,2)
                                                                        total_value = int_value + sexa_value
                                                                        total_value = '{:.2f}'.format(total_value)
                                                                        column_list.append(total_value)
                                                                    else:
                                                                        column_list.append(m)
                                                                elif item_settings[x] == "5":
                                                                    if type(m) is float:
                                                                        int_value = int(m)
                                                                        float_value = m - int_value
                                                                        float_value = float('{:.2f}'.format(float_value))
                                                                        decimal_value = round(float_value*10/6,2)
                                                                        total_value = int_value + decimal_value
                                                                        total_value = '{:.2f}'.format(total_value)
                                                                        column_list.append(total_value)
                                                                    else:
                                                                        column_list.append(m)
                                                                else:
                                                                    column_list.append(m)
                                                                    print(item_name,m,"here")

                                                        else: 
                                                            column_list.append(0)
                                                main_columns[index][item_name] = column_list
                                            else: 
                                                for m in column_item.iloc[:,-1]:
                                                    if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                        if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                            column_list.append(m.strftime('%Y/%m/%d'))
                                                        else:
                                                            column_list.append(m)
                                                    else:
                                                        if type(m) is int or type(m) is float:
                                                            if math.isnan(m):
                                                                column_list.append(0)  
                                                            else: 
                                                                if item_settings[x] == "1":
                                                                    column_list.append(round(m))
                                                                elif item_settings[x] == "2":
                                                                    column_list.append(math.floor(m))
                                                                elif item_settings[x] == "3":
                                                                    column_list.append(math.ceil(m))
                                                                elif item_settings[x] == "4":
                                                                    if type(m) is float:
                                                                        int_value = int(m)
                                                                        float_value = m - int_value
                                                                        float_value = float('{:.2f}'.format(float_value))
                                                                        sexa_value = round(float_value*0.6,2)
                                                                        total_value = int_value + sexa_value
                                                                        total_value = '{:.2f}'.format(total_value)
                                                                        column_list.append(total_value)
                                                                    else:
                                                                        column_list.append(m)
                                                                elif item_settings[x] == "5":
                                                                    if type(m) is float:
                                                                        int_value = int(m)
                                                                        float_value = m - int_value
                                                                        float_value = float('{:.2f}'.format(float_value))
                                                                        decimal_value = round(float_value*10/6,2)
                                                                        total_value = int_value + decimal_value
                                                                        total_value = '{:.2f}'.format(total_value)
                                                                        column_list.append(total_value)
                                                                    else:
                                                                        column_list.append(m)
                                                                else:
                                                                    column_list.append(m)
                                                                    print(item_name,m,"this")
                                                        else: 
                                                            column_list.append(0)                                 
                                                main_columns[index][item_name] = column_list

                                    else:
                                        main_columns[index][item_name] = empty_list
                                                        
                                    if item_name not in main_columns[index]:
                                        main_columns[index][item_name] = empty_list

                for i in range(max):
                    for name in item_names:
                        if name in main_columns[i]:
                            sub_columns[name] += main_columns[i][name]

        else:
            for p in pathnames:
                input_file = pd.ExcelFile(p)
                sheet_names = input_file.sheet_names
                
                for sheet_name in sheet_names:
                    main_columns = {}
                    for i in range(max):
                        main_columns[i] = {}
                    df = pd.read_excel(p,index_col=None,header=None,sheet_name=sheet_name)
                    filtered_df = df.iloc[1:,:end]
                    dfg = filtered_df.groupby((df.isnull().all(axis=1)).cumsum())

                    for index, g in dfg:
                        g = g.dropna(how="all")
                        if len(g) > 0:
                            g = g.values
                            d = pd.DataFrame(g[1:,:], columns=g[0])
                            index_num = 0
                            divisions = {}
                            for i,r in d.items():
                                if type(i) is str:
                                    divisions[index_num] = r
                                index_num += 1
                            divisions[len(d)] = ""
                            
                            num = 0
                            prev = 0
                            for i in divisions:
                                if i > 1:
                                    if num == 0:
                                        column_item = d.iloc[:,0:i]
                                    else:
                                        column_item = d.iloc[:,prev:i]
                                    num += 1
                                    prev = i

                                    x = 0
                                    for item_name in item_names:
                                        x += 1
                                        empty_list = [0]*len(column_item)
                                        
                                        column_list = []
                                        if item_name in d.columns:
                                            if item_name in column_item.columns:
                                                if column_item.shape[1] <= 1:
                                                    print(column_item)
                                                    for m in column_item[item_name]:
                                                        if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                            if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                column_list.append(m.strftime('%Y/%m/%d'))
                                                            else:
                                                                column_list.append(m)
                                                        else:
                                                            if type(m) is int or type(m) is float:
                                                                if math.isnan(m):
                                                                    column_list.append(0)
                                                                else: 
                                                                    if item_settings[x] == "1":
                                                                        column_list.append(round(m))
                                                                    elif item_settings[x] == "2":
                                                                        column_list.append(math.floor(m))
                                                                    elif item_settings[x] == "3":
                                                                        column_list.append(math.ceil(m))
                                                                    elif item_settings[x] == "4":
                                                                        if type(m) is float:
                                                                            int_value = int(m)
                                                                            float_value = m - int_value
                                                                            float_value = float('{:.2f}'.format(float_value))
                                                                            sexa_value = round(float_value*0.6,2)
                                                                            total_value = int_value + sexa_value
                                                                            total_value = '{:.2f}'.format(total_value)
                                                                            column_list.append(total_value)
                                                                        else:
                                                                            column_list.append(m)
                                                                    elif item_settings[x] == "5":
                                                                        if type(m) is float:
                                                                            int_value = int(m)
                                                                            float_value = m - int_value
                                                                            float_value = float('{:.2f}'.format(float_value))
                                                                            decimal_value = round(float_value*10/6,2)
                                                                            total_value = int_value + decimal_value
                                                                            total_value = '{:.2f}'.format(total_value)
                                                                            column_list.append(total_value)
                                                                        else:
                                                                            column_list.append(m)
                                                                    else:
                                                                        column_list.append(m)
                                                            else: 
                                                                column_list.append(0)
                                                    main_columns[index][item_name] = column_list
                                                else: 
                                                    print(column_item)
                                                    for m in column_item.iloc[:,-1]:
                                                        if type(m) is str or type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                            if type(m) is datetime.datetime or type(m) is pd._libs.tslibs.timestamps.Timestamp:
                                                                column_list.append(m.strftime('%Y/%m/%d'))
                                                            else:
                                                                column_list.append(m)
                                                        else:
                                                            if type(m) is int or type(m) is float:
                                                                if math.isnan(m):
                                                                    column_list.append(0)  
                                                                else: 
                                                                    if item_settings[x] == "1":
                                                                        column_list.append(round(m))
                                                                    elif item_settings[x] == "2":
                                                                        column_list.append(math.floor(m))
                                                                    elif item_settings[x] == "3":
                                                                        column_list.append(math.ceil(m))
                                                                    elif item_settings[x] == "4":
                                                                        if type(m) is float:
                                                                            int_value = int(m)
                                                                            float_value = m - int_value
                                                                            float_value = float('{:.2f}'.format(float_value))
                                                                            decimal_value = round(float_value*0.6,2)
                                                                            total_value = int_value + decimal_value
                                                                            total_value = '{:.2f}'.format(total_value)
                                                                            column_list.append(total_value)
                                                                        else:
                                                                            column_list.append(m)
                                                                    elif item_settings[x] == "5":
                                                                        if type(m) is float:
                                                                            int_value = int(m)
                                                                            float_value = m - int_value
                                                                            float_value = float('{:.2f}'.format(float_value))
                                                                            sexa_value = round(float_value*10/6,2)
                                                                            total_value = int_value + sexa_value
                                                                            total_value = '{:.2f}'.format(total_value)
                                                                            column_list.append(total_value)
                                                                        else:
                                                                            column_list.append(m)
                                                                    else:
                                                                        column_list.append(m)
                                                            else: 
                                                                column_list.append(0)                                 
                                                    main_columns[index][item_name] = column_list

                                        else:
                                            main_columns[index][item_name] = empty_list
                                                            
                                        if item_name not in main_columns[index]:
                                            main_columns[index][item_name] = empty_list
                    print(main_columns)
                    for i in range(max):
                        for name in item_names:
                            if name in main_columns[i]:
                                sub_columns[name] += main_columns[i][name]
    
            data = sub_columns

    # print(data)
    df_result = pd.DataFrame(data)
    df_result = df_result[df_result[list(item_names)[0]] != 0]
    print(len(df_result[list(item_names)[0]]))

    print(df_result)

    if not os.path.isdir("抽出結果"):
        os.mkdir("抽出結果")
        
    df_result.to_excel('抽出結果/result.xlsx', index=False)

def getCurrentItem():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    cursor.execute("SELECT selected_id FROM option")
    rows = cursor.fetchall()

    for row in rows:
        id = row

    cursor.execute("SELECT * FROM excel WHERE id=?",(id))
    rows = cursor.fetchall()

    item_names = {}
    num = 0

    if rows:
        for row in rows[0][2:]:
            num += 1
            item_names[row] = num

    conn.close()

    return (item_names)


def getCurrentSettings():
    dbname = ('excel.db')
    conn = sqlite3.connect(dbname, isolation_level=None)
    cursor = conn.cursor()

    cursor.execute("SELECT selected_id FROM option")
    rows = cursor.fetchall()

    for row in rows:
        for i in row:
            id = str(i)

    cursor.execute("SELECT * FROM settings WHERE id=?",(id))
    
    setting_rows = cursor.fetchall()
    if setting_rows:
        setting_rows = setting_rows[0][1:]

    conn.close()

    return (setting_rows)
