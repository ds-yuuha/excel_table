import wx
import os
import eel
import sqlite3
import pandas as pd

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

# wx
class FileDropTarget(wx.FileDropTarget):
    """ Drag & Drop Class """
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        self.window.trimming(filenames)

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
    
    def trimming(self,pathnames):
        trim_excel(pathnames, "個人コード", "氏　　名")
        self.Destroy()


@eel.expose
def openTrimWindow():
    app = wx.App()
    App(None, -1, 'エクセル整形')
    app.MainLoop()

@eel.expose
def trim_excel(pathnames,n,m):
    id = getCurrentId()
    path_ref = '対応表/'+ str(id) +'/ref_table.xlsx'
    
    r_df = pd.read_excel(path_ref)
    data_columnns = {}
    code_list = []

    if type(pathnames) == str:
        df = pd.read_excel(pathnames)
        for i,r in df.iterrows():
            item = r.iloc[0] #整形対象のエクセルの1列目の値
            if item in list(r_df[m]):
                match = r_df[r_df[m].str.contains(item)] #参照エクセルの列名/項目名mの中で値itemと一致する行
                for r in match.iloc[:,0:1][n]: #参照エクセル1列目の項目名nを抜き取る
                    personal_code = r
            else: personal_code = 0
            code_list.append(personal_code)
            
        data_columnns[n] = code_list

    else:
        for p in pathnames:
            df = pd.read_excel(p)
            for i,r in df.iterrows():
                item = r.iloc[0]
                if item in list(r_df[m]):
                    match = r_df[r_df[m].str.contains(item)]
                    for r in match.iloc[:,0:1][n]:
                        personal_code = r
                else: personal_code = 0
                code_list.append(personal_code)
                
            data_columnns[n] = code_list

    df_result = pd.DataFrame(data_columnns)
    df_result = pd.concat([df_result, df], axis=1)
    df_result.to_excel('抽出結果/trimmed_result.xlsx', index=False)


