import os
import wx
import eel
import shutil
import sqlite3

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
def postSettingFile(file_path):
    id = getCurrentId()    
    if not os.path.isdir("対応表"):
        os.mkdir("対応表")
    path_new_dir = "対応表/" + str(id)
    if not os.path.isdir(path_new_dir):
        os.mkdir(path_new_dir)

    path_new = path_new_dir + "/ref_table.xlsx" 
    
    print(file_path[0])
    if (file_path[0]):
        shutil.copy(file_path[0], path_new)


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
        postSettingFile(pathnames)
        self.Destroy()


@eel.expose
def openSettingFileWindow():
    app = wx.App()
    App(None, -1, '設定ファイル')
    app.MainLoop()


