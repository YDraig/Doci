#!/usr/bin/env python
#-------------------------------------------------------------------------------
# Name:        Doci
# Purpose:      Create html index for Documents
#
# Author:      Brin
#
# Created:     02/06/2012
# Copyright:   (c) Brinley Craig 2012
# Licence:     GPL V2
#-------------------------------------------------------------------------------
import sqlite3, traceback
import os, sys, wx, wx.stc
import pprint

class MyFrame(wx.Frame):
    def __init__(self, parent, title):

        self.docdir = "Test"
        self.docdb = "Doci.db"
        self.addfiles = False

        wx.Frame.__init__(self, parent, title=title)
        if os.path.splitext(sys.argv[0])[1] == ".exe":
            icon = wx.Icon(sys.argv[0], wx.BITMAP_TYPE_ICO)
        else:
            icon = wx.Icon(os.path.splitext(sys.argv[0])[0] + ".ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.SetMinSize((500,500))
        self.statusBar = self.CreateStatusBar()
        self.statusBar.SetFieldsCount(3)
        self.statusBar.SetStatusWidths([30,-1,-2])

        self.searchLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Find:')
        self.searchText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.searchButton = wx.Button(self.panel, wx.ID_ANY, 'Search')
        self.indexButton = wx.Button(self.panel, wx.ID_ANY, 'Index')
        self.openButton = wx.Button(self.panel, wx.ID_ANY, 'Open')

        self.buttonBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Controls')
        self.addButton = wx.Button(self.panel, wx.ID_ANY, 'Add')
        self.editButton = wx.Button(self.panel, wx.ID_ANY, 'Edit')
        self.updateButton = wx.Button(self.panel, wx.ID_ANY, 'Update')
        self.updateButton.Enable(False)
        self.prevButton = wx.Button(self.panel, wx.ID_ANY, 'Prev')
        self.nextButton = wx.Button(self.panel, wx.ID_ANY, 'Next')
        self.exitButton = wx.Button(self.panel, wx.ID_EXIT, 'Exit')
        
        self.idLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Id:')
        self.idText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21))
        self.idText.SetEditable(False)
        self.fileLabel = wx.StaticText(self.panel, wx.ID_ANY, 'File:')
        self.fileText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.fileText.SetEditable(False)
        self.extLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Ext:')
        self.extText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21))
        self.extText.SetEditable(False)
        
        self.dirLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Dir:')
        self.dirText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.dirText.SetEditable(False)
        self.dateLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Added:')
        self.dateText = wx.TextCtrl(self.panel, wx.ID_ANY)
        #self.dateText.SetEditable(False)
        self.dateText.Enable(False)

        self.descBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Description')
        self.descText = wx.stc.StyledTextCtrl(self.panel, wx.ID_ANY, style=wx.TE_MULTILINE|wx.TE_PROCESS_TAB ) #|wx.TE_DONTWRAP
        self.descText.SetReadOnly(True)
        self.descText.SetTabWidth(4)
        self.descText.SetUseTabs(0)
        self.descText.SetTabIndents(1)
        #self.descText.SetIndentationGuides(1)
        self.descText.SetEOLMode(wx.stc.STC_EOL_LF)

        self.searchSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.buttonBoxSizer = wx.StaticBoxSizer(self.buttonBox, wx.HORIZONTAL)
        self.buttonSizer.Add(self.buttonBoxSizer, 1, wx.ALL | wx.EXPAND, 5)
        self.propSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.dirSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.descSizer = wx.StaticBoxSizer(self.descBox, wx.HORIZONTAL)
        self.rootSizer = wx.BoxSizer(wx.VERTICAL)

        self.searchSizer.Add(self.searchLabel, 0, wx.ALL, 5)
        self.searchSizer.Add(self.searchText, 1, wx.ALL | wx.EXPAND, 5)
        self.searchSizer.Add(self.searchButton, 0, wx.ALL, 5)
        self.searchSizer.Add(self.indexButton, 0, wx.ALL, 5)
        self.searchSizer.Add(self.openButton, 0, wx.ALL, 5)

        self.buttonBoxSizer.Add(self.addButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.editButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.updateButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.prevButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.nextButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.exitButton, 0, wx.ALL, 5)
        
        self.propSizer.Add(self.idLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.idText, 0, wx.ALL, 5)
        self.propSizer.Add(self.fileLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.fileText, 1, wx.ALL | wx.EXPAND | wx.RIGHT | wx.LEFT, 5)
        self.propSizer.Add(self.extLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.extText, 0, wx.ALL, 5)
        
        self.dirSizer.Add(self.dirLabel, 0, wx.ALL, 5)
        self.dirSizer.Add(self.dirText, 2, wx.ALL | wx.EXPAND, 5)
        self.dirSizer.Add(self.dateLabel, 0, wx.ALL, 5)
        self.dirSizer.Add(self.dateText, 1, wx.ALL, 5)
        
        self.descSizer.Add(self.descText, 1, wx.ALL | wx.EXPAND | wx.RIGHT, 5)
        #self.descSizer.Add(self.descText, 1, wx.EXPAND|wx.LEFT|wx.RIGHT,5)

        self.rootSizer.Add(self.searchSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.buttonSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.propSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.dirSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.descSizer, 1, wx.EXPAND|wx.TOP|wx.BOTTOM,5)

        self.Bind(wx.EVT_BUTTON, self.onSearch, self.searchButton)
        self.Bind(wx.EVT_BUTTON, self.onAdd, self.addButton)
        self.Bind(wx.EVT_BUTTON, self.onEdit, self.editButton)
        self.Bind(wx.EVT_BUTTON, self.onUpdate, self.updateButton)
        self.Bind(wx.EVT_BUTTON, self.onPrev, self.prevButton)
        self.Bind(wx.EVT_BUTTON, self.onNext, self.nextButton)
        self.Bind(wx.EVT_BUTTON, self.onExit, self.exitButton)
        self.Bind(wx.EVT_TEXT, self.onDesc, self.descText)

        self.panel.SetSizer(self.rootSizer)
        self.rootSizer.Fit(self)

        if not os.path.isfile(self.docdb):
            # Create table
            try:
                self.con = sqlite3.connect(self.docdb)
                self.con.row_factory = sqlite3.Row
            except:
                self.onError("Failed to create Database")
            self.sql = self.con.cursor()
            self.sql.execute("create table docs (id INTEGER PRIMARY KEY, dir TEXT, file TEXT, ext TEXT, desc TEXT, date TEXT)")
            self.con.commit()
        else:
            # Load first entry
            try:
                self.con = sqlite3.connect(self.docdb)
                self.con.row_factory = sqlite3.Row
            except:
                self.onError("Failed to Open Database")
            self.sql = self.con.cursor()
            rows = self.sql.execute('select id, dir, file, ext, desc from docs order by id')
            for row in rows:
                print(row)
        

        self.Show()

    def onError(self, message):
        dlg = wx.MessageDialog(self, message="Line: " + str(sys.exc_info()[2].tb_lineno) + " - " + message + "\n\n" + traceback.format_exc(0), caption='Error', style=wx.OK|wx.CANCEL|wx.ICON_EXCLAMATION)
        result = dlg.ShowModal()
        dlg.Destroy()
        return(result)

    def onSearch(self, e):
        pass

    def onAdd(self, e):
        self.disableButtons()
        self.addfiles = True
        #filelist = [f for f in os.listdir(configdir)
        #       if os.path.isfile(os.path.join(".", f))]
        try:
            filelist = os.listdir(self.docdir)
        except:
            self.onError("Error reading Dir: " + self.docdir)
            self.addfiles = False
            self.enableButtons()
            return
        pprint.pprint(filelist)
        for file in filelist:
            filename = os.path.splitext(file)[0]
            fileext = os.path.splitext(file)[1]
            #print file
            # Insert a row of data
            # if doesnt exist insert
            self.sql.execute("""insert into docs (dir, file, ext, desc, date) values (?, ?, ?, ?, datetime())""", (self.docdir, filename, fileext,file))

        self.con.commit()
    
    def onEdit(self, e):
        self.disableButtons()
    
    def onUpdate(self, e):
        self.enableButtons()
        self.addfiles = False
        #Commit the changes after updating the Desc
        self.con.commit()
    
    def onPrev(self, e):
        pass
    
    def onNext(self, e):
        pass
    
    def onExit(self, e):
        self.con.commit()
        self.sql.close()
        self.Close(True)
        
    def onDesc(self, e):
        pass
        
    def disableButtons(self):
        self.searchButton.Enable(False)
        self.indexButton.Enable(False)
        self.updateButton.Enable(True)
        self.addButton.Enable(False)
        self.editButton.Enable(False)
        self.prevButton.Enable(False)
        self.nextButton.Enable(False)
        
    def enableButtons(self):
        self.searchButton.Enable(True)
        self.indexButton.Enable(True)
        self.updateButton.Enable(False)
        self.addButton.Enable(True)
        self.editButton.Enable(True)
        self.prevButton.Enable(True)
        self.nextButton.Enable(True)

if __name__ == '__main__':
    try:
        application = wx.PySimpleApp()
        frame = MyFrame(None, "Doci")
        application.MainLoop()
    finally:
        del application
