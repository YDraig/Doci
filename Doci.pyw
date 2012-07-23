#!/usr/bin/env python
# coding: utf-8
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
import sqlite3, traceback, hashlib
import os, sys, wx, time, datetime, random
import locale, codecs, subprocess
import ConfigParser, threading, Queue
import wx.lib.agw.multidirdialog as MDD
import wx.lib.agw.ultimatelistctrl as ULC
import wx.lib.mixins.inspection, wx.lib.inspection

class DisplayForm(wx.Frame):
    def __init__(self, parent, title):

        self.version = "0.0" # Grab real version from exe
        self.docini = "Doci.ini"
        self.docdb = ""
        self.dochtml = ""
        self.docdir = [os.getcwdu()]
        self.selectlimit = 0
        self.encoding = locale.getdefaultlocale()[1]
        self.addfiles = False
        self.addid = []
        self.results = []
        self.grey = wx.NamedColour("GREY")
        self.black = wx.NamedColour("BLACK")
        self.workerRun = threading.Event()
        self.workerAbort = threading.Event()
        self.workerDir = Queue.Queue(maxsize=0)
        self.messageQueue = Queue.Queue(maxsize=0)
        self.messageCount = 0
        self.maxidQueue = Queue.Queue(maxsize=1)
        self.addidQueue = Queue.Queue(maxsize=1)
        self.thread = None
        self.statusBarClear = None

        # Check that pywin32 is installed
        if os.name == "nt":
            try:
                from win32api import LoadResource
            except:
                wx.MessageBox(message="Missing dependancy pywin32\nhttp://sourceforge.net/projects/pywin32/", caption='Info', style=wx.OK|wx.ICON_INFORMATION)
                quit()

        # Get the Icon and Version string from the exe resources
        if os.path.splitext(sys.argv[0])[1] == ".exe":
            icon = wx.Icon(sys.argv[0], wx.BITMAP_TYPE_ICO)
            self.version = LoadResource(0, u'VERSION', 1)
        else:
            icon = wx.Icon("Doci.ico", wx.BITMAP_TYPE_ICO)

        wx.Frame.__init__(self, parent, title=title + " - v" + self.version)
        self.SetIcon(icon)
        wx.EVT_CLOSE(self, self.onClose)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.SetMinSize((500,500))

        self.statusBar = self.CreateStatusBar()
        self.statusBar.SetFieldsCount(3)
        self.statusBar.SetStatusWidths([-1,85,60])
        self.statusBar.SetStatusText("0/0", 1) # Initalise the index/records field
        self.messageTimer = wx.Timer(self)

        self.menuBar = wx.MenuBar()
        self.fileMenu = wx.Menu()
        self.menuBar.Append(self.fileMenu, "&File")
        self.fileOpen = self.fileMenu.Append(wx.ID_OPEN, '&Open', 'Open current file')
        self.fileIndex = self.fileMenu.Append(wx.ID_INDEX, '&Index', 'Create HTML Index')
        self.fileMenu.AppendSeparator()
        self.fileScan = self.fileMenu.Append(wx.ID_FIND, '&Scan', 'Scan Directories')
        self.fileMenu.AppendSeparator()
        self.fileExit = wx.MenuItem(self.fileMenu, wx.ID_EXIT, '&Exit\tCtrl+Q')
        self.fileMenu.AppendItem(self.fileExit)
        self.editMenu = wx.Menu()
        self.menuBar.Append(self.editMenu, "&Edit")
        self.editDirectories = self.editMenu.Append(wx.ID_ADD, '&Directories', 'Directories to scan')
        self.editCategories = self.editMenu.Append(wx.ID_EDIT, '&Categories', 'Categories and colors')
        self.editSettings = self.editMenu.Append(wx.ID_SETUP, '&Settings', 'Edit Settings')
        self.helpMenu = wx.Menu()
        self.menuBar.Append(self.helpMenu, "&Help")
        self.helpDebug = self.helpMenu.Append(wx.ID_CONTEXT_HELP, '&Debug', 'Widget Inspection Toool for Debugging')
        self.helpAbout = self.helpMenu.Append(wx.ID_ABOUT, '&About')

        self.SetMenuBar(self.menuBar)

        self.searchLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Find:')
        self.searchText = wx.TextCtrl(self.panel, wx.ID_ANY, style=wx.TE_PROCESS_ENTER)
        self.searchButton = wx.Button(self.panel, wx.ID_ANY, 'Search')
        self.indexButton = wx.Button(self.panel, wx.ID_ANY, 'Index')

        self.buttonBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Controls')
        self.scanButton = wx.Button(self.panel, wx.ID_ANY, 'Scan')
        self.editButton = wx.Button(self.panel, wx.ID_ANY, 'Edit')
        self.updateButton = wx.Button(self.panel, wx.ID_ANY, 'Update')
        self.updateButton.Enable(False)
        self.prevButton = wx.Button(self.panel, wx.ID_ANY, 'Prev')
        self.nextButton = wx.Button(self.panel, wx.ID_ANY, 'Next')
        self.openButton = wx.Button(self.panel, wx.ID_ANY, 'Open')

        self.idLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Id:', size=(30,20))
        self.idText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21), style=wx.TE_PROCESS_ENTER)
        self.idText.SetEditable(True)
        self.extLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Ext:')
        self.extText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21))
        self.extText.SetEditable(False)
        self.extText.SetForegroundColour(self.grey)
        self.dateLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Date:')
        self.dateText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(120,21))
        self.dateText.SetEditable(False)
        self.dateText.SetForegroundColour(self.grey)
        self.categoryChoice = wx.Choice( self.panel, wx.ID_ANY, style=0)
        self.categoryChoice.Enable(False)

        self.fileLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Name:', size=(30,20))
        self.fileText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.fileText.SetEditable(False)
        self.fileText.SetForegroundColour(self.grey)
        self.dirLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Dir:', size=(30,20))
        self.dirText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.dirText.SetEditable(False)
        self.dirText.SetForegroundColour(self.grey)

        self.descBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Description')
        self.descText = wx.TextCtrl(self.panel, wx.ID_ANY, style=wx.TE_MULTILINE ) #|wx.TE_PROCESS_TAB
        self.descText.SetForegroundColour(self.grey)
        self.descText.SetEditable(False)

        self.searchSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.buttonBoxSizer = wx.StaticBoxSizer(self.buttonBox, wx.HORIZONTAL)
        self.buttonSizer.Add(self.buttonBoxSizer, 1, wx.ALL | wx.EXPAND, 5)
        self.propSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.nameSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.dirSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.descSizer = wx.StaticBoxSizer(self.descBox, wx.HORIZONTAL)
        self.rootSizer = wx.BoxSizer(wx.VERTICAL)

        self.searchSizer.Add(self.searchLabel, 0, wx.ALL, 5)
        self.searchSizer.Add(self.searchText, 1, wx.ALL | wx.EXPAND, 5)
        self.searchSizer.Add(self.searchButton, 0, wx.ALL, 5)
        self.searchSizer.Add(self.indexButton, 0, wx.ALL, 5)

        self.buttonBoxSizer.Add(self.scanButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.editButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.updateButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.prevButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.nextButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.openButton, 0, wx.ALL, 5)

        self.propSizer.Add(self.idLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.idText, 0, wx.ALL, 5)
        self.propSizer.Add(self.extLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.extText, 0, wx.ALL, 5)
        self.propSizer.Add(self.dateLabel, 0, wx.ALL, 5)
        self.propSizer.Add(self.dateText, 0, wx.ALL, 5)
        self.propSizer.Add(self.categoryChoice, 1, wx.ALL, 5)

        self.nameSizer.Add(self.fileLabel, 0, wx.ALL, 5)
        self.nameSizer.Add(self.fileText, 1, wx.ALL | wx.EXPAND, 5)

        self.dirSizer.Add(self.dirLabel, 0, wx.ALL, 5)
        self.dirSizer.Add(self.dirText, 1, wx.ALL | wx.EXPAND, 5)

        self.descSizer.Add(self.descText, 1, wx.ALL | wx.EXPAND | wx.RIGHT, 5)

        self.rootSizer.Add(self.searchSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.propSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.nameSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.dirSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.descSizer, 1, wx.EXPAND|wx.TOP|wx.BOTTOM,5)
        self.rootSizer.Add(self.buttonSizer, 0, wx.ALL | wx.EXPAND, 5)

        self.Bind(wx.EVT_BUTTON, self.onSearch, self.searchButton)
        self.Bind(wx.EVT_BUTTON, self.onIndex, self.indexButton)
        self.Bind(wx.EVT_BUTTON, self.onOpen, self.openButton)
        self.Bind(wx.EVT_BUTTON, self.onScan, self.scanButton)
        self.Bind(wx.EVT_BUTTON, self.onEdit, self.editButton)
        self.Bind(wx.EVT_BUTTON, self.onUpdate, self.updateButton)
        self.Bind(wx.EVT_BUTTON, self.onPrev, self.prevButton)
        self.Bind(wx.EVT_BUTTON, self.onNext, self.nextButton)
        self.Bind(wx.EVT_TEXT_ENTER, self.onSearch, self.searchText)
        self.Bind(wx.EVT_TEXT_ENTER, self.onId, self.idText)
        self.Bind(wx.EVT_TIMER, self.setMessage, self.messageTimer)
        self.Bind(wx.EVT_MENU, self.onOpen, self.fileOpen)
        self.Bind(wx.EVT_MENU, self.onIndex, self.fileIndex)
        self.Bind(wx.EVT_MENU, self.onScan, self.fileScan)
        self.Bind(wx.EVT_MENU, self.onChangeDir, self.editDirectories)
        self.Bind(wx.EVT_MENU, self.onEditCategories, self.editCategories)
        self.Bind(wx.EVT_MENU, self.onEditSettings, self.editSettings)
        self.Bind(wx.EVT_MENU, self.onExit, self.fileExit)
        self.Bind(wx.EVT_MENU, self.onDebug, self.helpDebug)
        self.Bind(wx.EVT_MENU, self.onAbout, self.helpAbout)

        self.panel.SetSizer(self.rootSizer)
        self.rootSizer.Fit(self)

        self.messageTimer.Start(500)
        self.getIni()
        self.openDB()
        self.getCategories()
        self.setMaxid()
        self.searchRecords()
        self.displayRecord(1)
        self.Show()

        print "***Form Init*** (%s)" % self.encoding

    def openDB(self):
        # Create DB file if it doesnt exist
        if not os.path.isfile(self.docdb):
            print "Creating DB file"
            try:
                self.con = sqlite3.connect(self.docdb)
                self.con.row_factory = sqlite3.Row
                self.con.text_factory = unicode # Allow unicode conversion
            except:
                self.displayMessage("Failed to create Database")
            self.sql = self.con.cursor()
            self.sql.execute("""CREATE TABLE docs (id INTEGER PRIMARY KEY, dir TEXT, name TEXT, ext TEXT, desc TEXT, hash TEXT, size TEXT, date REAL,
                            categoriesid INTEGER, seen INTEGER, added TEXT, FOREIGN KEY(categoriesid) REFERENCES categories(id))""")
            self.sql.execute("""CREATE TABLE dupes (id INTEGER PRIMARY KEY, dir TEXT, name TEXT, ext TEXT, desc TEXT, hash TEXT, size TEXT, date REAL, seen INTEGER, added TEXT,
                            docsid INTEGER, FOREIGN KEY(docsid) REFERENCES docs(id), UNIQUE(dir, name, ext))""")
            self.sql.execute("CREATE TABLE categories (id INTEGER PRIMARY KEY, category TEXT, color TEXT, font TEXT)")
            for (cat,color, font) in (("Standards", "#fff0f0", "#000000"), ("Drawings", "#ecfbd4", "#000000")):
                self.sql.execute("INSERT INTO categories (category, color, font) VALUES (?, ?, ?)", (cat, color, font))
            self.sql.execute("CREATE TABLE colors (id INTEGER PRIMARY KEY, tag TEXT, color TEXT, font TEXT, UNIQUE(tag))")
            for (tag,color, font) in (("Header", "#328aa4", "#ffffff"), ("Default", "#e5f1f4", "#000000"), ("Highlight", "#ffffb2", "#000000")):
                self.sql.execute("INSERT INTO colors (tag, color, font) VALUES (?, ?, ?)", (tag, color, font))
            self.sql.execute("CREATE INDEX tag on colors (tag)")
            self.sql.execute("CREATE INDEX hash on docs (hash,size)")
            self.sql.execute("CREATE INDEX filename on docs (dir,name,ext)")
            # Check for FTS dependancy
            try:
                self.sql.execute("CREATE VIRTUAL TABLE search USING fts4(content='docs', name, desc)")
            except sqlite3.OperationalError:
                self.displayMessage("Missing dependancy sqlite (dll) > v3.7.10\nhttp://www.sqlite.org/download.html")
                self.onClose(None)
                quit()
            self.sql.execute("CREATE TRIGGER docs_bupdate BEFORE UPDATE ON docs BEGIN DELETE FROM search WHERE docid=old.rowid; END")
            self.sql.execute("CREATE TRIGGER docs_bdelete BEFORE DELETE ON docs BEGIN DELETE FROM search WHERE docid=old.rowid; END")
            self.sql.execute("CREATE TRIGGER docs_aupdate AFTER UPDATE ON docs BEGIN INSERT INTO search(docid, name, desc) VALUES(new.rowid, new.name, new.desc); END")
            self.sql.execute("CREATE TRIGGER docs_ainsert AFTER INSERT ON docs BEGIN INSERT INTO search(docid, name, desc) VALUES(new.rowid, new.name, new.desc); END")
            self.con.commit()
        else:
            # Open DB
            try:
                self.con = sqlite3.connect(self.docdb)
                self.con.row_factory = sqlite3.Row
                self.con.text_factory = unicode # Allow unicode conversion
            except:
                self.displayMessage("Failed to Open Database")
                quit()
            self.sql = self.con.cursor()
            # Check for FTS dependancy
            try:
                self.sql.execute("select count(*) from search")
            except sqlite3.OperationalError:
                self.displayMessage("Missing dependancy sqlite (dll) > v3.7.10\nhttp://www.sqlite.org/download.html")
                self.onClose(None)
                quit()

    def closeDB(self):
        if isinstance(self.con, sqlite3.Connection):
            self.con.commit()
            self.sql.close()

    def getIni(self):
        # Create ini file if it doesnt exist
        path = {'dirs':self.docdir, 'db':'Doci.db', 'html':'Doci.html'}
        settings = {'select':'30000'}

        if not os.path.isfile(self.docini):
            print "Creating ini file"
            currentdir = os.getcwdu()
            #http://xoomer.virgilio.it/infinity77/main/freeware.html
            dlg = MDD.MultiDirDialog(self, message=u'Choose directory(s) to scan', title=u'Browse For Folders', defaultPath=currentdir, agwStyle=MDD.DD_DIR_MUST_EXIST|MDD.DD_MULTIPLE, name='multidirdialog')
            if dlg.ShowModal() == wx.ID_OK:
                self.docdir = dlg.GetPaths()
                path['dirs'] = self.docdir
                print self.docdir
            else:
                self.displayMessage("Unable to continue without ini file", "Info")
                self.Destroy()
                return
            os.chdir(currentdir)

            self.config = ConfigParser.SafeConfigParser()
            self.config.add_section('path')
            for option in path:
                self.config.set('path', option, str(path[option]))
            self.config.add_section('settings')
            for option in settings:
                self.config.set('settings', option, str(settings[option]))

            try:
                with open(self.docini, 'w') as configfile:
                    self.config.write(configfile)
            except:
                self.displayMessage("Error Creating ini file")
                self.onExit(self)

        #defaults = dict(path.items() + limit.items())
        #self.config = ConfigParser.SafeConfigParser(defaults)
        self.config = ConfigParser.SafeConfigParser()
        try:
            self.config.read(self.docini)
            self.docdir = eval(self.config.get('path', 'dirs'))
            self.docdb = self.config.get('path', 'db')
            self.dochtml = self.config.get('path', 'html')
            self.selectlimit = int(self.config.get('settings', 'select'))
            try:
                self.encoding = self.config.get('settings', 'encoding')
            except ConfigParser.NoOptionError:
                pass
        except:
            self.displayMessage("Missing ini file")
            self.onExit(self)

    def onChangeDir(self, event):
        currentdir = os.getcwdu()
        if os.path.exists(self.docdir[0]):
            defaultdir = self.docdir[0]
        else:
            defaultdir = currentdir
        addorreplace = self.displayMessage("Add to exiting Directory list?", status="Query")
        dlg = MDD.MultiDirDialog(self, message=u'Choose directory(s) to scan', title=u'Browse For Folders', defaultPath=defaultdir, agwStyle=MDD.DD_DIR_MUST_EXIST|MDD.DD_MULTIPLE, name='multidirdialog')
        if dlg.ShowModal() == wx.ID_OK:
            if addorreplace == wx.ID_NO:
                self.docdir = []
            self.docdir.extend(dlg.GetPaths())
            print self.docdir
            os.chdir(currentdir)
            self.config.set('path', 'dirs', str(self.docdir))
            try:
                with open(self.docini, 'w') as configfile:
                    self.config.write(configfile)
            except:
                self.displayMessage("Error Saving ini file")
        else:
            return

    def getCategories(self):
        self.categoryChoice.Clear()
        self.categoryChoice.Append("")
        categories = self.sql.execute("select category from categories order by id")
        for category in categories.fetchall():
            self.categoryChoice.Append(category["category"])
        self.categoryChoice.SetSelection(0)

    def setIndex(self, index):
        (oldindex, results) = self.statusBar.GetStatusText(1).split("/")
        self.statusBar.SetStatusText(str(index) + "/" + str(results),1)

    def getIndex(self):
        if self.statusBar.GetStatusText(1) != 'None':
            (index, results) = self.statusBar.GetStatusText(1).split("/")
            return int(index)

    def getId(self, index=None):
        if not index:
            index = self.getIndex()
        if index and self.results:
            return self.results[index - 1]

    def setResults(self, results):
        (index, oldresults) = self.statusBar.GetStatusText(1).split("/")
        self.statusBar.SetStatusText(str(index) + "/" + str(results),1)

    def getResults(self):
        if self.statusBar.GetStatusText(1) != 'None':
            (index, results) = self.statusBar.GetStatusText(1).split("/")
            return int(results)

    def setMaxid(self, maxid=""):
        if maxid == "":
            maxid = self.sql.execute('select max(id) from docs').fetchone()[0]
        if maxid:
            self.statusBar.SetStatusText(str(maxid),2)

    def setMessage(self, event):
        message = None
        self.messageCount += 1
        queuedepth = self.messageQueue.qsize()
        try:
            message = self.messageQueue.get_nowait()
        except Queue.Empty:
            if self.messageCount > 10:
                message = u""
        try:
            if message != None:
                self.messageCount = 0
                if len(message):
                    message = str(queuedepth) + ": " + message
                self.statusBar.SetStatusText(message,0)
        except:
            self.displayMessage("Failure on messagebar")

    def displayMessage(self, message, status="Error"):
        if status == "Info":
            dlg = wx.MessageDialog(self, message=message, caption='Info', style=wx.OK|wx.ICON_INFORMATION)
        elif status == "Query":
            dlg = wx.MessageDialog(self, message=message, caption='Query', style=wx.YES_NO|wx.ICON_QUESTION)
        elif status == "Duplicate":
            dlg = wx.MessageDialog(self, message=message, caption='Delete Duplicate File?', style=wx.YES_NO|wx.CANCEL|wx.ICON_EXCLAMATION)
        elif status == "Warning":
            self.messageQueue.put_nowait(message)
            print "Line: " + str(sys.exc_info()[2].tb_lineno) + " - " + message
            return
        else:
            self.messageQueue.put_nowait(message)
            dlgMessage = "Line: " + str(sys.exc_info()[2].tb_lineno) + " - " + message + "\n\n" + traceback.format_exc(0)
            print dlgMessage
            dlg = wx.MessageDialog(self, message=dlgMessage, caption='Error', style=wx.OK|wx.CANCEL|wx.ICON_ERROR)
        result = dlg.ShowModal()
        dlg.Destroy()
        return(result)

    def searchRecords(self, search=""):
        #self.sql.execute('select docs.id as id, dir, name, ext, desc, date, category, color from docs left join categories on docs.categoriesid=categories.id where docs.id=?', (self.getId(index),))
        if search != "":
            self.results = [element[0] for element in self.sql.execute('select docid from search where search match ?', (search,)).fetchall()]
        else:
            self.results = [element[0] for element in self.sql.execute('select id from docs').fetchall()]
        self.setResults(len(self.results))

    def displayRecord(self, index=None):
        if self.results:
            if not index:
                index=1
            self.sql.execute('select id, dir, name, ext, desc, date, categoriesid from docs where id=?', (self.getId(index),))
            row = self.sql.fetchone()
            self.setIndex(index)
            self.categoryChoice.SetSelection(row["categoriesid"] if row["categoriesid"] else 0)
            self.idText.SetValue(str(row["id"]))
            self.fileText.SetValue(row["name"])
            self.extText.SetValue(row["ext"])
            self.dirText.SetValue(row["dir"])
            self.dateText.SetValue(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(row["date"])))
            self.descText.SetValue(row["desc"])

    def onPrev(self, event):
        if self.getIndex():
            prev = self.getIndex() - 1
            if prev < 1:
                self.displayMessage("Reached start, wraping to end", status="Info")
                prev = self.getResults()
            self.displayRecord(prev)

    def onNext(self, event):
        if self.getIndex():
            next = self.getIndex() + 1
            if next > self.getResults():
                self.displayMessage("Reached end, wraping to beginning", status="Info")
                next = 1
            self.displayRecord(next)

    def onSearch(self, event):
        self.searchRecords(self.searchText.GetValue())
        self.displayRecord()

    def onIndex(self, event):
        #http://www.tablefixedheader.com/fullpagedemo/
        #http://vikaskhera.wordpress.com/2008/11/06/5-easy-steps-to-create-a-fixed-header/
        if os.path.exists(self.dochtml):
            if self.displayMessage("File %s Already Exists!\nOverwite file?" % self.dochtml, "Query") == wx.ID_NO:
                return
        try:
            htmlfile = codecs.open(self.dochtml, 'w', self.encoding)
        except:
            self.displayMessage("Error Creating html file %s" % self.dochtml)
            
        # HTML Headers
        self.messageQueue.put("Creating Html")
        htmlfile.write(u'<!DOCTYPE html>\n<html><head><title>Doci</title>\n')
        htmlfile.write(u"""<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"> </script>
            <script src="jquery.fixedheader.js"> </script>
            <script type="text/javascript">
            $(document).ready(function(){ $("#data").fixedHeader({ width: 770,height: 500 }); })
            </script>\n""")
        htmlfile.write(u'<style type="text/css">\n')
        
        # Get Table Colors
        header = self.sql.execute('select color from colors where tag="Header"').fetchone()[0]
        headerfont = self.sql.execute('select font from colors where tag="Header"').fetchone()[0]
        defaultcolor = self.sql.execute('select color from colors where tag="Default"').fetchone()[0]
        defaultfont = self.sql.execute('select font from colors where tag="Default"').fetchone()[0]
        highlight = self.sql.execute('select color from colors where tag="Highlight"').fetchone()[0]
        highlightfont = self.sql.execute('select font from colors where tag="Highlight"').fetchone()[0]
        htmlfile.write(""".dataTable { font-family:Verdana, Arial, Helvetica, sans-serif; border-collapse: collapse; border:1px solid #999999; width: 750px; font-size:12px;}
            .dataTable td, .dataTable th {border: 1px solid #999999; padding: 3px 5px; margin:0px;}
            .dataTable thead th, td{text-align:left; padding:.5em; border:1px solid #fff;}
            .dataTable thead a:hover { text-decoration: underline;}\n""")
        htmlfile.write(u"th, tfoot td{background:%s; color:%s;}\ntd{background:%s; color:%s;}\n" % (header, headerfont, defaultcolor, defaultfont))
        htmlfile.write(u"a:link {color:%s;}\n" % defaultfont)
        
        # Create Category classes
        self.sql.execute("select category, color, font from categories")
        categories = self.sql.fetchall()
        for category in categories:
            htmlfile.write(u'tr.%s td{background:%s; color:%s;}\n' % (category["category"], category["color"], category["font"]))
        htmlfile.write(u'tr:hover>td{background:%s; color:%s;}\n' % (highlight, highlightfont))
        htmlfile.write(u'\n</style><body>\n')

        # Create Table
        self.sql.execute("select docs.id as id, dir, name, ext, desc, date, size, category, color from docs left join categories on docs.categoriesid=categories.id limit 0,?", (self.selectlimit,))
        docs = self.sql.fetchall()
        #htmlfile.write(u'<div align=center><h1>Document Indexer</h1></div>\n')
        htmlfile.write(u'<table cellspacing="0" cellpadding="0" id="data" class="dataTable">\n')
        htmlfile.write(u'<thead><th>Id</th><th>File</th><th>Category</th><th>Date</th><th>Size</th></thead>\n')
        if len(docs) == self.selectlimit:
            htmlfile.write(u'<tfoot><tr><td>%s</td><td colspan="4">Max Limit Reached!</td></tr></tfoot>\n' % self.selectlimit)
        htmlfile.write(u'<tbody>\n')
        for doc in docs:
            filename = os.path.join(doc["dir"], doc["name"]+doc["ext"])
            date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(doc["date"]))
            size = self.sizeof(float(doc["size"]))
            htmlfile.write(u'<tr class="%s"><td>%s</td>' % (doc["category"], doc["id"]))
            try:
                htmlfile.write(u'<td><a href="%s" title="%s" target="_blank">%s</a></td>' % (filename, filename, doc["desc"]))
            except:
                self.messageQueue.put("Error converting: %s" % repr(filename))
            htmlfile.write(u'<td>%s</td><td>%s</td><td>%s</td></tr>\n' % (doc["category"], date, size))
        htmlfile.write(u"</tbody></table></body></html>\n")
        htmlfile.close()
        self.messageQueue.put("Opening Html File")
        startfile(self.dochtml)

    def sizeof(self, num):
        for x in ['bytes','KB','MB','GB']:
            if num < 1024.0:
                return "%3.1f%s" % (num, x)
            num /= 1024.0
        return "%3.1f%s" % (num, 'TB')

    def onOpen(self, event):
        filename = os.path.join(self.dirText.GetValue(), self.fileText.GetValue() + self.extText.GetValue())
        self.messageQueue.put_nowait(u"Opening file: " + self.fileText.GetValue())
        startfile(filename)

    def onScan(self, event):
        self.disableButtons(everything=True)
        self.messageQueue.put_nowait(u"Started File Check")
        self.closeDB() # Best to be thread safe
        self.thread = startThread(self.addFiles, self.docdir)
        self.updateProgress()
        self.openDB()

        self.messageQueue.put_nowait(u"Check Duplicates")
        dupes = self.sql.execute("select max(id) from dupes").fetchone()[0]
        if dupes:
            if self.displayMessage("Found %s Duplicate Files, Ignore them?" % dupes, status="Query") == wx.ID_YES:
                self.sql.execute("delete from dupes")
                self.con.commit()
            else:
                self.removeFiles()
        if not self.workerAbort.isSet():
            self.messageQueue.put_nowait(u"Checking for Deleted Files")
            missing = self.sql.execute("select count(*) from docs where seen <> 1").fetchone()[0]
            if missing:
                if self.displayMessage("Found %s missing files\nPurge from database?" % missing, status="Query") == wx.ID_YES:
                        self.sql.execute("delete from docs where seen <>1")
        self.messageQueue.put_nowait(u"Updating Results")
        self.searchRecords() # Would be nice to do this later, but needs the results for below
        if self.addid and self.results:
            self.displayRecord(self.results.index(self.addid.pop(0))+1) # Index of 0 based array
            if self.displayMessage("Found %s new files\nBulk update files?" % (len(self.addid) + 1), status="Query") == wx.ID_YES:
                self.addfiles = True
            else:
                self.addfiles = False
                self.addid = []
                self.enableButtons()
        else:
            self.enableButtons()
        self.messageQueue.put_nowait(u"Cleaning up")
        self.sql.execute("update docs set seen=''")
        self.con.commit()
        self.sql.execute("INSERT INTO search(search) VALUES('optimize')")

    def removeFiles(self):
        self.sql.execute("""select dupes.id did, dupes.dir ddir, dupes.name dname, dupes.ext dext, dupes.docsid oid, docs.dir odir, docs.name oname, docs.ext oext
                            from dupes,docs where dupes.docsid=docs.id""")
        for dupe in self.sql.fetchall():
            message = "YES to Delete:\n" + dupe["ddir"] + "\n" + dupe["dname"] + dupe["dext"] + "\n\n"
            message += "NO to Delete:\n" + dupe["odir"] + "\n" + dupe["oname"] + dupe["oext"] + "\n\n"
            message += "***WARNING: This will PERMANENTLY delete the file from the PC!***"
            removeFile = self.displayMessage(message , "Duplicate")
            if removeFile == wx.ID_YES:
                filename = os.path.join(dupe["ddir"], dupe["dname"] + dupe["dext"])
                self.messageQueue.put_nowait(u"Deleting File: " + filename)
                try:
                    os.remove(filename)
                except:
                    self.displayMessage("Error removing file: " + filename)
                finally:
                    self.sql.execute("delete from dupes where id=?", (dupe["did"],))
            elif removeFile == wx.ID_NO:
                filename = os.path.join(dupe["odir"], dupe["oname"] + dupe["oext"])
                self.messageQueue.put_nowait(u"Deleting File: " + filename)
                try:
                    os.remove(filename)
                    self.sql.execute("update docs set dir=?, name=?, ext=? where id=?", (dupe["ddir"], dupe["dname"], dupe["dext"], dupe["oid"]))
                except:
                    self.displayMessage("Error removing file: " + filename)
                finally:
                    self.sql.execute("delete from dupes where id=?", (dupe["did"],))
            else:
                self.messageQueue.put_nowait(u"Ignoring Duplicates")
                self.sql.execute("delete from dupes")
                self.con.commit()
                return
        self.con.commit()

    def updateProgress(self):
        self.progress = wx.ProgressDialog('Checking Files', 'Please wait...',style=wx.PD_CAN_ABORT|wx.PD_ELAPSED_TIME)
        #No longer Modal so we can run in the main GUI thread :)~
        self.progress.Show()
        while self.thread.isAlive():
            try:
                maxId = self.maxidQueue.get_nowait()
                self.setMaxid(maxId)
                self.maxidQueue.task_done()
            except Queue.Empty:
                pass
            try:
                currentDir = self.workerDir.get_nowait()
                if len(currentDir) > 30:
                    currentDir = currentDir[:10] + "..." + currentDir[-20:]
                (run, skip) = self.progress.Pulse(currentDir)
                self.workerDir.task_done()
            except Queue.Empty:
                (run, skip) = self.progress.Pulse()
            if run == False:
                if self.workerRun.isSet():
                    self.workerRun.clear()
                    self.workerAbort.set()
                    self.messageQueue.queue.clear()
                    self.messageQueue.put_nowait(u"Aborted File Check")
                else:
                    print "Aborting..."
            time.sleep(0.2)
        else:
            self.workerDir.queue.clear()
        try:
            self.addid = self.addidQueue.get_nowait()
            self.addidQueue.task_done()
        except Queue.Empty:
            pass
        self.progress.Destroy()

    def chunkReader(self, fobj, chunk_size=8192):
        """Generator that reads a file in chunks of bytes"""
        while True:
            chunk = fobj.read(chunk_size)
            if not chunk:
                return
            yield chunk

    def addFiles(self, paths, hash=hashlib.sha1):
        self.workerRun.set()
        self.workerAbort.clear()
        addid=[]
        # Open DB (thread safe)
        self.openDB()
        for path in paths:
            if not self.workerRun.isSet():
                break
            if not os.path.isdir(path):
                self.displayMessage("Missing Directory: %s" % path, status="Info")
                continue
            for dirpath, dirnames, filenames in os.walk(path):
                self.workerDir.put_nowait(dirpath)
                if not self.workerRun.isSet():
                    break
                for file in filenames:
                    filepath = os.path.join(dirpath, file)
                    filebasename = os.path.splitext(file)[0]
                    fileext = os.path.splitext(file)[1]
                    filedate = os.path.getctime(filepath)
                    hashobj = hash()
                    for chunk in self.chunkReader(open(filepath, 'rb')):
                        hashobj.update(chunk)
                        if not self.workerRun.isSet():
                            break
                    filehash = hashobj.hexdigest()
                    filesize = os.path.getsize(filepath)
                    self.sql.execute('select id, dir, name, ext from docs where hash=? and size=?', (filehash,filesize))
                    duplicate = self.sql.fetchone()
                    if duplicate:
                        duplicateFile = os.path.join(duplicate["dir"], duplicate["name"] + duplicate["ext"])
                        if filepath != duplicateFile:
                            if not os.path.isfile(duplicateFile):
                                # File has moved update record
                                self.messageQueue.put_nowait(u"File Moved: %s" % filepath)
                                self.messageQueue.put_nowait(u"Old Path: %s" % duplicateFile)
                                try:
                                    self.sql.execute("update docs set dir=?,name=?,ext=? where id=?", (dirpath, filebasename, fileext, duplicate["id"]))
                                except:
                                    if self.displayMessage("Failed to update file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                                        self.workerRun.clear()
                                        break
                            else:
                                # Duplicate File
                                self.messageQueue.put_nowait(u"Found Duplicate: %s" % filepath)
                                self.messageQueue.put_nowait(u"Clashes with: %s" % duplicateFile)
                                try:
                                    self.sql.execute("insert into dupes (dir, name, ext, desc, hash, size, date, seen, added, docsid) values (?, ?, ?, ?, ?, ?, ?, 1, datetime(),?)",
                                                    (dirpath, filebasename, fileext, filebasename, filehash, filesize, filedate, duplicate["id"]))
                                except:
                                    if self.displayMessage("Failed to mark duplicate file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                                        self.workerRun.clear()
                                        break
                            continue
                        else:
                            # Found Existing File, update the seen value
                            self.sql.execute("update docs set seen=1 where id=?", (duplicate["id"],))
                            continue
                    try:
                        # New File, add to DB
                        self.sql.execute("insert into docs (dir, name, ext, desc, hash, size, date, seen, added) values (?, ?, ?, ?, ?, ?, ?, 1, datetime())",
                                        (dirpath, filebasename, fileext, filebasename, filehash, filesize, filedate))
                        addid.append(self.sql.lastrowid)
                        try:
                            self.maxidQueue.put_nowait(self.sql.lastrowid)
                        except Queue.Full:
                            pass
                    except:
                        if self.displayMessage("Failed to add file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                            self.workerRun.clear()
                            break
        self.closeDB()
        self.workerDir.queue.clear()
        self.messageQueue.queue.clear()
        self.workerRun.clear()
        self.addidQueue.put_nowait(addid)

    def onEdit(self, event):
        if self.editButton.GetLabel() == "Cancel":
            self.displayRecord(self.getId())
            self.enableButtons(True)
        else:
            self.disableButtons(cancel=True)

    def onUpdate(self, event):
        #Commit the changes after updating the Desc
        self.sql.execute("update docs set desc=?, categoriesid=? where id=?", (self.descText.GetValue(), self.categoryChoice.GetSelection(), self.getId()))
        self.con.commit()
        if self.addfiles:
            newid = self.addid.pop(0)
            if not self.addid:
                self.enableButtons()
                self.addfiles = False
            else:
                self.displayRecord(newid)
        else:
            self.enableButtons(True)
            self.addfiles = False

    def onExit(self, event):
        self.Close(True)

    def onClose(self, event):
        print "Closing..."
        self.closeDB()
        self.messageTimer.Stop()
        self.Destroy()

    def onId(self, event):
        id = self.idText.GetValue()
        if id.isdigit():
            id = int(id)
            if id in self.results:
                self.displayRecord(self.results.index(id + 1)) # Allow for 0 indexed array

    def disableButtons(self, cancel=False, everything=False):
        self.searchButton.Enable(False)
        self.menuBar.EnableTop(0,False)
        self.menuBar.EnableTop(1,False)
        self.indexButton.Enable(False)
        self.openButton.Enable(False)
        self.idText.SetEditable(False)
        self.idText.SetForegroundColour(self.grey)
        self.scanButton.Enable(False)
        self.prevButton.Enable(False)
        self.nextButton.Enable(False)
        if everything:
            self.descText.SetForegroundColour(self.grey)
            self.descText.SetEditable(False)
            self.updateButton.Enable(False)
            self.categoryChoice.Enable(False)
        else:
            self.descText.SetForegroundColour(self.black)
            self.descText.SetEditable(True)
            self.updateButton.Enable(True)
            self.categoryChoice.Enable(True)
        if cancel:
            self.editButton.SetLabel("Cancel")
        else:
            self.editButton.Enable(False)

    def enableButtons(self, cancel=False):
        self.menuBar.EnableTop(0,True)
        self.menuBar.EnableTop(1,True)
        self.searchButton.Enable(True)
        self.indexButton.Enable(True)
        self.openButton.Enable(True)
        self.idText.SetEditable(True)
        self.idText.SetForegroundColour(self.black)
        self.updateButton.Enable(False)
        self.scanButton.Enable(True)
        self.prevButton.Enable(True)
        self.nextButton.Enable(True)
        self.descText.SetForegroundColour(self.grey)
        self.descText.SetEditable(False)
        self.categoryChoice.Enable(False)
        if cancel:
            self.editButton.SetLabel("Edit")
        else:
            self.editButton.Enable(True)

    def onDebug(self, event):
        wx.lib.inspection.InspectionTool().Show()

    def onAbout(self, event):
        self.displayMessage("       Doci - v%s\nDocument indexer\n\n             by\n     Brinley Craig\n       Phill Slater" % str(self.version), "Info")

    def onEditCategories(self, event):
        changeCategories = EditCategories(self)
        changeCategories.ShowModal()
        changeCategories.Destroy()

    def onEditSettings(self, event):
        changeLimits = EditSettings(self)
        changeLimits.ShowModal()
        changeLimits.Destroy()

class EditCategories(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, wx.ID_ANY, 'Edit Categories')
        wx.EVT_CLOSE(self, self.onClose)

        panel = wx.Panel(self)
        rootsizer = wx.BoxSizer(wx.VERTICAL)
        self.lastrow = 0
        self.currentItem = 0
        self.currentId = 0
        self.currentList = ""

        #http://xoomer.virgilio.it/infinity77/AGW_Docs/ultimatelistctrl_module.html
        self.colorList = ULC.UltimateListCtrl(panel, wx.ID_ANY, agwStyle=ULC.ULC_REPORT | ULC.ULC_HRULES | ULC.ULC_NO_HIGHLIGHT |
                                                 ULC.ULC_HAS_VARIABLE_ROW_HEIGHT | ULC.ULC_USER_ROW_HEIGHT | ULC.ULC_SINGLE_SEL,
                                                 size=(350,110), name="colors")
        self.colorList.InsertColumn(0, 'Table', width=170)
        self.colorList.InsertColumn(1, 'Color', width=80)
        self.colorList.InsertColumn(2, 'Font', width=80)
        self.colorList.SetUserLineHeight(25)

        parent.sql.execute("select id,tag,color,font from colors")
        tags = parent.sql.fetchall()
        for tag in tags:
            index = self.colorList.InsertStringItem(tag["id"], tag["tag"])
            self.colorList.SetItemData(index, tag["id"])
            colorPicker = wx.ColourPickerCtrl(self.colorList, tag["id"], col=tag["color"], style=wx.CLRP_SHOW_LABEL, name='color')
            self.colorList.SetItemWindow(index, 1, wnd=colorPicker, expand=True)
            colorPicker = wx.ColourPickerCtrl(self.colorList, tag["id"], col=tag["font"], style=wx.CLRP_SHOW_LABEL, name='font')
            self.colorList.SetItemWindow(index, 2, wnd=colorPicker, expand=True)
        self.colorList.EnsureVisible(index)

        self.categoryList = ULC.UltimateListCtrl(panel, wx.ID_ANY, agwStyle=ULC.ULC_REPORT | ULC.ULC_HRULES | ULC.ULC_BORDER_SELECT |
                                                 ULC.ULC_HAS_VARIABLE_ROW_HEIGHT | ULC.ULC_USER_ROW_HEIGHT | ULC.ULC_SINGLE_SEL,
                                                 size=(350,300), name="categories")
        self.categoryList.InsertColumn(0, 'Category', width=170)
        self.categoryList.InsertColumn(1, 'Color', width=80)
        self.categoryList.InsertColumn(2, 'Font', width=80)
        self.categoryList.SetUserLineHeight(25)

        catcount = parent.sql.execute("select id,category,color,font from categories")
        categories = parent.sql.fetchall()
        for category in categories:
            index = self.categoryList.InsertStringItem(self.lastrow, category["category"])
            self.categoryList.SetItemData(index, category["id"])
            colorPicker = wx.ColourPickerCtrl(self.categoryList, category["id"], col=category["color"], style=wx.CLRP_SHOW_LABEL, name='color')
            self.categoryList.SetItemWindow(index, 1, wnd=colorPicker, expand=True)
            colorPicker = wx.ColourPickerCtrl(self.categoryList, category["id"], col=category["font"], style=wx.CLRP_SHOW_LABEL, name='font')
            self.categoryList.SetItemWindow(index, 2, wnd=colorPicker, expand=True)
            self.lastrow += 1

        maxid = parent.sql.execute('select max(id) from categories').fetchone()[0]
        index = self.categoryList.InsertStringItem(self.lastrow, '<New Category>')
        self.categoryList.SetItemData(index, maxid + 1)
        colorPicker = wx.ColourPickerCtrl(self.categoryList, self.lastrow, col=self.getRamdomColor(), style=wx.CLRP_SHOW_LABEL, name='color')
        self.categoryList.SetItemWindow(index, 1, wnd=colorPicker, expand=True)
        colorPicker = wx.ColourPickerCtrl(self.categoryList, self.lastrow, col="#000000", style=wx.CLRP_SHOW_LABEL, name='font')
        self.categoryList.SetItemWindow(index, 2, wnd=colorPicker, expand=True)
        self.categoryList.EnsureVisible(index)

        rootsizer.Add(self.colorList, 0, wx.ALL, 5)
        rootsizer.Add(self.categoryList, 1, wx.ALL | wx.EXPAND, 5)

        self.Bind(wx.EVT_COLOURPICKER_CHANGED, self.onColorCtrl)
        self.Bind(ULC.EVT_LIST_ITEM_SELECTED, self.OnItemSelected, self.categoryList)
        self.Bind(ULC.EVT_LIST_END_LABEL_EDIT, self.onEndEdit, self.categoryList)
        self.categoryList.Bind(wx.EVT_LEFT_DCLICK, self.onDoubleClick)

        panel.SetSizer(rootsizer)
        rootsizer.Fit(self)

    def OnItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        self.currentList = event.EventObject.GetLabel()
        self.currentId = self.categoryList.GetItemData(self.currentItem)
        print "OnItemSelected %s: row %s, id %s, %s" %(self.currentList, self.currentItem, self.currentId, self.categoryList.GetItemText(self.currentItem))

    def onDoubleClick(self, event):
        name = event.EventObject.GetLabel()
        if self.currentItem == self.lastrow:
            self.categoryList.SetItemText(self.currentItem, '')
        self.categoryList.EditLabel(self.currentItem)

    def onEndEdit(self, event):
        print "onEndEdit %s: row %s, id %s, %s" % (self.currentList, event.m_itemIndex, self.currentId, event.GetText())
        if self.currentItem == self.lastrow and event.GetText() != '': # Add
            self.lastrow += 1
            newcolor = self.getRamdomColor()
            index = self.categoryList.InsertStringItem(self.lastrow, '<New Category>')
            colorPicker = wx.ColourPickerCtrl(self.categoryList, self.lastrow, col=newcolor, style=wx.CLRP_SHOW_LABEL, name='color')
            self.categoryList.SetItemWindow(index, 1, wnd=colorPicker, expand=True)
            colorPicker = wx.ColourPickerCtrl(self.categoryList, self.lastrow, col="#000000", style=wx.CLRP_SHOW_LABEL, name='font')
            self.categoryList.SetItemWindow(index, 2, wnd=colorPicker, expand=True)
            self.Parent.sql.execute('insert into categories (category, color, font) values (?,?,?)', (event.GetText(), newcolor, '#000000'))
            self.categoryList.SetItemData(index, self.Parent.sql.lastrowid)
        elif self.currentItem != self.lastrow and event.GetText() == '': # Delete
            self.Parent.sql.execute('delete from categories where id=?', (self.currentId,))
            self.Parent.sql.execute('update docs set categoriesid="" where categoriesid=?', (self.currentId,))
            wx.CallAfter(self.categoryList.DeleteItem, self.currentItem)
            self.currentItem -= 1
            self.lastrow -= 1
        elif event.GetText() == '': # Last Row (dont Delete)
            self.categoryList.SetItemText(self.currentItem, '<New Category>')
        else: # Update
            self.Parent.sql.execute('update categories set category=? where id=?', (event.GetText(), self.currentId))

    def onColorCtrl(self, event):
        color = event.GetColour()
        name = event.EventObject.GetLabel()
        table = event.EventObject.GrandParent.GetName()
        catid = event.EventObject.GetId()
        colorHexStr = "#%02x%02x%02x" % color.Get()
        print "onColorCtrl %s: id %s, name %s, color %s" % (table, str(catid), name, colorHexStr)
        if table == "categories":
            if name == "color":
                self.Parent.sql.execute('update categories set color=? where id=?', (colorHexStr, catid))
            elif name == 'font':
                self.Parent.sql.execute('update categories set font=? where id=?', (colorHexStr, catid))
        elif table == "colors":
            if name == 'color':
                self.Parent.sql.execute('update colors set color=? where id=?', (colorHexStr, catid))
            elif name == 'font':
                self.Parent.sql.execute('update colors set font=? where id=?', (colorHexStr, catid))

    def getRamdomColor(self):
        randomcolor = "#%x" % (random.randint(1, 16777215))
        print "New Color: %s" % randomcolor
        return randomcolor

    def onClose(self, event):
        self.Parent.con.commit()
        self.Destroy()

class EditSettings(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, wx.ID_ANY, 'Edit Settings')
        wx.EVT_CLOSE(self, self.onClose)

        panel = wx.Panel(self)
        rootsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.selectLabel = wx.StaticText(panel, wx.ID_ANY, 'Select:', size=(30,20))
        self.selectText = wx.TextCtrl(panel, wx.ID_ANY, style=wx.TE_PROCESS_ENTER)
        self.selectText.SetLabel(str(parent.selectlimit))

        rootsizer.Add(self.selectLabel, 0, wx.ALL, 10)
        rootsizer.Add(self.selectText, 1, wx.ALL | wx.EXPAND, 10)

        panel.SetSizer(rootsizer)
        rootsizer.Fit(self)

    def onClose(self, event):
        self.Parent.selectlimit = self.selectText.GetLabel()
        self.Parent.config.set('settings', 'select', str(self.Parent.selectlimit))
        try:
            with open(self.Parent.docini, 'w') as configfile:
                self.Parent.config.write(configfile)
        except:
            self.Parent.displayMessage("Error Saving ini file")
        self.Destroy()

def startThread(func, *args): # helper method to run a function in another thread
    thread = threading.Thread(target=func, args=args)
    thread.setDaemon(True)
    thread.start()
    return thread

def startfile(filename):
    try:
        os.startfile(filename)
    except AttributeError:
        subprocess.Popen(['xdg-open', filename])

class MyApp(wx.App, wx.lib.mixins.inspection.InspectionMixin):
    def __init__(self):
        wx.App.__init__(self, redirect=False)

    def OnInit(self):
        self.Init()  # initialize the inspection tool
        frame = DisplayForm(None, title="Doci")
        frame.Show()
        self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    try:
        #application = wx.PySimpleApp(wx.lib.inspection.InspectionTool)
        #frame = MyFrame(None, "Doci")
        application = MyApp()
        application.MainLoop()
    finally:
        try:
            del application
        except:
            pass
