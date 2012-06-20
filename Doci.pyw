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
import os, sys, wx, time, datetime
import ConfigParser, threading, Queue
from win32api import LoadResource
import wx.lib.agw.multidirdialog as mdd
import pprint

class MyFrame(wx.Frame):
    def __init__(self, parent, title):

        version = "0.0" # Grab real version from exe
        self.docini = "Doci.ini"
        self.docdb = ""
        self.dochtml = ""
        self.docdir = []
        self.selectlimit = 0
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

        # Get the Icon and Version string from the exe resources
        if os.path.splitext(sys.argv[0])[1] == ".exe":
            icon = wx.Icon(sys.argv[0], wx.BITMAP_TYPE_ICO)
            version = LoadResource(0, u'VERSION', 1)
        else:
            icon = wx.Icon("Doci.ico", wx.BITMAP_TYPE_ICO)
            
        wx.Frame.__init__(self, parent, title=title + " - v" + version)
        self.SetIcon(icon)
        wx.EVT_CLOSE(self, self.onClose)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.SetMinSize((500,500))
        
        self.statusBar = self.CreateStatusBar()
        self.statusBar.SetFieldsCount(3)
        self.statusBar.SetStatusWidths([85,-1,60,])
        self.statusBar.SetStatusText("0/0", 0)
        self.messageTimer = wx.Timer(self)

        self.searchLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Find:')
        self.searchText = wx.TextCtrl(self.panel, wx.ID_ANY, style=wx.TE_PROCESS_ENTER)
        self.searchButton = wx.Button(self.panel, wx.ID_ANY, 'Search')
        self.indexButton = wx.Button(self.panel, wx.ID_ANY, 'Index')
        self.openButton = wx.Button(self.panel, wx.ID_ANY, 'Open')

        self.buttonBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Controls')
        self.checkButton = wx.Button(self.panel, wx.ID_ANY, 'Check')
        self.editButton = wx.Button(self.panel, wx.ID_ANY, 'Edit')
        self.updateButton = wx.Button(self.panel, wx.ID_ANY, 'Update')
        self.updateButton.Enable(False)
        self.prevButton = wx.Button(self.panel, wx.ID_ANY, 'Prev')
        self.nextButton = wx.Button(self.panel, wx.ID_ANY, 'Next')
        self.exitButton = wx.Button(self.panel, wx.ID_EXIT, 'Exit')

        self.idLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Id:', size=(30,20))
        self.idText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21), style=wx.TE_PROCESS_ENTER)
        self.idText.SetEditable(True)
        self.fileLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Name:', size=(30,20))
        self.fileText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.fileText.SetEditable(False)
        self.fileText.SetForegroundColour(self.grey)
        self.extLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Ext:')
        self.extText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21))
        self.extText.SetEditable(False)
        self.extText.SetForegroundColour(self.grey)
        #self.categoryList = wx.ListBox(self.panel, wx.ID_ANY, choices=[], name='category', style=0) #, pos=wx.Point(8, 48), size=wx.Size(184, 256), style=0)
        self.categoryChoice = wx.Choice( self.panel, wx.ID_ANY, style=0) 
        self.categoryChoice.Enable(False)
        #id1 = wx.NewId()
        #wx.RegisterId(id1)        
        self.dirLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Dir:', size=(30,20))
        self.dirText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.dirText.SetEditable(False)
        self.dirText.SetForegroundColour(self.grey)
        self.dateLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Date:')
        self.dateText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(120,21))
        self.dateText.SetEditable(False)
        self.dateText.SetForegroundColour(self.grey)

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
        self.searchSizer.Add(self.openButton, 0, wx.ALL, 5)

        self.buttonBoxSizer.Add(self.checkButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.editButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.updateButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.prevButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.nextButton, 0, wx.ALL, 5)
        self.buttonBoxSizer.Add(self.exitButton, 0, wx.ALL, 5)

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
        self.Bind(wx.EVT_BUTTON, self.onCheck, self.checkButton)
        self.Bind(wx.EVT_BUTTON, self.onEdit, self.editButton)
        self.Bind(wx.EVT_BUTTON, self.onUpdate, self.updateButton)
        self.Bind(wx.EVT_BUTTON, self.onPrev, self.prevButton)
        self.Bind(wx.EVT_BUTTON, self.onNext, self.nextButton)
        self.Bind(wx.EVT_BUTTON, self.onExit, self.exitButton)
        self.Bind(wx.EVT_TEXT_ENTER, self.onSearch, self.searchText)
        self.Bind(wx.EVT_TEXT_ENTER, self.onId, self.idText)
        self.Bind(wx.EVT_TIMER, self.setMessage, self.messageTimer)

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

        print "***Form Init***"

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
            self.sql.execute("CREATE TABLE categories (id INTEGER PRIMARY KEY, category TEXT, color TEXT, display INTEGER)")
            for (cat,color, display) in (("Standards", "#fff0f0", 1), ("Drawings", "#ecfbd4", 2)):
                self.sql.execute("INSERT INTO categories (category, color, display) VALUES (?, ?, ?)", (cat, color, display))
            self.sql.execute("CREATE TABLE colors (id INTEGER PRIMARY KEY, tag TEXT, color TEXT, UNIQUE(tag))")
            for (tag,color) in (("Header", "#328aa4"), ("Default", "#e5f1f4"), ("Highlight", "#ffffb2")):
                self.sql.execute("INSERT INTO colors (tag, color) VALUES (?, ?)", (tag, color))
            self.sql.execute("CREATE INDEX tag on colors (tag)")
            self.sql.execute("CREATE INDEX hash on docs (hash,size)")
            self.sql.execute("CREATE INDEX filename on docs (dir,name,ext)")
            self.sql.execute("CREATE VIRTUAL TABLE search USING fts4(content='docs', name, desc)")
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
            self.sql = self.con.cursor()
            
    def closeDB(self):
        if self.con:
            self.con.commit()
            self.sql.close()
            
    def getIni(self):
        # Create ini file if it doesnt exist
        #http://xoomer.virgilio.it/infinity77/main/freeware.html
        path = {'dirs':self.docdir, 'db':'Doci.db', 'html':'Doci.html'}
        limit = {'select':'10000'}

        if not os.path.isfile(self.docini):
            print "Creating ini file"
            currentDir = os.getcwd()
            dlg = mdd.MultiDirDialog(self, message=u'Choose directory(s) to scan', title=u'Browse For Folders', defaultPath=currentDir, agwStyle=mdd.DD_DIR_MUST_EXIST|mdd.DD_MULTIPLE, name='multidirdialog')
            if dlg.ShowModal() == wx.ID_OK:
                self.docdir.extend(dlg.GetPaths())
            else:
                self.displayMessage("Unable to continue without ini file", "Info")
                self.Destroy()
                return
            
            config = ConfigParser.ConfigParser()
            config.add_section('Path')
            for option in path:
                config.set('Path', option, path[option])
            config.add_section('Limit')
            for option in limit:
                config.set('Limit', option, limit[option])
                
            try:
                with open(self.docini, 'w') as configfile:
                    config.write(configfile)
                configfile.close()
            except:
                self.displayMessage("Error Creating ini file")
                self.onExit(self)
                
        defaults = dict(path.items() + limit.items())
        config = ConfigParser.SafeConfigParser(defaults)
        try:
            config.readfp(open(self.docini))
            self.docdir.extend(eval(config.get('Path', 'dirs')))
            self.docdb = config.get('Path', 'db')
            self.dochtml = config.get('Path', 'html')
            self.selectlimit = int(config.get('Limit', 'select'))
        except:
            self.displayMessage("Missing ini file")
            self.onExit(self)
            
    def getCategories(self):
        self.categoryChoice.Clear()
        self.categoryChoice.Append("")
        categories = self.sql.execute("select category from categories order by display")
        for category in categories.fetchall():
            self.categoryChoice.Append(category["category"])
        self.categoryChoice.SetSelection(0)

    def setIndex(self, index):
        (oldindex, results) = self.statusBar.GetStatusText(0).split("/")
        self.statusBar.SetStatusText(str(index) + "/" + str(results),0)

    def getIndex(self):
        if self.statusBar.GetStatusText(0) != 'None':
            (index, results) = self.statusBar.GetStatusText(0).split("/")
            return int(index)
        
    def getId(self, index=None):
        if not index:
            index = self.getIndex()
        if index and self.results:
            return self.results[index - 1]

    def setResults(self, results):
        (index, oldresults) = self.statusBar.GetStatusText(0).split("/")
        self.statusBar.SetStatusText(str(index) + "/" + str(results),0)

    def getResults(self):
        if self.statusBar.GetStatusText(1) != 'None':
            (index, results) = self.statusBar.GetStatusText(0).split("/")
            return int(results)

    def setMaxid(self, maxid=""):
        if maxid == "":
            maxid = self.sql.execute('select max(id) from docs').fetchone()[0]
        if maxid:
            self.statusBar.SetStatusText(str(maxid),2)

    def setMessage(self, e):
        message = None
        self.messageCount += 1
        queuedepth = self.messageQueue.qsize()
        try:
            message = self.messageQueue.get_nowait()
        except Queue.Empty:
            if self.messageCount > 5:
                message = u""
        try:
            if message != None:
                self.messageCount = 0
                if len(message):
                    message = str(queuedepth) + ": " + message
                self.statusBar.SetStatusText(message,1)
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

    def onPrev(self, e):
        if self.getIndex():
            prev = self.getIndex() - 1
            if prev < 1:
                self.displayMessage("Reached start, wraping to end", status="Info")
                prev = self.getResults()
            self.displayRecord(prev)

    def onNext(self, e):
        if self.getIndex():
            next = self.getIndex() + 1
            if next > self.getResults():
                self.displayMessage("Reached end, wraping to beginning", status="Info")
                next = 1
            self.displayRecord(next)

    def onSearch(self, e):
        self.searchRecords(self.searchText.GetValue())
        self.displayRecord()

    def onIndex(self, e):
        if os.path.exists(self.dochtml):
            if self.displayMessage("File %s Already Exists!\nOverwite file?" % self.dochtml, "Query") == wx.ID_NO:
                return
        try:
            htmlfile = open(self.dochtml, 'w')
        except:
            self.displayMessage("Error Creating html file %s" % self.dochtml)
        # HTML Headers
        self.messageQueue.put("Creating Html")
        htmlfile.write('<html><head><title>Doci</title>')
        htmlfile.write('<style type="text/css">\n')
        # Get Table Colors
        header = self.sql.execute('select color from colors where tag="Header"').fetchone()[0]
        defaultcolor = self.sql.execute('select color from colors where tag="Default"').fetchone()[0]
        highlight = self.sql.execute('select color from colors where tag="Highlight"').fetchone()[0]
        htmlfile.write("""table, td{ font:100% Arial, Helvetica, sans-serif; }
table{width:100%;border-collapse:collapse;margin:1em 0;}
th, td{text-align:left;padding:.5em;border:1px solid #fff;}""")
        htmlfile.write("th, tfoot td{background:%s; color:#fff;}\ntd{background:%s;}" % (header, defaultcolor))
        
        # Create Category classes
        self.sql.execute("select category, color from categories")
        categories = self.sql.fetchall()
        for category in categories:
            htmlfile.write('tr.%s td{background:%s;}' % (category["category"], category["color"]))
        htmlfile.write('tr:hover>td{background:%s;}' % highlight)
        htmlfile.write('</style><body>\n')
            
        self.sql.execute("select docs.id as id, dir, name, ext, desc, date, size, category, color from docs left join categories on docs.categoriesid=categories.id limit 0,?", (self.selectlimit,))
        docs = self.sql.fetchall()
        # Create Table
        htmlfile.write('<div align=center><h1>Document Indexer</h1></div>\n')
        htmlfile.write('<table cellspacing="0" cellpadding="0" >\n')
        htmlfile.write('<thead><th>Id</th><th>File</th><th>Category</th><th>Date</th><th>Size</th></thead>\n')
        if len(docs) == self.selectlimit:
            htmlfile.write('<tfoot><tr><td>%s</td><td colspan="4">Max Limit Reached!</td></tr></tfoot>\n' % self.selectlimit)
        htmlfile.write('<tbody>\n')
        # Create Table Body
        for doc in docs:
            filename = os.path.join(doc["dir"], doc["name"]+doc["ext"])
            date = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(doc["date"]))
            size = self.sizeof(float(doc["size"]))
            htmlfile.write('<tr class="%s" ><td>%s</td>' % (doc["category"], doc["id"]))
            htmlfile.write('<td><a href="%s" title="%s" target="_blank">%s</a></td>' % (filename, filename, doc["desc"]))
            htmlfile.write('<td>%s</td><td>%s</td><td>%s</td></tr>\n' % (doc["category"], date, size))
        htmlfile.write("</tbody></table></body></html>\n")
        htmlfile.close()
        self.messageQueue.put("Opening Html File")
        os.startfile(self.dochtml)
        
    def sizeof(self, num):
        for x in ['bytes','KB','MB','GB']:
            if num < 1024.0:
                return "%3.1f%s" % (num, x)
            num /= 1024.0
        return "%3.1f%s" % (num, 'TB')

    def onOpen(self, e):
        filename = os.path.join(self.dirText.GetValue(), self.fileText.GetValue() + self.extText.GetValue())
        self.messageQueue.put_nowait(u"Opening file: " + self.fileText.GetValue())
        os.startfile(filename)

    def onCheck(self, e):
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
        if self.workerAbort.isSet():
            self.sql.execute("delete from docs where seen <>1")
        else:
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

    def onEdit(self, e):
        if self.editButton.GetLabel() == "Cancel":
            self.displayRecord(self.getId())
            self.enableButtons(True)
        else:
            self.disableButtons(cancel=True)

    def onUpdate(self, e):
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

    def onExit(self, e):
        self.Close(True)

    def onClose(self, e):
        print "Closing..."
        self.closeDB()
        self.messageTimer.Stop()
        self.Destroy()

    def onId(self, e):
        id = self.idText.GetValue()
        if id.isdigit():
            id = int(id)
            if id in self.results:
                self.displayRecord(self.results.index(id + 1)) # Allow for 0 indexed array

    def disableButtons(self, cancel=False, everything=False):
        self.searchButton.Enable(False)
        self.indexButton.Enable(False)
        self.openButton.Enable(False)
        self.idText.SetEditable(False)
        self.idText.SetForegroundColour(self.grey)
        self.checkButton.Enable(False)
        self.prevButton.Enable(False)
        self.nextButton.Enable(False)
        if everything:
            self.descText.SetForegroundColour(self.grey)
            self.descText.SetEditable(False)
            self.updateButton.Enable(False)
            self.categoryChoice.Enable(False)
            self.exitButton.Enable(False)
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
        self.exitButton.Enable(True)
        self.searchButton.Enable(True)
        self.indexButton.Enable(True)
        self.openButton.Enable(True)
        self.idText.SetEditable(True)
        self.idText.SetForegroundColour(self.black)
        self.updateButton.Enable(False)
        self.checkButton.Enable(True)
        self.prevButton.Enable(True)
        self.nextButton.Enable(True)
        self.descText.SetForegroundColour(self.grey)
        self.descText.SetEditable(False)
        self.categoryChoice.Enable(False)
        if cancel:
            self.editButton.SetLabel("Edit")
        else:
            self.editButton.Enable(True)


def startThread(func, *args): # helper method to run a function in another thread
    thread = threading.Thread(target=func, args=args)
    thread.setDaemon(True)
    thread.start()
    return thread

if __name__ == '__main__':
    try:
        application = wx.PySimpleApp()
        frame = MyFrame(None, "Doci")
        application.MainLoop()
    finally:
        del application
