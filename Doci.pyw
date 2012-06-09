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
import sqlite3, traceback, hashlib
import os, sys, wx, time, datetime
import ConfigParser, threading
import pprint

class MyFrame(wx.Frame):
    def __init__(self, parent, title):

        self.docdb = "Doci.db"
        self.docini = "Doci.ini"
        self.docdir = []
        self.addfiles = False
        self.addid = []
        self.results = []
        self.grey = wx.NamedColour("GREY")
        self.black = wx.NamedColour("BLACK")
        self.workerRun = False
        self.workerDir = ""

        wx.Frame.__init__(self, parent, title=title)
        if os.path.splitext(sys.argv[0])[1] == ".exe":
            icon = wx.Icon(sys.argv[0], wx.BITMAP_TYPE_ICO)
        else:
            icon = wx.Icon("Doci.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.SetMinSize((500,500))
        self.statusBar = self.CreateStatusBar()
        self.statusBar.SetFieldsCount(4)
        self.statusBar.SetStatusWidths([50,50,50,-1])

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

        self.idLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Id:')
        self.idText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21), style=wx.TE_PROCESS_ENTER)
        self.idText.SetEditable(True)
        self.fileLabel = wx.StaticText(self.panel, wx.ID_ANY, 'File:')
        self.fileText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.fileText.SetEditable(False)
        self.fileText.SetForegroundColour(self.grey)
        self.extLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Ext:')
        self.extText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(50,21))
        self.extText.SetEditable(False)
        self.extText.SetForegroundColour(self.grey)

        self.dirLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Dir:')
        self.dirText = wx.TextCtrl(self.panel, wx.ID_ANY)
        self.dirText.SetEditable(False)
        self.dirText.SetForegroundColour(self.grey)
        self.dateLabel = wx.StaticText(self.panel, wx.ID_ANY, 'Date:')
        self.dateText = wx.TextCtrl(self.panel, wx.ID_ANY, size=(120,21))
        self.dateText.SetEditable(False)
        self.dateText.SetForegroundColour(self.grey)
        #self.dateText.Enable(False)

        self.descBox = wx.StaticBox(self.panel, wx.ID_ANY, 'Description')
        self.descText = wx.TextCtrl(self.panel, wx.ID_ANY, style=wx.TE_MULTILINE ) #|wx.TE_PROCESS_TAB
        self.descText.SetForegroundColour(self.grey)
        self.descText.SetEditable(False)

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

        self.buttonBoxSizer.Add(self.checkButton, 0, wx.ALL, 5)
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
        self.dirSizer.Add(self.dirText, 1, wx.ALL | wx.EXPAND, 5)
        self.dirSizer.Add(self.dateLabel, 0, wx.ALL, 5)
        self.dirSizer.Add(self.dateText, 0, wx.ALL, 5)

        self.descSizer.Add(self.descText, 1, wx.ALL | wx.EXPAND | wx.RIGHT, 5)

        self.rootSizer.Add(self.searchSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.buttonSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.propSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.dirSizer, 0, wx.ALL | wx.EXPAND, 5)
        self.rootSizer.Add(self.descSizer, 1, wx.EXPAND|wx.TOP|wx.BOTTOM,5)

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

        self.panel.SetSizer(self.rootSizer)
        self.rootSizer.Fit(self)

        self.openDB()
        self.setMaxid()
        self.searchRecords()
        self.displayRecord(1)
        self.Show()

        # Create ini file if it doesnt exist
        config = ConfigParser.ConfigParser()
        if not os.path.isfile(self.docini):
            print "Creating ini file"
            currentDir = os.getcwd()
            dlg = wx.DirDialog(self, "Choose documents directory", currentDir,style=wx.DD_DEFAULT_STYLE)
            if dlg.ShowModal() == wx.ID_OK:
                selectedDir = dlg.GetPath()
                if currentDir in selectedDir:
                    self.docdir.append(selectedDir.lstrip(currentDir))
                else:
                    self.docdir.append(selectedDir)
            else:
                self.onError("Unable to continue without ini file", "Info")
                self.onExit(self)
                return
                
            config.add_section('Path')
            config.set('Path', 'Dirs', self.docdir)
            try:
                with open(self.docini, 'w') as configfile:
                    config.write(configfile)
                configfile.close()
            except:
                self.onError("Error Creating ini file")
                self.onExit(self)
        else:
            try:
                config.readfp(open(self.docini))
                self.docdir.extend(eval(config.get("Path", "Dirs")))
            except:
                self.onError("Missing ini file")
                self.onExit(self)

        print "***Form Init***"

    def openDB(self):
        # Create DB file if it doesnt exist
        if not os.path.isfile(self.docdb):
            print "Creating DB file"
            try:
                self.con = sqlite3.connect(self.docdb)
                self.con.row_factory = sqlite3.Row
                self.con.text_factory = str # Allow unicode conversion
            except:
                self.onError("Failed to create Database")
            self.sql = self.con.cursor()
            self.sql.execute("create table docs (id INTEGER PRIMARY KEY, dir TEXT, name TEXT, ext TEXT, desc TEXT, hash TEXT, size TEXT, date REAL, seen INTEGER, added TEXT)")
            self.sql.execute("""create table dupes (id INTEGER PRIMARY KEY, dir TEXT, name TEXT, ext TEXT, desc TEXT, hash TEXT, size TEXT, date REAL, seen INTEGER, added TEXT,
                            docsid INTEGER, FOREIGN KEY(docsid) REFERENCES docs(id), UNIQUE(dir, name, ext))""")
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
                self.con.text_factory = str # Allow unicode conversion
            except:
                self.onError("Failed to Open Database")
            self.sql = self.con.cursor()
            
    def closeDB(self):
        if self.con:
            self.con.commit()
            self.sql.close()

    def setIndex(self, sb):
        self.statusBar.SetStatusText(str(sb),0)

    def getIndex(self):
        if self.statusBar.GetStatusText(0) != 'None':
            return int(self.statusBar.GetStatusText(0))

    def setResults(self, sb):
        self.statusBar.SetStatusText(str(sb),1)

    def getResults(self):
        if self.statusBar.GetStatusText(1) != 'None':
            return int(self.statusBar.GetStatusText(1))

    def setMaxid(self, maxid=""):
        if maxid == "":
            maxid = self.sql.execute('select max(id) from docs').fetchone()[0]
        if maxid:
            self.statusBar.SetStatusText(str(maxid),2)

    def setMessage(self, sb):
        # Unicode Sux, probably a better way to do this, but this currently works, REF:
        # http://docs.python.org/howto/unicode.html#unicode-filenames
        # https://github.com/wimleers/fileconveyor/issues/62
        # http://stackoverflow.com/questions/2392732/sqlite-python-unicode-and-non-utf-data
        # http://stackoverflow.com/questions/2838100/pysqlite2-programmingerror-you-must-not-use-8-bit-bytestrings
        try:
            self.statusBar.SetStatusText(sb.decode('latin-1', 'replace'),3)
        except:
            self.onError("Unicode failure on messagebar")

    def searchRecords(self, search=""):
        if search != "":
            self.results = [element[0] for element in self.sql.execute('select docid from search where search match ?', (search,)).fetchall()]
        else:
            self.results = [element[0] for element in self.sql.execute('select id from docs').fetchall()]
        self.setResults(len(self.results))        

    def displayRecord(self, index):
        if index and self.results:
            id = self.results[index - 1]
            self.sql.execute('select id, dir, name, ext, desc, date from docs where id=?', (id,))
            row = self.sql.fetchone()
            self.setIndex(index)
            self.idText.SetValue(str(row["id"]))
            self.fileText.SetValue(str(row["name"]))
            self.extText.SetValue(str(row["ext"]))
            self.dirText.SetValue(str(row["dir"]))
            self.dateText.SetValue(str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(row["date"]))))
            self.descText.SetValue(str(row["desc"]))

    def onPrev(self, e):
        if self.getIndex():
            prev = self.getIndex() - 1
            if prev < 1:
                self.onError("Reached start, wraping to end", status="Info")
                prev = self.getResults()
            self.displayRecord(prev)

    def onNext(self, e):
        if self.getIndex():
            next = self.getIndex() + 1
            if next > self.getResults():
                self.onError("Reached end, wraping to beginning", status="Info")
                next = 1
            self.displayRecord(next)

    def onError(self, message, status="Error"):
        if status == "Info":
            self.setMessage(message)
            dlg = wx.MessageDialog(self, message=message, caption='Info', style=wx.OK|wx.ICON_INFORMATION)
        elif status == "Query":
            self.setMessage(message)
            dlg = wx.MessageDialog(self, message=message, caption='Query', style=wx.YES|wx.NO|wx.ICON_QUESTION)
        else:
            self.setMessage(message)
            dlgMessage = "Line: " + str(sys.exc_info()[2].tb_lineno) + " - " + message + "\n\n" + traceback.format_exc(0)
            print dlgMessage
            dlg = wx.MessageDialog(self, message=dlgMessage, caption='Error', style=wx.OK|wx.CANCEL|wx.ICON_EXCLAMATION)
        result = dlg.ShowModal()
        dlg.Destroy()
        return(result)

    def onSearch(self, e):
        self.searchRecords(self.searchText.GetValue())
        self.displayRecord(1)

    def onIndex(self, e):
        pass

    def onOpen(self, e):
        filename = os.path.join(self.dirText.GetValue(), self.fileText.GetValue() + self.extText.GetValue())
        self.setMessage("Opening file: " + self.fileText.GetValue().encode('ascii', 'replace'))
        os.startfile(filename)

    def onCheck(self, e):
        self.disableButtons()
        self.closeDB() # Best to be thread safe
        self.progress = wx.ProgressDialog('Checking Files', 'Please wait...',style=wx.PD_CAN_ABORT|wx.PD_ELAPSED_TIME)
        startThread(self.addFiles, self.docdir)
        self.progress.ShowModal()
        self.openDB()
        
        dupes = self.sql.execute("select max(id) from dupes").fetchone()[0]
        if dupes:
            self.onError("Found %s Duplicate Files, Ignoring them for now" % dupes, status="Info")
            # We should list these and let the user delete them
            self.sql.execute("delete from dupes")
            self.con.commit()
        missing = self.sql.execute("select count(*) from docs where seen <> 1").fetchone()[0]
        if missing:
            if self.onError("Found %s missing files\nPurge from database?" % missing, status="Query") == wx.ID_YES:
                    self.sql.execute("delete from docs where seen <>1")
        self.sql.execute("update docs set seen=''")
        if self.addid:
            self.displayRecord(self.results.index(self.addid.pop(0))+1) # Idex of 0 based array
            if self.onError("Found %s new files\nBulk update files?" % (len(self.addid) + 1), status="Query") == wx.ID_YES:
                self.addfiles = True
            else:
                self.addfiles = False
                self.addid = []
                self.enableButtons()
        else:
            self.enableButtons()
        self.searchRecords()
        self.removeFiles()
        self.sql.execute("INSERT INTO search(search) VALUES('optimize')")

    def updateProgress(self):
        if self.workerRun == True:
            (run, skip) = self.progress.Pulse(self.workerDir)
            if run == False:
                self.workerRun = False
            else:
                threading.Timer(0.2, self.updateProgress).start()

    def chunkReader(self, fobj, chunk_size=8192):
        """Generator that reads a file in chunks of bytes"""
        while True:
            chunk = fobj.read(chunk_size)
            if not chunk:
                return
            yield chunk

    def addFiles(self, paths, hash=hashlib.sha1):
        self.workerRun = True
        threading.Timer(0.2, self.updateProgress).start()
        # Open DB (thread safe)
        self.openDB()
        for path in paths:
            if self.workerRun == False:
                break
            if not os.path.isdir(path):
                self.onError("Missing Directory: %s" % path, status="Info")
                continue
            for dirpath, dirnames, filenames in os.walk(path):
                self.workerDir = dirpath
                if self.workerRun == False:
                    break
                for file in filenames:
                    filepath = os.path.join(dirpath, file)
                    filebasename = os.path.splitext(file)[0]
                    fileext = os.path.splitext(file)[1]
                    filedate = os.path.getctime(filepath)
                    hashobj = hash()
                    for chunk in self.chunkReader(open(filepath, 'rb')):
                        hashobj.update(chunk)
                        if self.workerRun == False:
                            break                        
                    filehash = hashobj.hexdigest()
                    filesize = os.path.getsize(filepath)
                    self.sql.execute('select id, dir, name, ext from docs where hash=? and size=?', (filehash,filesize))
                    duplicate = self.sql.fetchone()
                    if duplicate:
                        duplicateFile = os.path.join(duplicate["dir"], str(duplicate["name"]) + str(duplicate["ext"]))
                        if filepath != duplicateFile:
                            if not os.path.isfile(duplicateFile):
                                # File has moved update record
                                self.setMessage("File Moved: %s" % filepath)
                                self.setMessage("Old Path: %s" % duplicateFile)
                                try:
                                    self.sql.execute("update docs set dir=?,name=?,ext=? where id=?", (dirpath, filebasename, fileext, duplicate["id"]))
                                except:
                                    if self.onError("Failed to update file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                                        return
                            else:
                                # Duplicate File
                                self.setMessage("Found Duplicate: %s" % filepath)
                                self.setMessage("Clashes with: %s" % duplicateFile)
                                #self.onError("Duplicate found:\n %s\n---matches existing file---\n%s\n" % (filepath, duplicateFile), status="Info")
                                try:
                                    self.sql.execute("insert into dupes (dir, name, ext, desc, hash, size, date, seen, added, docsid) values (?, ?, ?, ?, ?, ?, ?, 1, datetime(),?)",
                                                    (dirpath, filebasename, fileext, filebasename, filehash, filesize, filedate, duplicate["id"]))
                                except:
                                    if self.onError("Failed to mark duplicate file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                                        return
                            continue
                        else:
                            # Found Existing File, update the seen value
                            self.sql.execute("update docs set seen=1 where id=?", (duplicate["id"],))
                            continue
                    try:
                        self.sql.execute("insert into docs (dir, name, ext, desc, hash, size, date, seen, added) values (?, ?, ?, ?, ?, ?, ?, 1, datetime())",
                                        (dirpath, filebasename, fileext, filebasename, filehash, filesize, filedate))
                        self.addid.append(self.sql.lastrowid)
                        self.setMaxid(self.sql.lastrowid)
                    except:
                        if self.onError("Failed to add file.\n%s\n%s%s" % (dirpath, filebasename, fileext)) == wx.ID_CANCEL:
                            return
        self.closeDB()
        self.workerRun = False
        wx.CallAfter(self.progress.Destroy)

    def removeFiles(self):
        pass

    def onEdit(self, e):
        self.disableButtons()

    def onUpdate(self, e):
        #Commit the changes after updating the Desc
        self.sql.execute("update docs set desc=? where id=?", (self.descText.GetValue(), self.getIndex()))
        self.con.commit()
        if self.addfiles:
            newid = self.addid.pop(0)
            if not self.addid:
                self.enableButtons()
                self.addfiles = False
            else:
                self.displayRecord(newid)
        else:
            self.enableButtons()
            self.addfiles = False

    def onExit(self, e):
        self.closeDB()
        self.Close(True)

    def onId(self, e):
        id = self.idText.GetValue()
        if id.isdigit():
            id = int(id)
            if id in self.results:
                self.displayRecord(self.results.index(id + 1)) # Allow for 0 indexed array

    def disableButtons(self):
        self.searchButton.Enable(False)
        self.indexButton.Enable(False)
        self.idText.SetEditable(False)
        self.idText.SetForegroundColour(self.grey)
        self.updateButton.Enable(True)
        self.checkButton.Enable(False)
        self.editButton.Enable(False)
        self.prevButton.Enable(False)
        self.nextButton.Enable(False)
        self.descText.SetForegroundColour(self.black)
        self.descText.SetEditable(True)

    def enableButtons(self):
        self.searchButton.Enable(True)
        self.indexButton.Enable(True)
        self.idText.SetEditable(True)
        self.idText.SetForegroundColour(self.black)
        self.updateButton.Enable(False)
        self.checkButton.Enable(True)
        self.editButton.Enable(True)
        self.prevButton.Enable(True)
        self.nextButton.Enable(True)
        self.descText.SetForegroundColour(self.grey)
        self.descText.SetEditable(False)


def startThread(func, *args): # helper method to run a function in another thread
    thread = threading.Thread(target=func, args=args)
    thread.setDaemon(True)
    thread.start()

if __name__ == '__main__':
    try:
        application = wx.PySimpleApp()
        frame = MyFrame(None, "Doci")
        application.MainLoop()
    finally:
        del application
