from distutils.core import setup
import py2exe

exeversion = "0.25"
setup(
    options = {'py2exe': {'bundle_files': 2, "dll_excludes": ["w9xpopen.exe"]}},
    # windows = [
        # {
            # "script": "Doci.pyw",
            # "icon_resources": [(0,"Doci.ico")],
            # "other_resources": [(u"VERSION",1,exeversion)],
        # },
    # ],
	console = [
        {
            "script": "Doci.pyw",
            "icon_resources": [(0,"Doci.ico")],
            "other_resources": [(u"VERSION",1,exeversion)],
        },
    ],
    data_files=[('', ["msvcm90.dll", "msvcp90.dll", "msvcr90.dll", "Microsoft.VC90.CRT.manifest", "sqlite3.exe", "database.bat", "jquery.fixedheader.js"])],
    name="Doci",
    version=exeversion,
    author="Brinley Craig",
    description="Document Indexer",
    #zipfile = None
) 