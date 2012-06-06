from distutils.core import setup
import py2exe

setup(
    options = {'py2exe': {'bundle_files': 2, "dll_excludes": ["w9xpopen.exe"]}},
    windows = [
        {
            "script": "Doci.pyw",
            "icon_resources": [(0,"Doci.ico")],
        },
    ],
	console = [
        {
            "script": "Doci_Debug.py",
            "icon_resources": [(0,"Doci.ico")],
        },
    ],
    data_files=[('', ["msvcm90.dll", "msvcp90.dll", "msvcr90.dll", "Microsoft.VC90.CRT.manifest", "sqlite3.exe", "database.bat"])],
    name="Doci",
    version="0.4",
    author="Brinley Craig",
    description="Document Index Utility",
    #zipfile = None
) 