from distutils.core import setup
import py2exe

excludes = [
    "pywin",
    "pywin.debugger",
    "pywin.debugger.dbgcon",
    "pywin.dialogs",
    "pywin.dialogs.list",
    "win32com.server",
]

options = {
    "bundle_files": 1,
    "compressed"  : 1,                 # compress the library archive
    "excludes"    : excludes,
    "dll_excludes": ["w9xpopen.exe"],  # we don't need this
    "includes"    : ["lxml.etree", "lxml._elementpath", "gzip"],
}

t1 = dict(script="diffStudent.py",
          dest_base="diffStudent")
console = [t1]

setup(
    options = {"py2exe": options},
    console = console,
   )