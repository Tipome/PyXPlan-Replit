import sys
from cx_Freeze import setup, Executable

#Pour lancer la création de l'exe, taper dans la console "py setup.py build"
#inclus les autres fichiers non python
includefiles=["Biplan.ico","Biplan.png","LICENSE","lisezmoi.txt"]

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os",
                                  "openpyxl",
                                  "pyexcel",
                                  "pyexcel_xls",
                                  "pyexcel_xlsx",
                                  "datetime",
                                  "sys",
                                  "tkinter",
                                  "ics",
                                  "json",
                                  "openpyxl",
                                  "arrow",
                                  "pytz",
                                  "xls2xlsx",
                                  "re"],
                     "excludes": [],
                     "include_msvcr": True,  #skip error msvcr100.dll missing
                     "include_files" : includefiles
                     }

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"



setup(  name = "PyXPlan",
        version = "2024.1",
        description = "Exporte en ics le planning d'une promo depuis le planning FP",
        options = {"build_exe": build_exe_options},
        executables = [Executable("PyXPlan V4.py", icon="Biplan.ico", base=base)])
