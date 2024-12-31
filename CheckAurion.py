#!/usr/bin/env python3
# -*-coding:utf-8 -*

import os
import sys
#import json
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog
from datetime import date, timedelta, datetime
import pyexcel as pe
import openpyxl as op
from openpyxl.cell import MergedCell
import arrow
from ics import Calendar, Event
import datetime
import pytz



caurion=Calendar()
cpyxplan=Calendar()

repDest = str(os.getcwd())
filea=filedialog.askopenfilename(
    initialdir=repDest,
    title="Séléctionner le fichier ics exrait d'Aurion"
    )

with open(filea,encoding="utf-8",errors="ignore") as f:
    caurion=Calendar(f.read())
    
filep=filedialog.askopenfilename(
    initialdir=repDest,
    title="Séléctionner le fichier ics PyXPlan"
    )

with open(filep,encoding="utf-8",errors="ignore") as f:
    cpyxplan=Calendar(f.read())

listmatch=list()
listmis=list()

for ea in caurion.events :
    if ea.name.upper().find("FORMATION PRATIQUE")>-1 :
        astart=ea.begin
        match=False
        if astart<arrow.get(2023,11,29):
            for ep in cpyxplan.events :
                pstart=arrow.get(
                    ep.begin.astimezone(
                        pytz.timezone('Europe/Paris')
                        )
                    ) # convertit en heure locale pour comparer avec le fichier aurion qui est en local
                if astart.date()==pstart.date():
                    if astart.hour==pstart.hour:
                        match=True
                        listmatch.append(ep.name+str(ep.begin)+ea.name)
                        break
            if match==False :
                listmis.append(str(astart)+str(ea.begin))

print(listmatch,sep='/n')
print(listmis,sep="/n")
            
                    
            
            
    
