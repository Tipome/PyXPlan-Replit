#!/usr/bin/env python3
#
#  tkgui_pyxplan.py
#  
#  Copyright 2024 olivier <olivier@olivier-MS-7817>
#  
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#  
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#  
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#  
#  

import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog
import tkinter.scrolledtext as st
import PyXPlan_main as pxp


class MainWindow(tk.Tk):
  def __init__(self, titre, size, iconfile="", theme = "default"):
    
    # main setup
    super().__init__()
    self.style = ttk.Style()
    self.style.theme_use(theme)
    self.title(titre)
    self.geometry(f'{size[0]}x{size[1]}')
    self.minsize(600,600)
##        self.configure(bg = "black")
    ico = tk.PhotoImage(file = iconfile)
    self.iconphoto(True, ico)

    # Menu
    self.create_menu_bar()

    
    # widgets
    self.mainFrame = MainFrame(self)

    
    #run
    self.mainloop()

    
  def create_menu_bar(self):
    
    menuBar = tk.Menu(self)
    menuAide = tk.Menu(menuBar, tearoff = 0)
    menuBar.add_cascade(label ="Aide", menu = menuAide)
    menuAide.add_command(label="Mode d'emploi", command = self.showHelp)
    self.config(menu = menuBar)

  def showHelp(self):
##    afficheAide(self)
    top = Toplevel(self)
##    top.geometry('400x600')
    nomfichier = "lisezmoi.txt"
    f = open(nomfichier,'r')
    texte = f.read()
    lbltxt = ttk.Label(top, text = texte)
    lbltxt.grid(row = 0, column = 0, sticky = 'nswe', padx = 10, pady = 10)
    

    
class MainFrame(ttk.Frame):
  def __init__(self, parent):
    super().__init__(parent, relief=tk.RIDGE, borderwidth=5)
    self.place(relwidth = 1, relheight = 1)
    
    self.promo = ""
    self.fpfile = ""
    self.aurionfile = ""
    self.choix = tk.BooleanVar() #option pour éventuellement choisir les semaines à extraire du planning FP 
    #self.choix.set(value=False) # par défaut = False

    self.create_widgets()
    
  def create_widgets(self):
    btn_browse_xl = ttk.Button(self, text = "Modifier fichier FP", command = self.browse_xls)

    #définit une variable stringvar (pour pouvoir changer sa valeur facilement) et l'affecte à un label
    self.text_lbl_ex = tk.StringVar() 
    labl_excel = ttk.Label(self, textvariable = self.text_lbl_ex, background="white", wraplength = 400)

    #définit une checkbox qui changera la valeur de self.choix
    cbtn_fp = ttk.Checkbutton(self, text = "Choisir période", variable = self.choix, onvalue = True, offvalue = False)

    
    labl_nompromo = ttk.Label(self, text = 'Nom de la promo :')
    btn_promo = ttk.Button(self, text = 'Changer PROMO', command = self.choose_promo)

    #définit une variable stringvar (pour pouvoir changer sa valeur facilement) et l'affecte à un label
    self.text_lbl_promo = tk.StringVar()
    labl_promo = ttk.Label(self, textvariable = self.text_lbl_promo, background="white")
    
    btn_browse_aurion = ttk.Button(self, text = "Modifier fichier Aurion", command = self.browse_aurion)

    #définit une variable stringvar (pour pouvoir changer sa valeur facilement) et l'affecte à un label
    self.text_lbl_aurion = tk.StringVar()
    labl_aurion = ttk.Label(self, textvariable = self.text_lbl_aurion, background="white", wraplength = 400)
    
    btn_generer_ics = ttk.Button(self, text = 'Générer ics consolidé', command = self.generate_ics)
    
    self.txt_log = st.ScrolledText(self, height = 25, width = 50, bg = "white", wrap = 'word')
    # pour ajouter du texte dans le scrolledtext, faire self.txt_log.insert(tk.INSERT, '   '+\n)
    
    

    # create the grid
    self.columnconfigure((0,1,2), weight = 1, uniform = 'a')
    self.rowconfigure((0,1,2,3,4), weight = 1, uniform = 'a')
    self.rowconfigure((5), weight = 8, uniform = 'a')
    

    # place the widgets
    labl_excel.grid(row = 0, column = 0 , sticky = 'ew', columnspan = 2, padx = 10, pady = 10)
    btn_browse_xl.grid(row= 0, column = 2, sticky = 'nswe', columnspan = 1, padx = 10, pady = 10)

    cbtn_fp.grid(row = 1, column = 1, columnspan = 2, sticky ='e', padx = 10, pady =10)
    
    labl_nompromo.grid(row = 2, column = 0 , columnspan = 1, padx = 10, pady = 10)
    labl_promo.grid(row = 2, column = 1 , sticky = 'ew', columnspan = 1, padx = 10, pady = 10)
    btn_promo.grid(row = 2, column = 2 , sticky = 'nsew', columnspan = 1, padx = 10, pady = 10)

    labl_aurion.grid(row = 3, column = 0 , sticky = 'ew', columnspan = 2, padx = 10, pady = 10)
    btn_browse_aurion.grid(row = 3, column = 2, sticky = 'nswe', columnspan = 1, padx = 10, pady = 10)
    
    btn_generer_ics.grid(row = 4, column = 1, sticky = 'nswe', columnspan = 1, padx = 10, pady = 10)
    
    self.txt_log.grid(row = 5, column = 0, sticky = 'nswe', columnspan=3, padx = 20, pady = 20)

  def browse_xls(self):
    # lance la commande de PyXPlan_main pour aller chercher le fichier FP
    # et écrit le nom du fichier dans le le lbl correspondant
    fname = pxp.browseFileFP()
    self.text_lbl_ex.set(fname)
    self.fpfile = fname
    

  def browse_aurion(self):
    # lance la commande de PyXPlan_main pour aller chercher le fichier Aurion
    # et écrit le nom du fichier dans le le lbl correspondant
    fname = pxp.browseFileAurion()
    self.text_lbl_aurion.set(fname)
    self.aurionfile = fname

  def choose_promo(self):
    # lance la commande PyXplan_main pour choisir la promo
    # et écrit le résultat dans le lbl correspondant
    promo = pxp.choixPromo()
    self.text_lbl_promo.set(promo)
    self.promo = promo

  def generate_ics(self):
    # génère un planning ics consolidé si planning Fp et Aurion renseignés
    # génère uniquement un ics du planning FP si pas de promo Aurion renseigné
    if (self.fpfile!=None) and (self.fpfile!="") :
      if (self.aurionfile!=None) and (self.aurionfile!="") :
        # exporte en ics la promo du planning FP
        fpics = pxp.fp_to_ics(self.txt_log, self.fpfile, self.promo)
        # génère le ics consolidé 
        generate = pxp.check_aurion_fp(self.txt_log, self.aurionfile, fpics)
      else :
        # si aurion non spécifié, exporte uniquement les evts de la promo du planning FP
        # si self.choix = True, porpose de choisir la période
        generate = pxp.fp_to_ics(self.txt_log, self.fpfile, self.promo,self.choix.get())
        messagebox.showinfo(
          title="Fin de l'extraction",
          message="Extraction de la promo "+self.promo+" terminée.\n Appuyer sur ENTRÉE pour quitter.")
    else :
      # si pas de planning FP, ne fait rien et renvoie un message d'info
      messagebox.showinfo(
        title="Petit problème ...",
        message="Il faut au moins sélectionner un plannig FP et une promo.\n Appuyer sur ENTRÉE pour quitter.")
        

  

  
def main(args):
  iconfile=os.path.join(os.getcwd(),"Biplan.png")
  app=MainWindow('PyXPlan V4', (600,800), iconfile)
  return 0


if __name__ == '__main__':
  sys.exit(main(sys.argv[1:]))
