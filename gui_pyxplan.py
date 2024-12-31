import tkinter as tk
import tkinter.scrolledtext as st

window=tk.Tk() #instancie une fenêtre
window.title(
    "PyXPlan : exportation et vérification des plannings FP"
    )
#window.geometry("600x600")
icon_file="Biplan.png"
ico=tk.PhotoImage(file=icon_file)
window.iconphoto(True,ico)


##window.tk.call(
##    'wm',
##    'iconphoto',
##    window._w,
##    tk.PhotoImage(icon_file)
##    )

frame_ics=tk.Frame(
    master=window,
    relief=tk.RIDGE,
    borderwidth=5
    )

btn_export_to_ics=tk.Button(
    master=frame_ics,
    text="Créer un fichier .ics depuis le planning FP",
    width=50,
    height=5
    )

frame_aurion=tk.Frame(
    master=window,
    relief=tk.RIDGE,
    borderwidth=5
    )
btn_check_aurion=tk.Button(
    master=frame_aurion,
    text="Vérifier cohérence avec Planning Aurion des élèves",
    width=50,
    height=5
    )

frame_stext=tk.Frame(
    master=window,
    relief=tk.RIDGE,
    borderwidth=5
    )

txt_details=st.ScrolledText(
    master=frame_stext,
    width=50,
    height=25
    )
#pour ajouter du texte dans le scrolledtext, txt_details.insert(tk.INSERT," ")

    
    
#formatte la fenetre (3 lignes et 1 colonne)    
for i in range(2): #formatte la fenetre (3 lignes et 1 colonne
    window.rowconfigure(i,weight=1,minsize=50)
    
window.columnconfigure(0,weight=1, minsize=75)


#place les frame et les widgets associés
frame_ics.grid(
    row=0,
    column=0,
    padx=5,
    pady=5
    )

btn_export_to_ics.pack()

frame_aurion.grid(
    row=1,
    column=0,
    padx=5,
    pady=5
    )
btn_check_aurion.pack()

frame_stext.grid(
    row=2,
    column=0,
    padx=5,
    pady=5
    )
txt_details.pack()
#txt_details.configure(state="disabled") #pour ne pas pouvoir laisser le curseur dans la fenêtre (readonly)

window.mainloop()




