def planning_sorted(calaurion,calpp,ddeb,dfin):
  #tri les planning pour le réduire à la FP  
  #entre les dates ddeb et dfin et ordre chronologique
  #renvoie une liste d'évts aurion, une liste FP élèves et une liste d'evts réunions
  #typ doit être "aurion" ou "fp" en fonction du calendrier 
  l_evts_aurion=list()
  l_evts_fp=list()
  l_evts_reunion=list()

  if calaurion!=None :
      txt="FORMATION PRATIQUE"
      for e in sorted(cal.events) :
          if e.begin.date()>=ddeb and e.begin.date()<=dfin:
              if txt in e.name.upper():
                  #si le texte est trouvé
                  l_evts_aurion.append(e)

  if calpp!=None :
      l_txt=["BILAN","ASTRAL","HARMO"]
      for e in sorted(cal.events):
          test=False
          if e.begin.date()>=ddeb and e.begin.date()<=dfin:
              for t in l_txt:
                  if t in str(e.name).upper():
                      test=True
                      l_evts_reunion.append(e)
            #si le texte est trouvé, renvoie true et ajoute l'evt à la liste des réunions
              if test==False:
                  l_evts_fp.append(e)
  ###################################################################################################"
                   #ne sert que pour débug

  save_liste(l_evts_aurion,"aurion trié.txt") #ne sert que pour débug
  save_liste(l_evts_fp,"fp trié.txt") #ne sert que pour débug
  save_liste(l_evts_reunion,"reunions trié.txt") #ne sert que pour débug
#####################################################################################################      

  return l_evts_aurion,l_evts_fp,l_evts_reunion




########################## old version ####################################


##def planning_sorted(cal,typ,ddeb,dfin):
##  #tri le planning pour le réduire à la FP 
##  #entre les dates ddeb et dfin et ordre chronologique
##  #renvoie une liste d'évts
##  #typ doit être "aurion" ou "fp" en fonction du calendrier 
##  levts=list()
##  utyp=typ.upper()
##  if utyp=="AURION":
##    txt="FORMATION PRATIQUE"
##    for e in sorted(cal.events) :
##      if e.begin.date()>=ddeb and e.begin.date()<=dfin:
##        if txt in e.name.upper():
##          #si le texte est trouvé
##          levts.append(e)
##    
##  if utyp=="FP":
##    l_txt=["BILAN","ASTRAL","HARMO"]
##    for e in sorted(cal.events) :
##      test=False
##      if e.begin.date()>=ddeb and e.begin.date()<=dfin:
##        for t in l_txt:
##          if t in str(e.name).upper():
##            test=True
##            #si le texte est trouvé, renvoie true
##        if test==False:
##            levts.append(e)
##  save_liste(levts,typ+" trié.txt") #ne sert que pour débug
##  return levts
