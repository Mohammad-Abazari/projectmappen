import win32com.client
import sys, os
import string
import os
import win32com

print('-' *60)
print('--- Maak projectmappen aan in opgegeven map en Outlook ---')
print('-' * 60 + '\n')

print("Geef je persoonscode (kleine letters) op en druk op Enter:\n")

person = input()

if person == "pwl":
    mail = 'n.pouwels@abt.eu'
elif person == "rij":
    mail = 'j.reijers@abt.eu'
elif person == "mdk":
    mail = 'a.v.middelkoop@abt.eu'
elif person == "vbt":
    mail = 'm.verbaten@abt.eu'
elif person == "apn":
    mail = 'j.alphen@abt.eu'
elif person == "wnv":
    mail = 'l.wijnveld@abt.eu'
print('Geef de aanvrager op:')

print('[1] S&P \n[2] Balm\n[3] Kreeft\n[4] Vogel\n[5] Vogel_VVUV\n[6] Ervas\n\
[7] Overige opdrachtgevers\nKies nummer en druk op Enter:...')

keuze = int(input())

path_ = 'P:/158/1584%s/00_Kwaliteitsdossier/offerte'% keuze
project_code = "1584%s" % keuze


# Schrijf hyperlink om naar path
def change_path():
    global path_
    # change backslash
    a = ''
    for i in path_:
        if str(i) == "\\":
            a += '/'
        else:
            a += str(i)
    path_ = a


# os.walk loopt door alle mappen en bestanden. De eerste variabele is een
# list met alle mappen in een directory
def defwalk(_path_):
    b = []
    for i in os.walk(_path_):
        b.append(i)

    return b[0][1]

folders = defwalk(path_)   

# Zoekt laatste folder bepaalt het # en maakt nieuw #
try:
    last_folder = folders[-1]

    nr = ""
    for i in range(3):
        try:
            int(last_folder[i])
            nr += str(last_folder[i])
        except:
            pass

    nr = int(nr)
    nr += 1
    if nr <10:
        nr = "00" + str(nr)
    elif nr <100:
        nr = "0" + str(nr)
    else:
        nr = str(nr)


# Als er geen map is wordt nr. 001 aangemaakt.
except:
    nr = ""
    nr += '001'
 
print('\n\n' + '-'*60 + '\nNieuwe map krijgt fasenummer ' + nr + '.\n\n'\
'Als dit akkoord is, druk Enter.\n'\
'Klopt dit niet, geef dan het juiste fasenummer op.\n\n'\
'Druk op Enter of kies nieuw fasenummer:')

faseNummer = input()

if faseNummer:
    nr = faseNummer
    

print('-'*60 + '\nGeef titel nieuwe map op en druk op Enter.\nTitel:')
titel = input()

# Path naar nieuwe mapnummertje
folderpath = path_ + '/' + nr + ' - ' + titel      # Path is link naar specifieke map, of bestand
# Maak nieuwe map
os.mkdir(folderpath)

# Paths voor mappen in de nieuwe projectmap
fld_ontvangen = folderpath + '/' + 'Ontvangen bestanden'
fld_brk = folderpath + '/' + 'Berekeningen'
fld_kwaliteit = folderpath + '/' + 'Kwaliteit'
fld_indicatie = folderpath + '/' + 'Indicatie'

os.mkdir(fld_ontvangen)
os.mkdir(fld_brk)
os.mkdir(fld_kwaliteit)
os.mkdir(fld_indicatie)



############ Deel Outlook #############

# COM syntax om verbinding met Outlook te krijgen.
try:
    print("Verbinding met Outlook maken...")
    outlook = win32com.client.Dispatch("Outlook.Application")
   
    base = outlook.GetNamespace("MAPI")
   
except:
    print("Kan geen verbinding met Outlook krijgen."\
    '\nStart Outlook en probeer opnieuw.')
print("Verbonden met Outlook.\n")

# Iteratie door emailmappen
def findFolder(folderName,searchIn):
    try:
        lowerAccount = searchIn.Folders
        for x in lowerAccount:
            if x.Name == folderName:
                print('found it %s'%x.Name)
                objective = x
                return objective
        return None
    except Exception as error:
        print("Looks like we had an issue accessing the searchIn object")
        print (error)
        return None
        

one = 'Openbare mappen - %s' % mail
two = 'Alle openbare mappen'
three = 'ABT Projecten'
four = project_code[:3]
five = project_code
six = 'Postvak In'

Folder1 = findFolder(one, base)
Folder2 = findFolder(two, Folder1)
Folder3 = findFolder(three, Folder2)
Folder4 = findFolder(four, Folder3)
Folder5 = findFolder(five, Folder4)
Folder6 = findFolder(six, Folder5)

def addFolder(folderName_, current):
    # Maak nieuwe map aan
    current.Folders.Add(folderName_)
    # cd naar nieuwe map
    nieuw = findFolder(folderName_, current)
    nieuw.Folders.Add('Postvak IN')
    nieuw.Folders.Add('Postvak UIT')
    

folderName = nr + ' - ' + titel
addFolder(folderName, Folder6)

