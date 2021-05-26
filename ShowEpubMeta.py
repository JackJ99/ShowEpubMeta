from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import *
from tkinter import messagebox
import ebooklib
from ebooklib import epub
import os
import sys

global metadataLijst
global INDEXKEY
global INDEXNAMESPACE
global INDEXTYPE
global INDEXMANDOPT
global INDEXLIST
global INDEXDETAIL
global INDEXHEADING
global separator
global lastUsedPath
global opf
#                                                           HEADING
#                                            IN DETAIL VIEW   |
#                                       IN LIST VIEW    |     |
#                             MANDATORY/OPTIONAL   |    |     |
#                                     TYPE    |    |    |     |
#                           NAMESPACE    |    |    |    |     |
#                    KEY           |     |    |    |    |     |
metadataLijst = [
                  ['identifier',  'DC', 'M', 'M', 'Y', 'Y', 'Identificatie'], \
                  ['title',       'DC', 'E', 'M', 'Y', 'Y', 'Titel'], \
                  ['language',    'DC', 'E', 'M', 'Y', 'Y', 'Taal'], \
                  ['creator',     'DC', 'M', 'O', 'Y', 'Y', 'Maker'] ,\
                  ['contributor', 'DC', 'M', 'O', 'Y', 'Y', 'Bijdrage van'], \
                  ['publisher',   'DC', 'E', 'O', 'Y', 'Y', 'Uitgever'], \
                  ['rights',      'DC', 'E', 'O', 'Y', 'Y', 'Rechten'], \
                  ['coverage',    'DC', 'E', 'O', 'Y', 'Y', 'Dekking'], \
                  ['date',        'DC', 'M', 'O', 'Y', 'Y', 'Datum'], \
                  ['source',      'DC', 'E', 'O', 'Y', 'Y', 'Bron'], \
                  ['format',      'DC', 'E', 'O', 'Y', 'Y', 'Formaat'], \
                  ['type',        'DC', 'M', 'O', 'Y', 'Y', 'Soort'], \
                  ['relation',    'DC', 'E', 'O', 'Y', 'Y', 'Relatie'], \
                  ['description', 'DC', 'D', 'O', 'N', 'Y', 'Beschrijving'] \
                ]
INDEXKEY = 0
INDEXNAMESPACE = 1
INDEXTYPE = 2
INDEXMANDOPT = 3
INDEXLIST = 4
INDEXDETAIL= 5
INDEXHEADING = 6

separator = ';'
lastUsedPath = ''
opf = '{http://www.idpf.org/2007/opf}'

def geefHeading(mdkey):
    global metadataLijst
    global INDEXKEY
    global INDEXHEADING
    uit = ''
    #specials
    if mdkey == 'root':
        return ''
    if mdkey == 'bestand':
        return 'Bestand'
    for mdlitem in metadataLijst:
        if mdkey == mdlitem[INDEXKEY]:
            return mdlitem[INDEXHEADING]
    return uit

def maakLijstMetadata(bf, md):
    global opf
    try:
        book = epub.read_epub(bf)
    except:
        lijstmduit = [sys.exc_info()[0]]
        return lijstmduit
    lenopf = len(opf)
    lijstmduit = []
    for tag in md:
        lijst = book.get_metadata(tag[1], tag[0])
        # Dit maakt een LIST van TUPLES. Een tag kan meerdere voorkomens hebben.
        # Elk voorkomen is een voorkomen in de LIST.
        # Het TUPLE bestaat uit een STRING met de inhoud van de tag en een
        # - eventueel lege - DIRECTORY. Elke KEY-waarde in de DIRECTORY is een
        # attribuut van de tag. Bijv: de tag "creator" heeft als attribuut o.m.
        # "role", en dat attribuut kan dan weer de waarde "auteur" hebben, wat
        # als bij de KEY horende waarde in de DIRECTORY staat
        if len(lijst) > 0:
            # een lege tag slaan we over
            Occurrence = lijst[0]
            # voor de lijstoutput gebruiken we alleen het eerste LIST-voorkomen en
            # negeren de overige
            if not Occurrence[0] == None:
                # een lege tag-waarde kunnen we niet wegschrijven
                outstring = Occurrence[0]
            else:
                outstring = 'None'
            for key in list(Occurrence[1]):
                # de DIRECTORY vormen we om in een LIST die we kunnen doorlopen
                # op KEY's.
                plek = key.find(opf)
                # we verwijderen de opf-namespace uit de KEY
                if plek >= 0:
                    outstring = outstring + ' ' + key[lenopf:] + '=' + Occurrence[1][key]
                else:
                    outstring = outstring + ' ' + key + '=' + Occurrence[1][key]
                    # en we schrijven de KEY (zonder opf-namespace) en de
                    # bijbehorende waarde wegschrijven
            lijstmduit.append(outstring)
        else:
            lijstmduit.append('None')
    return lijstmduit

def OpenOutput():
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    global separator
    outstring = '\"bestand\"'
    for tag in metadataLijst:
        if tag[INDEXLIST] == 'Y':
            outstring = outstring + separator + '\"' + tag[INDEXHEADING] + '\"'
    return outstring

def SluitOutput():
    return '\"*** Einde lijst ***\"'

# accepteer een list en plaats de elementen daarin in een string, waarbij
# de elementen worden omgeven door aanhalingstekens en gescheiden door ;
def maakCsvRegelOp(t):
    global separator
    outstring = '\"' + os.path.join(t[0], t[1]) + '\"'
    for kolom in t:
        if t.index(kolom) > 1:
            outstring = outstring + separator + '\"' + kolom + '\"'
    return outstring

# neem de inhoud van het scherm over in een csv
# Geef het pad waarin de file is opgeslagen terug.
def fillMetadataList(file):
    f = open(file, 'w')
    dirf = os.path.dirname(file)
    f.write(OpenOutput() + '\n')
    for item in tv.get_children():
        f.write(maakCsvRegelOp(tv.item(item, 'values')) + '\n')
    f.write(SluitOutput() + '\n')
    f.close()
    return dirf

# Toon de help-info
def showHelp():
    print("nog niet geïmplementeerd" + '\n')
    return None

# Open een map
def openMap():
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    for item in tv.selection():
        tv.selection_remove(item)
    for item in tv.get_children():
        tv.delete(item)
    m = askdirectory()
    if not m: return None
    metadata = []
    for mdlitem in metadataLijst:
        if mdlitem[INDEXLIST] == 'Y':
            metadata.append((mdlitem[INDEXKEY], mdlitem[INDEXNAMESPACE]))
    # wijzig de cursor in een klokje
    tv.config(cursor="watch")
    tv.update()
    for root, mappen, bestanden in os.walk(m, topdown=False):
        for b in bestanden:
            (r, e) = os.path.splitext(b)
            if e.upper() == '.EPUB':
                # print(os.path.join(root, b))
                md = []
                md.append(root)
                md.append(b)
                boek = os.path.join(root, b)
                md = md + maakLijstMetadata(boek, metadata)
                tv.insert('', 'end', values=tuple(md))
    tv.config(cursor="arrow")
    tv.update()
    # cursor weer teruggezet
    return None


# sla de lijst met metadata op als CSV
def saveMetadataList():
    global lastUsedPath
    root = Tk()
    root.withdraw()
    # define options for opening
    options = {}
    options['parent'] = root
    options['defaultextension'] = ".csv"
    options['filetypes'] = [("Comma separated","*.csv"),("alle bestanden","*.*")]
    options['initialdir'] = lastUsedPath
    options['initialfile'] = ""
    options['title'] = "Sla lijst op"
    f = asksaveasfilename(**options)
    if len(f) == 0: return None
    lastUsedPath = fillMetadataList(f)
    return None

# haal de afzonderlijke opties uit de DIRECTORY
def GeefOpties(l):
    global opf
    if l[0]:
        uit = l[0]
    else:
        uit = 'None'
    lenopf = len(opf)
    for key in list(l[1]):
        # de DIRECTORY vormen we om in een LIST die we kunnen doorlopen
        # op KEY's.
        plek = key.find(opf)
        # we verwijderen de opf-namespace uit de KEY
        if plek >= 0:
            optie = key[lenopf:] + '=' + l[1][key]
        else:
            optie = key + '=' + l[1][key]
        uit = uit + ' ' + optie
    return uit

def bouwFrameEnkelvoudigeMeta(pu, md):
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    enkelvoudigeMetaFrame = Frame(pu)
    # de geometrie opzetten voor de enkelvoudige metadata m.b.v. een grid
    enkelvoudigeMetaFrame.columnconfigure(0, weight=1)
    enkelvoudigeMetaFrame.columnconfigure(1, weight=3)
    itemnr = 0
    boekvar = []
    boeklabel = []
    boekentry = []
    for mdlitem in metadataLijst:
        if mdlitem[INDEXDETAIL] == 'Y' and mdlitem[INDEXTYPE] == 'E':
            enkelvoudigeMetaFrame.rowconfigure([itemnr], weight=1)
            boekvar.append(StringVar(enkelvoudigeMetaFrame))
            # de labels en velden voor de enkelvoudige metadata
            boeklabel.append(Label(enkelvoudigeMetaFrame, text=mdlitem[INDEXHEADING]))
            boeklabel[itemnr].grid(column=0, row=itemnr, sticky='w')
            boekentry.append(Entry(enkelvoudigeMetaFrame, textvariable=boekvar[itemnr]))
            boekentry[itemnr].grid(column=1, row=itemnr, sticky='e')
            entry_lijst = md.get_metadata(mdlitem[INDEXNAMESPACE], mdlitem[INDEXKEY])
            if len(entry_lijst) > 0:
                entry = entry_lijst[0]
                boekvar[itemnr].set(entry[0])
            itemnr += 1
    # plaats frame voor enkelvoudige metadata in de geometrie
    enkelvoudigeMetaFrame.grid(column=0, row=0, sticky='nw')
    return None

def bouwMvSubFrame(rt, md, mdi, h, w, r):
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    mvSubframe = Frame(rt)
    mvSubframe_lb = Listbox(mvSubframe, height=h, width=w)
    mvSubframe_ybar = Scrollbar(mvSubframe, orient=VERTICAL)
    mvSubframe_ybar.pack(side=RIGHT, fill=Y)
    mvSubframe_xbar = Scrollbar(mvSubframe, orient=HORIZONTAL)
    mvSubframe_xbar.pack(side=BOTTOM, fill=X)
    mvSubframe_lb.pack(side=LEFT)
    mvSubframe_lb.configure(yscrollcommand=mvSubframe_ybar.set)
    mvSubframe_ybar.config(command=mvSubframe_lb.yview)
    mvSubframe_lb.configure(xscrollcommand=mvSubframe_xbar.set)
    mvSubframe_xbar.config(command=mvSubframe_lb.xview)
    items = md.get_metadata(mdi[INDEXNAMESPACE], mdi[INDEXKEY])
    for item in items:
        mvSubframe_lb.insert(END, GeefOpties(item))
    mvSubframe.grid(column=1, row=r, padx=5, pady=5, sticky='sw')
    return None

def bouwFrameMeervoudigeMeta(pu, md):
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    meervoudigeMetaFrame = Frame(pu)
    # de geometrie van het frame voor meervoudige data opzetten
    meervoudigeMetaFrame.columnconfigure(0, weight=1)
    meervoudigeMetaFrame.columnconfigure(1, weight=9)
    listboxHoogte = 2
    listboxBreedte = 25
    mvlabel = []
    itemnr = 0
    for mdlitem in metadataLijst:
        if mdlitem[INDEXDETAIL] == 'Y' and mdlitem[INDEXTYPE] == 'M':
            meervoudigeMetaFrame.rowconfigure([2 * itemnr, 2 * itemnr + 1], weight=1)
            mvlabel.append(Label(meervoudigeMetaFrame, text=mdlitem[INDEXHEADING]))
            mvlabel[itemnr].grid(column=0, row=(2 * itemnr), columnspan=2, sticky='w')
            bouwMvSubFrame(meervoudigeMetaFrame, md, mdlitem, listboxHoogte, listboxBreedte, 2 * itemnr + 1)
            itemnr += 1
    # plaats frame voor de meervoudige metadata in de geometrie
    meervoudigeMetaFrame.grid(column=0, row=1, sticky='sw')
    return None

def bouwFrameBeschrijving(pu, md):
    global metadataLijst
    global INDEXKEY
    global INDEXNAMESPACE
    global INDEXTYPE
    global INDEXMANDOPT
    global INDEXLIST
    global INDEXDETAIL
    global INDEXHEADING
    beschrijvingFrame = Frame(pu)
    # label en veld voor de beschrijving
    for mdlitem in metadataLijst:
        if mdlitem[INDEXTYPE] == 'D' and mdlitem[INDEXDETAIL] == 'Y' :
            label_beschrijving = Label(beschrijvingFrame, text=mdlitem[INDEXHEADING])
            namespace = mdlitem[INDEXNAMESPACE]
            key = mdlitem[INDEXKEY]
            break
    label_beschrijving.pack(side=TOP)
    text_frame = Frame(beschrijvingFrame)
    text_beschrijving = Text(text_frame, height=40, width=50)
    text_ybar = Scrollbar(text_frame, orient=VERTICAL)
    text_ybar.pack(side=RIGHT, fill=Y)
    text_xbar = Scrollbar(text_frame, orient=HORIZONTAL)
    text_xbar.pack(side=BOTTOM, fill=X)
    text_beschrijving.configure(yscrollcommand=text_ybar.set)
    text_ybar.config(command=text_beschrijving.yview)
    text_beschrijving.configure(xscrollcommand=text_xbar.set)
    text_xbar.config(command=text_beschrijving.xview)
    # de beschrijving
    description_lijst = md.get_metadata(namespace, key)
    if len(description_lijst) > 0:
        description = description_lijst[0]
        text_beschrijving.insert(END, description[0])
    text_beschrijving.pack(side=TOP)
    text_frame.pack(side=TOP)
    # plaats frame voor de beschrijving in de geometrie
    beschrijvingFrame.grid(column=1, row=0, rowspan=2, sticky='ne')
    return None

# haal alle metadata op van het geselecteerde boeken toon die op een
# pop-up dialoog
def ToonMetaData(event):
    geselecteerdItem = tv.selection()
    md = tv.item(geselecteerdItem, 'values')
    boek = os.path.join(md[0], md[1])
    try:
        boekmeta = epub.read_epub(boek)
    except:
        return None
    # popup opzetten
    boekpopup = Toplevel(window)
    boekpopup.geometry("750x700")
    boekpopup.title("Boek: " + md[1])
    # de geometrie van de popup opzetten
    boekpopup.columnconfigure(0, weight=1)
    boekpopup.columnconfigure(1, weight=2)
    boekpopup.rowconfigure(0, weight=1)
    boekpopup.rowconfigure(1, weight=1)
    # een frame voor de metadata die enkelvoudig voorkomen
    bouwFrameEnkelvoudigeMeta(boekpopup, boekmeta)
    # een frame voor de metadata die meervoudig voorkomen
    bouwFrameMeervoudigeMeta(boekpopup, boekmeta)
    # een frame voor de beschrijving
    bouwFrameBeschrijving(boekpopup, boekmeta)

    return None

# Sluit de applicatie af
def exitApp():
    window.quit()
    return None

# Top level window
window = Tk()

window.title("Epub metadata")
window.geometry('800x320')
# Zorg dat bij klik op het kruisje de standaard afsluiting wordt geforceerd
window.protocol("WM_DELETE_WINDOW", exitApp)
# De applicatie moet op de deleteknop kunnen reageren
# window.bind('<Delete>', removeFilename)

# create a menu
menu = Menu(window)
window.config(menu=menu)

filemenu = Menu(menu)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="Open map...", command=lambda: openMap())
filemenu.add_command(label="Save as...", command=lambda: saveMetadataList())
filemenu.add_separator()
filemenu.add_command(label="Exit", command=exitApp)

helpmenu = Menu(menu)
menu.add_cascade(label="Help", menu=helpmenu)
helpmenu.add_command(label="About...", command=showHelp)
# een frame toevoegen om de listbox in te plaatsen
frame = Frame(window)

# progressbar
#pb = Progressbar(frame, orient='horizontal', mode='indeterminate', length=100)
#pb.pack(expand=True)
#pb.grid(column=0, row=0, columnspan=2, padx=10, pady=20)
frame.pack(side = LEFT)

ybar = Scrollbar(frame, orient=VERTICAL)
ybar.pack(side=RIGHT, fill=Y)
xbar = Scrollbar(frame, orient=HORIZONTAL)
xbar.pack(side=BOTTOM, fill=X)

# Create treeview
metadata = ['root', 'bestand']
for mdlistitem in metadataLijst:
    if mdlistitem[INDEXMANDOPT] == 'M' and mdlistitem[INDEXLIST] == 'Y':
        metadata.append(mdlistitem[INDEXKEY])
for mdlistitem in metadataLijst:
    if mdlistitem[INDEXMANDOPT] == 'O' and mdlistitem[INDEXLIST] == 'Y':
        metadata.append(mdlistitem[INDEXKEY])
kolommen = tuple(metadata)
tvoptions = {}
tvoptions['columns'] = kolommen
tvoptions['displaycolumns'] = '#all'
tvoptions['height'] = 100
tvoptions['selectmode'] = 'browse'
tvoptions['show'] = 'headings'
tvoptions['yscrollcommand'] = ybar.set
tvoptions['xscrollcommand'] = xbar.set
tv = Treeview(frame, **tvoptions)
for mditem in metadata:
    if mditem == 'root':
        colwidth=0
    else:
        colwidth=100
    tv.column(mditem, width=colwidth, stretch=False, anchor='nw')
    tv.heading(mditem, text=geefHeading(mditem), anchor='nw')

tv.bind('<Double-Button-1>', ToonMetaData)

ybar.config(command=tv.yview)
xbar.config(command=tv.xview)

tv.pack(fill='y', expand=True)

# Start de loop
window.mainloop()
