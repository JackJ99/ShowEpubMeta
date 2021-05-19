from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import *
from tkinter import messagebox
import ebooklib
from ebooklib import epub
import os
import sys

global verplichteMetadata
global optioneleMetadata
global separator
global lastUsedPath
global opf

verplichteMetadata = ['identifier', \
                        'title', \
                        'language']
optioneleMetadata = ['creator', \
                        'contributor', \
                        'publisher', \
                        'rights', \
                        'coverage', \
                        'date', \
                        'source', \
                        'format', \
                        'type', \
                        'relation']
separator = ';'
lastUsedPath = ''
opf = '{http://www.idpf.org/2007/opf}'

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
        lijst = book.get_metadata('DC', tag)
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

def SchrijfMetadata(file, metadata, resultaatString):
    for kolom in maakLijstMetadata(file, metadata):
        resultaatString = resultaatString + ' ' + kolom
    return resultaatString

def GeefMetadata(bookfile):
    global verplichteMetadata
    global optioneleMetadata
    resultaat = ''
    resultaat = SchrijfMetadata(bookfile, verplichteMetadata, resultaat)
    resultaat = SchrijfMetadata(bookfile, optioneleMetadata, resultaat)
    return resultaat

def OpenOutput():
    global verplichteMetadata
    global optioneleMetadata
    global separator
    outstring = '\"bestand\"'
    for tag in verplichteMetadata:
        outstring = outstring + separator + '\"' + tag + '\"'
    for tag in optioneleMetadata:
        outstring = outstring + separator + '\"' + tag + '\"'
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

# def openProgressbar(rt):
#     pb_popup = Toplevel(rt)
#     pb_popup.geometry("300x50")
#     pb_popup.title("Bezig")
#     pb_popup.grid()
#     # progressbar
#     pbar = Progressbar(
#         pb_popup,
#         orient='horizontal',
#         mode='indeterminate',
#         length=280
#     )
#     # place the progressbar
#     pbar.grid(column=0, row=0, columnspan=2, padx=10, pady=20)
#     pbar.start()
#     return pbar
#
# def sluitProgressbar(pbar):
#     parent = pbar.winfo_parent()
#     pbar.stop()
#     return None

# Open een map
def openMap():
    global verplichteMetadata
    global optioneleMetadata
    for item in tv.selection():
        tv.selection_remove(item)
    for item in tv.get_children():
        tv.delete(item)
    m = askdirectory()
    if not m: return None
    # pb = openProgressbar(tv)
    for root, mappen, bestanden in os.walk(m, topdown=False):
        for b in bestanden:
            (r, e) = os.path.splitext(b)
            if e.upper() == '.EPUB':
                # print(os.path.join(root, b))
                md = []
                md.append(root)
                md.append(b)
                boek = os.path.join(root, b)
                md = md + maakLijstMetadata(boek, verplichteMetadata)
                md = md + maakLijstMetadata(boek, optioneleMetadata)
                tv.insert('', 'end', values=tuple(md))
    # sluitProgressbar(pb)
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
    enkelvoudigeMetaFrame = Frame(pu)
    # de geometrie opzetten voor de enkelvoudige metadata m.b.v. een grid
    enkelvoudigeMetaFrame.columnconfigure(0, weight=1)
    enkelvoudigeMetaFrame.columnconfigure(1, weight=3)
    enkelvoudigeMetaFrame.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7], weight=1)
    # alle tekstvariabelen voor de enkelvoudige metadata
    boektitel = StringVar(enkelvoudigeMetaFrame)
    boektaal = StringVar(enkelvoudigeMetaFrame)
    boekuitgever = StringVar(enkelvoudigeMetaFrame)
    boekrechten = StringVar(enkelvoudigeMetaFrame)
    boekdekking = StringVar(enkelvoudigeMetaFrame)
    boekbron = StringVar(enkelvoudigeMetaFrame)
    boekformaat = StringVar(enkelvoudigeMetaFrame)
    boekrelatie = StringVar(enkelvoudigeMetaFrame)
    # de labels en velden voor de enkelvoudige metadata
    label_titel = Label(enkelvoudigeMetaFrame, text='Titel:')
    label_titel.grid(column=0, row=0, sticky='w')
    entry_titel = Entry(enkelvoudigeMetaFrame, textvariable=boektitel)
    entry_titel.grid(column=1, row=0, sticky='e')
    title_lijst = md.get_metadata('DC', 'title')
    if len(title_lijst) > 0:
        title = title_lijst[0]
        boektitel.set(title[0])
    #
    label_taal = Label(enkelvoudigeMetaFrame, text='Taal:')
    label_taal.grid(column=0, row=1, sticky='w')
    entry_taal = Entry(enkelvoudigeMetaFrame, textvariable=boektaal)
    entry_taal.grid(column=1, row=1, sticky='e')
    language_lijst = md.get_metadata('DC', 'language')
    if len(language_lijst) > 0:
        language = language_lijst[0]
        boektaal.set(language[0])
    #
    label_uitgever = Label(enkelvoudigeMetaFrame, text='Uitgever:')
    label_uitgever.grid(column=0, row=2, sticky='w')
    entry_uitgever = Entry(enkelvoudigeMetaFrame, textvariable=boekuitgever)
    entry_uitgever.grid(column=1, row=2, sticky='e')
    publisher_lijst =  md.get_metadata('DC', 'publisher')
    if len(publisher_lijst) > 0:
        publisher = publisher_lijst[0]
        boekuitgever.set(publisher[0])
    #
    label_rechten = Label(enkelvoudigeMetaFrame, text='Rechten:')
    label_rechten.grid(column=0, row=3, sticky='w')
    entry_rechten = Entry(enkelvoudigeMetaFrame, textvariable=boekrechten)
    entry_rechten.grid(column=1, row=3, sticky='e')
    rights_lijst =  md.get_metadata('DC', 'rights')
    if len(rights_lijst) > 0:
        rights = rights_lijst[0]
        boekrechten.set(rights[0])
    #
    label_dekking = Label(enkelvoudigeMetaFrame, text='Dekking:')
    label_dekking.grid(column=0, row=4, sticky='w')
    entry_dekking = Entry(enkelvoudigeMetaFrame, textvariable=boekdekking)
    entry_dekking.grid(column=1, row=4, sticky='e')
    coverage_lijst =  md.get_metadata('DC', 'coverage')
    if len(coverage_lijst) > 0:
        coverage = coverage_lijst[0]
        boekdekking.set(coverage[0])
    #
    label_bron = Label(enkelvoudigeMetaFrame, text='Bron:')
    label_bron.grid(column=0, row=5, sticky='w')
    entry_bron = Entry(enkelvoudigeMetaFrame, textvariable=boekbron)
    entry_bron.grid(column=1, row=5, sticky='e')
    source_lijst =  md.get_metadata('DC', 'source')
    if len(source_lijst) > 0:
        source = source_lijst[0]
        boekbron.set(source[0])
    #
    label_formaat = Label(enkelvoudigeMetaFrame, text='Formaat:')
    label_formaat.grid(column=0, row=6, sticky='w')
    entry_formaat = Entry(enkelvoudigeMetaFrame, textvariable=boekformaat)
    entry_formaat.grid(column=1, row=6, sticky='e')
    format_lijst =  md.get_metadata('DC', 'format')
    if len(format_lijst) > 0:
        format = format_lijst[0]
        boekformaat.set(format[0])
    #
    label_relatie = Label(enkelvoudigeMetaFrame, text='Relatie:')
    label_relatie.grid(column=0, row=7, sticky='w')
    entry_relatie = Entry(enkelvoudigeMetaFrame, textvariable=boekrelatie)
    entry_relatie.grid(column=1, row=7, sticky='e')
    relation_lijst =  md.get_metadata('DC', 'relation')
    if len(relation_lijst) > 0:
        relation = relation_lijst[0]
        boekrelatie.set(relation[0])
    # plaats frame voor enkelvoudige metadata in de geometrie
    enkelvoudigeMetaFrame.grid(column=0, row=0, sticky='nw')
    return None

def bouwIdentifierFrame(rt, md, h, w):
    identifier_frame = Frame(rt)

    identifier_lb = Listbox(identifier_frame, height=h, width=w)

    identifier_ybar = Scrollbar(identifier_frame, orient=VERTICAL)
    identifier_ybar.pack(side=RIGHT, fill=Y)
    identifier_xbar = Scrollbar(identifier_frame, orient=HORIZONTAL)
    identifier_xbar.pack(side=BOTTOM, fill=X)

    identifier_lb.pack(side=LEFT)

    identifier_lb.configure(yscrollcommand=identifier_ybar.set)
    identifier_ybar.config(command=identifier_lb.yview)
    identifier_lb.configure(xscrollcommand=identifier_xbar.set)
    identifier_xbar.config(command=identifier_lb.xview)
    identifiers = md.get_metadata('DC', 'identifier')
    for identifier in identifiers:
        identifier_lb.insert(END, GeefOpties(identifier))

    identifier_frame.grid(column=1, row=1, padx=5, pady=5, sticky='sw')
    return None

def bouwCreatorFrame(rt, md, h, w):
    creator_frame = Frame(rt)

    creator_lb = Listbox(creator_frame, height=h, width=w)

    creator_ybar = Scrollbar(creator_frame, orient=VERTICAL)
    creator_ybar.pack(side=RIGHT, fill=Y)
    creator_xbar = Scrollbar(creator_frame, orient=HORIZONTAL)
    creator_xbar.pack(side=BOTTOM, fill=X)

    creator_lb.pack(side=LEFT)

    creator_lb.configure(yscrollcommand=creator_ybar.set)
    creator_ybar.config(command=creator_lb.yview)
    creator_lb.configure(xscrollcommand=creator_xbar.set)
    creator_xbar.config(command=creator_lb.xview)
    creators = md.get_metadata('DC', 'creator')
    for creator in creators:
        creator_lb.insert(END, GeefOpties(creator))

    creator_frame.grid(column=1, row=3, padx=5, pady=5)
    return None

def bouwContributorFrame(rt, md, h, w):
    contributor_frame = Frame(rt)

    contributor_lb = Listbox(contributor_frame, height=h, width=w)

    contributor_ybar = Scrollbar(contributor_frame, orient=VERTICAL)
    contributor_ybar.pack(side=RIGHT, fill=Y)
    contributor_xbar = Scrollbar(contributor_frame, orient=HORIZONTAL)
    contributor_xbar.pack(side=BOTTOM, fill=X)

    contributor_lb.pack(side=LEFT)

    contributor_lb.configure(yscrollcommand=contributor_ybar.set)
    contributor_ybar.config(command=contributor_lb.yview)
    contributor_lb.configure(xscrollcommand=contributor_xbar.set)
    contributor_xbar.config(command=contributor_lb.xview)
    contributors = md.get_metadata('DC', 'contributor')
    for contributor in contributors:
        contributor_lb.insert(END, GeefOpties(contributor))

    contributor_frame.grid(column=1, row=5, padx=5, pady=5)
    return None
def bouwDateFrame(rt, md, h, w):
    date_frame = Frame(rt)

    date_lb = Listbox(date_frame, height=h, width=w)

    date_ybar = Scrollbar(date_frame, orient=VERTICAL)
    date_ybar.pack(side=RIGHT, fill=Y)
    date_xbar = Scrollbar(date_frame, orient=HORIZONTAL)
    date_xbar.pack(side=BOTTOM, fill=X)

    date_lb.pack(side=LEFT)

    date_lb.configure(yscrollcommand=date_ybar.set)
    date_ybar.config(command=date_lb.yview)
    date_lb.configure(xscrollcommand=date_xbar.set)
    date_xbar.config(command=date_lb.xview)
    dates = md.get_metadata('DC', 'date')
    for date in dates:
        date_lb.insert(END, GeefOpties(date))

    date_frame.grid(column=1, row=7, padx=5, pady=5)
    return None

def bouwTypeFrame(rt, md, h, w):
    type_frame = Frame(rt)

    type_lb = Listbox(type_frame, height=h, width=w)

    type_ybar = Scrollbar(type_frame, orient=VERTICAL)
    type_ybar.pack(side=RIGHT, fill=Y)
    type_xbar = Scrollbar(type_frame, orient=HORIZONTAL)
    type_xbar.pack(side=BOTTOM, fill=X)

    type_lb.pack(side=LEFT)

    type_lb.configure(yscrollcommand=type_ybar.set)
    type_ybar.config(command=type_lb.yview)
    type_lb.configure(xscrollcommand=type_xbar.set)
    type_xbar.config(command=type_lb.xview)
    types = md.get_metadata('DC', 'type')
    for type in types:
        type_lb.insert(END, GeefOpties(type))

    type_frame.grid(column=1, row=9, padx=5, pady=5)
    return None

def bouwFrameMeervoudigeMeta(pu, md):
    meervoudigeMetaFrame = Frame(pu)
    # de geometrie van het frame voor meervoudige data opzetten
    meervoudigeMetaFrame.columnconfigure(0, weight=1)
    meervoudigeMetaFrame.columnconfigure(1, weight=9)
    meervoudigeMetaFrame.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], weight=1)
    listboxHoogte = 2
    listboxBreedte = 25
    identifier_label = Label(meervoudigeMetaFrame, text='Identifiers')
    identifier_label.grid(column=0, row=0, columnspan=2, sticky='w')
    bouwIdentifierFrame(meervoudigeMetaFrame, md, listboxHoogte, listboxBreedte)
    creator_label = Label(meervoudigeMetaFrame, text='Maker')
    creator_label.grid(column=0, row=2, columnspan=2, sticky='nw')
    bouwCreatorFrame(meervoudigeMetaFrame, md, listboxHoogte, listboxBreedte)
    contributor_label = Label(meervoudigeMetaFrame, text='Bijdrage van')
    contributor_label.grid(column=0, row=4, columnspan=2, sticky='w')
    bouwContributorFrame(meervoudigeMetaFrame, md, listboxHoogte, listboxBreedte)
    date_label = Label(meervoudigeMetaFrame, text='Data')
    date_label.grid(column=0, row=6, columnspan=2, sticky='w')
    bouwDateFrame(meervoudigeMetaFrame, md, listboxHoogte, listboxBreedte)
    type_label = Label(meervoudigeMetaFrame, text='Soort')
    type_label.grid(column=0, row=8, columnspan=2, sticky='w')
    bouwTypeFrame(meervoudigeMetaFrame, md, listboxHoogte, listboxBreedte)
    # plaats frame voor de meervoudige metadata in de geometrie
    meervoudigeMetaFrame.grid(column=0, row=1, sticky='sw')
    return None

def bouwFrameBeschrijving(pu, md):
    beschrijvingFrame = Frame(pu)
    # label en veld voor de beschrijving
    label_beschrijving = Label(beschrijvingFrame, text='Beschrijving')
    label_beschrijving.pack(side=TOP)
    text_frame = Frame(beschrijvingFrame)
    text_beschrijving = Text(text_frame, height=40, width=50)
    text_ybar = Scrollbar(text_frame, orient=VERTICAL)
    text_ybar.pack(side=RIGHT, fill=Y)
    text_xbar = Scrollbar(text_frame, orient=HORIZONTAL)
    text_xbar.pack(side=BOTTOM, fill=X)

    #text_beschrijving.pack(side=LEFT)

    text_beschrijving.configure(yscrollcommand=text_ybar.set)
    text_ybar.config(command=text_beschrijving.yview)
    text_beschrijving.configure(xscrollcommand=text_xbar.set)
    text_xbar.config(command=text_beschrijving.xview)
    # de beschrijving
    description_lijst = md.get_metadata('DC', 'description')
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
metadata = ['root', 'bestand'] + verplichteMetadata + optioneleMetadata
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
tv.column('root', width=0, stretch=False, anchor='nw')
tv.column('bestand', width=100, stretch=False, anchor='nw')
tv.column('identifier', width=100, stretch=False, anchor='nw')
tv.column('title',  width=100, stretch=False, anchor='nw')
tv.column('language', width=100, stretch=False, anchor='nw')
tv.column('creator',  width=100, stretch=False, anchor='nw')
tv.column('contributor',  width=100, stretch=False, anchor='nw')
tv.column('publisher',  width=100, stretch=False, anchor='nw')
tv.column('rights', width=100, stretch=False, anchor='nw')
tv.column('coverage', width=100, stretch=False, anchor='nw')
tv.column('date',  width=100, stretch=False, anchor='nw')
tv.column('source',  width=100, stretch=False, anchor='nw')
tv.column('format', width=100, stretch=False, anchor='nw')
tv.column('type',  width=100, stretch=False, anchor='nw')
tv.column('relation', width=100, stretch=False, anchor='nw')

tv.heading('root', text='', anchor='nw')
tv.heading('bestand', text='Bestand', anchor='nw')
tv.heading('identifier', text='Id.', anchor='nw')
tv.heading('title',  text='Titel', anchor='nw')
tv.heading('language', text='Taal', anchor='nw')
tv.heading('creator',  text='Maker', anchor='nw')
tv.heading('contributor',  text='Bijdrage van', anchor='nw')
tv.heading('publisher',  text='Uitgever', anchor='nw')
tv.heading('rights', text='Rechten', anchor='nw')
tv.heading('coverage', text='Dekking', anchor='nw')
tv.heading('date',  text='Datum', anchor='nw')
tv.heading('source',  text='Bron', anchor='nw')
tv.heading('format', text='Formaat', anchor='nw')
tv.heading('type',  text='Soort', anchor='nw')
tv.heading('relation', text='Relatie', anchor='nw')

tv.bind('<Double-Button-1>', ToonMetaData)

ybar.config(command=tv.yview)
xbar.config(command=tv.xview)

tv.pack(fill='y', expand=True)

# Start de loop
window.mainloop()
