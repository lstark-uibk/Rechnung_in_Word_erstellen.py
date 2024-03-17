from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx import Document
from docx.shared import Cm
import os
import numpy as np
from datetime import date
import datetime
from PyInquirer import prompt
import pprint
import Einfügen_Routine
import dateutil.parser
import tkinter as tk


base_dir = Path(__file__).parent

template_path = base_dir / "Vorlage.docx"
allhourdata_path = base_dir / "Stundendaten.xlsx"
allclientdata_path = base_dir / "PatientInneninformationen.xlsx"
outputdir_path = base_dir / "Rechnungen"
invoicenumber_path = base_dir / "Metadaten//Rechnungsnummern.txt"
archive_which_invoices_path = 0

outputdir_path.mkdir(exist_ok=True)

# read in
allhourdata = pd.read_excel(allhourdata_path, parse_dates=[0])
allclientdata = pd.read_excel(allclientdata_path, index_col=0, header=None, sheet_name=None)
invoicenumbers = pd.read_csv(invoicenumber_path, header=None)[0].values.tolist()

#select which client
allclientsnames = list(allclientdata.keys())
allclientsnames.sort()
newclienttext = "Keiner der obigen Personen -> mach neue Person"
allclientsnames.append(newclienttext)
# questions = [
#     {
#         'type': 'list',
#         'name': 'Name',
#         'message': 'Von welchem Klienten willst du die Rechnung erstellen?',
#         'choices': allclientsnames,
#     },
#     ]2023-079
# answers = prompt(questions)
# name = answers["Name"]


name = Einfügen_Routine.ask_many_multiple_choice_question(
    "Von welcher Person willst du die Rechnung ausdrucken?",
    allclientsnames
)
# name = "Emma Essl"
print("Ich mache die Rechnung für die Person:")
print(name)
newclient = False
if name == newclienttext:
    newclient = True
    #execute the input Routine
    print("Dann lass uns einen Eintrag in der KlientInnendaten Datei anlegen: ")
    clientdata = Einfügen_Routine.input_new_person(allclientdata_path)
    name = clientdata["Name"]
    print("Dann können wir auch hier gleich die letzten Stundendaten eintrage:")
    namehourdata = Einfügen_Routine.insert_hourdata(allhourdata_path, name)
    namehourdata["Minuten"] = [int(x) for x in namehourdata["Minuten"]]



# prepare data
if newclient == False:
    namehourdata = allhourdata[allhourdata["Name"] == name] #select only the hours of the given name
    #namehourdata["Datum"] = list(map(lambda x: dateutil.parser.parse(x), namehourdata["Datum"]))



#select with a calendar
invoice_start_date, invoice_end_date = Einfügen_Routine.get_date()
# invoice_start_date, invoice_end_date = datetime.datetime(2023,1,6), datetime.datetime(2023,2,27)

print("Ich nehme alle Termine von " + name + " ab: " + invoice_start_date.strftime("%d.%m.%Y") + " bis zum " + invoice_end_date.strftime("%d.%m.%Y") )
year_of_invoice = invoice_end_date.year
basedir_data = f"C:\\Users\\rosma\\Documents\\Rechnungen\\{year_of_invoice}"
if not os.path.isdir(basedir_data):
    os.mkdir(basedir_data)
else:
    print(f"We already have a directory {basedir_data}")

archive_which_invoices_path =f"{basedir_data}\\Rechnungen {year_of_invoice}.xlsx"
namehourdata = namehourdata[[invoice_end_date >= date >= invoice_start_date for date in namehourdata["Datum"]]]   #delete everything brfore lastinvoicegroup
# take out only the selecte dates

#get which invoicenumber
answer1 = "Nimm einfach die Nächste in der Reihe"
answer2 = "Ich möchte sie selber eingeben"
invoicenumberquestion_choices = [answer1, answer2]
# questions = [
#     {
#         'type': 'list',
#         'name': 'Name',
#         'message': 'Wie willst du die Rechnungsnummer bestimmen?',
#         'choices': invoicenumberquestion_choices,
#     },
#     ]
# answers = prompt(questions)
result = Einfügen_Routine.ask_multiple_choice_question(
    "Wie willst du die Rechnungsnummer bestimmen?",
    invoicenumberquestion_choices
)


# make invoicenumber
lastinvoicenumber = int(invoicenumbers[-1][5:])
lastinvoiceyear = int(invoicenumbers[-1][:4])
if result == answer1:

    thisinvoicenumber = f"{invoice_end_date.year}-{(lastinvoicenumber + 1):03}"
    print("Das ist die Rechnungsnummer: " + thisinvoicenumber)
if result == answer2:
    root = tk.Tk()
    Einfügen_Routine.change_place_of_window(root)
    # withdraw() will make the parent window disappear.
    root.withdraw()
    # shows a dialogue with a string input field
    thisinvoicenumber = tk.simpledialog.askstring('Rechnungsnummer',
                                               "Dann kannst du sie jetzt selber eingeben (in dem Format z.b. 2024-012):",
                                               parent=root)
    root.destroy()
#

if newclient == False:

    clientdata = allclientdata[name].to_dict()[1]
    print("Die Patientendaten sind:")
    pprint.pprint(clientdata)

#additional data on invoice
if clientdata["Kind"] == "nein":
    clientdata["BeideElternteile"] = clientdata["Name"]
    clientdata
elif clientdata["Kind"] == "ja":
    clientdata["BeideElternteile"] = str(clientdata["Elternteil1"]) + " und " + str(clientdata["Elternteil2"])
else:
    print("Nicht gegeben ob Kind oder Erwachsen")
    raise SystemExit

clientdata["Rechnungsnummer"] = thisinvoicenumber
clientdata["Heute"] = datetime.date.today().strftime("%d.%m.%Y")
if clientdata["Kind"] == "ja":
    if clientdata["Geschlecht"] == "m":
        clientfirstname = clientdata["Name"].split()[0]
        clientdata["Wordkindtext"] = "für ihren Sohn " + clientfirstname + ", geboren am " + clientdata["Geb."].strftime("%d.%m.%Y") + ","
    if clientdata["Geschlecht"] == "w":
        clientfirstname = clientdata["Name"].split()[0]
        clientdata["Wordkindtext"] = "für ihre Tochter " + clientfirstname + ", geboren am " + clientdata["Geb."].strftime("%d.%m.%Y") + ","



# input the client data  in word
doc = DocxTemplate(template_path)
doc.render(clientdata)
outputfile_path = os.path.join(base_dir / outputdir_path, "RE " + thisinvoicenumber + " " + name + " "+ datetime.date.today().strftime("%d.%m.%Y") + ".docx")
doc.save(outputfile_path)

## input the hour table in word
doc = Document(outputfile_path)
doc.tables #a list of all tables in document
# table nr. 0 is the data table and table nr. 1 is the sum table

# prepare the hour table
amountpersession = namehourdata["Minuten"].apply(lambda x: round(x * float(clientdata["Stundensatz"])/60, 1))
amountpersession.rename("Betrag pro Einheit")

wordtable = pd.concat([namehourdata["Datum"].apply(lambda x: x.strftime("%d.%m.%Y")), namehourdata["Minuten"].apply(lambda x: str(x) + " min"), amountpersession.apply(lambda x: "%0.2f" % x + " €") ], axis=1)
print("Die Stunden sind: " )
print(wordtable)


# insert the table in the Word document
for index, row in wordtable.iterrows():
    hourdatatable = doc.tables[0]   #so hourdatatable is the first table in the document
    data_row = hourdatatable.add_row().cells
    for i in range(0,len(row)):
            data_row[i].text = row[i]
#format it
for row in doc.tables[0].rows:
    row.height = Cm(0.8)
    row.alignment = WD_TABLE_ALIGNMENT.CENTER



#insert total amount into tables[1]
totalamount = sum(np.array(amountpersession))
totalamountstring = (str(totalamount)+"0").replace(".",",")
doc.tables[1].cell(0, 2).text = str(totalamount) + "0" + " €"

doc.save(outputfile_path)


# In the end we update the metadata
with open(invoicenumber_path, 'a') as f:
    if lastinvoiceyear == int(datetime.date.today().strftime("%Y")):
        f.write(thisinvoicenumber + "\n" )
    else:
        f.write(datetime.date.today().strftime("%Y") + "-" + str(1) + "\n")





#with open(lastinvoice_path, 'a') as f:  #add the current date to the txt file so that the invoices which we did now will not show up the next time
#     f.write(name + ", " + today.strftime("%Y-%m-%d") + "\n" )

#write what I did in the archive
summe = totalamountstring + " €"
Einfügen_Routine.save_to_archive(thisinvoicenumber,clientdata["Heute"],name,invoice_start_date,invoice_end_date,summe,archive_which_invoices_path)

print("\n Die Rechnung wurde in einem Word Dokument erstellt, zu finden unter Desktop -> Rechnungen-Verknüpfung -> Rechnungsprogramm -> Rechnung. ")


os.startfile(archive_which_invoices_path)
os.startfile(outputfile_path)
