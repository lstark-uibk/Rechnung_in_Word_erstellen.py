from pathlib import Path
from docxtpl import DocxTemplate
import openpyxl
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


base_dir = current_directory = os.getcwd()
parent_dir = os.path.dirname(base_dir)
supparentdir = os.path.dirname(parent_dir)
template_path = os.path.join(parent_dir,"Vorlage.docx")
excel_template_path = os.path.join(parent_dir,"Jahresübersicht_Vorlage.xlsx")
allhourdata_path = os.path.join(parent_dir ,"Stundendaten.xlsx")
allclientdata_path = os.path.join(parent_dir ,"PatientInneninformationen.xlsx")
invoicenumber_path = os.path.join(parent_dir ,"Metadaten//Rechnungsnummern.txt")
outputdir_path = 0
archive_which_invoices_path = 0


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


clientname = Einfügen_Routine.ask_many_multiple_choice_question(
    "Von welcher Person willst du die Rechnung ausdrucken?",
    allclientsnames
)
# clientname = "Emma Essl"
print("Ich mache die Rechnung für die Person:")
print(clientname)
newclient = False
if clientname == newclienttext:
    newclient = True
    #execute the input Routine
    print("Dann lass uns einen Eintrag in der KlientInnendaten Datei anlegen: ")
    clientdata = Einfügen_Routine.input_new_person(allclientdata_path)
    clientname = clientdata["name"]
    print("Dann können wir auch hier gleich die letzten Stundendaten eintrage:")
    namehourdata = Einfügen_Routine.insert_hourdata(allhourdata_path, clientname)
    namehourdata["Minuten"] = [int(x) for x in namehourdata["Minuten"]]



# prepare data
if newclient == False:
    namehourdata = allhourdata[allhourdata["Name"] == clientname] #select only the hours of the given name
    #namehourdata["Datum"] = list(map(lambda x: dateutil.parser.parse(x), namehourdata["Datum"]))



#select with a calendar
invoice_start_date, invoice_end_date = Einfügen_Routine.get_date(namehourdata)
invoice_start_date = pd.to_datetime(invoice_start_date)
invoice_end_date = pd.to_datetime(invoice_end_date)
# invoice_start_date, invoice_end_date = datetime.datetime(2023,1,6), datetime.datetime(2023,2,27)

print("Ich nehme alle Termine von " + clientname + " ab: " + invoice_start_date.strftime("%d.%m.%Y") + " bis zum " + invoice_end_date.strftime("%d.%m.%Y") )

namehourdata = namehourdata[(namehourdata.Datum >= invoice_start_date)&(namehourdata.Datum <= invoice_end_date)]   #delete everything brfore lastinvoicegroup
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
result = Einfügen_Routine.get_selection("Möchtest du die Rechnungsnummer selbst eingeben?")
print(result)
# result = answer1


# make invoicenumber
lastinvoicenumber = int(invoicenumbers[-1][5:])
lastinvoiceyear = int(invoicenumbers[-1][:4])
if not result:

    thisinvoicenumber = f"{invoice_end_date.year}-{(lastinvoicenumber + 1):03}"
    print("Das ist die Rechnungsnummer: " + thisinvoicenumber)
if result:
    root = tk.Tk()
    Einfügen_Routine.change_place_of_window(root)
    # withdraw() will make the parent window disappear.
    root.withdraw()
    # shows a dialogue with a string input field
    thisinvoicenumber = tk.simpledialog.askstring('Rechnungsnummer',
                                               "Dann kannst du sie jetzt selber eingeben (in dem Format z.b. 2024-012):",
                                               parent=root)
    root.destroy()

year_of_invoice = invoice_end_date.year
print(f"Year to link this invoice to: {year_of_invoice}")
outputdir_path = os.path.join(supparentdir,f"{year_of_invoice}")

if not os.path.isdir(outputdir_path):
    os.mkdir(outputdir_path)
else:
    print(f"We already have a directory {outputdir_path}")


outputfile_path = os.path.join(outputdir_path, f"RE {thisinvoicenumber} {clientname} {datetime.date.today().strftime('%d_%m_%Y')}.docx")
archive_which_invoices_path = os.path.join(outputdir_path,f"Rechnungen {year_of_invoice}.xlsx")

print(f"Now i can create the outputdata filepaths:   \n{archive_which_invoices_path}\n{outputfile_path}")

if newclient == False:

    clientdata = allclientdata[clientname].to_dict()[1]
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
doc.save(outputfile_path)

## input the hour table in word
doc = Document(outputfile_path)
doc.tables #a list of all tables in document
# table nr. 0 is the data table and table nr. 1 is the sum table

# prepare the hour table
amountpersession = namehourdata["Minuten"].apply(lambda x: round(x * float(clientdata["Stundensatz"])/60, 1))
amountpersession = amountpersession.rename("Betrag_pro_Einheit")

wordtable = pd.concat([namehourdata["Datum"].apply(lambda x: x.strftime("%d.%m.%Y")), namehourdata["Minuten"].apply(lambda x: str(x) + " min"), amountpersession.apply(lambda x: "%0.2f" % x + " €") ], axis=1)
print("Die Stunden sind: " )
print(wordtable)


# insert the table in the Word document
for index, row in wordtable.iterrows():
    hourdatatable = doc.tables[0]   #so hourdatatable is the first table in the document
    data_row = hourdatatable.add_row().cells
    for i,(name,entry) in enumerate(row.items()):
            data_row[i].text = entry
#format it
for row in doc.tables[0].rows:
    row.height = Cm(0.8)
    row.alignment = WD_TABLE_ALIGNMENT.CENTER



#insert total amount into tables[1]
totalamount = sum(np.array(amountpersession))
totalamountstring = (str(totalamount)+"0").replace(".",",")
doc.tables[1].cell(0, 2).text = str(totalamount) + "0" + " €"

clientdata["Stundeninfo"] = wordtable
save_or_not = Einfügen_Routine.ask_to_save_with_dict(clientdata)
print(save_or_not)
if save_or_not:
    print(f"Save word file to {outputfile_path}")
    doc.save(outputfile_path)
    # In the end we update the metadata
    with open(invoicenumber_path, 'a') as f:
        if lastinvoiceyear == int(datetime.date.today().strftime("%Y")):
            f.write(thisinvoicenumber + "\n" )
        else:
            f.write(datetime.date.today().strftime("%Y") + "-" + str(1) + "\n")
    #if there is no archive excel yet create one
    if not os.path.exists(archive_which_invoices_path):
        print(f"Because there was no Archive file of the year create one at {archive_which_invoices_path}")
        wb = openpyxl.load_workbook(excel_template_path)
        # Select the worksheet
        sheet = wb["Tabelle1"]
        # Modify the cell
        sheet["A1"] = f"Rechnungen {year_of_invoice}"
        wb.save(archive_which_invoices_path)
    #write what I did in the archive
    summe = totalamountstring + " €"
    Einfügen_Routine.save_to_archive(thisinvoicenumber,clientdata["Heute"],clientname,invoice_start_date,invoice_end_date,summe,archive_which_invoices_path)

    print(f"\n Die Rechnung wurde in einem Word Dokument erstellt, zu finden unter Desktop -> Rechnungen-Verknüpfung -> Jahr {year_of_invoice}")


    os.startfile(archive_which_invoices_path)
    os.startfile(outputfile_path)
else:
    print("Didnot save anything")