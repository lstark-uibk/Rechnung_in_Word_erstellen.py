from docxtpl import DocxTemplate
import openpyxl
from openpyxl.styles import Border
import pandas as pd
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx import Document
from docx.shared import Cm, Pt
import os
import numpy as np
import datetime
import pprint
from Helfer_Objekte import (check_invoice_archive,question_next_invoice_number, select_client, get_items, save_to_archive, ask_to_save,
                                     stringsandyear_topath, stringsandinvoicenumber_topath)
import tkinter as tk


def make_invoice_praxis(config,method):
    ### parameters:
    print("------------------------------------------------------------")
    print(f"Run method: {method} with user: {config['username']}")
    clientdata_path = config["data"]["clientdata"]
    servicedata_path =  config["data"]["servicedata"]
    template_path = config["outputmethods"][method]["template_path"]
    cashbookdir = config["cashbook"]["cashbookbuchdir"]
    cashbookfilenamestructure = config["cashbook"]["cashbookfilenamestructure"]
    cashbooktemplate_path = config["cashbook"]["template_path"]
    invoiceoutputdir_namestructure = config["output"]["directorynamestructure"]
    invoiceoutputdir_suppath = config["output"]["suppath"]
    invoice_name = config["outputmethods"][method]["invoice_name"]
    invoicenumber_pattern_names = config["invoicenumber"]["invoicenumber_pattern_names"]
    invoicenumber_pattern = config["invoicenumber"]["invoicenumber_pattern"]
    doctype = config["outputmethods"][method]["doctype"]

    #readin
    allhourdata = pd.read_excel(servicedata_path, parse_dates=[0],sheet_name="Stundendaten")
    services = pd.read_excel(servicedata_path, sheet_name = "Leistungen")

    allclientdata = pd.read_excel(clientdata_path, index_col=0, header=None, sheet_name=None)


    #select which client
    allclientsnames = list(allclientdata.keys())
    allclientsnames_sheetnames = [x for x in allclientsnames if x != "Vorlage"]
    def getname(x):
        try:
            return allclientdata[x][1]["Name"]
        except: pass
    allclientsnames = [getname(x) for x in allclientsnames_sheetnames if getname(x) is not None]

    print("All client names are:")
    pprint.pprint(allclientsnames)
    allclientsnames.sort()
    clientname = select_client(   allclientsnames)
    # clientname = "Martina Test"
    print("Selected client is:")
    print(clientname)
    print("------------------------------------------------------------")

    namehourdata = allhourdata[allhourdata["Name"] == clientname]

    # check whether there was already an invoice for this person in the last 3 years (we want to see it in the calendar)
    invoicetime_last = 0
    for yearback in [2,1,0]:
        yearcheck = datetime.datetime.today().year - yearback
        cashbookfilename = stringsandyear_topath(cashbookfilenamestructure,yearcheck)
        cashbook_path = os.path.join(cashbookdir,cashbookfilename)
        try:
            print(f"Check cashbook in {cashbook_path}")
            archive =pd.read_excel(cashbook_path)
            invoicetime_last = archive.iloc[:,1][archive.iloc[:,2] == clientname].values[-1]
            print(f"got last invoice time in year {yearcheck}: {invoicetime_last}")
        except:
            print(f"no last invoice times in year {yearcheck}")

    #select with a calendar
    loop = True
    while loop:
        print("------------------------------------------------------------")
        print("Select the items for invoice")
        selected_data,invoice_start_date,invoice_end_date = get_items(clientname,namehourdata,services,invoicetime_last, config["outputmethods"][method]["defaultservice"])
        # selected_data,invoice_start_date,invoice_end_date = namehourdata, namehourdata.Datum.min(),namehourdata.Datum.max()

        print(f"Ich nehme alle Termine von {clientname} ab: {invoice_start_date} bis zum {invoice_end_date}" )
        print("The data for this person is")
        pprint.pprint(selected_data)
        print("------------------------------------------------------------")

        # namehourdata = namehourdata[(namehourdata.Datum >= invoice_start_date)&(namehourdata.Datum <= invoice_end_date)]   #delete everything brfore lastinvoicegroup
        namehourdata = selected_data
        namehourdata = namehourdata.sort_values(by="Datum")

        # too many entries not important
        # if namehourdata.shape[0] > 10:
        #     print("too many entries")
        #
        #     def too_many_entries():
        #         # Display an alert message box with an "Okay" button
        #         tk.messagebox.showinfo("Fehler",
        #                                f"Es sind zu viele daten ausgewählt: {namehourdata.shape[0]}")
        #     too_many_entries()
        #else:
        loop = False

    # now fix the year of which the invoice is made
    year_of_invoice = invoice_end_date.year
    print(f"Year to link this invoice to: {year_of_invoice}")

    #now fix the output path and check whether it exists
    outputdir = stringsandyear_topath(invoiceoutputdir_namestructure, year_of_invoice)
    outputdir_path = os.path.join(invoiceoutputdir_suppath, outputdir)
    if not os.path.isdir(outputdir_path):
        os.mkdir(outputdir_path)
    else:
        print(f"We already have a output directory {outputdir_path}")

    #now get the invoice archive or make new archive and check for invoice numbers
    archive_which_invoices_name = stringsandyear_topath(cashbookfilenamestructure,year_of_invoice)
    if cashbookdir == "outputdir":
        archive_which_invoices_path = os.path.join(outputdir_path, archive_which_invoices_name)
    else:
        archive_which_invoices_path = os.path.join(cashbookdir, archive_which_invoices_name)

    lastinvoice_num = check_invoice_archive(year_of_invoice, outputdir_path, archive_which_invoices_path,
                                            cashbooktemplate_path, invoicenumber_pattern= invoicenumber_pattern)
    print("------------------------------------------------------------")
    print(f"Last invoice number: {lastinvoice_num}")

    # check whether invoice number is okay
    thisinvoicenumber = question_next_invoice_number(year_of_invoice,lastinvoice_num,invoicenumber_pattern,invoicenumber_pattern_names
                                                     )
    print(f"This invoicenumber: {thisinvoicenumber}")
    print("------------------------------------------------------------")

    # since we now have the year and the invoicenumber we set outputfilepath
    filename = stringsandinvoicenumber_topath(invoice_name,thisinvoicenumber,clientname, datetime.date.today().strftime('%d_%m_%Y'))
    outputfile_path = os.path.join(outputdir_path, filename)
    print(f"Now i can create the outputdata filepaths:   \ncassabook: {archive_which_invoices_path}\ninvoice: {outputfile_path}")

    # data processing
    clientdata = allclientdata[clientname].to_dict()[1]

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
    if pd.isna(clientdata["Versicherungsnummer"]):
        clientdata["Versicherungsnummertext"] = ""
    else:
        clientdata["Versicherungsnummertext"] = f"Versicherungsnummer: {clientdata["Versicherungsnummer"]}"
    clientdata["Heute"] = datetime.date.today().strftime("%d.%m.%Y")
    clientdata["Wordkindtext"] = ""
    clientdata["HerrFrau"] = ""
    clientdata["Behandlungszeitraum"] = f"{invoice_start_date.strftime('%d.%m.%Y')} - {invoice_end_date.strftime('%d.%m.%Y')}"
    clientdata["Ort"] = config["location"]

    if clientdata["Geschlecht"] == "m":
        clientdata["HerrFrau"] = "Herr"
    if clientdata["Geschlecht"] == "w":
        clientdata["HerrFrau"] = "Frau"

    clientdata["Geburtstag"] = clientdata["Geburtstag"].strftime('%d.%m.%Y')
    try:
        clientdata["Gültige Genehmigung Land Tirol ab"] = clientdata["Gültige Genehmigung Land Tirol ab"].strftime('%d.%m.%Y')
    except: pass
    clientdata["Typeinvoice"] = config["outputmethods"][method]["typeinvoicetext"]
    clientdata_list = [[key,clientdata[key]] for key in clientdata]
    addedhourdata = []
    save_or_not = ask_to_save(clientdata_list, namehourdata, services, addedhourdata)
    for new_hourdata in addedhourdata:
        namehourdata = pd.concat([namehourdata,new_hourdata])

    totalamount = 0
    # now to configure the input table
    if config["outputmethods"][method]["positionsinvoicetype"] == "all entries":
        # so here the wordtable should be: internal number,date, servicetext, hours, hourly rate, sum if
        positionsinvoice = namehourdata
        amountpersession = (positionsinvoice["Minuten"].astype(float) * positionsinvoice["Stundensatz"].astype(float)) / 60
        amountpersession_str = amountpersession.apply(lambda x: '{:.2f}'.format(x).replace('.', ',') + " €")
        amountpersession_str = amountpersession_str.to_frame(name="Betrag_pro_Einheit")
        description = positionsinvoice['Leistung'].map(services.set_index('Leistung')['Beschreibung'])
        description = description.to_frame(name="Beschreibung")

        datums =   positionsinvoice["Datum"].apply(lambda x: x.strftime("%d.%m.%Y"))

        positionsinvoice  = pd.concat([
            positionsinvoice["Leistung"],
            datums,
            description,
            positionsinvoice["Minuten"].apply(lambda x: str(x) + " min"),
            positionsinvoice["Stundensatz"].apply(lambda x: str(x) + " €"),
            amountpersession_str

        ], axis=1)
        print(f"++++++++++{positionsinvoice}")
        totalamount = sum(np.array(amountpersession))
    elif config["outputmethods"][method]["positionsinvoicetype"] == "summary":

        positionsinvoice = namehourdata
        description = positionsinvoice['Leistung'].map(services.set_index('Leistung')['Beschreibung'])
        description = description.to_frame(name="Beschreibung")
        positionsinvoice = pd.concat([positionsinvoice,description],axis = 1)

        gathersummarypositions = []
        totalamounts = []
        groupedbyservice = positionsinvoice.groupby(by="Leistung")
        for servicename,datathisservice in groupedbyservice:
            groupedbyminutes = datathisservice.groupby(by="Minuten")
            for minutes, datathisminutes in groupedbyminutes:
                groupedbyhourlyrate = datathisminutes.groupby(by="Stundensatz")
                amount_hours = groupedbyhourlyrate.size().iloc[0]
                for hourlyrate, datathishourlyrate in groupedbyhourlyrate:
                    descriptionthis = f"{datathishourlyrate['Beschreibung'].iloc[0]}" #{str(minutes)} min"
                    if datathishourlyrate.shape[0] < 2: # if there is only one entry for this, add the date
                        descriptionthis += f" ({datathishourlyrate['Datum'].iloc[0].strftime('%d.%m.%Y')})"
                    minutesthis = '{:.0f}'.format(round(datathishourlyrate["Minuten"].astype(float))[0]).replace('.', ',') + ' min'
                    hourlyratethis = '{:.2f}'.format(float(hourlyrate)).replace('.', ',') + ' €'
                    amountpersession = (datathishourlyrate["Minuten"].astype(float) * datathishourlyrate["Stundensatz"].astype(float)) / 60
                    totalamounts.append(amountpersession)
                    totalsumthis = '{:.2f}'.format(amountpersession.sum()).replace('.', ',') + ' €'
            gathersummarypositionslist = [servicename,descriptionthis,amount_hours,minutesthis,totalsumthis]
            gathersummarypositions.append(gathersummarypositionslist)

        # sehr maßgeschneidert!!
        positionsinvoicecols = config["outputmethods"][method]["positionsinvoicecols"]
        positionsinvoice = pd.DataFrame(gathersummarypositions,columns=positionsinvoicecols)

        for amount in totalamounts:
            totalamount += amount.sum()

        # Behandlung sollte minuten beinhalten

    if config["outputmethods"][method]["Ausgleichszulage"]["exists"]:

        ausgleichpercent = config["outputmethods"][method]["Ausgleichszulage"]["percentage"]

        descriptionausgleichszulage = '+ ' + '{:.1f}'.format(float(ausgleichpercent)).replace('.', ',') + ' % Ausgleichszulage'
        ausgleichamount = totalamount * ausgleichpercent / 100
        amountausgleichszulagestr = '+ ' + '{:.2f}'.format(float(ausgleichamount)).replace('.', ',') + ' €'

        ausgleichszulagerow = [["",descriptionausgleichszulage,"","",amountausgleichszulagestr]]
        ausgleichszulagerow = pd.DataFrame(ausgleichszulagerow, columns=positionsinvoicecols)
        positionsinvoice = pd.concat([positionsinvoice,ausgleichszulagerow])
        totalamount += ausgleichamount



    # insert total amount into tables[1]
    clientdata["Stundeninfo"] = positionsinvoice
    print("------------------------------------------------------------")
    print("Die Patientendaten sind:")
    pprint.pprint(clientdata)

    clientdata = {key.replace(" ", ""): value for key, value in clientdata.items()}
    print("------------------------------------------------------------")
    print("Preparing the output files before saving.")
    if doctype == "docx":
        # input the client data  in word
        doc = DocxTemplate(template_path)
        roundedtotal = round(totalamount, 2)
        totalamountstring = '{:.2f}'.format(float(roundedtotal)).replace('.', ',') + " €"

        clientdata["Endsumme"] = totalamountstring
        clientdata["Stundentabelle"] = positionsinvoice.to_dict(orient="records")
        print(f"++++++++++++++{clientdata['Stundentabelle']}")
        try:
            terminetable_cols = config["outputmethods"][method]["table_termine"]["cols"]
            termine_table = namehourdata[terminetable_cols]
            for col in termine_table.select_dtypes(include=["datetime64[ns]"]).columns:
                termine_table[col] = termine_table[col].dt.strftime('%d.%m.%Y')
            termine_table[terminetable_cols[1]] = pd.to_datetime(termine_table.loc[:,terminetable_cols[1]], format="%H:%M:%S").dt.strftime("%H:%M")
            clientdata["Termintabelle"] = termine_table.to_dict(orient="records")
        except: pass
        doc.render(clientdata)
        # outputfile_path = "/home/leander/Documents/pycharm_projects/Abrechnungsprogramm/Rechnungen 2025/Rechnung_test.docx"
        try_saving = True
        while try_saving:
            try:
                doc.save(outputfile_path)
                try_saving = False
            except Exception as e:

                def show_alert():
                    # Display an alert message box with an "Okay" button
                    tk.messagebox.showinfo("Fehler",
                                           f"Ich konnte das Word für die Rechnung nicht speichern. Schließe zuerst die Datei {outputfile_path}")

                print("error in saving archive")
                print(e)
                root = tk.Tk()
                root.withdraw()  # Hide the main window

                # Show the alert
                show_alert()

                print("The alert has been dismissed. Continuing with the rest of the code...")
                root.quit()



    if doctype == "xlsx":
        invoice = openpyxl.load_workbook(template_path)
        invoice_sheet = invoice ['Rechnung']
        excelsheet_locs = {"Name": ("C", 10),
                           "Adresse": ("C", 12),
                           "Stadt": ("C", 13),
                           "Heute": ("I", 14),
                           "Rechnungsnummer": ("I", 15),
                           "Versicherungsnummer": ("I", 17)}
        usevalues = ["Name", "Adresse", "Stadt", "Heute", "Rechnungsnummer"]
        for value in usevalues:
            location = f"{excelsheet_locs[value][0]}{excelsheet_locs[value][1]}"
            invoice_sheet[location] = f"{clientdata[value]}"
        if not np.isnan(clientdata["Versicherungsnummer"]):
            location = f"{excelsheet_locs['Versicherungsnummer'][0]}{excelsheet_locs['Versicherungsnummer'][1]}"
            invoice_sheet[location] = f"{clientdata['Versicherungsnummer']}"
        else:
            print("no insurance number")
            location = f"H{excelsheet_locs['Versicherungsnummer'][1]}"
            invoice_sheet[location] = ""
            location = f"{excelsheet_locs['Versicherungsnummer'][0]}{excelsheet_locs['Versicherungsnummer'][1]}"
            invoice_sheet[location].border = Border()
            location = f"{'K'}{excelsheet_locs['Versicherungsnummer'][1]}"
            invoice_sheet[location].border = Border()

        leistung_text = "Logopädie"
        if "Selbstbehalt" in clientdata.keys():
            if clientdata["Selbstbehalt"] == "ja":
                leistung_text = "Logopädie Selbstbehalt"

        firstrows_hourdata = {"Datum": ("B", 22), "Leistungsbez": ("D", 22), "Preis/Einh.": ("G", 22),"Preis/Einh.": ("G", 22),"Sum":("I", 22)}

        i = 0
        for row, session in positionsinvoice.iterrows():
            invoice_sheet[f"{firstrows_hourdata['Datum'][0]}{firstrows_hourdata['Datum'][1] + i}"] = session["Datum"]
            invoice_sheet[f"{firstrows_hourdata['Datum'][0]}{firstrows_hourdata['Datum'][1] + i}"].number_format = 'DD.MM.YYYY'
            invoice_sheet[f"{firstrows_hourdata['Leistungsbez'][0]}{firstrows_hourdata['Leistungsbez'][1] + i}"] = f"{leistung_text} {str(round(session['Minuten'])).replace('.',',')} min"
            invoice_sheet[f"{firstrows_hourdata['Preis/Einh.'][0]}{firstrows_hourdata['Preis/Einh.'][1] + i}"] = float(clientdata["Stundensatz"])
            invoice_sheet[f"{firstrows_hourdata['Sum'][0]}{firstrows_hourdata['Sum'][1] + i}"] = round((session['Minuten'] / 60)*float(clientdata["Stundensatz"]),2)

            i += 1

    # now ask to save



    if save_or_not:
        safe_docs = True
        if safe_docs:
            if doctype == "docx":
                doc.save(outputfile_path)
                print(f"Saved word file to {outputfile_path}")

            if doctype == "xlsx":
                invoice.save(outputfile_path)
                print(f"Saved excel file to {outputfile_path}")

        #write what I did in the archive
        try_saving = True
        while try_saving:
            try:
                print("------------------------------------------------------------")
                save_to_archive(thisinvoicenumber,datetime.datetime.today(),clientname,invoice_start_date,invoice_end_date,totalamount,archive_which_invoices_path)
                try_saving = False
            except Exception as e:
                def show_alert():
                    # Display an alert message box with an "Okay" button
                    tk.messagebox.showinfo("Fehler", f"Ich konnte die Excel nicht speichern. Schließe zuerst die Datei {archive_which_invoices_path}")
                print("error in saving archive")
                print(e)
                root = tk.Tk()
                root.withdraw()  # Hide the main window

                # Show the alert
                show_alert()

                # Continue with the rest of the code after the alert
                print("The alert has been dismissed. Continuing with the rest of the code...")

                # Close the Tkinter application
                root.quit()

        print(f"\n Die Rechnung wurde in einem Word Dokument erstellt, zu finden unter Desktop -> Rechnungen-Verknüpfung -> Jahr {year_of_invoice}")

        if os.name == 'posix':
            print("This system is Linux or another Unix-like system.")
            import subprocess
            from sys import platform
            if platform == 'darwin': #apple
                subprocess.call(['open', archive_which_invoices_path])
                subprocess.call(['open', outputfile_path])

            else:
                subprocess.run(['xdg-open', archive_which_invoices_path])
                subprocess.run(['xdg-open', outputfile_path])

        else:
            print("This system is not Linux.")
            os.startfile(archive_which_invoices_path)
            os.startfile(outputfile_path)
    else:
        print("Didnot save anything")
    from Rechnung_erstellen import main
    main()