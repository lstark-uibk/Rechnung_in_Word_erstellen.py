import numpy as np
import openpyxl
import pandas as pd
import os
import datetime
# from PyInquirer import prompt
from Helfer_Objekte import (check_invoice_archive, question_next_invoice_number,save_to_archive,
                            validate_input_int, stringsandyear_topath, stringsandinvoicenumber_topath)
import tkinter as tk

class Grid_Entry():
    def __init__(self,gui_widget,value):
        self.gui_widget = gui_widget
        self.value = value


def make_invoice_tirol(allclientdata_path,invoice_tirol_path,excel_template_path,outputdir_suppath,nameoutputdir,nameoutputarchivefile,
                       invoicenumber_pattern, invoicenumber_pattern_names,nameinvoicefile, kassabuchdir,user = "r"):
    if user == "r":
        amount_of_persons_LandTirol = 3
    if user == "b":
        amount_of_persons_LandTirol = 5



    #invoicenumber_pattern = r'(\d{4})-(\d+)'
    # invoicenumber_pattern = r'(T\d{4})-(\d+)'
    # invoicenumber_pattern_names = ["T","year","-","invoicenumber"]
    #invoicenumber_pattern_names = ["year","-","invoicenumber"]



    allclientdata = pd.read_excel(allclientdata_path, index_col=0, header=None, sheet_name=None)
    print(allclientdata)
    # select which client
    allclientsnames = list(allclientdata.keys())
    allclientsnames = [x for x in allclientsnames if x != "Vorlage"]
    allclientsnames.sort()


    # unausgefüllt!

    window = tk.Tk()
    window.geometry("1000x650+50+0")

    window.title("Wähle die Personen und gib die Stunden ein")
    window.resizable(width=False, height=False)
    if user == "r":
        showvalues = ["Name", "Geb.", "Gültige Genehmigung Land Tirol ab", "Anzahl Einzelstunden", "Anzahl Gruppenstunden","Anzahl Hausbesuche"]
        columns = [0,1,2,4,7,9]
        minuteslist = ["30 min", "45 min", "60 min"]

    if user == "b":
        showvalues = ["Name", "Geb.", "Gültige Genehmigung Land Tirol ab", "Anzahl Stunden", "Anzahl Hausbesuche"]
        columns = [0,1,2,3,4]
        minuteslist = ["60 min"]

    for col, columnname in zip(columns,showvalues):
        collabel = tk.Label(master=window, text=columnname)
        collabel.grid(row=0, column=col, padx=10)

    emptyclientdata = allclientdata["Vorlage"].copy()
    print(emptyclientdata)
    emptyclientdata[:] = ""
    emptyclientdata = emptyclientdata.to_dict()
    selected_clientdata = {}
    if user == "r":
        selected_clientdata = {1:emptyclientdata.copy(),
                               2:emptyclientdata.copy(),
                               3:emptyclientdata.copy()}
    if user == "b":
        selected_clientdata = {1:emptyclientdata.copy(),
                               2:emptyclientdata.copy(),
                               3:emptyclientdata.copy(),
                               4:emptyclientdata.copy(),
                               5:emptyclientdata.copy()}
    for clientindex in range(1,amount_of_persons_LandTirol+1):
        startrow = clientindex*4
        tk.ttk.Separator(
            master=window,
            orient=tk.HORIZONTAL,
            style='blue.TSeparator',
            takefocus=1,
            cursor='plus'
        ).grid(row=startrow, column=0, columnspan=10, ipadx=200, pady=10,sticky='ew')
        title = tk.Label(master=window, text=f"Person {clientindex}")
        title.grid(row = startrow,column =0,pady=10)

        pady = 5
        padx = 1

        def on_name_select(selected_name,clientindex,selected_clientdata):
            # selected_clientdata[clientindex]["Name"].gui_widget['menu'].entryconfig("Select", state="disabled") # make tha you cannot select Auswählen anymore
            print(f"{selected_name} number {clientindex}, {selected_clientdata}")
            thisclientdata = allclientdata[selected_name].to_dict()[1]
            for key in thisclientdata.keys():
                if key not in showvalues:
                    selected_clientdata[clientindex][key] = thisclientdata[key]
                print(selected_clientdata)
                selected_clientdata[clientindex]["Name"].value = selected_name
                selected_clientdata[clientindex]["Geb."].gui_widget["text"] = thisclientdata["Geb."].strftime("%d.%m.%Y")
                selected_clientdata[clientindex]["Geb."].value = thisclientdata["Geb."]
                try:
                    selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].gui_widget["text"] = thisclientdata["Gültige Genehmigung Land Tirol ab"].strftime("%d.%m.%Y")
                    selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].value = thisclientdata["Gültige Genehmigung Land Tirol ab"]
                except Exception as e:
                    print("Würde keine Gültige Genehmigung Land Tirol angegeben für diese Person?")
                    print(e)

        clicked = tk.StringVar()
        clicked.set("Auswählen")
        # print(clientindex)
        options = allclientsnames.copy()
        selected_clientdata[clientindex]["Name"] = Grid_Entry(tk.OptionMenu(window , clicked, *options,command= lambda selected_name, ci=clientindex, sc =selected_clientdata: on_name_select(selected_name,ci,sc)),"" )

        selected_clientdata[clientindex]["Geb."] = Grid_Entry(tk.Label(master=window, text=""), "")
        vcmd = window.register(validate_input_int)
        selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"] = Grid_Entry(tk.Label(master=window, text=""), "")
        if user == "r":
            selected_clientdata[clientindex]["Anzahl Einzelstunden"] =  {"30 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                         "45 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                         "60 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                            }
            selected_clientdata[clientindex]["Anzahl Gruppenstunden"] =  {"30 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                         "45 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                         "60 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0)

                                   }
        if user == "b":
            selected_clientdata[clientindex]["Anzahl Stunden"] = Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0)
        selected_clientdata[clientindex]["Anzahl Hausbesuche"] = Grid_Entry(tk.Entry(master=window, text="", validate="key", validatecommand=(vcmd, '%S', '%P')), "")
        column = 0
        for columnname in showvalues:
            if isinstance(selected_clientdata[clientindex][columnname], dict):
                for row, minutes in enumerate(selected_clientdata[clientindex][columnname]):
                    minlabel = tk.Label(master=window, text=minutes)
                    minlabel.grid(row = startrow + row +1, column = column, padx = padx, pady = pady)
                    selected_clientdata[clientindex][columnname][minutes].gui_widget.grid(row = startrow +row+1, column = column+1, padx = 10,pady = pady)
                column += 2 #because we have 2 columns
            else:
                selected_clientdata[clientindex][columnname].gui_widget.grid(row = startrow +2, column = column, padx = padx,pady = pady)
            column += 1


    def on_ok_buttonpress():
        print("Erstelle Rechnung")
        print(f"open excel vorlage tirol {invoice_tirol_path}")
        invoice_tirol = openpyxl.load_workbook(invoice_tirol_path)
        invoice_tirol_sheet = invoice_tirol["Rechnung Einrichtung"]
        if user == "r":
            kostenstruktur = {"Anzahl Einzelstunden": {},
                              "Anzahl Gruppenstunden": {},
                              "Anzahl Hausbesuche":invoice_tirol_sheet["H25"].value,
                              "Ausgleichzulage": float(invoice_tirol_sheet["H26"].value)}
            print(minuteslist)
            for i,min in zip(range(22, 25),minuteslist):
                print(min)
                kostenstruktur["Anzahl Einzelstunden"][min] = invoice_tirol_sheet[f"E{i}"].value

            for i,min in zip(range(22, 25),minuteslist):
                kostenstruktur["Anzahl Gruppenstunden"][min] = invoice_tirol_sheet[f"G{i}"].value
            print(kostenstruktur)
            cellsbetweenclients = 7
            excelsheet_locs = {"Name":("A",22),
                               "Geb.":("B",22),
                               "Gültige Genehmigung Land Tirol ab":("C",22),
                               "Anzahl Einzelstunden":{"30 min":("D",22),"45 min":("D",23),"60 min":("D",24)},
                               "Anzahl Gruppenstunden":{"30 min":("F",22),"45 min":("F",23),"60 min":("F",24)},
                               "Anzahl Hausbesuche":("H",22)}
            otherlocs = {"Ort, Datum":"I15","Rechnungsnummer":"I16"}
            costsdf = pd.DataFrame(columns=["Anzahl Einzelstunden", "Anzahl Gruppenstunden", "Anzahl Hausbesuche"])
            col_costdf = 3


        if user == "b":
            firstrowclients = 21
            cellsbetweenclients = 5
            excelsheet_locs = {"Name":("A",firstrowclients),
                               "Geb.":("B",firstrowclients),
                               "Gültige Genehmigung Land Tirol ab":("C",firstrowclients),
                               "Anzahl Stunden":("D",firstrowclients),
                               "Anzahl Hausbesuche":("D",firstrowclients+1)}
            otherlocs = {"Ort, Datum":"F15","Rechnungsnummer":"F16"}

            kostenstruktur = { "Anzahl Stunden": invoice_tirol_sheet["E21"].value, #anzahl sollte eigentlich Hier Kosten heissen
                              "Anzahl Hausbesuche":invoice_tirol_sheet["F22"].value,
                              "Ausgleichzulage": invoice_tirol_sheet["F23"].value}

            costsdf = pd.DataFrame(columns=["Anzahl Stunden", "Anzahl Hausbesuche"])
            col_costdf = 2
        print(kostenstruktur)
        for clientindex in range(1, amount_of_persons_LandTirol + 1):
            costsdf.loc[clientindex - 1] = [0]*col_costdf
            # print(selected_clientdata[clientindex]["Name"].gui_widget["text"])
            for key in showvalues:
                if isinstance(selected_clientdata[clientindex][key],dict):
                    costvalues = np.array([])
                    for min in selected_clientdata[clientindex][key]:
                        location = f"{excelsheet_locs[key][min][0]}{excelsheet_locs[key][min][1]+(clientindex-1)*cellsbetweenclients}"
                        # print(location)
                        amounthours = selected_clientdata[clientindex][key][min].gui_widget.get()
                        if amounthours:
                            amounthours = float(amounthours)
                            cost = kostenstruktur[key][min] * amounthours
                            costvalues = np.append(costvalues,float(cost))
                        invoice_tirol_sheet[location] = amounthours

                    costsdf.loc[clientindex - 1,key] = costvalues.sum()
                else:
                    location = f"{excelsheet_locs[key][0]}{excelsheet_locs[key][1]+(clientindex-1)*cellsbetweenclients}"
                    print(key, location)
                    if (key == "Anzahl Hausbesuche") or (key == "Anzahl Stunden"):
                        value = selected_clientdata[clientindex][key].gui_widget.get()
                        if value:
                            value = value.replace(",",".")
                            value = float(value)
                            costsdf.loc[clientindex - 1,key] = kostenstruktur[key]*value
                        invoice_tirol_sheet[location] = value
                    else:
                        if key == "Name":
                            input = selected_clientdata[clientindex][key].gui_widget["text"]
                            if input != "Auswählen":
                                invoice_tirol_sheet[location] = selected_clientdata[clientindex][key].gui_widget["text"]
                        else:
                            invoice_tirol_sheet[location] = selected_clientdata[clientindex][key].gui_widget["text"]
                invoice_tirol_sheet[otherlocs["Ort, Datum"]]=f"{datetime.datetime.today().strftime('%d.%m.%Y')}"

        # calculate total sum
        if user == "r":
            costsdf["Ausgleichzulage"] = (costsdf["Anzahl Einzelstunden"]+costsdf["Anzahl Gruppenstunden"])*kostenstruktur["Ausgleichzulage"]
        if user == "b":
            costsdf["Ausgleichzulage"] = (costsdf["Anzahl Stunden"] ) * kostenstruktur["Ausgleichzulage"]
        costsdf["Summen"] = costsdf.sum(axis=1)
        totalsum  = costsdf["Summen"].sum(axis = 0)

        names_this_invoice = []
        for clientindex in range(1, amount_of_persons_LandTirol + 1):
            name = selected_clientdata[clientindex]["Name"].gui_widget["text"]
            names_this_invoice.append(name)
        window.destroy()

        #get invoice number
        year_of_invoice = datetime.datetime.today().year
        outputdir = stringsandyear_topath(nameoutputdir, year_of_invoice)
        outputdir_path = os.path.join(outputdir_suppath, outputdir)
        archive_which_invoices_name = stringsandyear_topath(nameoutputarchivefile, year_of_invoice)
        if user == "b":
            archive_which_invoices_path = os.path.join(kassabuchdir, archive_which_invoices_name)
        if user == "r":
            archive_which_invoices_path = os.path.join(outputdir_path, archive_which_invoices_name)

        print(f"Year to link this invoice to: {year_of_invoice}")
        lastinvoice_num = check_invoice_archive(year_of_invoice,outputdir_path,archive_which_invoices_path,excel_template_path,invoicenumber_pattern= invoicenumber_pattern)
        print(f"lastinvoice_num: {lastinvoice_num}")
        thisinvoicenumber = question_next_invoice_number(year_of_invoice,lastinvoice_num,invoicenumber_pattern,invoicenumber_pattern_names)
        print(f"thisinvoicenumber{thisinvoicenumber}")

        invoice_tirol_sheet[otherlocs["Rechnungsnummer"]] = thisinvoicenumber
        names_this_invoice = [x for x in names_this_invoice if x !="Auswählen"]
        namesstring = ""
        for name in names_this_invoice:
            namesstring+= f"{name} "
        filename = stringsandinvoicenumber_topath(nameinvoicefile, thisinvoicenumber, "Land Tirol",
                                                  datetime.date.today().strftime('%d_%m_%Y'))
        outputfile_path = os.path.join(outputdir_path, filename)
        print(f"Now i can create the outputdata filepaths:   \n{archive_which_invoices_path}\n{outputfile_path}")
        save_docs = True
        if save_docs:
            print(f"Speichern der Rechnungn in {outputfile_path}")
            invoice_tirol.save(outputfile_path)
        print(f"Speichern des Kassabuchs in {archive_which_invoices_path}")
        try_saving = True
        while try_saving:
            try:
                save_to_archive(thisinvoicenumber, datetime.datetime.today(), namesstring, datetime.datetime.today(),
                                "", totalsum, archive_which_invoices_path)

                try_saving = False
            except Exception as e:
                def show_alert():
                    # Display an alert message box with an "Okay" button
                    tk.messagebox.showinfo("Fehler", f"Ich konnte die Datei nicht speichern. Schließe zuerst die Datei {archive_which_invoices_path}")
                print("error in saving archive")
                print(e)
                root = tk.Tk()
                root.withdraw()
                show_alert()
                print("The alert has been dismissed. Continuing with the rest of the code...")
                root.quit()

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
            os.startfile(outputfile_path)
            os.startfile(archive_which_invoices_path)
        from Rechnung_erstellen import main
        main()




    button = tk.Button( window , text = "Speichern", command= on_ok_buttonpress).grid(row = startrow +4, column = 3,pady=10)
    window.mainloop()
