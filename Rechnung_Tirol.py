from pathlib import Path
from docxtpl import DocxTemplate
import openpyxl
import pandas as pd
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx import Document
from docx.shared import Cm
import xlwings as xw
import os
import numpy as np
import re
import datetime
# from PyInquirer import prompt
import pprint
import Helfer_Objekte
from Helfer_Objekte import check_invoice_archive, question_next_invoice_number,save_to_archive, validate_input_int
import dateutil.parser
import tkinter as tk

class Grid_Entry():
    def __init__(self,gui_widget,value):
        self.gui_widget = gui_widget
        self.value = value



# parent_dir = "C:\\Users\\peaq\\Documents\\Programm Logo\\Programm"
# supparentdir = os.path.dirname(parent_dir)
# template_path = os.path.join(parent_dir,"Vorlage.docx")
# invoice_tirol_path = os.path.join(parent_dir,"Abrechnung_TherapeutInnen_mit_Ausgleichssatz_ab_01.01.2024-2.xlsx")
# excel_template_path = os.path.join(parent_dir,"Jahresübersicht_Vorlage.xlsx")
# allhourdata_path = os.path.join(parent_dir ,"Stundendaten.xlsx")
# allclientdata_path = os.path.join(parent_dir ,"PatientInneninformationen.xlsx")
# outputinvoice_path = 'C:\\Users\\peaq\\Documents\\Programm Logo\\Programm\\Rechnungtest.xlsx'
def make_invoice_tirol(allclientdata_path,invoice_tirol_path,supparentdir,excel_template_path):
    allclientdata = pd.read_excel(allclientdata_path, index_col=0, header=None, sheet_name=None)

    # select which client
    allclientsnames = list(allclientdata.keys())
    allclientsnames.sort()

    emptyclientdata = allclientdata[list(allclientdata.keys())[0]][1].copy()
    emptyclientdata[:] = ""
    emptyclientdata = emptyclientdata.to_dict()
    selected_clientdata = {1:emptyclientdata.copy(),
                           2:emptyclientdata.copy(),
                           3:emptyclientdata.copy()}
    # unausgefüllt!

    window = tk.Tk()
    window.title("Wähle die Personen und gib die Stunden ein")
    window.resizable(width=False, height=False)
    showvalues = ["Name", "Geb.", "Gültige Genehmigung Land Tirol ab", "Anzahl Einzelstunden", "Anzahl Gruppenstunden","Anzahl Hausbesuche"]
    columns = [0,1,2,4,7,9]
    for col, columnname in zip(columns,showvalues):
        collabel = tk.Label(master=window, text=columnname)
        collabel.grid(row=0, column=col, padx=10)


    for clientindex in range(1,4):
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
            selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].gui_widget["text"] = thisclientdata["Gültige Genehmigung Land Tirol ab"].strftime("%d.%m.%Y")
            selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].value = thisclientdata["Gültige Genehmigung Land Tirol ab"]

        clicked = tk.StringVar()
        clicked.set("Auswählen")
        print(clientindex)
        selected_clientdata[clientindex]["Name"] = Grid_Entry(tk.OptionMenu(window , clicked, *allclientsnames,command= lambda selected_name, ci=clientindex, sc =selected_clientdata: on_name_select(selected_name,ci,sc)),"" )

        selected_clientdata[clientindex]["Geb."] = Grid_Entry(tk.Label(master=window, text=""), "")
        vcmd = window.register(validate_input_int)
        selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"] = Grid_Entry(tk.Label(master=window, text=""), "")
        selected_clientdata[clientindex]["Anzahl Einzelstunden"] =  {"30 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                     "45 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                     "60 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                        }
        selected_clientdata[clientindex]["Anzahl Gruppenstunden"] =  {"30 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                     "45 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0),
                                                                     "60 min": Grid_Entry(tk.Entry(master=window, width=10, validate="key", validatecommand=(vcmd, '%S', '%P')),0)

                               }
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
        invoice_tirol_sheet = invoice_tirol['Rechnung mit AZ']

        cellsbetweenclients = 7
        excelsheet_locs = {"Name":("A",22),
                           "Geb.":("B",22),
                           "Gültige Genehmigung Land Tirol ab":("C",22),
                           "Anzahl Einzelstunden":{"30 min":("D",22),"45 min":("D",23),"60 min":("D",24)},
                           "Anzahl Gruppenstunden":{"30 min":("F",22),"45 min":("F",23),"60 min":("F",24)},
                           "Anzahl Hausbesuche":("H",22)}
        otherlocs = {"Ort, Datum":"E16","Rechnungsnummer":"E17"}
        for clientindex in range(1, 4):
            for key in showvalues:
                if isinstance(selected_clientdata[clientindex][key],dict):
                    for min in selected_clientdata[clientindex][key]:
                        location = f"{excelsheet_locs[key][min][0]}{excelsheet_locs[key][min][1]+(clientindex-1)*cellsbetweenclients}"
                        print(location)
                        invoice_tirol_sheet[location] = selected_clientdata[clientindex][key][min].gui_widget.get()
                else:
                    location = f"{excelsheet_locs[key][0]}{excelsheet_locs[key][1]+(clientindex-1)*cellsbetweenclients}"
                    print(location)
                    if key == "Anzahl Hausbesuche":
                        invoice_tirol_sheet[location] = selected_clientdata[clientindex][key].gui_widget.get()
                    else:
                        if key == "Name":
                            input = selected_clientdata[clientindex][key].gui_widget["text"]
                            if input != "Auswählen":
                                invoice_tirol_sheet[location] = selected_clientdata[clientindex][key].gui_widget["text"]
                        else:
                            invoice_tirol_sheet[location] = selected_clientdata[clientindex][key].gui_widget["text"]
            print(invoice_tirol_sheet["I27"].value)
            invoice_tirol_sheet[otherlocs["Ort, Datum"]]=f"Innsbruck, {datetime.datetime.today().strftime('%d.%m.%Y')}"

        #get invoice number
        year_of_invoice = datetime.datetime.today().year
        print(f"Year to link this invoice to: {year_of_invoice}")
        lastinvoice_num = check_invoice_archive(year_of_invoice,supparentdir,excel_template_path,invoicenumber_pattern= r'T(\d{4})-(\d+)')
        print(f"lastinvoice_num: {lastinvoice_num}")
        invoicenumber_pattern = r'T(\d{4})-(\d+)'
        thisinvoicenumber = question_next_invoice_number(year_of_invoice,lastinvoice_num,invoicenumber_pattern,Tirol=True)
        print(f"thisinvoicenumber{thisinvoicenumber}")

        invoice_tirol_sheet[otherlocs["Rechnungsnummer"]] = thisinvoicenumber

        outputdir_path = os.path.join(supparentdir, f"{datetime.datetime.today().year}")
        archive_which_invoices_path = os.path.join(outputdir_path, f"Rechnungen {datetime.datetime.today().year}.xlsx")
        outputfile_path = os.path.join(outputdir_path, f"RE {thisinvoicenumber} {datetime.date.today().strftime('%d_%m_%Y')}.xlsx")

        invoice_tirol.save(outputfile_path)
        app = xw.App(visible=True)  # Set visible=True to see the Excel window
        workbook = app.books.open(outputfile_path)

        sheet = workbook.sheets['Rechnung mit AZ']  # Access a sheet by name
        totalsum = sheet.range('I43').value
        print(sheet.range('A1').value)
        print(totalsum)

        response =tk.messagebox.askyesno("Abschließen?",
                                       "Bist du mit dem Ergebnis zufrieden? \nWenn ja dann schließe ich nun das Program (du kannst immer noch im Excel Sachen ändern und dann von Excel speichern). \nWenn nein kannst du nun noch weiter im Programm sachen ändern.")

        if response:
            print("Abschließen")
            try:
                workbook.save(outputfile_path)
            except Exception as e:
                print("Could not save and close the workbook, is it already closed?")
                print(e)
            window.destroy()
            save_to_archive(thisinvoicenumber, datetime.datetime.today(), "", datetime.datetime.today(),
                            datetime.datetime.today(), totalsum, archive_which_invoices_path)
            os.startfile(archive_which_invoices_path)

        else:
            print("Weiter machen")
            # continue doesnot work fully
            app.quit()




    button = tk.Button( window , text = "Speichern", command= on_ok_buttonpress).grid(row = startrow +4, column = 3,pady=10)
    window.mainloop()
