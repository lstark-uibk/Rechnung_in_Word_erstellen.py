from pathlib import Path
from docxtpl import DocxTemplate
import openpyxl
import pandas as pd
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx import Document
from docx.shared import Cm
import os
import numpy as np
import re
import datetime
# from PyInquirer import prompt
import pprint
import Helfer_Objekte
from Helfer_Objekte import change_place_of_window
import dateutil.parser
import tkinter as tk

from  Rechnung_Praxis import make_invoice_praxis
from Rechnung_Tirol import make_invoice_tirol

parent_dir = "C:\\Users\\peaq\\Documents\\Programm Logo\\Programm"
print( "Programm für Brigitte")

supparentdir = os.path.dirname(parent_dir)
template_praxis_path = os.path.join(parent_dir,"Vorlage.docx")
template_tirol_path = os.path.join(parent_dir,"Abrechnung_TherapeutInnen_mit_Ausgleichssatz_ab_01.01.2024-2.xlsx")
excel_template_path = os.path.join(parent_dir,"Jahresübersicht_Vorlage.xlsx")
allhourdata_path = os.path.join(parent_dir ,"Stundendaten.xlsx")
allclientdata_path = os.path.join(parent_dir ,"PatientInneninformationen.xlsx")
outputdir_path = 0
archive_which_invoices_path = 0

Tirolinvoice = False

# Create the main window
root = tk.Tk()
root.title("Welche Art von Rechnung?")
def set_answer(answer):
    root.destroy()
    if answer:
        print("Tirol")
        make_invoice_tirol(allclientdata_path,template_tirol_path,supparentdir,excel_template_path)
    else:
        make_invoice_praxis(allhourdata_path, allclientdata_path, supparentdir, excel_template_path, template_praxis_path)


# Ask a question
question_label = tk.Label(root, text="Welche Art von Rechnung möchtest du erstellen?")
question_label.pack(pady=10)

button1 = tk.Button(root, text="Praxis (nach meiner Vorlage)", command=lambda: set_answer(False))
button1.pack(pady=5)

button2 = tk.Button(root, text="Land Tirol", command=lambda: set_answer(True))
button2.pack(pady=5)
root.mainloop()

