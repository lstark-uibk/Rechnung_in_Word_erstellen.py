import os
import tkinter as tk
from  Rechnung_Praxis import make_invoice_praxis
from Rechnung_Tirol import make_invoice_tirol
from  Neue_Person import make_new_Person


def main():
    parent_dir = os.path.dirname(os.path.realpath(__file__))
    #user = "r"
    user = "b"
    supparentdir = os.path.dirname(parent_dir)
    supsupparentdir = os.path.dirname(supparentdir)
    if user == "r":
        template_praxis_path = os.path.join(supparentdir, "Vorlagen/VorlageRosmarie.docx")
        template_tirol_path = os.path.join(supparentdir, "Vorlagen/Vorlage_LandTirol_Rosmarie.xlsx")

    if user == "b":
        template_praxis_path = os.path.join(supparentdir, "Vorlagen/Vorlage_Brigitte_2025.xlsx")
        template_tirol_path = os.path.join(supparentdir, "Vorlagen/Vorlage_LandTirol_Brigitte.xlsx")


    excel_template_path = os.path.join(supparentdir, "Vorlagen/Jahresübersicht_Vorlage.xlsx")
    allhourdata_path = os.path.join(supsupparentdir, "Daten/Stundendaten.xlsx")
    allclientdata_path = os.path.join(supsupparentdir, "Daten/PatientInneninformationen.xlsx")
    outputdir_suppath = supsupparentdir
    #nameoutputdir = ["Rechnungen ", "year"]
    nameoutputdir = ["Rechnungen_","year"]
    nameoutputarchivefile = ["Kassabuch_", "year",".xlsx"]
    kassbuchdir = ""
    if user == "r":
        nameinvoicefile = ["RE ", "invoicenumber", "clientname", " " , "date", ".docx"]
        invoicenumber_pattern = r'(\d{4})-(\d+)'
        invoicenumber_pattern_names = ["year", "-", "invoicenumber"]
    if user == "b":
        nameinvoicefile = ["clientname","-","invoicenumber","-", "date", ".xlsx"]
        invoicenumber_pattern = r'(\d{4})1(\d+)'
        invoicenumber_pattern_names = ["year","1", "invoicenumber"]
        kassbuchdir = "/Users/brigittemacbookpro/Brigitte Aron/Logopädie/Buchhaltung_Praxis/Kassabuch"

    #invoicenumber_pattern_names_T = ["T","year","-","invoicenumber"]


    archive_which_invoices_path = 0

    # Create the main window
    root = tk.Tk()
    root.title("Was möchtest du tun?")
    def set_answer(answer):
        root.destroy()
        if answer == "Tirol":
            print("Tirol")
            make_invoice_tirol(allclientdata_path,template_tirol_path,excel_template_path,outputdir_suppath,nameoutputdir,nameoutputarchivefile,
                               invoicenumber_pattern, invoicenumber_pattern_names, nameinvoicefile,kassbuchdir,user = user)
        elif answer == "Praxis":
            print("Praxis")
            make_invoice_praxis(allhourdata_path, allclientdata_path, excel_template_path, template_praxis_path,outputdir_suppath,
                                nameoutputdir,nameoutputarchivefile, invoicenumber_pattern,
                                invoicenumber_pattern_names,nameinvoicefile,kassbuchdir,user = user)
        elif answer == "Neu":
            print("Neu")
            make_new_Person(allclientdata_path)


# What do we want to do?
    question_label = tk.Label(root, text="Was möchtest du tun?")
    question_label.pack(pady=10,padx=10)

    button1 = tk.Button(root, text="Praxisrechnung nach meiner Vorlage erstellen", command=lambda: set_answer("Praxis"))
    button1.pack(pady=10,padx=10)

    button2 = tk.Button(root, text="Rechnung ans Land Tirol erstellen", command=lambda: set_answer("Tirol"))
    button2.pack(pady=10,padx=10)

    button3 = tk.Button(root, text="Neue Person anlegen", command=lambda: set_answer("Neu"))
    button3.pack(pady=10,padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()