import os
import tkinter as tk
from  Rechnung_Praxis import make_invoice_praxis
from Rechnung_Tirol import make_invoice_tirol
from  Neue_Person import make_new_Person


def main():
    parent_dir = "C:\\Users\\peaq\\Documents\\Programm Logo\\Programm"

    supparentdir = os.path.dirname(parent_dir)
    template_praxis_path = os.path.join(parent_dir, "Vorlage.docx")
    template_tirol_path = os.path.join(parent_dir, "Abrechnung_TherapeutInnen_mit_Ausgleichssatz_ab_01.01.2024-2.xlsx")
    excel_template_path = os.path.join(parent_dir, "Jahresübersicht_Vorlage.xlsx")
    allhourdata_path = os.path.join(supparentdir, "Daten\\Stundendaten.xlsx")
    allclientdata_path = os.path.join(supparentdir, "Daten\\PatientInneninformationen.xlsx")
    outputdir_path = 0
    archive_which_invoices_path = 0

    # Create the main window
    root = tk.Tk()
    root.title("Was möchtest du tun?")
    def set_answer(answer):
        root.destroy()
        if answer == "Tirol":
            print("Tirol")
            make_invoice_tirol(allclientdata_path,template_tirol_path,supparentdir,excel_template_path)
        elif answer == "Praxis":
            print("Praxis")
            make_invoice_praxis(allhourdata_path, allclientdata_path, supparentdir, excel_template_path, template_praxis_path)
        elif answer == "Neu":
            print("Neu")
            make_new_Person()



    # Ask a question
    question_label = tk.Label(root, text="Was möchtest du tun?")
    question_label.pack(pady=10,padx=10)

    button1 = tk.Button(root, text="Praxisrechnung nach meiner Vorlage erstellen", command=lambda: set_answer("Praxis"))
    button1.pack(pady=10,padx=10)

    button2 = tk.Button(root, text="Rechnung ans Land Tirol erstellen", command=lambda: set_answer("Tirol"))
    button2.pack(pady=10,padx=10)

    button3 = tk.Button(root, text="Neuen Klienten anlegen", command=lambda: set_answer("Neu"))
    button3.pack(pady=10,padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()