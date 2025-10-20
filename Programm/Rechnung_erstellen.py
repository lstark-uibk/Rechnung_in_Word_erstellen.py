import os
import tkinter as tk
from  Rechnung_Praxis import make_invoice_praxis
from Rechnung_Tirol import make_invoice_tirol
from  Neue_Person import make_new_Person
from functools import partial
import json

# this is branch for maria
def main():
    with open(r"C:\Users\goglm\Desktop\Ergotherapie\Buchhaltung\Abrechnungsprogramm\Programm\config.json", 'r') as f:
        config = json.load(f)

    # Create the main window
    root = tk.Tk()
    root.title("Was möchtest du tun?")
    def set_method(method):
        root.destroy()
        if method == "Neu":
            print("I use method Neu")
            make_new_Person(config)
        else:
            print(f"I use method {method}")
            function_to_run = globals()[config["outputmethods"][method]["function_to_evoce"]]
            function_to_run(config,method)



# What do we want to do?
    question_label = tk.Label(root, text="Was möchtest du tun?")
    question_label.pack(pady=10,padx=10)
    buttons = []

    for index,method in enumerate(config["outputmethods"].keys()):
        buttontext = f"Rechnung {method} erstellen"
        # print(method,buttontext,config["outputmethods"][method]["function_to_evoce"])
        def function_to_run(root,config,method):
            root.destroy()
            globals()[config["outputmethods"][method]["function_to_evoce"]](config,method)

        funct_with_inputs = partial(function_to_run,root = root,config = config,method = method)
        buttons.append(tk.Button(root, text=buttontext, command=funct_with_inputs))
        buttons[index].pack(pady=10,padx=10)

    buttonnewperson = tk.Button(root, text="Neue Person anlegen", command=lambda: set_method("Neu"))
    buttonnewperson.pack(pady=10,padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()