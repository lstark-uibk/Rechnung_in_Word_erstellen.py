from tkinter import *

import numpy as np
from lxml.etree import clear_error_log
from numpy.testing.print_coercion_tables import print_new_cast_table
from tkcalendar import Calendar, DateEntry
from functools import partial
import pandas as pd
import openpyxl
import os
import dateutil.parser
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
# from PyInquirer import prompt
import pprint
import re
from tkinter import messagebox
from tkinter.font import Font


def on_name_select(selected_name, clientindex, selected_clientdata, allclientdata, showvalues):
    # selected_clientdata[clientindex]["Name"].gui_widget['menu'].entryconfig("Select", state="disabled") # make tha you cannot select Auswählen anymore
    print(f"{selected_name} number {clientindex}, {selected_clientdata}")
    thisclientdata = allclientdata[selected_name].to_dict()[1]
    for key in thisclientdata.keys():
        if key not in showvalues:
            selected_clientdata[clientindex][key] = thisclientdata[key]
        selected_clientdata[clientindex]["Name"].value = selected_name
        selected_clientdata[clientindex]["Geb."].gui_widget["text"] = thisclientdata["Geb."].strftime("%d.%m.%Y")
        selected_clientdata[clientindex]["Geb."].value = thisclientdata["Geb."]
        selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].gui_widget["text"] = thisclientdata[
            "Gültige Genehmigung Land Tirol ab"].strftime("%d.%m.%Y")
        selected_clientdata[clientindex]["Gültige Genehmigung Land Tirol ab"].value = thisclientdata[
            "Gültige Genehmigung Land Tirol ab"]


def stringsandyear_topath(stringsandyear,year):
    string = ""
    for x in stringsandyear:
        if "year" not in x:
            string += x
        else:
            string += f"{year}"
    return string

def stringsandinvoicenumber_topath(stringsandinvoicenumber, invoicenumber, clientname, date):
    string = ""
    for x in stringsandinvoicenumber:
        if ("invoicenumber" in x) :
            string += invoicenumber
        if ("clientname" in x) :
            string += clientname
        if ("date" in x):
            string += date
        if ("date" not in x) and ("clientname" not in x) and ("invoicenumber" not in x) :
            string += x
    return string
def check_invoice_archive(year_of_invoice,outputdir_path,archive_which_invoices_path,invoice_achive_template_path,invoicenumber_pattern):

    if not os.path.exists(archive_which_invoices_path):
        print(f"Because there was no Archive file of the year create one at {archive_which_invoices_path}")
        wb = openpyxl.load_workbook(invoice_achive_template_path)
        # Select the worksheet
        sheet = wb["Tabelle1"]
        # Modify the cell
        sheet["A1"] = f"Rechnungen {year_of_invoice}"
        wb.save(archive_which_invoices_path)
        lastinvoice_year_num = f"{year_of_invoice}-001"
        lastinvoice_year = year_of_invoice
        lastinvoice_num = 1

    invoicenumbers = pd.read_excel(archive_which_invoices_path).iloc[:-2,0] # get the invoice numbers out of the invoice archive
    # search for the first invoicenumer which fits the pattern
    invoicenumber_pattern = invoicenumber_pattern
    lastinvoice_num = 0
    for index, entry in invoicenumbers.iloc[::-1].items():
        if not pd.isnull(entry):
            entry = str(entry)
            if re.match(invoicenumber_pattern, entry):
                lastinvoice_year_num = re.match(invoicenumber_pattern, entry)[0]
                lastinvoice_year = int(re.match(invoicenumber_pattern, entry)[1])
                lastinvoice_num = int(re.match(invoicenumber_pattern, entry)[2])
                break
    return lastinvoice_num

def ask_right_invoicenumber(question):
    root = tk.Tk()
    root.title("Frage")

    Label(root, text=question).pack(padx=10, pady=10)
    yes_no_frame = tk.Frame(root)
    yes_no_frame.pack(pady=10,expand=True)
    ttk.Style().configure('Treeview', rowheight=30)

    answer = [True]
    def button_press_y():
        print("Invoicenumber OK")
        answer[0] = True
        root.destroy()
    def button_press_n():
        print("Change Invoicenumber")
        answer[0] = False
        root.destroy()
    # Pack widgets side by side inside the last row frame
    tk.Button(yes_no_frame, text="OK",command=button_press_y).pack(side="left", padx=10)
    tk.Button(yes_no_frame, text="Ändern",command=button_press_n).pack(side="left", padx=10)
    root.mainloop()
    return answer[0]
def question_next_invoice_number(invoiceyear,lastinvoice_num,invoicenumber_pattern,invoicenumber_pattern_names):
    #get which invoicenumber
    answer1 = "Nimm einfach die Nächste in der Reihe"
    answer2 = "Ich möchte sie selber eingeben"
    invoicenumberquestion_choices = [answer1, answer2]
    thisinvoicenumber = ""
    while not thisinvoicenumber:
        last_inv_numb_str_sugg = ""
        this_inv_numb_str_sugg = ""
        for x in invoicenumber_pattern_names:
            if x in "year":
                last_inv_numb_str_sugg += f"{invoiceyear}"
                this_inv_numb_str_sugg += f"{invoiceyear}"
            elif x in "invoicenumber":
                last_inv_numb_str_sugg += f"{(lastinvoice_num):03}"
                this_inv_numb_str_sugg += f"{(lastinvoice_num + 1):03}"
            else:
                last_inv_numb_str_sugg += x
                this_inv_numb_str_sugg += x


        result = ask_right_invoicenumber(f"Die letzte Rechnungsnummer war {last_inv_numb_str_sugg}. \nSomit wäre die nächste Rechnungsnummer {this_inv_numb_str_sugg}.")
        if result:

            thisinvoicenumber = this_inv_numb_str_sugg
            print("Das ist die Rechnungsnummer: " + thisinvoicenumber)
        if not result:
            root = tk.Tk()
            # withdraw() will make the parent window disappear.
            root.withdraw()
            # shows a dialogue with a string input field
            print("input a new invoicenumber")
            thisinvoicenumber = tk.simpledialog.askstring('Rechnungsnummer',
                                                       f"Dann kannst du sie jetzt selber eingeben (in dem Format z.b. {last_inv_numb_str_sugg}):",
                                                       parent=root)
            root.destroy()
            if not thisinvoicenumber:
                continue

            while not re.match(invoicenumber_pattern,thisinvoicenumber):
                print(f"The input {thisinvoicenumber} didnot match the pattern {invoicenumber_pattern}")
                root = tk.Tk()
                change_place_of_window(root)
                # withdraw() will make the parent window disappear.
                root.withdraw()
                # shows a dialogue with a string input field
                thisinvoicenumber = tk.simpledialog.askstring('Rechnungsnummer',
                                                           f"Die letze eingetragen Rechnungsnummer hatte nicht das richtige Format. \nGib sie in dem Format ein wie {this_inv_numb_str_sugg} wobei die erste Nummer mit dem Jahr der Rechnung ersetzt wird und die zweite mit der Rechnungsnummer:",
                                                           parent=root)
                root.destroy()
    return thisinvoicenumber
def validate_input_int(char, input_value):
    """Function to validate input - allows only integer values."""
    # If the input is empty, it's valid (so the user can delete the input).
    if input_value == "":
        return True
    try:
        # Try to convert the input value to an integer
        input_value = input_value.replace('.','')
        input_value = input_value.replace(',','.')
        float(input_value)
        return True
    except ValueError:
        return False
def change_place_of_window(root):
    w = 800  # width for the Tk root
    h = 650  # height for the Tk root

    # get screen width and height
    ws = root.winfo_screenwidth()  # width of the screen
    hs = root.winfo_screenheight()  # height of the screen
    # calculate x and y coordinates for the Tk root window
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)

    # set the dimensions of the screen
    # and where it is placed
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))



def get_selection(title):

    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Prompt the user with a message box
    response = messagebox.askyesno("Eingeben?", title)

    # Return the user's response
    return response




def select_client(options):
    root = tk.Tk()
    root.title('Patientenauswahl:')
    # prompt =  "Von welcher Person willst du die Rechnung ausdrucken?",
    # tk.Label(root, text=prompt).pack()

    search_entry = tk.Entry(root, width=80)
    search_entry.pack(pady=10)
    searched_options = [options]

    def filter_list(event,so):
        searching_for = search_entry.get()
        if isinstance(searching_for,str):
            search_term = search_entry.get().lower()
            filtered_options = [option for option in options if search_term in option.lower()]
            so.append(filtered_options)

            # Clear the current listbox
            listbox.delete(0, tk.END)

            # Add filtered options to the listbox
            for option in filtered_options:
                listbox.insert(tk.END, option)
    search_entry.bind('<KeyRelease>', lambda event, so= searched_options: filter_list(event,so))

    frame = tk.Frame(root)
    frame.pack(pady=10)

    # Create a scrollable listbox
    listbox = tk.Listbox(frame, height=15, width=80, selectmode=tk.SINGLE)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH)

    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Link the scrollbar to the listbox
    listbox.config(yscrollcommand=scrollbar.set)

    # Populate the listbox with all options initially
    for option in options:
        listbox.insert(tk.END, option)

    selected_name = [""]
    def on_ok():
        selected = listbox.curselection()
        if selected:
            index, = listbox.curselection()
            last_selected = searched_options[-1]
            selected_name[0] =  last_selected[index]
            # selected_name = options[index]
            root.destroy()


    tk.Button(root,text="OK", command=on_ok,font=("Helvetica", 14) ).pack()
    root.mainloop()


    return selected_name[0]


def save_to_archive(invoicenumber,datetoday,clientname,invoice_start_date,invoice_end_date,summe,archive_which_invoices_path):
    # invoicenumber = int(invoicenumber)
    print("Reading in the cassabook")
    ws_archive_which_invoices = openpyxl.load_workbook(archive_which_invoices_path)
    archive_which_invoices = ws_archive_which_invoices.worksheets[0]
    invoiceduration = invoice_start_date.strftime("%d.%m.%Y")
    if invoice_end_date:
        invoiceduration += f" - {invoice_end_date.strftime('%d.%m.%Y')}"
    #datetoday = datetime.datetime.strptime(datetoday, "%d.%m.%Y")
    inputdata = [invoicenumber,datetoday, clientname, invoiceduration, summe]

    last_row_of_data = 0
    # so first row from bottom which is empty designates finish of data
    for row in archive_which_invoices:
        if any(cell.value is not None for cell in row):
            last_row_of_data += 1
        else:
            break
    # print(f"Last row of data in cassabook = {last_row_of_data}")
    # first row with Summe: designates Sum row
    sum_row = archive_which_invoices.max_row
    for index,row in enumerate(archive_which_invoices):
        if row[0].value == "Summe:":
            sum_row = index + 1
    # print(f"sum_row = {sum_row}")

    archive_which_invoices.insert_rows(last_row_of_data+1)
    sum_row += 1
    last_row_of_data += 1
    last_row = archive_which_invoices.max_row
    for col, value in zip(range(1, len(inputdata) + 1), inputdata):
        archive_which_invoices.cell(row=last_row_of_data, column=col, value=value)
        archive_which_invoices.cell(last_row_of_data,5).number_format = '€* #,##0.00'
        archive_which_invoices.cell(last_row_of_data,7).number_format = '€* #,##0.00'

        archive_which_invoices.cell(last_row_of_data,2).number_format = 'DD.MM.YYYY'

    cell_sum_invoiced = archive_which_invoices.cell(sum_row, 5)
    cell_sum_paid = archive_which_invoices.cell(sum_row, 7)
    cell_diff_inv_paid = archive_which_invoices.cell(sum_row+1, 7)

    cell_sum_invoiced.value = f"=SUM(E2:E{last_row_of_data})"
    cell_sum_paid.value = f"=SUM(G2:G{last_row_of_data})"
    cell_diff_inv_paid.value = f"=E{sum_row}-G{sum_row}"

    # sort the entries to invoice number

    rows = list(archive_which_invoices.iter_rows(values_only = True))
    header, data = rows[0], rows[1:last_row_of_data]
    print("The data of the cassabook is:")
    print(data)
    # sort to the first column
    # data.sort(key = lambda x: int(x[0]))
    # # overwrite the rows
    # for row_idx, row_data in enumerate(data,start=2): # start at two to skip the header
    #     for col_idx, value in enumerate(row_data, start=1):
    #         archive_which_invoices.cell(row= row_idx, column=col_idx, value= value)



    print(f"Saved the cassabook to: {archive_which_invoices_path}")
    ws_archive_which_invoices.save(archive_which_invoices_path)




def show_matrix_window(frame, matrix , head = ("",""),defaultservice = None):
    treeview = ttk.Treeview(frame, columns=head, show="headings",selectmode="extended")

    for colname in head:
        treeview.heading(colname, text=colname)

    item_ids = []

    for row in matrix:
        if defaultservice:
            row.append(defaultservice)
        rowtupel = tuple(row)
        if isinstance(rowtupel[1],pd.DataFrame):
            temp_list = list(rowtupel)
            temp_list[1] = rowtupel[1].to_string(index=False)
            rowtupel = tuple(temp_list)

        item_id = treeview.insert("", tk.END,values=rowtupel)
        item_ids.append(item_id)


    # place_comboboxes(treeview, item_ids, combobox_column=2)

    #
    # def motion_handler(tree, event):
    #     f = Font(font='TkDefaultFont')
    #
    #     # A helper function that will wrap a given value based on column width
    #     def adjust_newlines(val, width, pad=10):
    #         if not isinstance(val, str):
    #             return val
    #         else:
    #             words = val.split()
    #             lines = [[], ]
    #             for word in words:
    #                 line = lines[-1] + [word, ]
    #                 if f.measure(' '.join(line)) < (width - pad):
    #                     lines[-1].append(word)
    #                 else:
    #                     lines[-1] = ' '.join(lines[-1])
    #                     lines.append([word, ])
    #
    #             if isinstance(lines[-1], list):
    #                 lines[-1] = ' '.join(lines[-1])
    #
    #             return '\n'.join(lines)
    #
    #     if (event is None) or (tree.identify_region(event.x, event.y) == "separator"):
    #         # You may be able to use this to only adjust the two columns that you care about
    #         # print(tree.identify_column(event.x))
    #
    #         col_widths = [tree.column(cid)['width'] for cid in tree['columns']]
    #
    #         for iid in tree.get_children():
    #             new_vals = []
    #             for (v, w) in zip(tree.item(iid)['values'], col_widths):
    #                 new_vals.append(adjust_newlines(v, w))
    #             tree.item(iid, values=new_vals)
    #
    # def calculate_row_height(tree):
    #     """Calculate and adjust row height to fit the text."""
    #     # Retrieve the existing Treeview font
    #     style = ttk.Style()
    #     font = Font(name="TkDefaultFont", exists=True)  # Use the existing font
    #
    #     # Determine the required height for the tallest text
    #     max_text_height = 0
    #     for item in tree.get_children():
    #         row_values = tree.item(item, "values")
    #         for text in row_values:
    #             # Measure the height of the text
    #             max_text_height = max(max_text_height, font.metrics("linespace"))
    #
    #     # Adjust the row height dynamically
    #     style.configure("Treeview", rowheight=max_text_height + 10)  # Add padding
    #
    # treeview.bind('<B1-Motion>', partial(motion_handler, treeview))
    # motion_handler(treeview, None)   # Perform initial wrapping
    # calculate_row_height(treeview)
    return treeview, item_ids


def inquire_new_services(popup,output,services,datatoinquire):


    entryrowsidx = [0]

    labels = [tk.Label(popup, text=onedatalabel) for onedatalabel in datatoinquire]

    for colnumber, label in enumerate(labels):
        label.grid(column=colnumber, row=0)

    def on_combobox_change(e, combobox, descriptionlabel, hourlyratelabel):
        selected = combobox.get()
        descriptionlabel.config(text=services["Beschreibung"][services["Leistung"] == selected].values[0])
        hourlyratelabel.config(text=services["Stundensatz"][services["Leistung"] == selected].values[0])

    internal_refs = []
    descriptions = []
    hourlyrates = []
    dateentries = []
    timeentries = []
    hourentries = []

    def is_valid_float(value):
        """
        Accepts a single float using either a comma or dot as decimal separator.
        Examples: '1.2', '1,2', '-3.0', '4'
        """
        if value:
            raw = value.replace(",", ".")
            try:
                value = float(raw)
                return True
            except ValueError:
                return False
        else:
            return True

    vcmd = (popup.register(is_valid_float), "%P")

    def make_new_inputrow():
        if max(entryrowsidx) < 3:
            row = max(entryrowsidx) + 1
            popup.grid_rowconfigure(row, minsize=60)
            description = tk.Label(popup, text=services["Beschreibung"][0])
            description.grid(column=1, row=row)
            hourlyrate = tk.Label(popup, text=services["Stundensatz"][0])
            hourlyrate.grid(column=2, row=row)
            dateentry = DateEntry(popup, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern="d.m.yyyy")
            dateentry.grid(column=3, row=row)
            timeentry = tk.Entry(popup, validate="key", validatecommand=vcmd)
            timeentry.grid(column=4, row=row)
            hourentry = ttk.Entry(popup)
            hourentry.grid(column=5, row=row)

            cb = ttk.Combobox(popup, values=services["Leistung"].tolist(), state="readonly")
            cb.set(services["Leistung"].tolist()[0])
            cb.grid(column=0, row=row)
            cb.bind("<<ComboboxSelected>>",
                    lambda e: on_combobox_change(e, cb, descriptions[row - 1], hourlyrates[row - 1]))

            descriptions.append(description)
            hourlyrates.append(hourlyrate)
            dateentries.append(dateentry)
            timeentries.append(timeentry)
            hourentries.append(hourentry)
            internal_refs.append(cb)

            entryrowsidx.append(row)
        else:
            moreinputsbutton.config(text="Mehr gehen nicht")

    make_new_inputrow()

    def get_input():
        if timeentries[0].get():
            for row in entryrowsidx:
                if row > 0:
                    rowoutput = []
                    for data in [internal_refs, descriptions, hourlyrates, dateentries, timeentries, hourentries]:
                        entry = data[row - 1]
                        if isinstance(entry, tk.Label):
                            outputthis = entry.cget("text")
                        else:
                            outputthis = entry.get()

                        rowoutput.append(outputthis)
                    output.append(rowoutput)

            popup.destroy()
        else:
            errortext.config(text = "Anzahl Minuten darf nicht leer sein")


    def insert_row_above_button():
        print(f"New entry row {entryrowsidx}")
        nt_entryrowsthis = max(entryrowsidx) + 1
        for widget in popup.grid_slaves():
            info = widget.grid_info()
            r = info['row']
            if r >= nt_entryrowsthis:
                widget.grid_forget()
                widget.grid(row=r + 1, column=info['column'], sticky=info.get('sticky', ''))

        make_new_inputrow()

    moreinputsbutton = tk.Button(popup, text="Mehr Inputs", height=2, width=20,
                                 command=lambda: insert_row_above_button())
    #moreinputsbutton.grid(row=4, column=2, columnspan=2)
    errortext = Label(popup, text = "")
    errortext.grid(row=4, column=3, columnspan=2)
    tk.Button(popup, text="Hinzufügen", height=2, width=20, command=lambda: get_input()).grid(row=5, column=3, columnspan=2)


def ask_to_save(data_list, hourdata, services,added_hourdata):
    root = tk.Tk()
    root.title("Überprüfung")
    root.geometry("1600x1000+50+30")

    default_font = tk.font.nametofont("TkDefaultFont")
    bigger_font = default_font.copy()
    bigger_font.configure(size=16)

    root.grid_rowconfigure(1, weight=1)  # Middle row (frames) expands
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(2, weight=0)
    root.grid_rowconfigure(3, weight=0)
    # Top label (row 0)
    top_label = tk.Label(root, text="Hier sind alle Daten nochmal zusammengefasst", bg="white", font=bigger_font)
    top_label.grid(row=0, column=0, columnspan=2, sticky="ew")

    # Left frame (row 1, col 0)
    left_frame = tk.Frame(root)
    left_frame.grid(row=1, column=0, sticky="nsew")
    left_frame.grid_rowconfigure(1, weight=1)
    left_frame.grid_columnconfigure(0, weight=1)

    # Right frame (row 1, col 1)
    right_frame = tk.Frame(root)
    right_frame.grid(row=1, column=1, sticky="nsew")
    right_frame.grid_rowconfigure(1, weight=1)
    right_frame.grid_columnconfigure(0, weight=1)

    # Treeview in left frame
    left_label = tk.Label(left_frame, text="Daten PatientIn",font = bigger_font)
    left_label.grid(row=0, column=0, sticky="ew")
    data_list_without_hours = [x for x in data_list if "Stundeninfo" not in x[0] ]
    datalist, datalist_items = show_matrix_window(left_frame,data_list_without_hours,  head = ("","Wert"))
    datalist.grid(row = 1, column=0, sticky="nsew")

    # Treeview in right frame
    right_label = tk.Label(right_frame, text="Stundendaten",font = bigger_font)
    right_label.grid(row=0, column=0, sticky="ew")
    hourlist, hourlist_items = show_matrix_window(right_frame,list(hourdata.values),head=tuple(hourdata.columns))
    hourlist.grid(row=1, column=0, sticky="nsew")
    
    def add_services(root,hourdata,hourlist, hourlist_items, services):
        print("Add another service")
        popup = tk.Toplevel(root)
        popup.title("Zusätzliche Leistung")
        popup.geometry("900x300+150+50")

        datatoinquire = services.columns.tolist()
        datatoinquire.append("Datum")
        datatoinquire.append("Minuten")
        datatoinquire.append("Uhrzeit")
        added_data = []
        inquire_new_services(popup, added_data,services, datatoinquire)
        popup.wait_window()

        added_data = pd.DataFrame(added_data, columns=datatoinquire)
        added_data["Datum"] = pd.to_datetime(added_data["Datum"])
        added_data["Name"] = data_list[0][1]

        hourdata = pd.concat([hourdata,added_data])
        added_hourdata.append(added_data)
        for item in hourlist.get_children():
            hourlist.delete(item)

        for idx,row in hourdata.iterrows():
             hourlist.insert("", tk.END, values=tuple(row))




    tk.Button(right_frame, text="Füge noch andere Leistungen hinzu", command=lambda : add_services(root,hourdata,hourlist, hourlist_items, services)).grid(row=2, column=0, columnspan=2, sticky="ew")



    spaceframe = tk.Frame(root, height=50).grid(row=2, column=0, columnspan=2, sticky="ew")

    # Label above bottom frame (row 2)
    middle_label = tk.Label(root, text="Soll ich nun einen Rechnung mit diesen Daten erstellen?", font = bigger_font)
    middle_label.grid(row=3, column=0, columnspan=2, sticky="ew")
    # Bottom frame (row 3)
    yes_no_frame = tk.Frame(root)
    yes_no_frame.grid(row=4, column=0, columnspan=2, sticky="ew")
    yes_no_frame.grid_propagate(False)

    ttk.Style().configure('Treeview', rowheight=30)

    answer = [False]

    def button_press(y_n):
        if y_n == "Y":
            print("Selected saving: yes")
            answer[0] = True
        elif y_n == "N":
            answer[0] = False
            print("Selected saving: no \\return without saving")
        root.destroy()

    # Pack widgets side by side inside the last row frame
    button_container = tk.Frame(yes_no_frame)
    button_container.pack(anchor="center", pady=20)
    tk.Button(button_container, text="Ja", command=lambda: button_press("Y")).pack(side="left", padx=10)
    tk.Button(button_container, text="Nein", command=lambda: button_press("N")).pack(side="left", padx=10)

    # Start the app
    root.mainloop()
    return answer[0]



def get_items(clientname,hourdata,services,lastdate,defaultservice = None):
    defaulthourlyrate = services.loc[services["Leistung"] == defaultservice,"Stundensatz"].values[0]
    hourdatacopy = hourdata.copy()
    hourdatacopy  =hourdatacopy.sort_values(by="Datum", ascending= False)

    returndata = []
    somedateselected = [False]

    root = tk.Tk()
    root.geometry("1600x700+50+0")
    root.title(f"Auswahl des Rechnungszeitraums für die Rechnung von {clientname}")

    left_frame = tk.Frame(root)
    right_frame = tk.Frame(root)

    left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)


    label1 = Label(left_frame, text='Rechnung ab: ', font=("Helvetica", 14) )
    label1.pack(ipadx=10, ipady=10)

    cal1 = Calendar(left_frame,
                   font="Arial 14", selectmode='day')
    cal1.pack(fill="both", expand=True)
    label2 = Label(left_frame, text='bis: ', font=("Helvetica", 14) )
    label2.pack(ipadx=10, ipady=10)
    cal2 = Calendar(left_frame,
                   font="Arial 14", selectmode='day')
    cal2.pack(fill="both", expand=True)


    data_list =   hourdatacopy.values.tolist()
    if lastdate:
        Title = tk.Label(right_frame, text=f"Die letzte Rechnung für {clientname} wurde am {pd.to_datetime(lastdate).strftime('%d.%m.%Y')} erstellt").pack(pady=10)
    head = hourdatacopy.columns.tolist()
    head.append("Leistung")
    head.append("Stundensatz")
    datelist,datelist_item_ids = show_matrix_window(right_frame,data_list, head = head, defaultservice = defaultservice )
    comboboxes_services = []
    inputs_prices = []

    def place_comboboxes_services(treeview, treeview_item_ids, combobox_options,comboboxes_column):
        for index, item_id in enumerate(treeview_item_ids):
            bbox = treeview.bbox(item_id, column=comboboxes_column)
            if not bbox:
                continue
            x, y, width, height = bbox
            value = treeview.set(item_id, comboboxes_column)

            cb = ttk.Combobox(treeview, values=combobox_options.tolist(), state="readonly")
            cb.set(value)
            cb.place(x=x, y=y, width=width, height=height)
            #dont need to update the tree, all variable are taken then form the comboboxes
            comboboxes_services.append(cb)

    def is_valid_float(value):
        """
        Accepts a single float using either a comma or dot as decimal separator.
        Examples: '1.2', '1,2', '-3.0', '4'
        """
        if value:
            raw = value.replace(",", ".")
            try:
                value = float(raw)
                return True
            except ValueError:
                return False
        else: return True
    vcmd = (root.register(is_valid_float), "%P")

    def place_inputs_prices(treeview, treeview_item_ids, inputs_column,defaulthourlyrate):
        for index, item_id in enumerate(treeview_item_ids):
            bbox = treeview.bbox(item_id, column=inputs_column)
            if not bbox:
                continue
            x, y, width, height = bbox
            entry = ttk.Entry(treeview, validate="key", validatecommand=vcmd)
            entry.insert(0,defaulthourlyrate)
            entry.place(x=x, y=y, width=width, height=height)
            inputs_prices.append(entry)
    datelist.pack()
    # if dropdown menu doesnot show, make waittime longer (this is a bad workaround)

    # def place_comboboxes_inputs_on_treeview_after_loading():
    #     if len(datelist.get_children()) > 0:
    #         print("Place comboboxes")
    #         place_comboboxes_services(datelist,datelist_item_ids, services["Leistung"],"Leistung")
    #         place_inputs_prices(datelist, datelist_item_ids, "Stundensatz", defaulthourlyrate)
    #         return
    #     root.after(100,place_comboboxes_inputs_on_treeview_after_loading)
    # place_comboboxes_inputs_on_treeview_after_loading()
    waittimetoloadinputoverlay = 1500

    root.after(waittimetoloadinputoverlay, lambda: place_comboboxes_services(datelist,datelist_item_ids, services["Leistung"],"Leistung"))  # Wait for Treeview to render
    root.after(waittimetoloadinputoverlay, lambda: place_inputs_prices(datelist,datelist_item_ids, "Stundensatz",defaulthourlyrate))  # Wait for Treeview to render

    def on_date_change(e,somedateselected):
        somedateselected.append(True)
        startdate = cal1.selection_get()
        enddate = cal2.selection_get()
        print(f"Daterange changed, {startdate} - {enddate}")
        selected_dates = (hourdatacopy['Datum'].dt.date > startdate) & (hourdatacopy['Datum'].dt.date < enddate)
        item_ids = np.array(datelist_item_ids)
        datelist.selection_set(item_ids[selected_dates].tolist())
        # print(x)

    cal1.bind("<<CalendarSelected>>", lambda e: on_date_change(e, somedateselected))
    cal2.bind("<<CalendarSelected>>", lambda e: on_date_change(e, somedateselected))

    def on_ok(root,datelist,comboboxes_services,inputs_prices,hourdata,returndata):
        all_items = datelist.get_children()
        selected_items = datelist.selection()  # returns a tuple of selected item IDs

        if selected_items:
            matrix = hourdata.copy()
            matrix = matrix.sort_values(by="Datum", ascending=False)
            comboboxinput = [combobox.get() for combobox in comboboxes_services]
            floatinputs = [input.get() for input in inputs_prices]

            matrix["Leistung"] = comboboxinput
            matrix["Stundensatz"] = floatinputs


            for treerow, (index,matrixrow) in zip(all_items,matrix.iterrows()):
                if treerow in selected_items:
                    returndata.append(matrixrow.tolist())
            root.destroy()
        else:
            errorlabel.config(text="Wähle mindestens ein Datum aus")
            print("No dates selected")
    errorlabel = tk.Label(left_frame, text="")
    errorlabel.pack(pady=10)
    tk.Button(left_frame, text="ok",height=2, width=20, font="Arial 14", command=lambda: on_ok(root,datelist,comboboxes_services,inputs_prices,hourdatacopy,returndata)).pack(pady=10)

    root.mainloop()

    #this happens after root.destroy
    returndata = pd.DataFrame(returndata, columns=head)
    # check whether something was selected
    if not np.any(somedateselected):
        date1 = returndata.Datum.min()
        date2 = returndata.Datum.max()
    else:
        date1 = cal1.get_date()
        date2 = cal2.get_date()
        date1 = pd.to_datetime(date1)
        date2 = pd.to_datetime(date2)

    return returndata, date1, date2


def input_new_person(allclientdata_path):
    allclientdata = pd.read_excel(allclientdata_path, index_col=0, header=None, sheet_name=None)

    datatoinquire = list(allclientdata["Vorlage"].index)
    #

    root = Tk()
    # initialise the boxes
    labels = [Label(root, text = onedatalabel) for onedatalabel in datatoinquire]
    entries = [Entry(root) for x in range(0,len(datatoinquire))]


    #position the inquiries in a nice table
    for rownumber, (label, entry) in enumerate(zip(labels, entries)):
        label.grid(column=0, row=rownumber)
        if rownumber != 1 and rownumber != 2:
            entry.grid(column=1, row=rownumber)

    #make the dropdowns
    sexoptions = ["w","m"]
    childoptions =["ja", "nein"]


    child = StringVar(root)
    child.set(childoptions[0])
    childoptiondropdown = OptionMenu(root, child, *childoptions)
    childoptiondropdown.grid(column=1, row=1)

    sex = StringVar(root)
    sex.set(sexoptions[0])
    sexoptiondropdown = OptionMenu(root, sex, *sexoptions)
    sexoptiondropdown.grid(column=1, row=2)


    userinputs = []
    def command():
        for entry in entries:
            userinputs.append(entry.get())
        userinputs[1] = child.get()
        userinputs[2] = sex.get()
        root.destroy()

    Button(root, text="Speichern", command=command).grid(column=1,row =len(datatoinquire)+1)
    root.mainloop()
    # parse datetime inputs
    try:
        userinputs[3] = dateutil.parser.parse(userinputs[3])
    except:
        print("Das Datum ist falsch eingegeben")

    try:
        userinputs[12] = dateutil.parser.parse(userinputs[12])
    finally:

        userinputsdict = dict(zip(datatoinquire, userinputs))


        excelsheet_with_added_person = openpyxl.load_workbook(allclientdata_path)#
        excelsheet_with_added_person.iso_dates = True
        sheet_new_person = excelsheet_with_added_person.create_sheet(userinputsdict["Name"])


        for row, (dataname,userinput) in enumerate(zip(datatoinquire,userinputs)):
            sheet_new_person.cell(row=row+1, column=1).value = dataname
            sheet_new_person.cell(row=row+1, column=2).value = userinput
        excelsheet_with_added_person.save(allclientdata_path)

        print("Ich habe eine neues Blatt für " + userinputsdict["Name"] + " zur PatienInneninformations Exceldatei hinzugefügt")
        print("Mit diesen Einträgen: ")
        pprint.pprint(userinputsdict)
        return userinputsdict


def insert_hourdata(allhourdata_path,clientname):
    root = Tk()
    root.title("Therapiedaten für " + clientname)
    root.geometry("650x500+120+120")

    # empty arrays for your Entrys and StringVars
    text_var = []
    entries = []

    # callback function to get your StringVars
    clienthourdata = []
    def command():
        matrix = []
        for i in range(rows):
            matrix.append([])
            for j in range(cols):
                matrix[i].append(text_var[i][j].get())
        clienthourdata.append(matrix)
        root.destroy()

    labelnames = ["Datum Therapie (im Format wie 1.1.2023)", "Einheitslänge in min"]
    for column in range(0,2):
        Label(root, text=labelnames[column], font=('arial', 10, 'bold'),
          bg="bisque2").place(x=20 + 110*column, y=20)

    x2 = 0
    y2 = 0
    rows, cols = (10,2)
    for i in range(rows):
        # append an empty list to your two arrays
        # so you can append to those later
        text_var.append([])
        entries.append([])
        for j in range(cols):
            # append your StringVar and Entry
            text_var[i].append(StringVar())
            entries[i].append(Entry(root, textvariable=text_var[i][j],width=10))
            entries[i][j].place(x=60 + x2, y=50 + y2)
            x2 += 100

        y2 += 30
        x2 = 0
    button= Button(root,text="Daten speichern", bg='bisque3', width=15, command=command)
    button.place(x=160,y=350)
    root.mainloop()


    clienthourdata = clienthourdata[0]
    datestherapy = list(filter(None,[row[0] for row in clienthourdata] ))
    lengththerapy = list(filter(None,[row[1] for row in clienthourdata]))
    datestherapy = list(map(lambda x: dateutil.parser.parse(x, dayfirst = True), datestherapy))
    lengththerapy = list(map(lambda x: float(x), lengththerapy))

    excelsheet_hourdata = openpyxl.load_workbook(allhourdata_path)  #
    excelsheet_hourdata.iso_dates = True
    sheet = excelsheet_hourdata["Stundendaten"]

    for dateonetherapy, lengthonetherapy in zip(datestherapy, lengththerapy):
        newRowLocation = sheet.max_row + 1
        sheet.cell(row=newRowLocation, column=1).value = dateonetherapy
        sheet.cell(row=newRowLocation, column=2).value = clientname
        sheet.cell(row=newRowLocation, column=3).value = lengthonetherapy


    excelsheet_hourdata.save(allhourdata_path)
    namehourdata = pd.DataFrame([datestherapy,[clientname for x in range(0,len(datestherapy))],lengththerapy])
    namehourdata = namehourdata.transpose()
    namehourdata.columns = ['Datum', 'Name', 'Minuten']

    return(namehourdata)


# start,end = get_date()