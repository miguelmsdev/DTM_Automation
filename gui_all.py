import customtkinter
from customtkinter import filedialog
import os
from tkcalendar import Calendar
from datetime import datetime, date
from tkinter import messagebox
from tkcalendar import DateEntry
import tkinter as tk
import automation_all
from functools import partial
import time

company_name = None

entry_end_date = None
entry_start_date = None

entry_end_date_var = None
entry_start_date_var = None

relatorio_path = ""
scans_path = ""

file_name_scans = None
file_name_relatorio = None

start_date_value = None
end_date_value = None

def get_value_button_select_relatorio(selection):
        global company_name
        company_name = selection

def get_value_button_select_scans(selection):
        global company_name
        company_name = selection

def get_value_optionmenu_1(selection):
        global company_name
        company_name = selection

def run_automation():
    print("sou eu")
    print(scans_path)
    print(company_name)
    print(start_date_value)
    print(end_date_value)
    print(relatorio_path)
    automation_all.run_automation(company_name, relatorio_path, scans_path, start_date_value, end_date_value)

def select_file_scans(label_selected_scans_name):
    global file_name_scans
    global scans_path

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])

    if file_path:
        _, file_ext = os.path.splitext(file_path)
        if file_ext.lower() not in [".xlsx", ".xlsm"]:
            messagebox.showerror(
                "Error",
                "Invalid file format. Please select an Excel file (.xlsx or .xlsm).",
            )
            return
        _, file_name_scans = os.path.split(file_path)
        if file_name_scans != None:
            scans_path = file_path
            label_selected_scans_name.configure(text=file_name_scans)
    else:
        scans_path = ""
        label_selected_scans_name.configure(text="")



def select_file_relatorio(label_selected_relatorio_name):
    global file_name_relatorio
    global relatorio_path

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
    if file_path:
        _, file_ext = os.path.splitext(file_path)
        if file_ext.lower() not in [".xlsx", ".xlsm"]:
            messagebox.showerror(
                "Error",
                "Invalid file format. Please select an Excel file (.xlsx or .xlsm).",
            )
            return

        _, file_name_relatorio = os.path.split(file_path)
        if file_name_relatorio != None:
            relatorio_path = file_path
            label_selected_relatorio_name.configure(text=file_name_relatorio)
    else:
        relatorio_path = ""
        label_selected_relatorio_name.configure(text="")


def format_end_date(start_date, start_date_var, end_date, end_date_var):
    text = end_date_var.get()
    cursor_position = end_date.index("insert")
    # Remove non-digit characters
    text = "".join(filter(str.isdigit, text))

    original_length = len(text)  # Store the original text length

    # Reformat the date as dd/mm/yyyy
    if len(text) >= 1:
        text = text[:2] + "/" + text[2:]
        
        dd_first_num = int(text[0])  # Extract the first digit of the day number as an integer
        if dd_first_num < 0 or dd_first_num > 3:
            text = text[0:0]

    if len(text) >= 2:
        if text[1].isdigit():
            dd_second_num = int(text[1])  # Extract the second digit of the day number as an integer
            if dd_first_num == 0 and dd_second_num == 0:
                text = text[0:0]
            if dd_first_num == 3 and dd_second_num not in [0,1]:
                text = text[0:0]

    if len(text) >= 4:
        text = text[:5] + "/" + text[5:]
        cursor_position += 1  # Move cursor one position forward after inserting first slash
        
        if text[3].isdigit():
            mm_first_num = int(text[3])  # Extract the first digit of the day number as an integer
            if mm_first_num not in [0,1]:
                text = text[:3]
    if len(text) >= 5:
        if text[4].isdigit():
            mm_second_num = int(text[4])  # Extract the second digit of the day number as an integer
            if mm_first_num == 0 and mm_second_num == 0:
                text = text[:4]
            if mm_first_num == 1 and mm_second_num > 2:
                text = text[:4]

    if len(text) >= 7:
        cursor_position += 1
        if int(text[6]) != 2 and text[6].isdigit():
            text = text[:6]
        cursor_position += 1
    if len(text) >= 8:
        if int(text[7]) != 0:
            text = text[:7]
        cursor_position += 1
    if len(text) >= 9:
        if int(text[8]) != 2:
            text = text[:8]
        cursor_position += 1
    if len(text) >= 10:
        if int(text[9]) not in [3,4,5,6,7,8]:
            text = text[:9]

    # Limit the character length to 10
    text = text[:10]
    cursor_position = min(
        cursor_position, 10
    )  # Ensure the cursor position is within the limit

    end_date_var.set(text)

    # Set the cursor position at the end of the text
    end_date.icursor(cursor_position)

    print(text)
    return text

def format_start_date(start_date, start_date_var, end_date, end_date_var):
    text = start_date_var.get()
    cursor_position = start_date.index("insert")

    # Remove non-digit characters
    text = "".join(filter(str.isdigit, text))

    # Reformat the date as xx/xx/xxxx
    if len(text) >= 1:
        text = text[:2] + "/" + text[2:]
        
        dd_first_num = int(text[0])  # Extract the first digit of the day number as an integer
        if dd_first_num < 0 or dd_first_num > 3:
            text = text[0:0]

    if len(text) >= 2:
        if text[1].isdigit():
            dd_second_num = int(text[1])  # Extract the second digit of the day number as an integer
            if dd_first_num == 0 and dd_second_num == 0:
                text = text[0:0]
            if dd_first_num == 3 and dd_second_num not in [0,1]:
                text = text[0:0]

    if len(text) >= 4:
        text = text[:5] + "/" + text[5:]
        cursor_position += 1  # Move cursor one position forward after inserting first slash
        
        if text[3].isdigit():
            mm_first_num = int(text[3])  # Extract the first digit of the day number as an integer
            if mm_first_num not in [0,1]:
                text = text[:3]
    if len(text) >= 5:
        if text[4].isdigit():
            mm_second_num = int(text[4])  # Extract the second digit of the day number as an integer
            if mm_first_num == 0 and mm_second_num == 0:
                text = text[:4]
            if mm_first_num == 1 and mm_second_num > 2:
                text = text[:4]

    if len(text) >= 7:
        cursor_position += 1
        if int(text[6]) != 2 and text[6].isdigit():
            text = text[:6]
        cursor_position += 1
    if len(text) >= 8:
        if int(text[7]) != 0:
            text = text[:7]
        cursor_position += 1
    if len(text) >= 9:
        if int(text[8]) != 2:
            text = text[:8]
        cursor_position += 1
    if len(text) >= 10:
        if int(text[9]) not in [3,4,5,6,7,8]:
            text = text[:9]

    # Limit the character length to 10
    text = text[:10]
    cursor_position = min(
        cursor_position, 10
    )  # Ensure the cursor position is within the limit

    start_date_var.set(text)

    # Set the cursor position at the end of the text
    start_date.icursor(cursor_position)

    print(text)
    return


def check_dates(start_date, start_date_var, end_date, end_date_var):
    global start_date_value
    global end_date_value
    if start_date_var is None or end_date_var is None:
        pass
    else:
        start_date_value = start_date_var.get()
        end_date_value = end_date_var.get()
        if len(start_date_value) == len(end_date_value) == 10:
            try:
                start_date_value = datetime.strptime(start_date_value, "%d/%m/%Y").strftime("%d/%m/%Y")
                end_date_value = datetime.strptime(end_date_value, "%d/%m/%Y").strftime("%d/%m/%Y")
                print(start_date_value)
                print(end_date_value)
            except ValueError:
                if not hasattr(check_dates, "error_shown") or (time.time() - check_dates.error_shown) > 5:
                    messagebox.showerror(
                        "Dia inexistente",
                        "Por favor, insira um dia que esse mês possua",
                    )
                    check_dates.error_shown = time.time()
                if entry_end_date.get():
                    entry_end_date.set_date(None)
                if entry_start_date.get():
                    entry_start_date.set_date(None)
                return

            if start_date_value > end_date_value:
                if not hasattr(check_dates, "error_shown") or (time.time() - check_dates.error_shown) > 5:
                    messagebox.showerror(
                        "Data inválida", "Data Final não pode ser maior do que Data Inicial"
                    )
                    check_dates.error_shown = time.time()
                if entry_end_date.get():
                    entry_end_date.set_date(None)
                if entry_start_date.get():
                    entry_start_date.set_date(None)
                return
        
def trace_add_write(start_date, start_date_var, end_date, end_date_var, functions):
    def wrapped_functions(*args):
        for func in functions:
                func(start_date, start_date_var, end_date, end_date_var)

    start_date_var.trace_add("write", wrapped_functions)
    end_date_var.trace_add("write", wrapped_functions)

def create_focus_in_wrapper(start_date, start_date_var, end_date, end_date_var):
    def wrapper(event):
        if isinstance(start_date_var, DateEntry) or start_date_var == None or (isinstance(end_date_var, DateEntry) or end_date_var == None) or (isinstance(end_date, DateEntry) or end_date == None) or(isinstance(start_date, DateEntry) or start_date == None):
            pass
        else:
            start_date_var.set('')
            end_date_var.set('')
            start_date.configure(foreground='black')
            end_date.configure(foreground='black')
    return wrapper

def create_focus_out_wrapper(start_date, start_date_var, end_date, end_date_var):
    def wrapper(event):
        if isinstance(start_date_var, DateEntry) or start_date_var == None or (isinstance(end_date_var, DateEntry) or end_date_var == None) or (isinstance(end_date, DateEntry) or end_date == None) or(isinstance(start_date, DateEntry) or start_date == None):
            pass
        else:
            if not start_date_var.get():
                start_date_var.set('Start Date')
                start_date.configure(foreground='grey')
            if not end_date_var.get():
                end_date_var.set('End Date')
                end_date.configure(foreground='grey')
    return wrapper


def remove_default_end_date(start_date, start_date_var, end_date, end_date_var):
    if end_date_var.get() == "dd/mm/yyyy":
        end_date_var.set("")
        end_date.configure(foreground="black")


def remove_default_start_date(start_date, start_date_var, end_date, end_date_var):
    if start_date_var.get() == "dd/mm/yyyy":
        start_date_var.set("")
        start_date.configure(foreground="black")

def run_main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    main_file = os.path.join(current_dir, "main.py")
    os.system(f'python "{main_file}"')

def run_gui():
    global company_name
    global entry_end_date
    global entry_start_date
    global entry_end_date_var
    global entry_start_date_var

    functions_involving_date = [format_start_date, check_dates, format_end_date, remove_default_start_date, remove_default_end_date]
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("dark-blue")

    root = customtkinter.CTk()
    root.title("Recebidor de MTR")
    root.geometry("1100x800")

    current_year = datetime.now().year

    frame = customtkinter.CTkFrame(master=root)
    frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.7, relheight=0.90)

    # Configure center column to expand and center the widgets
    frame.columnconfigure(1, weight=2)
    frame.columnconfigure(2, weight=1)
    frame.columnconfigure(3, weight=2)


    label = customtkinter.CTkLabel(
        master=frame, text="Recebidor de MTR", font=("Verdana", 24)
    )
    label.grid(row=0, column=2, pady=12, padx=70, sticky="nsew")

    selected_file2 = ""

    label_selected_scans_name = customtkinter.CTkLabel(
        master=frame, text=selected_file2, font=("Verdana", 12, "bold")
    )
    label_selected_scans_name.grid(row=3, column=2, pady=(0, 6), padx=30, sticky="n")

    button_select_scans = customtkinter.CTkButton(
        master=frame,
        text="Selecionar MTR Escaneados",
        font=("Verdana", 12, "bold"),
        command=lambda: select_file_scans(label_selected_scans_name)
    )
    button_select_scans.grid(row=1, column=2, pady=(26, 0), padx=70, sticky="nsew")


    label_selected_scans = customtkinter.CTkLabel(
        master=frame, text="Planilha Selecionada:", font=("Verdana", 12, "bold")
    )
    label_selected_scans.grid(row=2, column=2, pady=(8, 0), padx=30, sticky="nsew")




    company_label = customtkinter.CTkLabel(
        master=frame, text="Empresa:", font=("Verdana", 12)
    )
    company_label.grid(row=4, column=2, pady=(4, 0), padx=70, sticky="nsew")

    optionmenu_1 = customtkinter.CTkOptionMenu(
        master=frame,
        dynamic_resizing=False,
        values=["Telar", "Comér", "Cinco Estrelas", "Consórcio DBO"],
        font=("Verdana", 12, "bold"),
        width=230,
        anchor="center",
        command=get_value_optionmenu_1
    )
    
    if company_name is None:
        company_name = optionmenu_1.get()

    optionmenu_1.grid(row=5, column=2, pady=6, padx=50)

    label_start_date = customtkinter.CTkLabel(
        master=frame, text="Data Inicial:", font=("Verdana", 12, "bold")
    )
    label_start_date.grid(
        row=6,
        column=2,
        pady=(14, 0),
        padx=70,
    )


    entry_start_date_var = customtkinter.StringVar()
    entry_start_date = DateEntry(
        master=frame,
        textvariable=entry_start_date_var,
        selectmode="day",
        date_pattern="dd/mm/yyyy",
        showweeknumbers=False,
        showothermonthdays=False,
        weekendbackground="#ededed",
        weekendforeground="black",
        placeholder_text="dd/mm/yyyy",
        font=("Verdana", 11),
        selectbackground="gray",
        selectforeground="black",
        foreground="gray"
    )



    entry_start_date.grid(row=7, column=2, padx=(0, 0), pady=10, sticky="ns")

    wrapper_focus_in = create_focus_in_wrapper(entry_start_date, entry_start_date_var, entry_end_date, entry_end_date_var)
    wrapper_focus_out = create_focus_out_wrapper(entry_start_date, entry_start_date_var, entry_end_date, entry_end_date_var)

    entry_start_date.bind("<FocusIn>", wrapper_focus_in)
    entry_start_date.bind("<FocusOut>", wrapper_focus_out)
    entry_start_date_var.set("dd/mm/yyyy")

    label_end_date = customtkinter.CTkLabel(
        master=frame, text="Data Final:", font=("Verdana", 12, "bold")
    )
    label_end_date.grid(row=8, column=2, pady=(14, 0), padx=70, sticky="nsew")



    entry_end_date_var = customtkinter.StringVar()
    entry_end_date = DateEntry(
        master=frame,
        textvariable=entry_end_date_var,
        selectmode="day",
        date_pattern="dd/mm/yyyy",
        showweeknumbers=False,
        showothermonthdays=False,
        weekendbackground="#ededed",
        weekendforeground="black",
        placeholder_text="dd/mm/yyyy",
        font=("Verdana", 11),
        selectbackground="gray",
        selectforeground="black",
        foreground="gray"
    )

    entry_end_date.grid(row=9, column=2, padx=(0, 0), pady=10, sticky="ns")
    entry_end_date.bind("<FocusIn>", wrapper_focus_in)
    entry_end_date.bind("<FocusOut>", wrapper_focus_out)
    entry_end_date_var.set("dd/mm/yyyy")

    trace_add_write(entry_start_date, entry_start_date_var, entry_end_date, entry_end_date_var, functions_involving_date)

    button_select_relatorio = customtkinter.CTkButton(
        master=frame,
        text="Selecionar Relatório",
        font=("Verdana", 12, "bold"),
        command=lambda: select_file_relatorio(label_selected_relatorio_name)
    )
    button_select_relatorio.grid(row=10, column=2, pady=(26, 6), padx=70, sticky="nsew")

    selected_file = ""
    label_selected_relatorio = customtkinter.CTkLabel(
        master=frame, text="Relátorio Selecionado:", font=("Verdana", 12, "bold"))
    label_selected_relatorio.grid(row=11, column=2, pady=(10, 0), padx=30, sticky="nsew")

    label_selected_relatorio_name = customtkinter.CTkLabel(
        master=frame, text=selected_file, font=("Verdana", 12, "bold")
    )
    label_selected_relatorio_name.grid(row=12, column=2, pady=(4, 6), padx=30, sticky="ns")


    button_run_main = customtkinter.CTkButton(master=frame, text="Iniciar Recebimento", command=run_automation)

    button_run_main.grid(row=13, column=2)
    root.mainloop()