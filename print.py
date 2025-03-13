import os
import tkinter
from tkinter import filedialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui
import time
import win32api
import win32print

# Upload files from select button
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    for file in files:
        file_list.insert(tkinter.END, file)

def drag_files(event):
    files = root.tk.splitlist(event.data)
    for file in files:
        file_list.insert(tkinter.END, file)

def open_files():
    default_printer = win32print.GetDefaultPrinter()
    total_files = file_list.size()
    progress['maximum'] = total_files

    printer_name = get_select_printer()
    if printer_name:
        try:
            win32print.SetDefaultPrinter(printer_name)
            print(f'Se ha cambiado la impresora a {printer_name}')
        except Exception as e:
            print(f'Error al cambiar la impresora: {e}')

    for i in range(file_list.size()):
        file = file_list.get(i)
        try:
            path = file.replace("/", "\\")
            command = f'start msedge "{path}"'
            os.system(command)
            time.sleep(2)

            pyautogui.hotkey('ctrl', 'shift', 'p')
            time.sleep(2)

            pyautogui.press('tab', presses=6, interval=0.1)
            pyautogui.press('enter')
            pyautogui.write(printer_name) 
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('ctrl', 'w')
        except Exception as e:
            print(f'Error opening file {file}: {e}')

        progress['value'] = i + 1
        root.update_idletasks() # Update GUI

    progress['value'] = 0 # Reset progressbar
    
    win32print.SetDefaultPrinter(default_printer)
    print(f'Regresando a la impresora predefinida: {default_printer}')
    file_list.delete(0, tkinter.END)

def remove_selected_files():
    selected_files = file_list.curselection()
    for index in reversed(selected_files):
        file_list.delete(index)

def update_selected_count(event=None):
    selected_files = file_list.curselection()
    number_files_selected = len(selected_files)
    remove_button.config(text=f'Borrar ({number_files_selected}) seleccionados')

def get_printers():
    printers = []
    for printer in win32print.EnumPrinters(2):
        printer_name = printer[2]
        printers.append(printer_name)
    return printers

def update_printer_list():
    printers = get_printers()
    printer_combobox['value'] = printers
    if printers:
        printer_combobox.current(0)

def get_select_printer():
    return printer_combobox.get()

# Window configuration
root = TkinterDnD.Tk()
root.title('Asistente de impresión en masa')
root.geometry('600x450')

# Frame to select printer
frame_printers = tkinter.Frame(root)
frame_printers.pack(pady=5)
tkinter.Label(frame_printers, text='Selecciona una impresora').pack(side=tkinter.LEFT, padx=5)
printer_combobox = ttk.Combobox(frame_printers, state='readonly', width=40)
printer_combobox.pack(side=tkinter.RIGHT, padx=5)
update_printer_list()

# Button container
frame_buttons = tkinter.Frame(root)
frame_buttons.pack(pady=10)

select_button = tkinter.Button(frame_buttons, text="Seleccionar Archivos", command=select_files, width=20)
select_button.pack(side=tkinter.LEFT, padx=10)

open_button = tkinter.Button(frame_buttons, text="Imprimir Archivos", command=open_files, width=20)
open_button.pack(side=tkinter.RIGHT, padx=10)

remove_button = tkinter.Button(frame_buttons, text="Borrar (0) seleccionados", command=remove_selected_files, width=20)
remove_button.pack(side=tkinter.LEFT, padx=10)

# Config progressbar
progress = ttk.Progressbar(root, length=400, mode='determinate')
progress.pack(pady=10)

# List where the files are displayed
file_list = tkinter.Listbox(root, width=80, height=15, selectmode=tkinter.MULTIPLE)
file_list.pack(expand=True, fill="both", padx=10, pady=10)

file_list.bind('<<ListboxSelect>>', update_selected_count)

# Enable drag files
file_list.drop_target_register(DND_FILES)
file_list.dnd_bind("<<Drop>>", drag_files)

root.mainloop()