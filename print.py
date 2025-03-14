import os
import tkinter
from tkinter import filedialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32print
import win32api
import time


ghostscript = 'GHOSTSCRIPT\\gs10050w64.exe'
gsprint = 'GSPRINT\\gsprint.exe'

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    for file in files:
        file_list.insert(tkinter.END, file)

def drag_files(event):
    files = root.tk.splitlist(event.data)
    for file in files:
        file_list.insert(tkinter.END, file)

def get_select_printer():
    return printer_combobox.get()

def get_printers():
    return [printer[2] for printer in win32print.EnumPrinters(2)]

def update_printer_list():
    printers = get_printers()
    printer_combobox['value'] = printers
    if printers:
        printer_combobox.current(0)

def open_files():
    total_files = file_list.size()
    progress['maximum'] = total_files

    printers = get_printers()
    printer_name = get_select_printer()

    if printer_name not in printers:
        print(f"[ERROR] Impresora '{printer_name}' no encontrada en la lista de impresoras disponibles.")
        return

    for i in range(total_files):
        file = file_list.get(i)
        try:
            path = file.replace("/", "\\")
            command = f'"{gsprint}" -ghostscript "{ghostscript}" -printer "{printer_name}" "{path}"'

            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Ejecutando comando:")
            print(command)

            result = win32api.ShellExecute(0, 'open', gsprint, command, '.', 0)

            if result > 32:
                print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Impresión enviada correctamente: {file} en {printer_name}")
            else:
                print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Error al ejecutar el comando. Código de retorno: {result}")

            progress['value'] = i + 1
            root.update_idletasks()

        except Exception as e:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] ⚠️ ERROR al imprimir {file}: {e}")

    progress['value'] = 0
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 🏁 Todas las impresiones han finalizado.")

def remove_selected_files():
    selected_files = file_list.curselection()
    for index in reversed(selected_files):
        file_list.delete(index)

def update_selected_count(event=None):
    selected_files = file_list.curselection()
    number_files_selected = len(selected_files)
    remove_button.config(text=f'Borrar ({number_files_selected}) seleccionados')

root = TkinterDnD.Tk()
root.title('Asistente de impresión en masa')
root.geometry('600x450')

frame_printers = tkinter.Frame(root)
frame_printers.pack(pady=5)
tkinter.Label(frame_printers, text='Selecciona una impresora').pack(side=tkinter.LEFT, padx=5)
printer_combobox = ttk.Combobox(frame_printers, state='readonly', width=40)
printer_combobox.pack(side=tkinter.RIGHT, padx=5)
update_printer_list()

frame_buttons = tkinter.Frame(root)
frame_buttons.pack(pady=10)

select_button = tkinter.Button(frame_buttons, text="Seleccionar Archivos", command=select_files, width=20)
select_button.pack(side=tkinter.LEFT, padx=10)

open_button = tkinter.Button(frame_buttons, text="Imprimir Archivos", command=open_files, width=20)
open_button.pack(side=tkinter.RIGHT, padx=10)

remove_button = tkinter.Button(frame_buttons, text="Borrar (0) seleccionados", command=remove_selected_files, width=20)
remove_button.pack(side=tkinter.LEFT, padx=10)

progress = ttk.Progressbar(root, length=400, mode='determinate')
progress.pack(pady=10)

file_list = tkinter.Listbox(root, width=80, height=15, selectmode=tkinter.MULTIPLE)
file_list.pack(expand=True, fill="both", padx=10, pady=10)

file_list.bind('<<ListboxSelect>>', update_selected_count)

file_list.drop_target_register(DND_FILES)
file_list.dnd_bind("<<Drop>>", drag_files)

root.mainloop()
