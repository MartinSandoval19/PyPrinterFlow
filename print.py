import os
import tkinter
from tkinter import ttk, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import win32api
import win32print

# Config main panel
root = TkinterDnD.Tk()
root.title('Asistente de impresión')
root.geometry('600x400')
root.config(bg='#2E2E2E')

# Funciones
def get_all_printers():
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
    return printers

def get_selected_printer():
    selected_printer = printer_combobox.get()
    return selected_printer

def handle_dropped_files(event):
    files = root.tk.splitlist(event.data)
    for file in files:
        file_list.insert(tkinter.END, file)

def remove_selected_files():
    selected_items = file_list.curselection()
    for index in reversed(selected_items):
        file_list.delete(index)

sumatra_pdf = 'SumatraPDF\\SumatraPDF-3.5.2-64.exe'
def print_files():
    files = file_list.get(0, tkinter.END)
    selected_printer = get_selected_printer()

    if not files:
        messagebox.showwarning("Advertencia", "No se han seleccionado archivos")
        return
        return
    
    print_settings = "duplex" if duplex_var.get == 'Una Cara' else 'simplex'
    
    for file in files:
        if os.path.exists(file):
            command = f'-print-to "{selected_printer}" -print-settings {print_settings} -silent "{file}"'
            win32api.ShellExecute(0, 'open', sumatra_pdf, command, '.', 0)
        else:
            messagebox.showerror("Error", f"El archivo {file} no existe")
    file_list.delete(0, tkinter.END)
    messagebox.showinfo("Impresión completada", "Todas las impresiones se han enviado.")

# widgets style
style = ttk.Style()
style.theme_use('clam')

# Config buttons
style.configure('TButton', background='#444', foreground='white', font=('Arial', 12), padding=5)
style.map('TButton', background=[('active', '#666')])  # Hover

# Top
frame_top = tkinter.Frame(root, bg='#333', padx=10, pady=10)
frame_top.pack(fill='x')
label_printer = tkinter.Label(frame_top, text='Configuración de impresión', fg='white', bg='#333')
label_printer.pack(side='left', padx=5)
duplex_options = ['Una Cara', 'Doble Cara']
duplex_var = tkinter.StringVar(value=duplex_options[0])
duplex_combobox = ttk.Combobox(frame_top, state='readonly', values=duplex_options, textvariable=duplex_var, width=15)
duplex_combobox.pack(side='right', padx=10)
printers = get_all_printers()
printer_combobox = ttk.Combobox(frame_top, state='readonly', values=printers, width=40)
printer_combobox.pack(side='right', padx=5)

if printers:
    printer_combobox.current(0)

# Center
frame_center = tkinter.Frame(root, bg='#2E2E2E', padx=10, pady=10)
frame_center.pack(expand=True, fill='both')
file_list = tkinter.Listbox(frame_center, bg='#444', fg='white', selectbackground='#00A3E0', selectforeground='white', selectmode=tkinter.MULTIPLE)
file_list.pack(expand=True, fill='both', padx=10, pady=10)
file_list.drop_target_register(DND_FILES)
file_list.dnd_bind('<<Drop>>', handle_dropped_files)

# Bottom
frame_bottom = tkinter.Frame(root, bg='#333', padx=10, pady=10)
frame_bottom.pack(fill='x')
btn_print = ttk.Button(frame_bottom, text='Imprimir', command=print_files)
btn_print.pack(side='right', padx=10)
btn_remove = ttk.Button(frame_bottom, text='Eliminar Seleccionados', command=remove_selected_files)
btn_remove.pack(side='left', padx=10)

root.mainloop()