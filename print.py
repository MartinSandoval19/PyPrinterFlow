import os
import tkinter
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui
import time

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
    for i in range(file_list.size()):
        file = file_list.get(i)
        try:
            path = file.replace("/", "\\")
            command = f'start msedge "{path}"'
            os.system(command)
            time.sleep(1)

            pyautogui.hotkey('ctrl', 'p')
            time.sleep(2)

            pyautogui.press('enter')
            time.sleep(2)
        except Exception as e:
            print(f'Error opening file {file}: {e}')
    
    file_list.delete(0, tkinter.END)

def remove_selected_files():
    selected_files = file_list.curselection()
    for index in reversed(selected_files):
        file_list.delete(index)

def update_selected_count(event=None):
    selected_files = file_list.curselection()
    number_files_selected = len(selected_files)
    remove_button.config(text=f'Borrar ({number_files_selected}) seleccionados')

# Window configuration
root = TkinterDnD.Tk()
root.title('Asistente de impresión en masa')
root.geometry('600x450')

# Button container
frame_buttons = tkinter.Frame(root)
frame_buttons.pack(pady=10)

select_button = tkinter.Button(frame_buttons, text="Seleccionar Archivos", command=select_files, width=20)
select_button.pack(side=tkinter.LEFT, padx=10)

open_button = tkinter.Button(frame_buttons, text="Imprimir Archivos", command=open_files, width=20)
open_button.pack(side=tkinter.RIGHT, padx=10)

remove_button = tkinter.Button(frame_buttons, text="Borrar (0) seleccionados", command=remove_selected_files, width=20)
remove_button.pack(side=tkinter.LEFT, padx=10)

# List where the files are displayed
file_list = tkinter.Listbox(root, width=80, height=15, selectmode=tkinter.MULTIPLE)
file_list.pack(expand=True, fill="both", padx=10, pady=10)

file_list.bind('<<ListboxSelect>>', update_selected_count)

# Enable drag files
file_list.drop_target_register(DND_FILES)
file_list.dnd_bind("<<Drop>>", drag_files)

root.mainloop()
