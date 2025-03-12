import os
import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui
import time

def seleccionar_archivos():
    archivos = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    for archivo in archivos:
        lista_archivos.insert(tk.END, archivo)

def arrastrar_archivos(event):
    archivos = root.tk.splitlist(event.data)  # Convierte en lista
    for archivo in archivos:
        lista_archivos.insert(tk.END, archivo)

def abrir_archivos():
    for i in range(lista_archivos.size()):
        archivo = lista_archivos.get(i)
        try:
            archivo_correcto = archivo.replace("/", "\\")  
            comando = f'start msedge "{archivo_correcto}"'  
            os.system(comando)  
            time.sleep(1)  

            pyautogui.hotkey('ctrl', 'p')
            time.sleep(2)  

            pyautogui.press('enter')
            time.sleep(1)  

        except Exception as e:
            print(f"Error al abrir {archivo}: {e}")

    lista_archivos.delete(0, tk.END)

# Configuración de la ventana
root = TkinterDnD.Tk()  
root.title("Arrastra archivos para abrir")
root.geometry("600x450")  

# Crear un frame para los botones
frame_botones = tk.Frame(root)
frame_botones.pack(pady=10)

btn_seleccionar = tk.Button(frame_botones, text="Seleccionar archivos", command=seleccionar_archivos, width=20)
btn_seleccionar.pack(side=tk.LEFT, padx=10)

btn_abrir = tk.Button(frame_botones, text="Imprimir Archivos", command=abrir_archivos, width=20)
btn_abrir.pack(side=tk.RIGHT, padx=10)

# Lista donde se muestran los archivos
lista_archivos = tk.Listbox(root, width=80, height=15)
lista_archivos.pack(expand=True, fill="both", padx=10, pady=10)

# Habilitar arrastrar y soltar archivos en la lista
lista_archivos.drop_target_register(DND_FILES)
lista_archivos.dnd_bind("<<Drop>>", arrastrar_archivos)

root.mainloop()





