import os
import io
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image
import fitz  # PyMuPDF
import win32print
import subprocess

class PrintAssistantApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # Global configuration
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("Asistente de Impresión")
        self.geometry("1000x700")
        self.minsize(900, 600)
        self.configure(bg="#1E1E1E")

        # Grid configuration
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Files and selectors lists
        self.pdf_files = []
        self.page_selectors = []

        # Build UI
        self._build_header()
        self._build_center()
        self._build_footer()

    # ----------- Print methods -----------
    def _get_all_printers(self):
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
            return printers if printers else ["Sin impresoras"]
        except Exception:
            return ["Sin impresoras"]

    # ----------- Build UI methods -----------
    def _build_header(self):
        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#2D2D2D", height=80)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(3, weight=1)

        lbl_config = ctk.CTkLabel(header, text="Configuración de Impresión", font=("Arial", 18))
        lbl_config.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        self.printer_cb = ctk.CTkComboBox(header, values=self._get_all_printers(), width=200)
        self.printer_cb.grid(row=0, column=1, padx=10, pady=20)

        self.duplex_cb = ctk.CTkComboBox(header, values=["Una cara", "Doble cara"])
        self.duplex_cb.set("Una cara")
        self.duplex_cb.grid(row=0, column=2, padx=10, pady=20)

        lbl_sign = ctk.CTkLabel(header, text="Firmas:", font=("Arial", 14))
        lbl_sign.grid(row=0, column=3, padx=(10, 0), pady=20, sticky="w")

        self.global_page_selector = ctk.CTkComboBox(
            header,
            values=["Mantener Certificados", "Quitar certificados"],
            width=200,
            command=self._apply_to_all
        )
        self.global_page_selector.set("Mantener Certificados")
        self.global_page_selector.grid(row=0, column=3, padx=10, pady=20)

    def _build_center(self):
        center = ctk.CTkFrame(self, corner_radius=0, fg_color="#1E1E1E")
        center.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

        self.files_frame = ctk.CTkScrollableFrame(center, label_text="Arrastra aquí tus PDFs")
        self.files_frame.pack(expand=True, fill="both", padx=20, pady=20)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._on_drop_files)

        self.files_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="col")

    def _build_footer(self):
        footer = ctk.CTkFrame(self, corner_radius=0, fg_color="#2D2D2D", height=60)
        footer.grid(row=2, column=0, sticky="ew")

        btn_delete_all = ctk.CTkButton(footer, text="Borrar Todo", fg_color="#A33", command=self._delete_all)
        btn_delete_all.pack(side="left", padx=20, pady=10)

        btn_print = ctk.CTkButton(footer, text="Imprimir", fg_color="#00A36C", command=self._print_files)
        btn_print.pack(side="right", padx=20, pady=10)

    # ----------- Methods for Managing PDF Files -----------
    # Drag and Drop Handler
    def _on_drop_files(self, event):
        files = self.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith(".pdf") and file not in self.pdf_files:
                self.pdf_files.append(file)
                self._add_file_card(file)

    def _add_file_card(self, file_path):
        card = ctk.CTkFrame(self.files_frame, fg_color="#333333", corner_radius=10, width=250, height=300)

        thumb_img = self._generate_thumbnail(file_path, size=(180, 220))
        thumb_label = ctk.CTkLabel(card, image=thumb_img, text="")
        thumb_label.image = thumb_img
        thumb_label.pack(pady=10)

        filename = os.path.basename(file_path)
        lbl_file = ctk.CTkLabel(card, text=filename, font=("Arial", 14), wraplength=200)
        lbl_file.pack(pady=5)

        options = ["Mantener Certificados", "Quitar certificados"]
        page_selector = ctk.CTkOptionMenu(
            card,
            values=options,
            fg_color="#444444",
            button_color="#555555",
            button_hover_color="#666666",
            width=200
        )
        page_selector.set("Mantener Certificados")
        page_selector.pack(pady=5)

        self.page_selectors.append(page_selector)

        btn_delete = ctk.CTkButton(card, text="Eliminar", fg_color="#A33", command=lambda sel=page_selector: self._remove_file(file_path, card, sel))
        btn_delete.pack(pady=5)

        self._reorder_files()

    def _remove_file(self, file_path, card, selector=None):
        if file_path in self.pdf_files:
            self.pdf_files.remove(file_path)
        if selector and selector in self.page_selectors:
            self.page_selectors.remove(selector)
        card.destroy()
        self._reorder_files()

    def _delete_all(self):
        for child in self.files_frame.winfo_children():
            child.destroy()
        self.pdf_files.clear()
        self.page_selectors.clear()

    def _reorder_files(self):
        for idx, child in enumerate(self.files_frame.winfo_children()):
            row = idx // 3
            col = idx % 3
            child.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

    # ----------- Page and certificate options -----------
    def _apply_to_all(self, value):
        alive_selectors = []
        for selector in self.page_selectors:
            if str(selector) in self.tk.call("winfo", "children", str(selector.master)):
                try:
                    selector.set(value)
                    alive_selectors.append(selector)
                except Exception:
                    pass
        self.page_selectors = alive_selectors

    # ----------- PDF and Thumbnail Utilities -----------
    def _generate_thumbnail(self, file_path, size=(180, 220)):
        try:
            doc = fitz.open(file_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Zoom x2
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            img.thumbnail(size)
            return ctk.CTkImage(light_image=img, dark_image=img, size=size)
        except Exception as e:
            print(f"Error generando miniatura: {e}")
            placeholder = Image.new("RGB", size, (85, 85, 85))
            return ctk.CTkImage(light_image=placeholder, dark_image=placeholder, size=size)

    # ----------- Print Files -----------
    def _print_files(self):
        if not self.pdf_files:
            CTkMessagebox(title="Advertencia", message="No se han agregado archivos")
            return

        selected_printer = self.printer_cb.get()
        duplex_option = "duplex" if self.duplex_cb.get() == "Doble cara" else "simplex"
        sumatra_pdf = 'SumatraPDF\\SumatraPDF-3.5.2-64.exe'

        for idx, file_path in enumerate(self.pdf_files):
            if not os.path.exists(file_path):
                CTkMessagebox(title="Error", message=f"El archivo {file_path} no existe", icon="cancel")
                continue

            option = self.page_selectors[idx].get()
            temp_file = file_path

            if option == "Quitar certificados":
                try:
                    doc = fitz.open(file_path)
                    total_pages = doc.page_count
                    if total_pages > 2:
                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=0, to_page=total_pages - 3)
                        temp_file = os.path.join(os.getenv('TEMP'), os.path.basename(file_path))
                        new_doc.save(temp_file, garbage=4, deflate=True)
                        new_doc.close()
                    else:
                        temp_file = file_path
                    doc.close()
                except Exception as e:
                    print(f"Error quitando certificados: {e}")
                    continue

            command = [
                sumatra_pdf,
                "-print-to", selected_printer,
                "-print-settings", duplex_option,
                "-silent",
                temp_file
            ]
            subprocess.Popen(command)

        self._delete_all()
        CTkMessagebox(title="Impresión completada", message="Todas las impresiones se han enviado", icon="check")

def main():
    app = PrintAssistantApp()
    app.mainloop()

if __name__ == "__main__":
    main()