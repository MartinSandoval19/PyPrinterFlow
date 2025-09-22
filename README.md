# PyPrinterFlow  
**PyPrinterFlow** is a lightweight and modern assistant designed to streamline the batch printing of PDF files, with preview thumbails, smart certificate page removal, and flexible printing options. 

## 🚀 Features  
- 📌 **Modern and intuitive graphical interface** built with CustomTkinter   
- 🔄 **Drag and drop** support for quickly adding files
- 🖼️ **PDF thumbnails preview** for visual identification
- ❌ **Remove individual files** or **clear all** before printing  
- 🖨️ **Printer selection** with duplex (double-sided) support  
- 📄 **Certificate management:** keep or remove certificate pages automatically (last two pages)
- 📦 **Portable standalone executable** (no installation required)

## 📦 Installation & Usage  
### 🖥️ Requirements  
- Python **3.12.9+**  
- **customtkinter**
- **CTkMessagebox**  
- **tkinterdnd2**  
- **pywin32** (win32api, win32print)
- **pillow**
- **PyMuPDF** (fitz)

### ▶️ Running the Application  
1. Download the executable (`.exe`) from the [Releases](https://github.com/MartinSandoval19/PyPrinterFlow/releases) section.  
2. Extract the files if necessary.  
3. Run the program.

### ▶️ Running from Source  
To run PyPrinterFlow from source, ensure you have all dependencies installed:  
```bash
pip install customtkinter CTkMessagebox tkinterdnd2 pywin32 pillow PyMuPDF
```

Then, execute the script:
```bash
python src/print.py
```

### 📦 Generating an Executable (Standalone Installer)
To create a standalone executable, use PyInstaller:

Install PyInstaller if you haven't already:
```bash
pip install pyinstaller
```

Run the following command to generate a .exe file
```bash
pyinstaller --onefile --noconsole \
--add-data "path/to/tkinterdnd2;tkinterdnd2" \
--add-data "path/to/SumatraPDF.exe;SumatraPDF" \
--hidden-import=tkinter \
--hidden-import=tkinter.ttk \
--hidden-import=tkinterdnd2 \
--hidden-import=win32api \
--hidden-import=win32print \
--collect-all tkinterdnd2 \
src/print.py
```

- Replace "path/to/tkinterdnd2" with the correct path where tkinterdnd2 is installed.
- Replace "path/to/SumatraPDF.exe" with the actual path of SumatraPDF on your system.
