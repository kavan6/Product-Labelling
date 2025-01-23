### Kavan Heppenstall 21/01/2025
### Product labelling software for excel spreadsheets
# Use "pyinstaller --onefile --add-data "Helvetica.ttf;." --add-data "FRE3OF9X.ttf;." main.py" to build
# Produces just over 5 pdfs a second

import threading
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw
import fpdf as fpdf

bookName = ""

base_dir = os.path.dirname(os.path.abspath(__file__))

helvet_path = os.path.join(base_dir, "Helvetica.ttf")
barcode_path = os.path.join(base_dir, "FRE3OF9X.ttf")

def UploadAction(event=None):
    global bookName 
    bookName = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if(bookName):
        messagebox.showinfo("File Uploaded", f"Spreadsheet loaded: {bookName}")

def create_label(product_name, product_price, product_SKU, product_barcode, isShow = False):
    pdf = fpdf.FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_margins(3, 2.5, 3)
    pdf.b_margin = 2.5

    with open(os.devnull, 'w') as devnull:
        old_stderr = sys.stderr
        sys.stderr = devnull
        try:
            pdf.add_font('barcode', style='', fname=barcode_path)
            pdf.add_font('helvet', style='', fname=helvet_path)
        finally:
            sys.stderr = old_stderr

    curr_font = 28
    pdf.set_font('helvet', size=curr_font)

    page_width = pdf.w - pdf.l_margin - pdf.r_margin

    if(not isShow):
        pname_width = pdf.get_string_width(product_name) + 10
        pprice_width = pdf.get_string_width(product_price)

        while(pname_width + pprice_width > page_width):
            curr_font -= 1
            pdf.set_font('helvet', size=curr_font)

            pname_width = pdf.get_string_width(product_name) + 10
            pprice_width = pdf.get_string_width(product_price)

        pdf.set_x(pdf.l_margin)
        pdf.cell(pname_width, 85, product_name, align="L")
        pdf.set_x(pdf.w - pdf.r_margin - pprice_width)
        pdf.cell(pprice_width, 85, product_price, align="R")
    else:
        curr_font_SKU = 44
        pdf.set_font('helvet', size=curr_font_SKU)
        SKU_width = pdf.get_string_width(product_SKU) + 10
        curr_font = 72
        pdf.set_font('helvet', size=curr_font)
        pprice_width = pdf.get_string_width(product_price)

        while(SKU_width + pprice_width > page_width):
            curr_font_SKU -= 1
            pdf.set_font('helvet', size=curr_font_SKU)
            SKU_width = pdf.get_string_width(product_SKU) + 10

            curr_font -= 1
            pdf.set_font('helvet', size=curr_font)
            pprice_width = pdf.get_string_width(product_price)

        pdf.set_x(pdf.l_margin)
        pdf.set_font('helvet', size=curr_font_SKU)
        pdf.cell(SKU_width, 85, product_SKU, align="L")
        pdf.set_x(pdf.w - pdf.r_margin - pprice_width)
        pdf.set_font('helvet', size=curr_font)
        pdf.cell(pprice_width, 85, product_price, align="R")

    curr_font = 156
    pdf.set_font('barcode', size=curr_font)
    barcode_width = pdf.get_string_width(product_barcode)

    while(barcode_width > page_width):
        curr_font -= 24
        pdf.set_font('barcode', size=curr_font)

        barcode_width = pdf.get_string_width(product_barcode)

    pdf.set_y(90)
    pdf.cell(page_width, 9, product_barcode, align="C")

    if(not isShow):
        curr_font = 56
        pdf.set_font('helvet', size=curr_font)
        SKU_width = pdf.get_string_width(product_SKU)

        while(SKU_width > page_width):
            curr_font -= 1
            pdf.set_font('helvet', size=curr_font)
            
            SKU_width = pdf.get_string_width(product_SKU)

        pdf.set_y(130)
        pdf.cell(page_width, 9, product_SKU, align="C")
    else:
        curr_font = 32
        pdf.set_font('helvet', size=curr_font)
        pname_width = pdf.get_string_width(product_name)

        while(pname_width > page_width):
            curr_font -= 1
            pdf.set_font('helvet', size=curr_font)
            
            pname_width = pdf.get_string_width(product_name)

        pdf.set_y(130)
        pdf.cell(page_width, 9, product_name, align="C")

    if(not isShow):
        os.makedirs("Warehouse Labels", exist_ok=True)
        pdf.output(f"Warehouse Labels/{product_SKU}.pdf")
    else:
        os.makedirs("Show Labels", exist_ok=True)
        pdf.output(f"Show Labels/{product_SKU}.pdf")

def CreateLabels(event=None, isShow=False):
    def task():
        global bookName
        if not bookName:
            messagebox.showwarning("No file", "Please upload a spreadsheet first")
            return
        
        try:
            wb = xw.Book(bookName) 
            sheet = wb.sheets[0]

            row = 3
            loading_bar.start()
            while True:

                product_name = sheet[f"B{row}"].value
                product_price = sheet[f"C{row}"].value
                product_SKU = sheet[f"D{row}"].value
                product_barcode = sheet[f"E{row}"].value

                if not product_name:
                    break

                create_label(
                    str(product_name), 
                    f"£{'{:.2f}'.format(product_price)}" if product_price else "£0.00", 
                    str(product_SKU) if product_SKU else "", 
                    f"*{product_barcode}*" if product_barcode else "",
                    isShow=isShow
                )
                row += 1

                root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            loading_bar.stop()
            root.update_idletasks()

        messagebox.showinfo("Success", f"Labels successfully created: see labels folder")

    loading_bar.start()
    threading.Thread(target=task).start()

root = tk.Tk()
root.title("Product Labelling")

root.minsize(width=400, height=150)

root.grid_columnconfigure(0, weight=1, minsize=40)
root.grid_columnconfigure(1, weight=2, minsize=80)


upload_button = tk.Button(root, text="Upload Spreadsheet", command=UploadAction)
upload_button.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

show_labels_button = tk.Button(root, text="Create Show Labels", command=lambda: CreateLabels(isShow=True))
show_labels_button.grid(row=1, column=0, pady=10, padx=10, sticky="ew")

warehouse_labels_button = tk.Button(root, text="Create Warehouse Labels", command=lambda: CreateLabels(isShow=False))
warehouse_labels_button.grid(row=2, column=0, pady=10, padx=10, sticky="ew")

loading_bar = ttk.Progressbar(root, mode="indeterminate", length=75)
loading_bar.grid(row=1, column=1, rowspan=1, pady=10, padx=10)

root.update_idletasks()
button_width = max(upload_button.winfo_width(), show_labels_button.winfo_width(), warehouse_labels_button.winfo_width())
upload_button.config(width=button_width)
show_labels_button.config(width=button_width)
warehouse_labels_button.config(width=button_width)

root.mainloop()

