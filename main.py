### Kavan Heppenstall 21/01/2025
### Product labelling software for excel spreadsheets
# Use "pyinstaller --onefile --add-data "Helvetica.ttf;." --add-data "FRE3OF9X.ttf;." main.py" to build

import os
import tkinter as tk
from tkinter import filedialog, messagebox
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

def create_label(product_name, product_price, product_SKU, product_barcode):
    curr_font = 28

    pdf = fpdf.FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_margins(3, 2.5, 3)
    pdf.b_margin = 2.5
    pdf.add_font('barcode', style='', fname=barcode_path)
    pdf.add_font('helvet', style='', fname=helvet_path)
    pdf.set_font('helvet', size=curr_font)

    page_width = pdf.w - pdf.l_margin - pdf.r_margin

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

    curr_font = 156
    pdf.set_font('barcode', size=curr_font)
    barcode_width = pdf.get_string_width(product_barcode)

    while(barcode_width > page_width):
        curr_font -= 24
        pdf.set_font('barcode', size=curr_font)

        barcode_width = pdf.get_string_width(product_barcode)

    pdf.set_y(90)
    pdf.cell(page_width, 9, product_barcode, align="C")

    curr_font = 56
    pdf.set_font('helvet', size=curr_font)
    SKU_width = pdf.get_string_width(product_SKU)

    while(SKU_width > page_width):
        curr_font -= 1
        pdf.set_font('helvet', size=56)
        
        SKU_width = pdf.get_string_width(product_SKU)

    pdf.set_y(130)
    pdf.cell(page_width, 9, product_SKU, align="C")

    #product_SKU = product_SKU.replace("/", "_").replace("\\", "_")

    os.makedirs("labels", exist_ok=True)
    pdf.output(f"labels/{product_SKU}.pdf")

def CreateLabels(event=None):
    global bookName
    if not bookName:
        messagebox.showwarning("No file", "Please upload a spreadsheet first")
        return
    
    try:
        wb = xw.Book(bookName) 
        sheet = wb.sheets[0]

        row = 3
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
                f"*{product_barcode}*" if product_barcode else ""
            )
            row += 1
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

    messagebox.showinfo("Success", f"Labels successfully created: see labels folder")
    root.destroy()

root = tk.Tk()
root.title("Product Labelling")

upload_button = tk.Button(root, text="Upload Spreadsheet", command=UploadAction)
upload_button.pack(pady=10, padx=10)

create_button = tk.Button(root, text="Create Labels", command=CreateLabels)
create_button.pack(pady=10, padx=10)

root.mainloop()

