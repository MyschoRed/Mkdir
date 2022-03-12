import pdfplumber, os, xlsxwriter, shutil, re
import subprocess
from tkinter import *
from texts import *
from settings import *
from tkinter import ttk, filedialog

################ MAIN FRAME ################
root = Tk()
root.title(window_title)
root.geometry(f"{WIDTH}x{HEIGHT}+{POS_X}+{POS_Y}")
root.resizable(0, 0) # resizable window on/off


def openPdf():
    root.filename = filedialog.askopenfilename(initialdir="./")
    opened_BOM_label = Label(text=root.filename).place(x=450, y=70)

def getTextInput():
    ############# RETURN ORDER NUMBER ###############
    result = order_input.get(0.0, "end").rstrip()
    info_label.config(text=f'PO {result}')
    pdf_file = f'PO {result}.pdf'

    ################## PDF to TXT ##################
    # open and read PDF
    text_array = []
    with pdfplumber.open(pdf_file) as pdf:
        pages = pdf.pages
        for page in pages:
            text_array.append(page.extract_text())
            text_array = [el.replace('\xa0', ' ') for el in text_array]
    text_array = "".join(text_array)

    # generate and fill TXT
    with open(generated_txt, "w") as output:
        output.write(str(text_array))

    # waiting alert
    alert_label = Label(root, fg="#B33030", text=alert_text).place(x=30, y=100)

    # open TXT
    # os.system('Notepad.exe data.txt') # for Windows
    subprocess.call(['open', '-a', 'TextEdit',
                     generated_txt])  # for MacOs

    # confirm TXT changed
    buttonClicked = False
    def callback():
        buttonClicked = True
        if buttonClicked:
            with open(generated_txt, "r") as cleared_txt:
                data = cleared_txt.readlines()

        # Arrays
        text_array = []
        folder_name = []
        id_array = []
        item_array = []
        quantity_array = []

        # read TXT
        for line in range(len(data)):
            ids = re.search(id_regex, data[line])
            item = re.search(item_regex, data[line])
            quantity = re.search(quantity_regex, data[line])
            folder_name.append(item.group() + " " + quantity.group() + "x")
            id_array.append(ids.group())
            item_array.append(item.group())
            quantity_array.append(quantity.group())

        # creating folders
        for i in range(len(folder_name)):
            if os.path.exists(folder_name[i]):
                shutil.rmtree(folder_name[i])
            os.mkdir(folder_name[i])

################### XLSX ####################
# create XLSX
        outBOM = xlsxwriter.Workbook(generated_bom)
        outSheet = outBOM.add_worksheet()

        # column names
        outSheet.write("A1", "ID")
        outSheet.write("B1", "POLOZKA")
        outSheet.write("C1", "KS")

        # write XLSX
        for i in range(len(data)):
            outSheet.write(i + 1, 0, id_array[i])
            outSheet.write(i + 1, 1, item_array[i])
            outSheet.write(i + 1, 2, quantity_array[i])
        outBOM.close()
        end_label = Label(root, text=done)
        end_label.place(x=180, y=200)

        # generate log
        with open(log_txt, "w") as output:
            output.write(log_text1)
            output.write(str(len(id_array)))
            output.write(log_text2)

    confirm_btn = Button(root,
                         text=confirm_text,
                         command=callback)
    confirm_btn.place(x=120, y=160)

#################### WIDGETS ###################
# Labels
title1_label = Label(text=title1).place(x=100, y=10)
title2_label = Label(text=title2).place(x=530, y=10)
order_label = Label(text=input_order_number).place(x=20, y=40)
info_label = Label(root, text="")
info_label.place(x=180, y=80)
sig_label = Label(root, text=signature).place(x=10, y=270)
version_label = Label(root, text=version).place(x=WIDTH - 90, y=270)

# Inputs
order_input = Text(root, width=10, height=1)
order_input.place(x=170, y=42)

# Buttons
order_btn = Button(root, text=load_order, command=getTextInput).place(x=250, y=37)
open_BOM = Button(root, text=open_bom, command=openPdf).place(x=600, y=37)

# Others
sep = ttk.Separator(root, orient=VERTICAL).place(x=WIDTH / 2, y=0, relheight=1)

root.mainloop()