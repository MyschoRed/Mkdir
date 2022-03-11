import pdfplumber, os, xlsxwriter, shutil
from tkinter import *
import subprocess
from texts import *
from settings import *

################ Main frame ################
root = Tk()
root.title(window_title)
root.geometry(f"{WIDTH}x{HEIGHT}+{POS_X}+{POS_Y}")
# Witgets
order_label = Label(text=input_order_number)
order_label.place(x=20, y=20)
order_input = Text(root, width=10, height=1)
order_input.place(x=170, y=22)
alert = alert_text

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
    with open("data.txt", "w") as output:
        output.write(str(text_array))

    # waiting alert
    alert_label = Label(root, fg="#B33030", text=alert_text)
    alert_label.place(x=30, y=80)

    # open TXT
    # os.system('Notepad.exe data.txt') # for Windows
    subprocess.call(['open', '-a', 'TextEdit',
                     "data.txt"])  # for MacOs

    # confirm TXT changed
    buttonClicked = False
    def callback():
        buttonClicked = True
        if buttonClicked:
            with open("data.txt", "r") as cleared_txt:
                data = cleared_txt.readlines()
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
        outBOM = xlsxwriter.Workbook('bom.xlsx')
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

    confirm_btn = Button(root,
                         text=confirm_text,
                         command=callback)
    confirm_btn.place(x=120, y=160)

order_btn = Button(root, text=load_order, command=getTextInput)
order_btn.place(x=250, y=17)
info_label = Label(root, text="")
info_label.place(x=180, y=50)

root.mainloop()