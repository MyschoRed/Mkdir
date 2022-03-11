import pdfplumber, os, re, xlsxwriter, shutil
from tkinter import *
import subprocess

################ Main frame ################
root = Tk()
root.title("Creator 2000 :)")
root.geometry("430x300+1100+400")
root.positionfrom()
# Witgets
order_label = Label(text="Zadaj číslo objednávky:")
order_label.place(x=20, y=20)

order_input = Text(root, width=10, height=1)
order_input.place(x=170, y=22)
alert = "!!!!!!! Skontroluj, ocisti a uloz vygenerovany TXT subor. !!!!!!!"
# Return order number
def getTextInput():
    result = order_input.get(0.0, "end").rstrip()
    info_label.config(text=f'PO {result}')
    pdf_file = f'data/PO {result}.pdf'
################## PDF to TXT ##################
# open and read PDF
    text_array = []
    with pdfplumber.open(pdf_file) as pdf:
        pages = pdf.pages
        for page in pages:
            text_array.append(page.extract_text())
            text_array = [el.replace('\xa0', ' ') for el in text_array]
    text_array = "".join(text_array)
#otvori a vygeneruje TXT
    with open("toto_skontroluj_a_uloz.txt", "w") as output:
        output.write(str(text_array))
#################### cakacia hlaska ######################
    alert_label = Label(root, fg="#B33030", text=alert)
    alert_label.place(x=30, y=80)

    def openTxt():
        # os.system('Notepad.exe toto_skontroluj_a_uloz.txt') # pre Windows
        subprocess.call(['open', '-a', 'TextEdit', "toto_skontroluj_a_uloz.txt"]) # pre MacOs
    open_txt_btn = Button(root, text="Otvor TXT súbor", fg="#B33030", command=openTxt)
    open_txt_btn.place(x=150, y=120)

    buttonClicked = False
    def callback():

        buttonClicked = True

        if buttonClicked:
            with open("toto_skontroluj_a_uloz.txt", "r") as cleared_txt:
                data = cleared_txt.readlines()
        text_array = []
        folder_name = []
        id_array = []
        item_array = []
        pcs_array = []
# regex
        id_re = re.compile(r'\d{1,4} ')
        item_re = re.compile(r' \d{6,7}')
        quantity_re = re.compile(r'(?<=2022 )\w+')
# nacitanie udajov z TXT
        for line in range(len(data)):
            ids = re.search(id_re, data[line])
            item = re.search(item_re, data[line])
            quantity = re.search(quantity_re, data[line])
            folder_name.append(item.group() + " " + quantity.group() + "x")
            id_array.append(ids.group())
            item_array.append(item.group())
            pcs_array.append(quantity.group())
# vytvorenie adresarov
        for i in range(len(folder_name)):
            if os.path.exists(folder_name[i]):
                shutil.rmtree(folder_name[i])
            os.mkdir(folder_name[i])
################### XLSX ####################
# vytvorenie zosita a harku
        outBOM = xlsxwriter.Workbook('bom.xlsx')
        outSheet = outBOM.add_worksheet()
# nazvy stlpcov
        outSheet.write("A1", "ID")
        outSheet.write("B1", "POLOZKA")
        outSheet.write("C1", "KS")
# zapis dat do zosita
        for i in range(len(data)):
            outSheet.write(i + 1, 0, id_array[i])
            outSheet.write(i + 1, 1, item_array[i])
            outSheet.write(i + 1, 2, pcs_array[i])
        outBOM.close()

        end_label = Label(root, fg="#A1B57D", text="HOTOVO 😜")
        end_label.place(x=180, y=200)

    confirm_btn = Button(root, text="Podtvrď zmeny a pokračuj", command=callback)
    confirm_btn.place(x=120, y=160)

order_btn = Button(root, text="Načítaj objednávku", command=getTextInput)
order_btn.place(x=250, y=17)
info_label = Label(root, fg="#B33030", text="")
info_label.place(x=180, y=50)

root.mainloop()