from datetime import datetime
from msilib.schema import Error
import os
import re
import tkinter as tk
from tkinter import DISABLED, NORMAL, Label, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo, showerror
import pandas as pd
from screeninfo import get_monitors
import categories
import sys
from more_itertools import pairwise
from sys import platform


root = tk.Tk()

if platform == "darwin":
    userName = os.environ.get('USER')
elif platform == "win32":
    userName = os.getenv('username')

userName = os.getenv('username')


def resource_path(relative_path):
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


WIN_WIDTH = 542
WIN_HEIGHT = 309

root.geometry(
    f"{WIN_WIDTH}x{WIN_HEIGHT}+{(get_monitors()[0].width - WIN_WIDTH)//2}+{(get_monitors()[0].height - WIN_HEIGHT)//2}")
root.title('Papa Cleaner')
root.resizable(False, False)
root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

path = resource_path("papaj.png")
bg = tk.PhotoImage(file=path)
frame = tk.Frame(root)
frame.grid(row=1, column=0, sticky='nsew')
label1 = tk.Label(frame, image=bg)
label1.place(x=0, y=0)

for i in range(4):
    frame.columnconfigure(i, weight=1)

frame.rowconfigure(1, weight=1)

myLabel = Label(root, text="filepath:")
myLabel2 = Label(root, text="Status: Idle")


def select_file():
    filepath = fd.askopenfilename(
        title='Open file',
        initialdir='/',
        filetypes=[("Excel files", ".xlsx .xls")]
    )

    root.focus_set()

    if filepath:
        global userFile
        myLabel.config(text=(f"filepath: {filepath}"))
        userFile = filepath
        run_button['state'] = NORMAL

    if pb['value'] != 0:
        updateStatusBar(0)

    if myLabel2['text'] != "Status: Idle":
        updateStatus("Status: Idle")


dataList = []


def processData():
    data = pd.read_excel(userFile)
    data.columns = map(str.lower, data.columns)
    dataList = [asin[1].to_dict() for asin in data.iterrows()]
    updateStatus("Status: Processing file...")
    updateStatusBar(len(dataList))
    cleanData(dataList)
    buildTitle(dataList)


open_button = ttk.Button(
    frame,
    text='Open File',
    command=select_file
)

run_button = ttk.Button(
    frame,
    text='Create Titles',
    command=processData,
    state=DISABLED
)

s = ttk.Style()
s.theme_use('clam')
s.configure("red.Horizontal.TProgressbar",
            foreground='yellow', background='yellow')

pb = ttk.Progressbar(
    frame,
    orient='horizontal',
    mode='determinate',
    length=300,
    style="red.Horizontal.TProgressbar"

)

pb.grid(column=1, row=1, columnspan=2)

open_button.grid(row=2, column=1, pady=80)
run_button.grid(row=2, column=2, pady=80)
myLabel.grid(row=3, column=0, columnspan=3, sticky=tk.W,)
myLabel2.grid(row=0, column=0, columnspan=3, sticky=tk.N,)

titleList = []
cleanTitleList = []


def cleanData(list):
    for dic in list:
        for k, v in dic.items():
            if isinstance(v, float):
                v = '{0:g}'.format(v).strip().replace('nan', '')
            dic[k] = str(v).strip().replace('nan', '')


def lowerCaseSubStr(string):
    stringToList = string.split()
    subStrToLower = ['Bez', 'Dla', 'Do', 'I', 'Jak' 'Lub', 'Na',
                     'Nad', 'O', 'Od', 'Po', 'Pod', 'Przed', 'Się', 'W', 'Z', 'Ze']
    result = [
        substr.lower() if substr in subStrToLower else substr for substr in stringToList]
    return " ".join(result)


def updateStatus(txt):
    myLabel2.config(text=txt)


def updateStatusBar(num):
    if num < 1:
        pb['value'] = 0
        return
    for idx in range(num):
        amount = (idx + 1) / num
        rounded = int(round(amount * 100))
    for current, next in pairwise(range(rounded + 1)):
        if current == next:
            return
        pb['value'] = next
        root.after(1, root.update())


def buildTitle(list):
    def group1(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {department} {modelName} {itemName}, {color}, {size}"})

    def group11(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {modelNum} {itemName}, {color}, {size}"})

    def group14(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelNum} {itemName}, {color}, {size}"})

    def group2(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {department} {modelName} {modelNum} {itemName}, {color}, {size}"})

    def group3(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} sub_brand_name_if_any {itemName}, {color}, {size}"})

    def group5(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {color}, {size}"})

    def group6(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {color}, {size}, {wattage} {voltage}"})

    def group8(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {flavour}, {size}"})

    def group9(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {material}, {color}, {size}"})

    def group10(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {size}"})

    def group15(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {partNum} {itemName}, {color}, {size}"})

    def group7(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {color}, {wattage} {voltage}"})

    def group12(*args):
        titleList.append(
            {'asin': asin, 'title': f"{brand} {modelName} {modelNum} {itemName}, processor, memory(ram), storage(dysk), graphic card, operating system, {color}, {size}"})

    try:
        for dic in list:
            asin = dic['asin']
            brand = dic['brand.value'].title()
            color = dic['color.value'].title()
            department = dic['department.value'].title()
            GLname = dic['gl_product_group_type.value']
            itemName = dic['item_type_name.value'].title()
            flavour = dic['flavour.value'].title()
            material = dic['material.value'].title()
            modelName = dic['model_name.value'].title()
            modelNum = dic['model_number.value']
            partNum = dic['part_number.value']
            prodType = dic['product_type.value']
            size = dic['size.value']
            wattage = dic['wattage.value']
            voltage = dic['voltage.value']

            if wattage != "":
                wattage = wattage + " W"
            if voltage != "":
                voltage = voltage + " V"

            itemName = lowerCaseSubStr(itemName)

            if GLname in categories.gl_group1:
                group1(asin, brand, department,
                       modelName, itemName, color, size)

            elif GLname in categories.gl_group14:
                group14(asin, brand, modelNum, itemName, color, size)

            elif GLname in categories.gl_group2:
                group2(asin, brand, department, modelName,
                       modelNum, itemName, size, color)

            elif GLname in categories.gl_group3:
                group3(asin, brand, modelName, itemName, color, size)

            elif GLname in categories.gl_group6:
                group6(asin, brand, modelName, itemName,
                       color, size, wattage, voltage)

            elif GLname in categories.gl_group8:
                group8(asin, brand, modelName, itemName, flavour, size)

            elif GLname in categories.gl_group9:
                group9(asin, brand, modelName,
                       itemName, material, color, size)

            elif GLname in categories.gl_group10 and prodType in categories.product_group5:
                group5(asin, brand, modelName, itemName, color, size)

            elif GLname in categories.gl_group10:
                group10(asin, brand, modelName, itemName, size)

            elif GLname in categories.gl_group15 and prodType in categories.product_group1:
                group1(asin, brand, department,
                       modelName, itemName, color, size)

            elif GLname in categories.gl_group15:
                group15(asin, brand, partNum, itemName, color, size)

            elif GLname in categories.gl_group5 and prodType in categories.product_group11:
                group11(asin, brand, modelName,
                        modelNum, itemName, color, size)

            elif GLname in categories.gl_group5 and prodType in categories.product_group7:
                group7(asin, brand, modelName, itemName,
                       color, wattage, voltage)

            elif GLname in categories.gl_group5 and prodType in categories.product_group9:
                group9(asin, brand, modelName,
                       itemName, material, color, size)

            elif GLname in categories.gl_group5:
                group5(asin, brand, modelName, itemName, color, size)

            elif GLname in categories.gl_group11 and prodType in categories.product_group12:
                group12(asin, brand, modelName,
                        modelNum, itemName, color, size)

            elif GLname in categories.gl_group11:
                group11(asin, brand, modelName,
                        modelNum, itemName, color, size)

        cleanMissingData(titleList)
        buildFiles()
    except KeyError as colName:
        showerror(message=f"Error! Missing column: {colName}")
        updateStatus("Status: Idle")
        updateStatusBar(0)
        root.focus_set()


def cleanMissingData(list):
    for item in list:
        title = str(item['title'])
        # replace multiple spaces/white chars with single space
        title = re.sub('\s\s+', ' ', title)
        # replace more than one comma with comma space and trim
        title = re.sub('[, ]{2,}', ', ', title).strip()
        asin = item['asin']
        cleanTitleList.append({'asin': asin, 'item_name.value': title.rstrip(',')
                               if title.endswith(',') else title, 'sc_vendor_name': 'AmazonPl/NM5V9', 'login': userName})


def buildFiles():
    try:
        pd.io.formats.excel.ExcelFormatter.header_style = None
        outPath = createDirectory(userFile)
        now = datetime.now()
        currentDate = now.strftime("%m%d%Y")
        fileName = f'FLEX_TCU {currentDate}_{userName}'
        output = pd.DataFrame(cleanTitleList)
        output.to_excel(f"{outPath}/{fileName}_qa.xlsx", index=False)
        output = output.loc[:, :'sc_vendor_name']
        filepath = f"{outPath}/{fileName}.xlsx"
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
        output.to_excel(writer, sheet_name='Sheet1',
                        index=False, startrow=1)
        worksheet = writer.sheets['Sheet1']
        worksheet.write_string(0, 0, 'version=1.0.0')
        writer.save()
        showinfo(message=f"Files created successfully in: {outPath}")
        updateStatus("Status: Done!")
        updateStatusBar(0)
        root.focus_set()
    except:
        showerror(message="Error writing file!")
        updateStatus("Status: Idle")
        updateStatusBar(0)
        root.focus_set()


def createDirectory(filepath):
    directory = filepath.split("/")
    directory.pop()
    return "/".join(directory)


root.mainloop()
