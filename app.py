from datetime import datetime
import os
import re
import tkinter as tk
from tkinter import DISABLED, NORMAL, Label, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo, showerror
import pandas as pd
import time
from screeninfo import get_monitors
import traceback

try:
    root = tk.Tk()

    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    WIN_WIDTH = 542
    WIN_HEIGHT = 309

    root.geometry(
        f"{WIN_WIDTH}x{WIN_HEIGHT}+{(get_monitors()[0].width - WIN_WIDTH)//2}+{(get_monitors()[0].height - WIN_HEIGHT)//2}")
    root.title('Papa Cleanup')
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

        if filepath:
            global userFile
            myLabel.config(text=(f"filepath: {filepath}"))
            userFile = filepath
            run_button['state'] = NORMAL

        if pb['value'] != 0:
            pb['value'] = 0
        
        if myLabel2['text'] != "Status: Idle":
            myLabel2['text'] = "Status: Idle"

    dataList = []

    def processData():
        data = pd.read_excel(userFile)
        data.columns = map(str.lower, data.columns)
        dataList = [asin[1].to_dict() for asin in data.iterrows()]
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

    pb = ttk.Progressbar(
        frame,
        orient='horizontal',
        mode='determinate',
        length=280
    )

    pb.grid(column=1, row=1, columnspan=2)

    open_button.grid(row=2, column=1, pady=80)
    run_button.grid(row=2, column=2, pady=80)
    myLabel.grid(row=3, column=0, columnspan=3, sticky=tk.W,)
    myLabel2.grid(row=0, column=0, columnspan=3, sticky=tk.N,)

    # userName = os.environ.get('USER') - for macOs
    userName = os.getenv('username')

    titleList = []
    cleanTitleList = []

    genericGLlist = ["gl_jewelry", "gl_watch",
                     "gl_pet_products", "gl_drugstore"]

    def cleanData(list):
        for dic in list:
            for k, v in dic.items():
                if isinstance(v, float):
                    v = '{0:g}'.format(v).strip().replace('nan', '')
                dic[k] = str(v).strip().replace('nan', '')

    def lowerCaseSubStr(string):
        stringToList = string.split()
        subStrToLower = ["I", "Na", "Z", "Do", "W", "Ze"]
        result = [
            substr.lower() if substr in subStrToLower else substr for substr in stringToList]
        return " ".join(result)

    def buildTitle(list):
        def home(*args):
            titleList.append(
                {'asin': asin, 'title': f"{brand} {modelName} {itemName}, {material}, {color}, {size}"})

        def homeImprovement(*args):
            titleList.append(
                {'asin': asin, 'title': f"{brand} {modelNum} {itemName}, {wattage}, {voltage}, {color}, {size}"})

        def kitchen(*args):
            titleList.append(
                {'asin': asin, 'title': f"{brand} {modelNum} {modelName} {itemName}, {material}, {wattage}, {size}, {color}"})

        def luggage(*args):
            titleList.append(
                {'asin': asin, 'title': f"{brand} {itemName}, {color}"})

        def generic(*args):
            titleList.append(
                {'asin': asin, 'title': f"{brand} {modelNum} {itemName}, {size}, {color}"})

        try:
            for dic in list:
                myLabel2.config(text=(f"Status: Processing titles..."))
                pb['value'] += 1
                root.update()
                time.sleep(.0010)
                asin = dic['asin']
                brand = dic['brand.value']
                color = dic['color.value']
                GLname = dic['gl_product_group_type.value']
                itemName = dic['item_type_name.value']
                material = dic['material.value']
                modelName = dic['model_name.value']
                modelNum = dic['model_number.value']
                size = dic['size.value']
                wattage = dic['wattage.value']
                voltage = dic['voltage.value']

                if wattage != "":
                    wattage = wattage + " W"
                if voltage != "":
                    voltage = voltage + " V"

                modelName = modelName.title()
                itemName = itemName.title()
                itemName = lowerCaseSubStr(itemName)

                if GLname == "gl_home":
                    home(asin, brand, modelName, itemName, material, color, size)

                elif GLname == "gl_home_improvement":
                    homeImprovement(asin, brand, modelNum, itemName,
                                    wattage, voltage, color, size)

                elif GLname == "gl_kitchen":
                    kitchen(asin, brand, modelNum, modelName,
                            itemName, material, wattage, size, color)

                elif GLname == "gl_luggage":
                    luggage(asin, brand, modelName, itemName, color)

                if GLname in genericGLlist:
                    generic(asin, brand, modelNum, itemName, size, color)

            pb['value'] = 0
            cleanMissingData(titleList)
            buildFiles()

        except KeyError as err:
            showerror(message=f"Error! Missing column: {err}")
            pb['value'] = 0
            myLabel2.config(text=(f"Status: Idle"))

    def cleanMissingData(list):
        for item in list:
            myLabel2.config(text=(f"Status: Cleaning up..."))
            pb['value'] += 1
            time.sleep(.0010)
            root.update()
            title = str(item['title'])
            title = re.sub('[, ]{2,}', ', ', title).strip()
            asin = item['asin']
            cleanTitleList.append({'asin': asin, 'item_name.value': title.rstrip(',')
                                   if title.endswith(',') else title, 'sc_vendor_name': 'AmazonPl/NM5V9', 'login': userName})

        myLabel2.config(text=(f"Status: Building files..."))

    def buildFiles():
        outPath = createDirectory(userFile)
        now = datetime.now()
        currentDate = now.strftime("%m%d%Y")
        fileName = f'{userName}_{currentDate}'
        output = pd.DataFrame(cleanTitleList)
        output.to_excel(f"{outPath}/{fileName}_qa.xlsx", index=False)
        output = output.loc[:, :'sc_vendor_name']
        output.to_excel(f"{outPath}/{fileName}.xlsx", index=False)
        showinfo(message=f"Files created successfully in: {outPath}")
        myLabel2.config(text=(f"Status: Done!"))
        pb['value'] = 0

    def createDirectory(filepath):
        directory = filepath.split("/")
        directory.pop()
        return "/".join(directory)

    root.mainloop()
except:
    traceback.print_exc()
    print("Something went wrong, please send screenshot of this error to: @krolikma")
