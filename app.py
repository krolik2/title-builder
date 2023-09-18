from datetime import datetime
import os
import re
import tkinter as tk
from tkinter import DISABLED, NORMAL, Label, ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo, showerror
import pandas as pd
from screeninfo import get_monitors
import sys
from more_itertools import pairwise
from sys import platform
from categories import categories


root = tk.Tk()

if platform == "darwin":
    userName = os.environ.get("USER")
elif platform == "win32":
    userName = os.getenv("username")


def resource_path(relative_path):
    try:
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


WIN_WIDTH = 542
WIN_HEIGHT = 309

root.geometry(
    f"{WIN_WIDTH}x{WIN_HEIGHT}+{(get_monitors()[0].width - WIN_WIDTH)//2}+{(get_monitors()[0].height - WIN_HEIGHT)//2}"
)
root.title("Papa Cleaner - v1.0.4.2")
root.resizable(False, False)
root.columnconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

path = resource_path("papaj.png")
bg = tk.PhotoImage(file=path)
frame = tk.Frame(root)
frame.grid(row=1, column=0, sticky="nsew")
label1 = tk.Label(frame, image=bg)
label1.place(x=0, y=0)

for i in range(4):
    frame.columnconfigure(i, weight=1)

frame.rowconfigure(1, weight=1)

myLabel = Label(root, text="filepath:")
myLabel2 = Label(root, text="Status: Idle")

filepath = ""


def select_file():
    global filepath
    filepath = fd.askopenfilename(
        title="Open file", initialdir="/", filetypes=[("Excel files", ".xlsx .xls")]
    )
    root.focus_set()

    if filepath:
        myLabel.config(text=(f"filepath: {filepath}"))
        run_button["state"] = NORMAL

    if pb["value"] != 0:
        updateStatusBar(0)

    if myLabel2["text"] != "Status: Idle":
        updateStatus("Status: Idle")


dataList = []


def processData():
    global filepath
    data = pd.read_excel(filepath, sheet_name=0, dtype="string")
    data.columns = map(str.lower, data.columns)
    required_columns = {
        "asin",
        "brand.value",
        "color.value",
        "department.value",
        "gl_product_group_type.value",
        "flavor.value",
        "material.value",
        "model_name.value",
        "model_number.value",
        "part_number.value",
        "size.value",
        "wattage.value",
        "voltage.value",
        "product_type.value",
        "item_type_name.value",
        "cpu_model.value",
        "computer_memory.value",
        "memory_storage_capacity.value",
        "hard_disk.description.value",
        "graphics_coprocessor.value",
        "operating_system.value",
        "keyboard_layout.value",
        "sub_brand.value",
    }

    def validate_template(required, template):
        if set(required) <= set(template.columns):
            template.fillna("", inplace=True)
            dataList = [asin[1].to_dict() for asin in template.iterrows()]
            updateStatus("Status: Processing file...")
            updateStatusBar(len(dataList))
            build_titles(dataList)
            filepath = ""
            run_button["state"] = DISABLED
            myLabel.config(text=(f"filepath: {filepath}"))
        else:
            missing = set(required) - set(template.columns)
            msg = ("\n").join(sorted(missing))
            filepath = ""
            run_button["state"] = DISABLED
            myLabel.config(text=(f"filepath: {filepath}"))
            root.focus_set()
            raise Exception(showerror(message=f"Missing column:\n{msg}"))

    validate_template(required_columns, data)


open_button = ttk.Button(frame, text="Open File", command=select_file)

run_button = ttk.Button(
    frame, text="Create Titles", command=processData, state=DISABLED
)

s = ttk.Style()
s.theme_use("clam")
s.configure("red.Horizontal.TProgressbar", foreground="yellow", background="yellow")

pb = ttk.Progressbar(
    frame,
    orient="horizontal",
    mode="determinate",
    length=300,
    style="red.Horizontal.TProgressbar",
)

pb.grid(column=1, row=1, columnspan=2)

open_button.grid(row=2, column=1, pady=80)
run_button.grid(row=2, column=2, pady=80)
myLabel.grid(
    row=3,
    column=0,
    columnspan=3,
    sticky=tk.W,
)
myLabel2.grid(
    row=0,
    column=0,
    columnspan=3,
    sticky=tk.N,
)


title_list = []


def updateStatus(txt):
    myLabel2.config(text=txt)


def updateStatusBar(num):
    if num < 1:
        pb["value"] = 0
        return
    for idx in range(num):
        amount = (idx + 1) / num
        rounded = int(round(amount * 100))
    for current, next in pairwise(range(rounded + 1)):
        if current == next:
            return
        pb["value"] = next
        root.after(1, root.update())


class TitleBuilder:
    def __init__(
        self,
        asin,
        brand,
        department,
        model_name,
        model_num,
        color,
        size,
        flavor,
        material,
        part_num,
        wattage,
        voltage,
        item_name,
        cpu_model,
        computer_memory,
        memory_storage_capacity,
        hard_disk,
        graphics_description,
        operating_system,
        keyboard_layout,
        sub_brand,
    ):
        self.asin = asin
        self.brand = brand.title()
        self.department = department.title()
        self.model_name = model_name.title()
        self.color = color.title()
        self.size = size
        self.model_num = model_num
        self.flavor = flavor.title()
        self.material = material.title()
        self.part_num = part_num
        self.wattage = wattage + "W" if wattage else ""
        self.voltage = voltage + "V" if voltage else ""
        self.item_name = self.lower_case_sub_str(item_name.title())
        self.cpu_model = cpu_model
        self.computer_memory = computer_memory + " RAM" if computer_memory else ""
        self.memory_storage_capacity = memory_storage_capacity
        self.hard_disk = hard_disk
        self.graphics_description = graphics_description
        self.operating_system = operating_system
        self.keyboard_layout = keyboard_layout
        self.sub_brand = sub_brand.title()

    def lower_case_sub_str(self, string):
        string_to_list = string.split()
        sub_str_to_lower = [
            "Bez",
            "Dla",
            "Do",
            "I",
            "Jak",
            "Lub",
            "Na",
            "Nad",
            "O",
            "Od",
            "Po",
            "Pod",
            "Przed",
            "Się",
            "W",
            "Z",
            "Za",
            "Ze",
        ]
        result = [
            substr.lower() if substr in sub_str_to_lower else substr
            for substr in string_to_list
        ]
        return " ".join(result)

    def create_title_dictionary(self, title):
        return {"asin": self.asin, "title": title}

    def build_title_group1(self):
        title = f"{self.brand} {self.department} {self.model_name} {self.item_name}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group2(self):
        title = f"{self.brand} {self.department} {self.model_name} {self.model_num} {self.item_name}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group3(self):
        title = f"{self.brand} {self.model_name} {self.sub_brand} {self.item_name}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group5(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group6(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.color}, {self.size}, {self.wattage} {self.voltage}"
        return self.create_title_dictionary(title)

    def build_title_group7(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.color}, {self.wattage} {self.voltage}"
        return self.create_title_dictionary(title)

    def build_title_group8(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.flavor}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group9(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.material}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group10(self):
        title = f"{self.brand} {self.model_name} {self.item_name}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group11(self):
        title = f"{self.brand} {self.model_name} {self.model_num} {self.item_name}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group12(self):
        title = f"{self.brand} {self.model_name} {self.model_num} {self.item_name}, {self.cpu_model}, {self.computer_memory}, {self.memory_storage_capacity} {self.hard_disk}, {self.graphics_description}, {self.operating_system}, {self.keyboard_layout}, {self.color}, {self.size}"
        return self.create_title_dictionary(title)

    def build_title_group14(self):
        title = (
            f"{self.brand} {self.model_num} {self.item_name}, {self.color}, {self.size}"
        )
        return self.create_title_dictionary(title)

    def build_title_group15(self):
        title = (
            f"{self.brand} {self.part_num} {self.item_name}, {self.color}, {self.size}"
        )
        return self.create_title_dictionary(title)


def build_titles(list):
    try:
        for dic in list:
            asin = dic["asin"]
            brand = dic["brand.value"]
            color = dic["color.value"]
            department = dic["department.value"]
            product_group_type = dic["gl_product_group_type.value"]
            flavor = dic["flavor.value"]
            material = dic["material.value"]
            model_name = dic["model_name.value"]
            model_num = dic["model_number.value"]
            part_num = dic["part_number.value"]
            size = dic["size.value"]
            wattage = dic["wattage.value"]
            voltage = dic["voltage.value"]
            product_type = dic["product_type.value"]
            item_name = dic["item_type_name.value"]
            cpu_model = dic["cpu_model.value"]
            computer_memory = dic["computer_memory.value"]
            memory_storage_capacity = dic["memory_storage_capacity.value"]
            hard_disk = dic["hard_disk.description.value"]
            graphics_description = dic["graphics_coprocessor.value"]
            operating_system = dic["operating_system.value"]
            keyboard_layout = dic["keyboard_layout.value"]
            sub_brand = dic["sub_brand.value"]

            title_builder = TitleBuilder(
                asin,
                brand,
                department,
                model_name,
                model_num,
                color,
                size,
                flavor,
                material,
                part_num,
                wattage,
                voltage,
                item_name,
                cpu_model,
                computer_memory,
                memory_storage_capacity,
                hard_disk,
                graphics_description,
                operating_system,
                keyboard_layout,
                sub_brand,
            )

            def assign_to_build_function(category_name, sub_category_name=None):
                function_name = "build_title_group"
                default_group = "11"
                group_id = default_group
                for category in categories:
                    if category["name"] == category_name:
                        group_id = str(category["group_id"])
                        for sub_category in category.get("sub_categories", []):
                            if sub_category_name in sub_category["name"]:
                                group_id = str(sub_category["group_id"])
                                break
                        break
                return title_list.append(
                    getattr(title_builder, function_name + group_id)()
                )

            assign_to_build_function(product_group_type, product_type)

        buildFiles()

    except Exception as e:
        showerror(message=f"Error! {e}")
        updateStatus("Status: Idle")
        updateStatusBar(0)
        root.focus_set()


def cleanMissingData(list):
    clean_title_list = []
    for item in list:
        title = str(item["title"])
        asin = item["asin"]
        # replace multiple spaces/white chars with single space
        title = re.sub("\s\s+", " ", title)
        # replace more than one comma with comma space and trim
        title = re.sub("[, ]{2,}", ", ", title).strip()
        clean_title_list.append(
            {
                "ASIN": asin,
                "item_name.value": title.rstrip(",") if title.endswith(",") else title,
                "sc_vendor_name": "AmazonPl/NM5V9",
                "login": userName,
            }
        )
    return clean_title_list


def createDirectory(path):
    directory = path.split("/")
    directory.pop()
    return "/".join(directory)


def buildFiles():
    global title_list
    try:
        pd.io.formats.excel.ExcelFormatter.header_style = None
        outPath = createDirectory(filepath)
        now = datetime.now()
        currentDate = now.strftime("%m%d%Y")
        fileName = f"FLEX_TCU {currentDate}_{userName}"
        output = pd.DataFrame(cleanMissingData(title_list))
        output.to_excel(f"{outPath}/{fileName}_qa.xlsx", index=False, sheet_name="TCU")
        output = output.loc[:, :"sc_vendor_name"]
        path = f"{outPath}/{fileName}.xlsx"
        writer = pd.ExcelWriter(path, engine="xlsxwriter")
        output.to_excel(writer, sheet_name="TCU", index=False, startrow=1)
        worksheet = writer.sheets["TCU"]
        worksheet.write_string(0, 0, "version=1.0.0")
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
    finally:
        title_list = []


root.mainloop()
