class BuildTitle:
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
