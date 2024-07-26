import tkinter as tk
from tkinter import filedialog, Label
import os
import pandas as pd
import json


# Get the directory of the script
script_dir = os.path.dirname(os.path.realpath(__file__))

# Set the current working directory to the script's directory
os.chdir(script_dir)

class FolderAndExcelAnalyzer:
    def __init__(self, root):
        self.root = root
        self.folders = {
            "Damaged": "",
            "Inventory": "",
            "Personal": "",
            "Sold": "",
            "To Sell": ""
        }
        self.excel_file = ""
        self.sheet_name = ""

        for name in self.folders.keys():
            button = tk.Button(root, text=f"Select {name}", command=lambda n=name: self.select_folder(n))
            button.pack()
            label = Label(root, text="")
            label.pack()
            self.folders[name] = {'path': '', 'path_label': label}

        self.excel_button = tk.Button(root, text="Select Excel File", command=self.select_excel_file)
        self.excel_button.pack()
        self.excel_label = Label(root, text="")
        self.excel_label.pack()

        self.inventory_button = tk.Button(root, text="Select Inventory Excel File", command=self.select_inventory_file)
        self.inventory_button.pack()
        self.inventory_label = Label(root, text="")
        self.inventory_label.pack()

        self.analyze_button = tk.Button(root, text="Analyze Folders and Excel", command=self.analyze_folders_and_excel, state=tk.DISABLED)
        self.analyze_button.pack()
        self.load_settings()

    def save_settings(self):
        settings = {
            'folders': {name: info['path'] for name, info in self.folders.items()},
            'excel_file': self.excel_file,
            'sheet_name': self.sheet_name,
            'inventory_file': self.inventory_file,
            'inventory_sheet_name': self.inventory_sheet_name
        }
        with open('settings.json', 'w') as f:
            json.dump(settings, f)

    def load_settings(self):
        try:
            with open('settings.json', 'r') as f:
                settings = json.load(f)
                for name, path in settings['folders'].items():
                    if name in self.folders:
                        self.folders[name]['path'] = path
                        self.folders[name]['path_label'].config(text=path)
                self.excel_file = settings.get('excel_file', '')
                self.excel_label.config(text=self.excel_file)
                self.sheet_name = settings.get('sheet_name', '')

                self.inventory_file = settings.get('inventory_file', '')
                self.inventory_label.config(text=self.inventory_file)
                self.inventory_sheet_name = settings.get('inventory_sheet_name', '')

                self.check_all_selected()
        except FileNotFoundError:
            print("Settings file not found. Please select folders and Excel file.")

    def select_folder(self, name):
        selected_path = filedialog.askdirectory()
        if selected_path:
            self.folders[name]['path'] = selected_path
            self.folders[name]['path_label'].config(text=selected_path)
            print(f"Folder {name} path selected: {selected_path}")
        self.check_all_selected()

    def select_excel_file(self):
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.excel_file:
            self.excel_label.config(text=self.excel_file)
            self.show_sheet_selection_window()

    def show_sheet_selection_window(self):
        window = tk.Toplevel(self.root)
        window.title("Select a Sheet")
        xls = pd.ExcelFile(self.excel_file)
        for sheet in xls.sheet_names:
            button = tk.Button(window, text=sheet, command=lambda s=sheet: self.set_sheet_name(s, window))
            button.pack()

    def set_sheet_name(self, sheet_name, window):
        self.sheet_name = sheet_name
        self.excel_label.config(text=f"{self.excel_file} ({self.sheet_name})")
        window.destroy()
        print(f"Excel file selected: {self.excel_file}, Sheet: {self.sheet_name}")
        self.check_all_selected()

    def select_inventory_file(self):
        self.inventory_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.inventory_file:
            self.inventory_label.config(text=self.inventory_file)  # Corrected line
            self.show_inventory_sheet_selection_window()

    def show_inventory_sheet_selection_window(self):
        window = tk.Toplevel(self.root)
        window.title("Select a Sheet for Inventory")
        xls = pd.ExcelFile(self.inventory_file)
        for sheet in xls.sheet_names:
            button = tk.Button(window, text=sheet, command=lambda s=sheet: self.set_inventory_sheet_name(s, window))
            button.pack()

    def set_inventory_sheet_name(self, sheet_name, window):
        self.inventory_sheet_name = sheet_name
        window.destroy()
        self.check_all_selected()

    def check_all_selected(self):
        all_folders_selected = all(folder['path'] for folder in self.folders.values())
        all_files_selected = all([self.excel_file, self.sheet_name, self.inventory_file, self.inventory_sheet_name])
        if all_folders_selected and all_files_selected:
            self.analyze_button.config(state=tk.NORMAL)
        else:
            self.analyze_button.config(state=tk.DISABLED)

    def extract_product_ids_from_folders(self):
        product_ids = []
        for folder_info in self.folders.values():
            path = folder_info['path']
            if path:
                product_ids.extend(self.extract_product_ids(path))
        print("Folder product IDs:", sorted(product_ids))
        return product_ids

    def extract_product_ids_from_excel(self):
        if not self.excel_file or not self.sheet_name:
            return []
        df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
        product_ids = sorted(df['Product ID'].dropna().astype(str).tolist())
        print("Excel product IDs:", product_ids)
        return product_ids

    def custom_sort_key(self, id):
        return (len(id), id.upper())

    def analyze_folders_and_excel(self):
        self.save_settings()
        print("Starting analysis...")

        folder_product_ids = self.extract_product_ids_from_folders()
        excel_product_ids = self.extract_product_ids_from_excel()

        highest_folder_id = max(folder_product_ids, key=self.custom_sort_key) if folder_product_ids else 'A0'
        highest_excel_id = max(excel_product_ids, key=self.custom_sort_key) if excel_product_ids else 'A0'

        print("Highest Folder Product ID:", highest_folder_id)
        print("Highest Excel Product ID:", highest_excel_id)

        folder_sequence = self.generate_complete_sequence(highest_folder_id)
        excel_sequence = self.generate_complete_sequence(highest_excel_id)

        missing_folder_ids = []
        for id in folder_sequence:
            if id.upper() not in [pid.upper() for pid in folder_product_ids]:
                missing_folder_ids.append(id)
            elif id.upper() in [pid.upper() for pid in folder_product_ids] and id not in folder_product_ids:
                missing_folder_ids.append(f"{id} - (found in lowercase)")

        missing_excel_ids = []
        for id in excel_sequence:
            if id.upper() not in [pid.upper() for pid in excel_product_ids]:
                missing_excel_ids.append(id)
            elif id.upper() in [pid.upper() for pid in excel_product_ids] and id not in excel_product_ids:
                missing_excel_ids.append(f"{id} - (found in lowercase)")

        # Write the missing product IDs into the text file
        with open("missing_product_ids.txt", "w") as file:
            file.write("-----------------Missing Folder Product IDs-----------\n")
            file.write("\n".join(sorted(missing_folder_ids, key=lambda x: x.upper())) + "\n")
            file.write("-----------------Missing Excel Product IDs-----------\n")
            file.write("\n".join(sorted(missing_excel_ids, key=lambda x: x.upper())) + "\n")

        print("Analysis completed for folders and main Excel file.")

        missing_inventory_ids, duplicate_notes = self.analyze_inventory()

        with open("missing_product_ids.txt", "a") as file:  # Append to the existing file
            file.write("\n-----------------Missing Inventory Product IDs-----------\n")
            file.write("\n".join(sorted(missing_inventory_ids, key=lambda x: x.upper())) + "\n")
            file.write("\n-----------------Duplicate Inventory Product IDs-----------\n")
            file.write("\n".join(sorted(duplicate_notes, key=lambda x: x.upper())) + "\n")

        print("Inventory analysis completed and file updated.")

    def analyze_inventory(self):
        if not self.inventory_file or not self.inventory_sheet_name:
            return [], []
        df = pd.read_excel(self.inventory_file, sheet_name=self.inventory_sheet_name)
        inventory_product_ids = sorted(df['Product ID'].dropna().astype(str).tolist())

        print("Inventory product IDs:", inventory_product_ids)

        highest_inventory_id = max(inventory_product_ids, key=self.custom_sort_key) if inventory_product_ids else 'A0'
        print("Highest Inventory Product ID:", highest_inventory_id)

        # Identifying duplicates and their respective rack IDs
        rack_ids = df['Rack ID'].fillna('Unknown').tolist()
        duplicate_notes = [f"{id} - (Duplicate entry / product present in rack {rack_ids[i]})"
                           for i, id in enumerate(inventory_product_ids) if inventory_product_ids.count(id) > 1]

        # Finding missing IDs
        inventory_sequence = self.generate_complete_sequence(highest_inventory_id)
        missing_inventory_ids = sorted([id for id in inventory_sequence if id.upper() not in map(str.upper, inventory_product_ids)])

        return missing_inventory_ids, duplicate_notes

    def extract_product_ids(self, folder_path):
        product_ids = []
        for folder_name in os.listdir(folder_path):
            if folder_name.startswith("-"):  # skip folders starting with "-"
                continue

            folder_path_full = os.path.join(folder_path, folder_name)
            if os.path.isdir(folder_path_full):
                product_id = folder_name.split(' ')[0]
                product_ids.append(product_id)

        return product_ids

    def generate_complete_sequence(self, highest_id):
        sequence = []
        chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        alphanumeric = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        for first_char in chars:
            for second_char in alphanumeric:
                id = f"{first_char}{second_char}"
                sequence.append(id)
                if id.upper() >= highest_id.upper():
                    break  # Stop correctly based on highest ID

        print("Generated sequence up to highest ID:", sequence)
        return sequence

    def identify_missing_product_ids(self, existing_ids, highest_id):
        all_ids = self.generate_complete_sequence(highest_id)
        missing_ids = [id for id in all_ids if id.upper() not in map(str.upper, existing_ids)]
        return missing_ids


def main():
    root = tk.Tk()
    root.title("Folder and Excel Product ID Analyzer")
    app = FolderAndExcelAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
