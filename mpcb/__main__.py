import tkinter as tk
import tkinter.filedialog as fd
import openpyxl
import re
import os

class MainWindow:
    def __init__(self, root):
        self.root = root
        root.title("Company Data Extractor")
        root.minsize(600, 400)

        self.testfile_label = tk.Label(root, text="Test File Location:")
        self.testfile_textbox = tk.Entry(root)
        self.browse_testfile_button = tk.Button(root, text="Browse", command=self.browse_testfile)

        self.company_name_label = tk.Label(root, text="Company Name:")
        self.company_name_textbox = tk.Entry(root)

        self.generate_button = tk.Button(root, text="Generate", command=self.generate_file)

        self.testfile_label.grid(row=0, column=0)
        self.testfile_textbox.grid(row=0, column=1)
        self.browse_testfile_button.grid(row=0, column=2)

        self.company_name_label.grid(row=1, column=0)
        self.company_name_textbox.grid(row=1, column=1)
        self.generate_button.grid(row=2, column=0, columnspan=2)

    def browse_testfile(self):
        filename = fd.askopenfilename(initialdir=".", title="Select Test File", filetypes=(("Excel Files", "*.xlsx"),))

        if filename:
            self.testfile_textbox.insert(0, filename)

    def generate_file(self):
        testfile_location = self.testfile_textbox.get()
        company_name = self.company_name_textbox.get()

        if not testfile_location:
            tk.messagebox.showwarning("Error", "Please select a test file.")
            return

        new_wb = openpyxl.Workbook()
        new_sheet = new_wb.active

        workbook = openpyxl.load_workbook(os.path.realpath(testfile_location))

        def write_to_file(row, name):
            data = []
            for cell in row:
                if cell.value:
                    data.append(cell.value)
            data.append(name)
            new_sheet.append(data)

        for sheet in workbook.worksheets:
            for i, cell in enumerate(sheet["B"]):
                val = cell.value
                try:
                    if val and re.search(company_name, val, re.IGNORECASE):
                        write_to_file(sheet[i+1], sheet.title)
                except Exception:
                    continue

        self.new_wb = new_wb
        self.company_name = company_name

        self.save_file()

    def save_file(self):
        filename = fd.asksaveasfilename(initialdir=".", title="Save File", filetypes=(("Excel Files", "*.xlsx"),))

        if filename:
            self.new_wb.save(filename=filename)
            tk.messagebox.showinfo("Success", "File saved successfully.")

def main():
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()

