import sys
from PyQt6 import QtWidgets, QtGui
from openpyxl import load_workbook, Workbook
import re
import os

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Company Data Extractor")
        self.setMinimumSize(600, 400)

        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QtWidgets.QGridLayout()
        self.central_widget.setLayout(self.layout)

        self.company_name_label = QtWidgets.QLabel("Company Name:")
        self.company_name_textbox = QtWidgets.QLineEdit()

        self.generate_button = QtWidgets.QPushButton("Generate")
        self.generate_button.clicked.connect(self.generate_file)

        self.save_file_button = QtWidgets.QPushButton("Save File")
        self.save_file_button.clicked.connect(self.save_file)

        self.layout.addWidget(self.company_name_label, 0, 0)
        self.layout.addWidget(self.company_name_textbox, 0, 1)
        self.layout.addWidget(self.generate_button, 1, 0, 1, 2)
        self.layout.addWidget(self.save_file_button, 2, 0, 1, 2)

    def generate_file(self):
        company_name = self.company_name_textbox.text()

        new_wb = Workbook()
        new_sheet = new_wb.active

        workbook = load_workbook(os.path.realpath("files/testfile.xlsx"))

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

    def save_file(self):
        filename, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save File", f"{self.company_name}.xlsx")

        if filename:
            self.new_wb.save(filename=filename)

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()

if __name__ == "__main__":
    main()
