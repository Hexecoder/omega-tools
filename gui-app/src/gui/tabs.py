from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QComboBox
import webbrowser
from utils.excel_handler import get_task_numbers, get_sheet_names

class ExcelTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout()

        # Excel dosyası seçimi
        self.label = QLabel("Select an Excel file:")
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Browse")
        self.select_button.clicked.connect(self.open_file_dialog)
        self.layout.addWidget(self.select_button)

        self.file_path_label = QLabel("")
        self.layout.addWidget(self.file_path_label)

        # Sheet seçimi
        self.sheet_dropdown = QComboBox()
        self.sheet_dropdown.setEnabled(False)  # Başlangıçta devre dışı
        self.layout.addWidget(self.sheet_dropdown)

        # İşlemi Başlat Butonu
        self.start_button = QPushButton("Start Process")
        self.start_button.setEnabled(False)  # Başlangıçta devre dışı
        self.start_button.clicked.connect(self.start_process)
        self.layout.addWidget(self.start_button)

        self.setLayout(self.layout)

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_path_label.setText(file_path)
            self.load_sheets(file_path)

    def load_sheets(self, file_path):
        try:
            sheet_names = get_sheet_names(file_path)  # Excel dosyasındaki sheet isimlerini al
            self.sheet_dropdown.clear()
            self.sheet_dropdown.addItems(sheet_names)
            self.sheet_dropdown.setEnabled(True)  # Dropdown'u etkinleştir
            self.start_button.setEnabled(True)  # İşlemi Başlat Butonunu etkinleştir
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading sheets: {e}")

    def start_process(self):
        file_path = self.file_path_label.text()
        if not file_path:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
            return

        try:
            selected_sheet = self.sheet_dropdown.currentText()  # Seçilen sheet ismini al
            if not selected_sheet:
                QMessageBox.warning(self, "Warning", "Please select a sheet first.")
                return

            task_numbers = get_task_numbers(file_path, selected_sheet)  # Sheet ismini argüman olarak gönder
            for task_number in task_numbers:
                url = f"http://example.com/task/{task_number}"
                webbrowser.open(url)
            QMessageBox.information(self, "Success", "Process completed successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")