import sys
from PyQt5.QtWidgets import QApplication, QTabWidget, QWidget, QVBoxLayout
from gui.tabs import ExcelTab

class MainApp(QTabWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Task Number Retriever")
        self.setGeometry(100, 100, 600, 400)
        
        self.addTab(ExcelTab(), "Excel File")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_app = MainApp()
    main_app.show()
    sys.exit(app.exec_())