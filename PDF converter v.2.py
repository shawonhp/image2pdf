import os
import img2pdf
from docx2pdf import convert
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QMessageBox, QInputDialog
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()

        self.files = []
        self.output_file = ""
        self.mode = ""

        self.setWindowTitle("PDF Converter by Shawon")
        self.resize(300, 300)

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.input_button = QPushButton("Select or Drag Files Here")
        self.input_button.clicked.connect(self.select_files)
        self.layout.addWidget(self.input_button)

        self.input_label = QLabel()
        self.layout.addWidget(self.input_label)

        self.convert_button = QPushButton("Convert")
        self.convert_button.clicked.connect(self.convert_files_to_pdf)
        self.layout.addWidget(self.convert_button)

        self.back_button = QPushButton("Back")
        self.back_button.clicked.connect(self.reset_app)
        self.layout.addWidget(self.back_button)

        self.exit_button = QPushButton("Exit")
        self.exit_button.clicked.connect(QApplication.instance().quit)
        self.layout.addWidget(self.exit_button)

        self.setAcceptDrops(True)

        self.mode, ok = QInputDialog.getItem(self, "Select mode", "Choose a conversion mode:", ["Images to PDF", "Docx to PDF"], 0, False)
        if ok and self.mode:
            self.mode = self.mode.replace(" ", "").lower()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.endswith(self.get_extensions()):
                self.files.append(path)
        self.input_label.setText("Files Selected Successfully!")

    def select_files(self):
        files = QFileDialog.getOpenFileNames(self, "Select Files", filter=self.get_filter())[0]
        if files:
            self.files = files
            self.input_label.setText("Files Selected Successfully!")

    def convert_files_to_pdf(self):
        try:
            self.output_file = QFileDialog.getSaveFileName(self, "Select Output File", filter="PDF Files (*.pdf)")[0]
            if self.output_file:
                if self.mode == "imagestopdf":
                    pdf_bytes = img2pdf.convert(self.files, dpi=300, x=None, y=None)
                    with open(self.output_file, "wb") as f:
                        f.write(pdf_bytes)
                elif self.mode == "docxtopdf":
                    for file in self.files:
                        output_file = self.output_file.replace('.pdf', f'_{os.path.basename(file)}.pdf')
                        convert(file, output_file)
                QMessageBox.information(self, "Success", "Files successfully converted to PDF!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def get_extensions(self):
        if self.mode == "imagestopdf":
            return (".png", ".jpg", ".jpeg", ".bmp", ".gif")
        elif self.mode == "docxtopdf":
            return (".docx", ".doc")

    def get_filter(self):
        if self.mode == "imagestopdf":
            return "Images (*.png *.jpg *.jpeg *.bmp *.gif)"
        elif self.mode == "docxtopdf":
            return "Documents (*.docx *.doc)"

    def reset_app(self):
        self.files = []
        self.output_file = ""
        self.mode, ok = QInputDialog.getItem(self, "Select mode", "Choose a conversion mode:", ["Images to PDF", "Docx to PDF"], 0, False)
        if ok and self.mode:
            self.mode = self.mode.replace(" ", "").lower()
        self.input_label.setText("")

app = QApplication([])
window = ConverterApp()
window.show()
app.exec_()
