import sys
import fitz  # PyMuPDF
import re
import pandas as pd
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QTextEdit, QLabel
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QFont

class Worker(QThread):
    update_message = pyqtSignal(str)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        self.process_pdfs(self.folder_path)

    def process_pdfs(self, folder_path):
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        if not pdf_files:
            self.update_message.emit("No PDF files found in the selected folder.")
            return

        all_info = []
        for pdf_file in pdf_files:
            try:
                self.update_message.emit(f"Processing {pdf_file}...")
                pdf_path = os.path.join(folder_path, pdf_file)
                text = self.extract_text_from_pdf(pdf_path)
                info = self.extract_information(text)
                all_info.append(info)
            except Exception as e:
                self.update_message.emit(f"Error processing {pdf_file}: {e}")

        if all_info:
            self.update_message.emit("All PDFs processed. Ready to export CSV.")
            self.all_info = all_info

    def extract_text_from_pdf(self, pdf_path):
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text

    def extract_information(self, text):
        # Regular expressions to find name, email, phone number, and address
        name_pattern = r"Name:\s*(.*)"
        email_pattern = r"[\w\.-]+@[\w\.-]+"
        phone_pattern = r"\+?\d[\d -]{8,12}\d"
        address_pattern = r"Address:\s*(.*)"
        
        name = re.search(name_pattern, text)
        email = re.search(email_pattern, text)
        phone = re.search(phone_pattern, text)
        address = re.search(address_pattern, text)
        
        return {
            "Name": name.group(1) if name else None,
            "Email": email.group(0) if email else None,
            "Phone Number": phone.group(0) if phone else None,
            "Address": address.group(1) if address else None,
        }

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Resume Extractor")
        self.setGeometry(100, 100, 800, 600)
        self.setFont(QFont('Arial', 12))

        layout = QVBoxLayout()

        self.message_box = QTextEdit()
        self.message_box.setReadOnly(True)
        self.message_box.setFont(QFont('Arial', 10))
        layout.addWidget(self.message_box)

        self.select_folder_button = QPushButton("Select Folder")
        self.select_folder_button.setFont(QFont('Arial', 12))
        self.select_folder_button.setFixedHeight(40)
        self.select_folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.select_folder_button)

        self.export_csv_button = QPushButton("Export CSV")
        self.export_csv_button.setFont(QFont('Arial', 12))
        self.export_csv_button.setFixedHeight(40)
        self.export_csv_button.clicked.connect(self.export_csv)
        layout.addWidget(self.export_csv_button)

        self.info_label = QLabel("Select a folder containing PDF resumes to extract information.")
        self.info_label.setFont(QFont('Arial', 10))
        layout.addWidget(self.info_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.worker = Worker(folder_path)
            self.worker.update_message.connect(self.append_message)
            self.worker.start()

    def append_message(self, message):
        self.message_box.append(message)

    def export_csv(self):
        if hasattr(self.worker, 'all_info') and self.worker.all_info:
            save_path, _ = QFileDialog.getSaveFileName(self, "Save CSV", filter="CSV files (*.csv)")
            if save_path:
                df = pd.DataFrame(self.worker.all_info)
                df.to_csv(save_path, index=False)
                self.append_message("Information saved to CSV.")
        else:
            self.append_message("No data to export. Please process the PDFs first.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
