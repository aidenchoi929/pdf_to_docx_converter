import os
import sys
import time
import psutil
from multiprocessing import cpu_count
from pdf2docx import Converter
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QTextEdit


class PdfConvertSeq(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        
        self.pdf_folder = "prefer"
        self.docx_folder = "prefer"
        
    def convPdfToDocx(self, pdfFilePath, docxFilePath):
        
        docxConv = Converter(pdfFilePath)
        docxConv.convert(docxFilePath, start=0, end=None)
        docxConv.close()

    def pdfFileList(self):
        try:
            return [f for f in os.listdir(self.pdf_folder) if f.endswith('.pdf')]
        except (Exception):
            return
    
    def on_convert(self, converter):
        self.pdf_folder = self.pdf_folder_input.text()
        self.docx_folder = self.docx_folder_input.text()
        self.convert()
        sys.stdout = sys.__stdout__
    
    def log(self, text):
        self.log_window.append(text)
    
    def convert(self):
        total_cpu_before = 0
        total_cpu_after = 0
        total_time = 0
        pdf_files = self.pdfFileList()
        if not pdf_files:
            self.log("PDF file does not exist in the folder.")
            return
        num_processes = cpu_count()
        self.log(f'---------------------------------------------------------------------------------------------')
        self.log(f'Total number of files : {len(pdf_files)}')
        self.log(f'Total number of logical cpu core : {num_processes}')
        self.log(f'---------------------------------------------------------------------------------------------')

        for pdf_file_name in pdf_files:
            pdf_file_path = os.path.join(self.pdf_folder, pdf_file_name)
            docx_file_name = os.path.splitext(pdf_file_name)[0] + '.docx'
            docx_file_path = os.path.join(self.docx_folder, docx_file_name)


            befCpuUsage = psutil.cpu_percent()

            convertStartTime = time.time()

            self.convPdfToDocx(pdf_file_path, docx_file_path)

            convertEndTime = time.time()

            afCpuUsage = psutil.cpu_percent()

            self.log(f'{pdf_file_name} has converted into {docx_file_name} successfully')
            self.log(f'Elapsed time: {convertEndTime - convertStartTime:.2f} seconds')
            self.log(f'CPU Usage : Before={befCpuUsage}% / After={afCpuUsage}%')

            total_cpu_before += befCpuUsage
            total_cpu_after += afCpuUsage
            total_time += convertEndTime - convertStartTime

        avg_cpu_before = total_cpu_before / len(pdf_files)
        avg_cpu_after = total_cpu_after / len(pdf_files)
        self.log(f'---------------------------------------------------------------------------------------------')
        self.log(f'PDF to DOCX conversion processing final outcome')
        self.log(f'Total elapsed time: {total_time:.2f} seconds')
        self.log(f'Average CPU Usage: Before={avg_cpu_before:.2f}%, After={avg_cpu_after:.2f}%')
        self.log(f'---------------------------------------------------------------------------------------------')

    def initUI(self):
        layout = QVBoxLayout()

        self.pdf_folder_input = QLineEdit(self)
        self.pdf_folder_input.setPlaceholderText("Enter the PDF folder path you wish to convert")
        layout.addWidget(self.pdf_folder_input)

        self.docx_folder_input = QLineEdit(self)
        self.docx_folder_input.setPlaceholderText("Enter the DOCX folder path you wish to be stored in")
        layout.addWidget(self.docx_folder_input)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.on_convert)
        layout.addWidget(self.convert_button)

        self.log_window = QTextEdit(self)
        self.log_window.setReadOnly(True)
        layout.addWidget(self.log_window)

        self.setLayout(layout)
        self.setWindowTitle('PDF to DOCX converter(Sequential)')
        self.setGeometry(500, 500, 500, 300)

def main():
    app = QApplication(sys.argv)
    ex = PdfConvertSeq()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()