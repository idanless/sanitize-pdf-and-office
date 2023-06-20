import os
from CDRModel import SANITIZE_DOCX,SANITIZE_XLSX,SANITIZE_PPTX,PDF_Main
import time
import psutil


def is_process_running(process_name):
    for process in psutil.process_iter(['name']):
        if process.info['name'] == process_name:
            return True
    return False

# Check if any Office application is running
office_running = any(is_process_running(proc) for proc in ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"])
print('office_running',office_running)

def sanitize_xlsx(file_path,office_running=office_running):
    cdr_xlsx = SANITIZE_XLSX(file_path,office_running)
    cdr_xlsx.remove_hyperlinks()
    cdr_xlsx.remove_formula()
    cdr_xlsx.remove_binary()
    cdr_xlsx.save()
    if file_path.endswith('xls'):
        cdr_xlsx.remove_macros()

def sanitize_docx(file_path,office_running=office_running):
    cdr_docx = SANITIZE_DOCX(file_path,office_running)
    cdr_docx.remove_hyperlinks()
    cdr_docx.remove_binary_content()
    cdr_docx.remove_revisions()
    cdr_docx.save_doc()
    time.sleep(1)
    #fix bug with docx files
    cdr_docx = SANITIZE_DOCX(file_path,office_running)
    cdr_docx.remove_hyperlinks()
    cdr_docx.remove_binary_content()
    cdr_docx.remove_revisions()
    cdr_docx.save_doc()


def sanitize_pptx(file_path):
    cdr_pptx = SANITIZE_PPTX(file_path)
    cdr_pptx.remove_hyperlinks()
    cdr_pptx.remove_binary_content()
    cdr_pptx.save()

def sanitize_pdf(file_path):
    cdr_pdf = PDF_Main(file_path)
    cdr_pdf.remove_link()
    cdr_pdf.remove_Emb()
    cdr_pdf.runpdfimg()
    cdr_pdf.img2pdf()
    cdr_pdf.remove_tmp()


def sanitize_file(file_path):
    if file_path.endswith(".pptx"):
        sanitize_pptx(file_path)
    elif file_path.endswith(".xlsx") or file_path.endswith(".xls")or file_path.endswith(".xltm"):
        sanitize_xlsx(file_path)
    elif file_path.endswith(".docx"):
        sanitize_docx(file_path)
    elif file_path.endswith(".pdf"):
        sanitize_pdf(file_path)



def sanitize_directory(directory_path):
    for file_name in os.listdir(directory_path):
        #print(file_name)
        file_path = os.path.join(directory_path, file_name)
        sanitize_file(file_path)

# Example usage
directory_path = r"C:\tmp"
sanitize_directory(directory_path)
