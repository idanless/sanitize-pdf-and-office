from pdf2image import convert_from_path
import img2pdf
import os
import shutil
import pdfrw
import PyPDF2 as pf
import getpass
HighQ = 'PNG'
LowQ = 'JPEG'
dpi = 130

path="c:/Users/"+str(getpass.getuser())+"/AppData/Local/temp/pdf_tmp"


try:
    os.mkdir(os.path.realpath(path))
except FileExistsError:
    pass


class PDF2IMG:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        original_file_size = os.path.getsize(pdf_path)
        target_file_size = original_file_size
        images = convert_from_path(pdf_path, first_page=0, last_page=1,poppler_path=r'.\Poppler_for_Windows\poppler-23.05.0\Library\bin')
        image = images[0]
        original_width, original_height = image.size
        estimated_dpi = int((target_file_size / (original_width * original_height)) ** 0.9)
        self.pdfimg = convert_from_path(pdf_path, poppler_path=r'.\Poppler_for_Windows\poppler-23.05.0\Library\bin',dpi=estimated_dpi)
        self.list_page = []
        self.dic = pdf_path.split('\\')[-1].split('.')[0]
        self.dirname = os.path.dirname(pdf_path)

    def runpdfimg(self):
        try:
            new_dic = path+f'\{self.dic}'
            os.mkdir(new_dic)
        except OSError:
            shutil.rmtree(new_dic)
            os.mkdir(new_dic)
        for i, image in enumerate(self.pdfimg):
            dic = new_dic + f'\page_{i}.jpg'
            image.save(dic, 'JPEG')
            self.list_page.append(dic)


class IMG2PDF(PDF2IMG):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)
        self.out_pdf = f'{self.dic}_full_CDR.pdf'

    def print_list(self):
        print(self.list_page)

    def img2pdf(self):
        self.dirname = f'{self.dirname}\{self.out_pdf}'
        with open(self.dirname, "wb") as pdf_file:
            pdf_file.write(img2pdf.convert(self.list_page))

    def remove_tmp(self):
        new_dic = path + f'\{self.dic}'
        shutil.rmtree(new_dic)

class LINK_REMOVE(IMG2PDF):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)
        self.file_path = pdf_path
        self.link = "http://blockby.CDR/"
    def remove_link(self):
        pdf = pdfrw.PdfReader(self.file_path)  # Load the pdf
        new_pdf = pdfrw.PdfWriter()  # Create an empty pdf
        for page in pdf.pages:  # Go through the pages
            # Links are in Annots, but some pages don't have links so Annots returns None
            for annot in page.Annots or []:
                old_url = annot.A.URI
                new_url = pdfrw.objects.pdfstring.PdfString(f"({self.link})")
                # Override the URL with ours
                annot.A.URI = new_url
            new_pdf.addpage(page)
        new_pdf.write(self.file_path)
        # Load the PDF document

class REMOVE_Embedded(LINK_REMOVE):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)

    def remove_Emb(self):
        # Read the input PDF into a variable
        pdf = pf.PdfFileReader(self.file_path)
        # Create a new PDF writer
        writer = pf.PdfFileWriter()
        # Copy pages from the input PDF to the new PDF
        for page_num in range(pdf.getNumPages()):
            page = pdf.getPage(page_num)
            writer.addPage(page)
        # Save the new PDF without embedded files
        with open(self.file_path, 'wb') as output_pdf:
            writer.write(output_pdf)


class PDF_Main(REMOVE_Embedded):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)

