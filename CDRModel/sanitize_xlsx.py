import xlrd
from xlutils.copy import copy as xl_copy
import win32com.client as win32

class SANITIZE_XLSX:
    def __init__(self, file_path,office_flag):
        self.file_path = file_path
        self.excel = win32.Dispatch("Excel.Application")
        if office_flag:
            self.excel.Visible = True
            self.excel.DisplayAlerts = False
        else:
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

        self.workbook = self.excel.Workbooks.Open(self.file_path)
        # Remove macros
        try:
            self.workbook.VBProject.VBComponents.Remove(self.workbook.VBProject.VBComponents("Module1"))
        except:
            pass

    def remove_macros(self):
        rb = xlrd.open_workbook(self.file_path)
        wb = xl_copy(rb)
        wb.save(self.file_path)

    def remove_hyperlinks(self):
        for sheet in self.workbook.Sheets:
            # Remove hyperlinks in the sheet
            sheet.Hyperlinks.Delete()

    def remove_formula(self):
        for worksheet in self.workbook.Worksheets:
            used_range = worksheet.UsedRange
            for row in used_range.Rows:
                for cell in row.Cells:
                    if cell.Formula:
                        cell.Formula = cell.Value

    def remove_binary(self):
        for worksheet in self.workbook.Worksheets:
            used_range = worksheet.UsedRange
            for row in used_range.Rows:
                for cell in row.Cells:
                    if cell.Formula:
                        cell.Formula = cell.Value
            # Remove binary content
            for shape in worksheet.Shapes:
                if shape.Type == 13:  # 13 corresponds to binary data (e.g., embedded objects, images)
                    shape.Delete()


    def save(self):
        self.workbook.Save()
        self.workbook.Close()
