import win32com.client as win32
import win32gui
import win32con


class SANITIZE_PPTX:
    def __init__(self,file_path):
        self.listofbin = []
        self.file_path = file_path
        self.pptx_app = win32.Dispatch("PowerPoint.Application")
        hwnd = win32gui.FindWindow(None, "Microsoft PowerPoint")
        self.pptx_app.Visible = True
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        self.pptx = self.pptx_app.Presentations.Open(self.file_path)


    def remove_hyperlinks(self):
        for slide in self.pptx.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text_frame = shape.TextFrame
                    text_frame.TextRange.Text = text_frame.TextRange.Text.replace("<w:hyperlink", "")

    def remove_binary_content(self):
        for slide in self.pptx.Slides:
            for shape in slide.Shapes:
                if shape.Type == 7:  # 13 corresponds to binary data (e.g., embedded objects, images)
                    self.listofbin.append(shape)
        for d in self.listofbin:
            d.Delete()

    def save(self):
        self.pptx.Save()
        self.pptx.Close()
        self.pptx_app.Quit()


