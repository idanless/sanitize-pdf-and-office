import win32com.client as win32


class SANITIZE_PPTX:
    def __init__(self,file_path):
        self.listofbin = []
        self.file_path = file_path
        self.pptx_app = win32.Dispatch("PowerPoint.Application")
        self.pptx_app.Visible = True
        self.pptx = self.pptx_app.Presentations.Open( self.file_path)

    def remove_hyperlinks(self):
        for slide in self.pptx.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text_frame = shape.TextFrame
                    text_frame.TextRange.Text = text_frame.TextRange.Text.replace("<w:hyperlink", "")

    def remove_binary_content(self):
        for slide in self.pptx.Slides:
            for shape in slide.Shapes:
                print(shape.Type)
                if shape.Type == 7:  # 13 corresponds to binary data (e.g., embedded objects, images)
                    self.listofbin.append(shape)
        for d in self.listofbin:
            d.Delete()

    def save(self):
        self.pptx.Save()
        self.pptx.Close()
        self.pptx.Quit()


def sanitize_pptx(file_path):
    listofbin = []
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = False
    presentation = ppt_app.Presentations.Open(file_path)

    # Remove hyperlinks from text frames
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                text_frame.TextRange.Text = text_frame.TextRange.Text.replace("<w:hyperlink", "")

    # Remove binary content
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            print(shape.Type)
            if shape.Type == 7:  # 13 corresponds to binary data (e.g., embedded objects, images)
               listofbin.append(shape)
    for d in listofbin:
        d.Delete()
    presentation.Save()
    presentation.Close()
    ppt_app.Quit()