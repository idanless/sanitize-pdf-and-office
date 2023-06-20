import win32com.client as win32
import os

#os.system("taskkill /f /im WINWORD.EXE > nul 2>&1")
class SANITIZE_DOCX:
    def __init__(self, file_path,office_flag):
        cover_note_start = 1  # Start position of the cover note section
        cover_note_end = 1  # End position of the cover note section
        self.file_path = file_path
        self.word = win32.Dispatch("Word.Application")
        if office_flag:
            self.word.Visible = True
        else:
            self.word.Visible = False
        self.doc = self.word.Documents.Open(self.file_path)
        self.cover_note_range = self.doc.Range(self.doc.Paragraphs(cover_note_start).Range.Start,
                                               self.doc.Paragraphs(cover_note_end).Range.End)
    def remove_hyperlinks(self):
        has_hyperlinks = False
        # Remove hyperlinks from the header of each section
        for section in self.doc.Sections:
            if section.Headers.Count > 0:
                header_range = section.Headers(1).Range
        # Remove hyperlinks from the footer of each section
        for section in self.doc.Sections:
            if section.Footers.Count > 0:
                footer_range = section.Footers(1).Range
                if 'http' or 'https' in footer_range:
                    section.Footers(1).Range.Delete()
                    for section in self.doc.Sections:
                        if section.Headers.Count > 0:
                            print(section.Headers(1))
                            section.Headers(1).Range.Delete()
        for shape in self.cover_note_range.InlineShapes:
            if shape.Type == 1:  # Embedded OLE object
                has_embedded_files = True
                break

        for field in self.cover_note_range.Fields:
            if field.Type == 8:  # Hyperlink field
                has_hyperlinks = True
                break

        self.cover_note_range.Delete()
        # Remove hyperlinks
        for t in range(0, 3):
            for hyperlink in self.doc.Hyperlinks:
                hyperlink.Delete()
            # Remove binary content

    def remove_binary_content(self):
        for i in range(0, 3):
            has_embedded_files = False
            for shape in self.cover_note_range.InlineShapes:
                if shape.Type == 1:  # Embedded OLE object
                    has_embedded_files = True
                    break
            if has_embedded_files:
                # Remove embedded files
                for shape in self.cover_note_range.InlineShapes:
                    if shape.Type == 1:  # Embedded OLE object
                        shape.Delete()
            for t in range(0, 3):
                for inline_shape in self.doc.InlineShapes:
                    inline_shape.Delete()
                # Delete the cover note range


    def remove_revisions(self):
        # Disable track changes
        #self.doc.TrackRevisions = False
        # Accept all revisions
        #self.doc.AcceptAllRevisions()
        revision_notes = []
        comments = self.doc.Comments
        # Print the comments
        for index, comment in enumerate(comments):
            author = comment.Author
            text = comment.Range.Text.strip()



            # Save the document to apply the changes


    def save_doc(self):
        # Save the modified document as plain text
        text_file_path = self.file_path.replace(".docx", "_plaintext.docx")
        self.doc.SaveAs2(text_file_path, FileFormat=16)  # FileFormat=16 corresponds to plain text format
        self.doc.Close()
        self.word.Quit()
        # Delete the original file and rename the plain text file to the original name
        os.remove(self.file_path)
        os.rename(text_file_path, self.file_path)


