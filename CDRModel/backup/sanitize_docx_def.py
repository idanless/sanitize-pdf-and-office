import win32com.client as win32
import os

word_app = win32.Dispatch("Word.Application")
    word_app.Visible = False
    doc = word_app.Documents.Open(file_path)

    # Remove hyperlinks from the header of each section
    for section in doc.Sections:
        if section.Headers.Count > 0:
            header_range = section.Headers(1).Range
            print(header_range)
            print(header_range.Text)


    # Remove hyperlinks from the footer of each section
    for section in doc.Sections:
        if section.Footers.Count > 0:
            footer_range = section.Footers(1).Range
            if 'http' or 'https' in footer_range:
                section.Footers(1).Range.Delete()
                for section in doc.Sections:
                    if section.Headers.Count > 0:
                        print(section.Headers(1))
                        section.Headers(1).Range.Delete()

    # Specify the start and end positions of the cover note section
    cover_note_start = 1  # Start position of the cover note section
    cover_note_end = 1  # End position of the cover note section

    cover_note_range = doc.Range(doc.Paragraphs(cover_note_start).Range.Start,
                                 doc.Paragraphs(cover_note_end).Range.End)
    #print(cover_note_range)

    # Check if the cover note range contains hyperlinks or embedded files
    has_hyperlinks = False
    has_embedded_files = False

    for shape in cover_note_range.InlineShapes:
        if shape.Type == 1:  # Embedded OLE object
            has_embedded_files = True
            break

    for field in cover_note_range.Fields:
        if field.Type == 8:  # Hyperlink field
            has_hyperlinks = True
            break

    # Remove the cover note section, embedded files, and hyperlinks if found
    if has_embedded_files or has_hyperlinks:
        # Remove embedded files
        for shape in cover_note_range.InlineShapes:
            if shape.Type == 1:  # Embedded OLE object
                shape.Delete()

        # Remove hyperlinks
        for field in cover_note_range.Fields:
            if field.Type == 8:  # Hyperlink field
                field.Delete()

        # Delete the cover note range
        cover_note_range.Delete()

    # Remove hyperlinks
    for t in range(0,3):
        for hyperlink in doc.Hyperlinks:
            hyperlink.Delete()
        # Remove binary content
    # Remove shapes containing binary content (e.g., embedded objects, images)
    for t in range(0, 3):
        for inline_shape in doc.InlineShapes:
            inline_shape.Delete()


    # Save the modified document as plain text
    text_file_path = file_path.replace(".docx", "_plaintext.docx")
    doc.SaveAs2(text_file_path, FileFormat=16)  # FileFormat=16 corresponds to plain text format

    doc.Close()
    word_app.Quit()

    # Delete the original file and rename the plain text file to the original name
    os.remove(file_path)
    os.rename(text_file_path, file_path)
# Remove macros from PPTX files