import os
import PyPDF2

# Open the PDF file in read-binary mode
with open('Z311746_B.pdf', 'rb') as file:
    # Create a PDF object
    pdf = PyPDF2.PdfFileReader(file)
    # Get the first page
    page = pdf.getPage(0)
    # Search for the text on the page
    matches = page.extractText().find('Z311746')
    # Check if the text was found
    if matches != -1:
        # Get the bounding box of the text
        x, y, width, height = page.getTextWords()[matches][:4]
        # Create a new text annotation
        annotation = PyPDF2.generic.TextAnnotation(
            rect=[x + width, y, x + width + 50, y + height],
            text='This is an annotation',
            wrap=True,
            border=PyPDF2.generic.Rectangle(2, 2)
        )
        # Add the annotation to the page
        page.addAnnotation(annotation)
        # Save the modified page back to the PDF
        pdf.addPage(page)
    # Save the output PDF file
    with open('output.pdf', 'wb') as outfile:
        pdf.write(outfile)
