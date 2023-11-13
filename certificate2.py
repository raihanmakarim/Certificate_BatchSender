from PyPDF2 import PdfWriter, PdfReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(
    TTFont('Montserrat-ExtraBold', 'Montserrat-ExtraBold.ttf'))

packet = io.BytesIO()
# Create a new PDF with Reportlab

x_position = 400
y_position = 370
# rgb values divide by 256
text_color = (0.0390625, 0.60546875, 0.60546875)


can = canvas.Canvas(packet, pagesize=letter)
can.setFont('Montserrat-ExtraBold', 48)
can.setFillColorRGB(*text_color)
can.drawString(x_position, y_position, "Hello world qwkjkjqw  qkwdkqwd")
can.showPage()
can.save()

# Move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfReader(packet)
# Read your existing PDF
existing_pdf = PdfReader(open("Sertifikatrev.pdf", "rb"))
output = PdfWriter()
# Add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.pages[0]  # Corrected syntax
page.merge_page(new_pdf.pages[0])
output.add_page(page)
# Finally, write "output" to a real file
outputStream = open("testq.pdf", "wb")
output.write(outputStream)
outputStream.close()
