import os
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from docx import Document
from docx.shared import RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx2pdf import convert  # Import the convert function


# Dummy user data (name, email)
user_data = [
    {'name': 'yp', 'email': 'raihanmakarim21@gmail.com'},
    {'name': 'yoyo', 'email': 'example@example.com'},
    {'name': 'ehehe', 'email': 'example@example.com'},
    {'name': 'wowo', 'email': 'example@example.com'},
    {'name': 'oi', 'email': 'example@example.com'},
]


# Iterate through user data and generate documents
for user in user_data:
    # Create a new Word document for each user
    doc = Document('Sertifikat.docx')

    # Replace text in the Word document and set text color to dark grey
    for paragraph in doc.paragraphs:
        if 'Nama Peserta' in paragraph.text:
            paragraph.text = paragraph.text.replace(
                'Nama Peserta', user['name'])
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(89, 89, 90)  # Dark grey color

    # Define the modified document filename for the current user
    modified_docx = f'Sertifikat_{user["name"]}.docx'

    # Save the modified Word document with the new filename
    doc.save(modified_docx)

    pdf_name = f'Sertifikat_{user["name"]}.pdf'
    # Convert the modified Word document to PDF
    convert(modified_docx, pdf_name)
# Send the PDF as an email (You will need to fill in the email details)
# Please use a valid SMTP server and credentials for actual email sending.

# Example code for sending emails
# server = smtplib.SMTP('smtp.gmail.com', 587)
# server.starttls()
# server.login('raihanmakarim21@gmail.com', 'macfinalraihan21')


# port = 465  # For SSL
# password = 'macfinalraihan21'
# email = 'raihanmakarim21@gmail.com'
# # Create a secure SSL context
# context = ssl.create_default_context()

# with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
#     server.login(email, password)

#     # Iterate through user data and send emails
#     for user in user_data:
#         msg = MIMEMultipart()
#         msg['From'] = email
#         msg['To'] = user['email']
#         msg['Subject'] = 'Your Certificate'

#         body = f"Dear {user['name']},\n\nPlease find your certificate attached."
#         msg.attach(MIMEText(body, 'plain'))

#         with open(pdf_file, 'rb') as attachment:
#             part = MIMEApplication(attachment.read(), _subtype="pdf")
#             part.add_header('Content-Disposition', 'attachment', filename='certificate.pdf')
#             msg.attach(part)

#         server.sendmail(email, user['email'], msg.as_string())
