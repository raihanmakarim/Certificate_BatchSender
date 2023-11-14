from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import utils
from reportlab.pdfbase.pdfmetrics import stringWidth
import os
import io


def get_text_width(text, font_name, font_size):
    pdfmetrics.registerFont(TTFont(font_name, font_name + '.ttf'))
    width = stringWidth(text, font_name, font_size)
    return width


def adjust_font_size(can, text, font_name, target_width, min_font_size=6):
    font_size = 48
    current_width = get_text_width(text, font_name, font_size)

    while current_width > target_width and font_size > min_font_size:
        font_size -= 1
        current_width = get_text_width(text, font_name, font_size)

    can.setFont(font_name, font_size)


def generate_certificate(username):
    pdfmetrics.registerFont(
        TTFont('Montserrat-ExtraBold', 'Montserrat-ExtraBold.ttf'))

    # Create a new PDF with Reportlab
    text_color = (0.0390625, 0.60546875, 0.60546875)

    # Open the existing PDF to get its size
    existing_pdf_path = "Sertifikatrev.pdf"
    existing_pdf_file = open(existing_pdf_path, "rb")

    with existing_pdf_file:
        existing_pdf = PdfReader(existing_pdf_file)
        existing_page = existing_pdf.pages[0]
        existing_width, existing_height = existing_page.mediabox.upper_right

        # Set the canvas size to match the existing PDF
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(existing_width, existing_height))
        can.setFillColorRGB(*text_color)

        x_margin = 20  # Minimum margin from text to page edge
        target_width = existing_width - 2 * x_margin  # Target width for centered text

        text = username

        # Automatically adjust font size and center text horizontally
        adjust_font_size(can, text, 'Montserrat-ExtraBold', target_width)
        text_width = get_text_width(
            text, 'Montserrat-ExtraBold', can._fontsize)
        x_position = (float(existing_width) - float(text_width)) / 2
        y_position = 330

        # Draw centered text on the canvas
        can.drawString(x_position, y_position, text)

        can.showPage()
        can.save()

        # Move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfReader(packet)

        # Read the existing PDF again
        existing_pdf_file.seek(0)
        existing_pdf = PdfReader(existing_pdf_file)
        output = PdfWriter()

        # Add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        folder_path = "/Users/r17/Documents/python/certificate/"
        os.makedirs(folder_path, exist_ok=True)
        output_file_path = f"{folder_path}Sertifikat_{username}.pdf"

        output_stream = open(output_file_path, "wb")
        output.write(output_stream)
        output_stream.close()


def send_email(username, email, attachment_path):
    # Email server details (replace with your own)
    smtp_server = 'your_smtp_server'
    smtp_port = 587
    smtp_username = 'info@r17.co.id'
    smtp_password = 'Info@2023-11'

    sender_email = 'info@r17.co.id'
    receiver_email = email
    subject = 'Certificate Attached'
    body = f"Dear {username},\n\nPlease find your certificate attached.\n\nBest regards,\nYour Organization"

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Attach the certificate file
    with open(attachment_path, 'rb') as attachment:
        part = MIMEText(body)
        message.attach(part)

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        part.add_header('Content-Disposition',
                        f'attachment; filename={os.path.basename(attachment_path)}')
        message.attach(part)

    # Connect to the SMTP server and send the email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        return True
    except Exception as e:
        print(f"Error sending email to {email}: {e}")
        return False


def send_emails(users):
    failed_emails = []

    for user in users:
        username = user['username']
        email = user['email']
        certificate_path = generate_certificate(username)

        if os.path.exists(certificate_path):
            if send_email(username, email, certificate_path):
                print(f"Email sent successfully to {email}")
            else:
                print(f"Failed to send email to {email}")
                failed_emails.append(email)
        else:
            print(f"Failed to generate certificate for {username}")

    return failed_emails


if __name__ == "__main__":
    user_list = [
        {'username': 'User1', 'email': 'user1@example.com'},
        {'username': 'User2', 'email': 'user2@example.com'},
        # Add more users as needed
    ]

    failed_emails = send_emails(user_list)

    if failed_emails:
        print("Failed to send emails to the following addresses:")
        for email in failed_emails:
            print(email)
