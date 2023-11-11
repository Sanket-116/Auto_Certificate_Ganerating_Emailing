import csv
import io
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import logging
import os


#to track progress
logging.basicConfig(level=logging.INFO)

#new certificate generator fun()
def generate_certificate(participant_name, participant_email, template_path):

    # create a new PDF with the participant's information
    output = PdfFileWriter()
    template = PdfFileReader(template_path)
    page = template.getPage(0)  

    packet = io.BytesIO()
    can = canvas.Canvas(packet)

    #text co-ordinates on template
    text_color = 'black'
    x = 230
    y = 440  
    font_size = 36

    can.setFont("Helvetica", font_size)

    can.drawString(x, y, f"{participant_name}")

    can.save()

    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    page.merge_page(new_pdf.getPage(0))
    output.addPage(page)

    # Save the certificate
    certificate_path = f"{participant_name}_certificate.pdf"
    with open(certificate_path, "wb") as output_pdf:
        output.write(output_pdf)

    return certificate_path


# Function to send email
def send_email(to_address, subject, body, attachment_path):

    #sender details
    sender_email = "sanketmaske16@gmail.com"
    sender_password = "iavrzaungzqcagrq"  # Used App Password here

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_address
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        with open(attachment_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), Name="certificate.pdf")
            part['Content-Disposition'] = f'attachment; filename="GDSC_Certificate.pdf"'
            msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to_address, msg.as_string())

        logging.info(f"Email sent successfully to {to_address} with certificate attached.\n")

    except Exception as e:
        logging.error(f"Error sending email to {to_address}: {e}")





# reads participant names and email from CSV
with open('participants.csv', 'r') as file:
    reader = csv.reader(file)
    next(reader)
    for row in reader:
        #check for missing data
        if len(row) >= 2:
            participant_name = row[0]
            participant_email = row[1]

            #template_path 
            template_path = "Certificate_temp.pdf"
            
            # calling PDF generator fun()
            certificate_path = generate_certificate(participant_name, participant_email, template_path)

            print(f"\nCertificate generated for {participant_name}. Saved at {certificate_path}\n")

            #calling Email-send fun()
            send_email(participant_email, "\nCertificate of Participation", f"Dear {participant_name},\n\nThank you for participating in our event. Attached is your certificate of participation.", certificate_path)

        else:
            logging.warning(f"Skipping row {row}, as it doesn't have enough elements.")
