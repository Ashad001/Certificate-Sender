import os
import pandas as pd
import ssl
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import letter, landscape
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
import random

def overlay_text_on_template(template_path, names, output_directory, email):
    os.makedirs(output_directory, exist_ok=True)

    for name in names:
        template = PdfReader(open(template_path, "rb"))
        output_path = f"{output_directory}/{email}_certificate.pdf"
        output = PdfWriter()
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=landscape(letter))
        font_size = 40
        try:
            can.setFont("Times-Roman", font_size)
        except:
            can.setFont("Helvetica", font_size)

        page_width, page_height = landscape(letter)

        text = name
        text_width = can.stringWidth(text)
        text_height = font_size
        text_x = ((page_width - text_width) / 2) + 24
        text_y = ((page_height - text_height) / 2) - 10
        can.drawString(text_x, text_y, text)

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)
        template_page = template.pages[0]
        template_page.merge_page(overlay.pages[0])
        output.add_page(template_page)

        with open(output_path, "wb") as output_file:
            output.write(output_file)


def send_email(receiver_email, subject, message, attachment_path=None):
    sender_email = "acm.khi@nu.edu.pk"
    sender_password = "ypxlcdfoecnsobej"
    smtp_server = "smtp.gmail.com"
    smtp_port = 465

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.attach(MIMEText(message, "html"))

    if attachment_path:
        with open(attachment_path, "rb") as pdf_file:
            pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
            pdf_attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(attachment_path),
            )
            msg.attach(pdf_attachment)

    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, receiver_email, msg.as_string())
    server.quit()
    
def update_log(log_file, data):
    with open(log_file, "a") as log:
        log.write(data + "\n")

def read_log(log_file):
    try:
        with open(log_file, "r") as log:
            data = log.read().splitlines()
        return data
    except FileNotFoundError:
        return []



excel_file_path = "./batchwise/batchwise_distinct/test.xlsx"
xls = pd.ExcelFile(excel_file_path)
sheet_names = xls.sheet_names
print(sheet_names)
certificate_template_path = (
    r"CODERSCUP_CERTIFICATE.pdf"
)

output_directory = "./cert"
i = 0
log_file = "sent_emails_log.txt"
error_file = "error_log.csv"
error_data = []
sent_emails_log = read_log(log_file)
os.makedirs(output_directory, exist_ok=True)
for sheet_name in sheet_names:
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    for _, row in df.iterrows():
        member_name = row["Name"]
        section = row["Section"]
        roll_number = row["NU ID"]
        email = row["Email"]
        if email in sent_emails_log:
            print(f"Email already sent to {member_name}, {roll_number}, {email}. Skipping...")
            continue
        names = [f"{member_name}"]
        overlay_text_on_template(
            certificate_template_path, names, output_directory, email
        )

        pdf_certificate_path = f"{output_directory}/{email}_certificate.pdf"

        subject = "Certificate for Participation"
        html = f"""

    <html>
  <head>
    <title>Report</title>
    <meta name="viewport" content="width=device-width, initial-scale=1"> <!-- Added viewport meta tag -->

    <style>
       .contain {{
        width: 80%;
        margin: 0 auto;
        text-align: right;
        font-family: Verdana, Arial, sans-serif; /* Change font to Verdana */
        padding-top: 3%;
      }}
      .container {{
        max-width: 80%;
        margin: 0 auto;
        text-align: justify;
        font-family: Verdana, Arial, sans-serif;
        font-size: 15px;
        color: white;
        padding: 0% 2%; /* Adjusted padding */
      }}

      .logo {{
        max-width: 70%; /* Ensure the logo is responsive */
        height: auto;
        display: block; /* Center the logo */
        margin: 0 auto;
      }}

      .content {{
        font-family: Verdana, Arial, sans-serif;
      }}

      .DTL{{
        margn-right: 0px;
      }}
      p {{
        color: white;
      }}

      img.g-img + div {{
        display:none;
      }}

      /* Media query for smaller screens */
      @media (max-width: 768px) {{
        .container {{
          font-size: 14px; /* Adjusted font size for smaller screens */
        }}
      }}
    </style>
  </head>
  <body style="color:white;">
    <div class="content">
      <div style="margin: 20px auto; width: 90%; padding: 5px 0"> <!-- Adjusted width -->

        <div style="color: white; text-align: center; background: -webkit-linear-gradient(0deg, #590909, #300203 100%); padding: 0% 0%; padding-bottom: 5%">
          <p class="DTL" style="font-size: 20px; text-align: center; padding-top: 0%; margin: 0 auto; width: 50%; margin-bottom: 4%">
            <b>Dear {member_name},</b>
          </p>
          <div class="container">

                <p>Thank You for participating in ACM Coders Cup 2023! Your dedication and enthusiasm have shone brightly, and we're thrilled to present your well-deserved participation certificates. Each one of you contributed to the success of this event, and your hard work is truly appreciated.</p>

                <p>It's been an incredible experience engaging with such a talented and passionate group. Your innovative solutions and problem-solving skills have left a lasting impression.<strong> See you in Developer's Day 2024 for more exciting challenges and opportunities! </strong></p>

                <p>Once again, congratulations on your achievements. We look forward to witnessing more incredible feats in the upcoming competitions.</p>

                <p><strong>Thank you for being part of this journey!</strong></p>
<br><br>
          </p></div>
          <p class="contain" style="font-size: 15px;">
            Best regards,<br>
                Team ACM
          </p>
        </div>
        <hr style="border: none; border-top: 5px solid #eee" />
        <div style="float: left; padding: 8px 0; color: #aaa; font-size: 0.8em; line-height: 1; font-weight: 300">
          <p style="color:black;">Contact Us</p>
          <p>
            <a href="mailto:acm.khi@nu.edu.pk">acm.khi@nu.edu.pk</a>.
          </p>
        </div>
      </div>
    </div>
  </body>
</html>"""
        try:
            send_email(email, subject, html, pdf_certificate_path)
            message = member_name + ', ' + email + ', ' + roll_number + ', ' + section
            print("Email sent for ", message)
            update_log(log_file, email)
            time.sleep(random.randint(1, 3))
        except Exception as e:
            error_data.append({
                "Member Name": member_name,
                "Roll Number": roll_number,
                "Email": email,
                "Section": section
            })
            error_df = pd.DataFrame(error_data)
            error_df.to_csv(error_file, mode='a', header=not os.path.exists(error_file), index=False)
            print(f"Error sending email to {member_name}, {roll_number}, {email}: {e}")


