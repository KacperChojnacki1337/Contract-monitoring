import os
import openpyxl
import datetime
from datetime import timedelta
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from openpyxl.styles.numbers import FORMAT_PERCENTAGE
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import Font, PatternFill

#Files part

#date
today = datetime.datetime.now().strftime('%Y-%m-%d')
yesterday = (datetime.datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

# path to excel files
folder_path = r""
if not os.path.exists(r"" + today):
    os.makedirs(r"" + today)
folder_path2 = r"" + today
print("Folder creation completed")

# List of files in the folder
file_name = 'Monitoring-SC.xlsx' 
workbook = load_workbook(filename=os.path.join(folder_path, file_name),data_only=True)
ws = workbook.active
unique_clients = set(cell.value for cell in ws['A'][1:])
unique_clients.discard(None)
print("A list of unique customers has been created")

#variables for font
bold_font = Font(bold=True)
grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# Convert Polish characters to their ASCII equivalents
def sanitize_filename(filename):
    filename = filename.replace("ą", "a").replace("ć", "c").replace("ę", "e")
    filename = filename.replace("ł", "l").replace("ń", "n").replace("ó", "o")
    filename = filename.replace("ś", "s").replace("ź", "z").replace("ż", "z")
    filename = filename.replace("Ą", "A").replace("ć", "c").replace("ę", "e")
    filename = filename.replace("Ł", "L").replace("Ń", "N").replace("Ó", "O")
    filename = filename.replace("Ś", "S").replace("Ź", "Z").replace("Ż", "Z")
    return filename

#create a new excel file for each client
for client in unique_clients:
    new_workbook = openpyxl.Workbook()
    new_ws = new_workbook.active
        # headers
    for col_num, cell in enumerate(ws[1], start=1):
        if col_num < len(ws[1]): 
            new_ws.cell(row=1, column=col_num, value=cell.value)
        # Copy data for a given client only
    for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if row[0].value == client:
            for col_num, cell in enumerate(row, start=1):
                if col_num < len(row): 
                    new_ws.cell(row=row_num, column=col_num, value=cell.value)
        # delete empty cells       
    indexes_to_delete = []
    for i, row in enumerate(new_ws.iter_rows(), start=1):
        if all(cell.value is None for cell in row):
            indexes_to_delete.append(i)
    for index in reversed(indexes_to_delete):
        new_ws.delete_rows(index)

        #changing number format
    for row in new_ws.iter_rows(min_row=2, min_col=8, max_col=8):
        for cell in row:
            cell.number_format = '0.00%'
            if cell.value is not None:
                cell.value = float(cell.value)

      # Calculate the maximum text length in each column
      column_widths = []
      for row in new_ws:
          for i, cell in enumerate(row):
              if len(column_widths) > i:
                  if len(str(cell)) > column_widths[i]:
                      column_widths[i] = len(str(cell))
              else:
                  column_widths.append(len(str(cell)))

        #adjust column width
      for i, column_width in enumerate(column_widths, 1):
          column_letter = get_column_letter(i)
          new_ws.column_dimensions[column_letter].width = column_width
      for row in new_ws.iter_rows():
          for cell in row:
              cell.alignment = openpyxl.styles.Alignment(horizontal='left')     
      for cell in new_ws[1]:
          cell.font = bold_font
          cell.fill = grey_fill

      # Save excel file
      output_file_name = f'{client}.xlsx'
      output_file_name_sanitize = sanitize_filename(output_file_name)
      new_workbook.save(os.path.join(folder_path2, output_file_name_sanitize))
      print(f"File saved for client {client} as {output_file_name_sanitize}")
    
print("File splitting is complete.")

#Email sending part

import smtplib
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from openpyxl import load_workbook


#creating client-mail assignment
file_to_email = {}
file_to_email_xlsx = {}
for row in ws.iter_rows(min_row=2, values_only=True):
    key_pl = row[0]
    key2 = ".xlsx"
    value = row[9]
    if key_pl is not None and value is not None:
        key = sanitize_filename(key_pl)
        file_to_email[key] = value
        file_to_email_xlsx[key+key2] = value
    elif key_pl is not None and value is None:
        print("There is no email to send")

#login details
email = input("Enter your email login: ")
password = getpass.getpass("Password: ")
email2 = input("Enter the @ address from which the email is to be sent: ")
email3 = "example@gmail.com"

#connecting to email
try:
    # Connect to SMTP server
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(email, password)
    print("Login was successful.")
except smtplib.SMTPAuthenticationError:
    print("Login failed. Please check your login details.")

#email
for filename in os.listdir(folder_path2):
    if filename.endswith(".xlsx") and filename in file_to_email_xlsx:
        try:
            # create an email
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = file_to_email_xlsx[filename]
            msg['Cc'] = ""
            filename_without_extension = filename.split(".xlsx")[0]
            msg['Subject'] = f"Monitoring the implementation of Permanent Contracts {filename_without_extension}"
            body = f"Hello,\nw attached monitoring for {filename_without_extension}.The monitoring also includes placed orders.\n\nBest Regards\nOffice of Risk and Market Analysis"
            msg.attach(MIMEText(body, 'plain'))

            # attach files
            with open(os.path.join(folder_path2, filename), 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % filename)
            msg.attach(part)

            # Send e-mail
            server.sendmail(email, [file_to_email_xlsx[filename]], msg.as_string())
            print(f"Sent successfully {filename} do {file_to_email_xlsx[filename]}")
        except smtplib.SMTPRecipientsRefused:
            print(f"Error while sending {filename} do {file_to_email_xlsx[filename]}")
    else:
        print(f"No recipient found for {filename}")

# Close SMTP connection
server.quit()



