# Monitoring and Email Automation Script

## Description
This Python script automates the process of monitoring and sending emails with Excel attachments. The script performs the following steps:
1. Creates a folder for today's date.
2. Loads an Excel file containing client data.
3. Creates separate Excel files for each unique client.
4. Sends emails with the attached Excel files to the respective clients.

## Requirements
- Python 3.x
- Libraries:
  - `os`
  - `openpyxl`
  - `datetime`
  - `smtplib`
  - `getpass`
  - `email`

## Installation
To install the required libraries, use the following command:
```bash
pip install openpyxl

Usage
Ensure you have an Excel file named Monitoring-SC.xlsx in the appropriate folder.
Run the script:
python monitoring_script.py

Enter your email login details when prompted.

## Code Structure
Importing Libraries: Importing necessary libraries for working with file paths, Excel files, and emails.
File Handling Part:
Creating a folder for today’s date.
Loading the Excel file and creating a list of unique clients.
Creating separate Excel files for each client.
Email Sending Part:
Creating client-email assignments.
Logging into the SMTP server.
Sending emails with the attached Excel files.

Author
Kacper Chojnacki

