import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Step 1: Scrape the website for email addresses
def extract_emails(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    email_pattern = r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}'
    emails = re.findall(email_pattern, str(soup))
    return emails

websites = [
    'https://www.powerfilmsolar.com/',
    'https://solarstik.com/',
    # Add more websites as needed
]

all_emails = []
for url in websites:
    emails = extract_emails(url)
    all_emails.extend(emails)

# Step 2: Create an Excel workbook and add the emails
workbook = Workbook()
sheet = workbook.active
sheet.title = 'Emails'

# Add a header row
sheet['A1'] = 'Email'

# Add emails to the sheet
for index, email in enumerate(all_emails, start=2):  # Assuming the first row is for headers
    sheet.cell(row=index, column=1, value=email)

# Step 3: Save the workbook to an Excel file
output_file = 'emails.xlsx'
workbook.save(output_file)

print(f"{len(all_emails)} emails found and added to the Excel file '{output_file}'.")
