import smtplib
import os
import cgi
import openpyxl

class Contact():
    def __init__(self, company_name, fname, lname, email):
        self.company = company_name
        self.fname = fname
        self.lname = lname
        self.email = email

    



def parseExcelSheet(sheet):
    # Load the workbook
    workbook = openpyxl.load_workbook(sheet)

    # Select the worksheet by name
    worksheet = workbook['Sheet1']

    # Access a cell value
    cell_value = worksheet['A1'].value
    print("TESTING CELL A1: ",cell_value)


    # access each cell and create a class that contains the contact info
    # Access a range of cells
    contact_list = []
    for row in worksheet:
        c = Contact()
        for cell in row:
            
            print(cell.value)


def sendEmail():
    # email address and password of the sender
    email_address = 'example@gmail.com'
    email_password = 'google-idAppCode-goes-here'

    # email address of the recipient
    to_email_address = 'recipient@gmail.com'

    FROM = email_address
    TO = to_email_address

    subject = "Subject Test"
    body = "I messed up the first test. Here we go. Body test. Please standby."
    msg = f"Subject: {subject} \n\n{body}"

    # create SMTP session
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo() #re-identify as an encrypted connection
        server.login(email_address, email_password)
        server.sendmail(email_address, to_email_address, msg)
        server.quit()

