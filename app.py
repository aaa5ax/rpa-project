import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory = single_input("What is the directory where your files reside?") # ASK ABOUT THIS, should it be allowed to include quotes

    # if """'""" in directory: # Remove Quotes if there are any
    #     direct1 = directory.split("'")
    #     directory = direct1[1]
    #     display(directory)

    direct2 = directory.split('/') # Check directory and add the / symbol
    if direct2[-1] != '':
        directory = directory + '/'

    directory_list = os.listdir(directory) # Check to make sure directory has the right files
    if ('ProductData.xlsx' in directory_list) and ('MasterInvoice.pdf' in directory_list) and ('Order1.xlsx' in directory_list) and ('Order2.xlsx' in directory_list) and ('Order3.xlsx' in directory_list) and ('Order4.xlsx' in directory_list):
        break
    else:
        show_message('Error', 'Re-enter a directory that contains ProductData.xlsx, MasterInvoice.pdf, Order1.xlsx, Order2.xlsx, Order3.xlsx, and Order4.xlsx')

descriptions = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

import datetime as dt

def get_order_data(invoice_num, filename, directory):
    order_form = openpyxl.load_workbook(directory+filename) # get basic customer info and date
    selected_sheet = order_form['OrderForm']
    customer_name = selected_sheet['C14'].value
    customer_email = selected_sheet['C13'].value
    order_date = selected_sheet['C9'].value
    
    update_dict = {} # Create update dictionary to fill out the pdf
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = str(order_date.month) + '/' + str(order_date.day) + '/' + str(order_date.year)
    update_dict['CustomerName'] = customer_name
    update_dict['CustomerCompany'] = selected_sheet['C15'].value
    update_dict['CustomerAddress1'] = selected_sheet['C16'].value
    update_dict['CustomerAddress2'] = selected_sheet['C17'].value + ', ' + selected_sheet['C18'].value + ' ' + str(selected_sheet['C19'].value)

    increment = 0 
    TOTAL_AMOUNT = 0
    while True:
        if increment <= 5 and selected_sheet[f'B{24+increment}'].value is not None:
            update_dict[f'Description{increment}'] = product_descriptions[selected_sheet[f'B{24+increment}'].value]
            update_dict[f'Price{increment}'] = f'${product_prices[selected_sheet[f'B{24+increment}'].value]:,.2f}'
            update_dict[f'Quantity{increment}'] = str(selected_sheet[f'C{24+increment}'].value)
            update_dict[f'Amount{increment}'] = f'${product_prices[selected_sheet[f'B{24+increment}'].value] * selected_sheet[f'C{24+increment}'].value:,.2f}'
            TOTAL_AMOUNT = TOTAL_AMOUNT + product_prices[selected_sheet[f'B{24+increment}'].value] * selected_sheet[f'C{24+increment}'].value
        else:
            break
        
        increment = increment + 1
            
    update_dict['Total'] = f'${TOTAL_AMOUNT:,.2f}'

    return update_dict, customer_name, customer_email, order_date


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
# print(order_data)
# print(name)
# print(email)
# print(date)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

create_invoice(invoice_num=8515, order_data=order_data, directory=directory)

#added up here instead of below
gmail_password_path = select_file('Please select the gmail password file')

def send_email(invoice_num, name, email, date, directory):
    import smtplib
    from email.message import EmailMessage

    # Enter your gmail user name and password

    gmail_user = "abrahamnielson555@gmail.com"

    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [f"{email}"]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name

    body = f"""Dear {name},
    
Attached is your invoice dated {date.month}/{date.day}/{date.year}.

Thank you for your business.

Python Tech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'/Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

# send_email(invoice_num=8515, name=name, email=email, date=date, directory=directory)

filenames = os.listdir(directory)
filenames = sorted(filenames)

use_invoice_no = 8515
for f, file in enumerate(filenames):
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(use_invoice_no, file, directory)
        create_invoice(use_invoice_no, order_data=order_data, directory=directory)
        send_email(use_invoice_no, name=name, email=email, date=date, directory=directory)
        use_invoice_no = use_invoice_no + 1