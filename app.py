# Insert your python code here
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



# 1
required_files = ['ProductData.xlsx', 'MasterInvoice.pdf',
                  'Order1.xlsx', 'Order2.xlsx',
                  'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input("Please enter the directory where the files are located.")

    if not directory.endswith(("\\", "/")):
        directory = directory + "\\"

    files_found = True

    for file in required_files:
        if not os.path.exists(directory + file) == True:
            files_found = False

    if files_found == True:
        break
    else:
        show_message("Error", "The directory is missing required files. Please try again.")

print(directory)

# 2
df_descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')
product_descriptions = {}
for index,row in df_descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df_prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')
product_prices = {}
for index,row in df_prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_descriptions)
print(product_prices)

#3
import datetime
 
 
def get_order_data (invoice_num, filename, directory):
    order_data = {}
    template = openpyxl.load_workbook(directory + filename, data_only= True)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value #this is retrevieing customer info
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    Unit1 = selected_sheet['B24'].value
    Quantity1 = selected_sheet['C24'].value
    Unit2 = selected_sheet['B25'].value
    Quantity2 = selected_sheet['C25'].value
    Unit3 = selected_sheet['B26'].value
    Quantity3 = selected_sheet['C26'].value
    
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date


    if Unit3 is not None:
        order_data['Total'] = (product_prices. get(Unit1) * Quantity1) + (product_prices.get(Unit2) * Quantity2) + (product_prices.get(Unit3) * Quantity3)
        order_data['Total'] = f"${order_data["Total"]:,.2f}"
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19']. value}"
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f"${order_data['Price0']:,.2f}"
        order_data['Quantity0'] = str(Quantity1)
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f"${order_data['Amount0']:,.2f}"
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1'] = product_prices.get(Unit2)
        order_data['Price1'] = f"${order_data['Price1']:,.2f}"
        order_data['Quantity1'] = str(Quantity2)
        order_data['Amount1'] = product_prices.get(Unit2) * Quantity2
        order_data['Amount1'] = f"${order_data['Amount1']:,.2f}"
        order_data['Description1'] = product_descriptions.get(Unit2)
        order_data['Price2'] = product_prices.get(Unit3)
        order_data['Price2'] = f"{order_data['Price2']:,.2f}"
        order_data['Quantity2'] = str(Quantity3)
        order_data['Amount2'] = product_prices.get(Unit3) * Quantity3
        order_data['Amount2'] = f"${order_data['Amount2']:,.2f}"
        order_data['Description2'] = product_descriptions.get(Unit3)

    elif Unit2 is not None:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1) + ((product_prices.get(Unit2) * Quantity2))
        order_data['Total'] = f"${order_data["Total"]:,.2f}"
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}"
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f"${order_data['Price0']:,.2f}"
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f"${order_data['Amount0']:,.2f}"
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1'] = product_prices.get(Unit2)
        order_data['Price1'] = f"${order_data['Price1']:,.2f}"
        order_data['Quantity1'] = Quantity2
        order_data['Amount1'] = product_prices.get(Unit2) * Quantity2
        order_data['Amount1'] = f"${order_data['Amount1']:,.2f}"
        order_data['Description1'] = product_descriptions.get(Unit2)

    else:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1)
        order_data['Total'] = f"${order_data['Total']:,.2f}"
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18']. value}, {selected_sheet['C19'].value}"
        order_data['Price'] = product_prices.get(Unit1)
        order_data['Price'] = f"${order_data['Price']:,.2f}"
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] =  f"${order_data['Amount0']:,.2f}"
        order_data['Description0']= product_descriptions.get(Unit1)
    return order_data, name, email, date 


get_order_data(8515, 'Order1.xlsx', directory)

#4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)


order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
create_invoice(8515, order_data, directory)

#5
gpass = open('gmail_password.txt','w')
gpass.write('gzrmcwwqyjgeicij')
gpass.close()
import smtplib
from email.message import EmailMessage
gmail_user = "sydneyberg44@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date,directory):
    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
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

    path_to_file = directory +f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()
send_email(
    8515,
    'Thomas Holt',
    gmail_user,
    '11/13/2023',
    directory)

#6
filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data,directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num +1

#7





