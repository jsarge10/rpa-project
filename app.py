# Insert your python code here
# Question 1
# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    before_cut_directory = select_file("Select 'ProductData.xlsx' in the folder with 'MasterInvoice.pdf', 'Order.xlsx', etc.")
    directory = before_cut_directory.replace('ProductData.xlsx', '')
    try:
        count = 0
        for file in os.listdir(directory):
            if file in ["ProductData.xlsx", "MasterInvoice.pdf", "Order1.xlsx", "Order2.xlsx", "Order3.xlsx", "Order4.xlsx"]:
                count = count + 1
        if count == 6:
            break
        else:
            show_message('Error', 'Choose a working directory')
    except:
        show_message('Error', 'Invalid input; Choose a working directory')
# Question 2
# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')

for i, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')
for i, row in df1.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

product_descriptions, product_prices
# Question 3
# INSERT ANSWER FOR QUESTION 3 HERE
import datetime as dt

def get_order_data(invoice_num : int,filename : str,directory : str):
    wb = openpyxl.load_workbook(directory + filename)
    selected_sheet = wb['OrderForm']
    name = selected_sheet['C14'].value
    date = selected_sheet['C9'].value
    date = date.strftime('%m/%d/20%y')
    email = selected_sheet['C13'].value
    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = date
    update_dict['CustomerAddress1'] = selected_sheet['C16'].value
    update_dict['CustomerName'] = name
    update_dict['CustomerCompany'] = selected_sheet['C15'].value
    update_dict['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}'
    total = 0
    if not selected_sheet['B24'].value is None:
        price0 = product_prices[selected_sheet['B24'].value]
        update_dict['Price0'] = f'${price0:,.2f}'
        quantity0 = selected_sheet['C24'].value
        update_dict['Quantity0'] = str(quantity0)
        amount0 = price0 * quantity0
        update_dict['amount0'] = f'${amount0:,.2f}'
        update_dict['Description0'] = product_descriptions[selected_sheet['B24'].value]
        total += amount0
    if not selected_sheet['B25'].value is None:
        price1 = product_prices[selected_sheet['B25'].value]
        update_dict['Price1'] = f'${price1:,.2f}'
        quantity1 = selected_sheet['C25'].value
        update_dict['Quantity1'] = str(quantity1)
        amount1 = price1 * quantity1
        update_dict['amount1'] = f'${amount1:,.2f}'
        update_dict['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total += amount1
    if not selected_sheet['B26'].value is None:
        price2 = product_prices[selected_sheet['B26'].value]
        update_dict['Price2'] = f'${price2:,.2f}'
        quantity2 = selected_sheet['C26'].value
        update_dict['Quantity2'] = str(quantity2)
        amount2 = price2 * quantity2
        update_dict['Amount2'] = f'${amount2:,.2f}'
        update_dict['Description2'] = product_descriptions[selected_sheet['B26'].value]
        total += amount2
    if not selected_sheet['B27'].value is None:
        price3 = product_prices[selected_sheet['B27'].value]
        update_dict['Price3'] = f'${price3:,.2f}'
        quantity3 = selected_sheet['C27'].value
        update_dict['Quantity3'] = str(quantity3)
        amount3 = price3 * quantity3
        update_dict['Amount3'] = f'${amount3:,.2f}'
        update_dict['Description3'] = product_descriptions[selected_sheet['B27'].value]
        total += amount3
    if not selected_sheet['B28'].value is None:
        price4 = product_prices[selected_sheet['B28'].value]
        update_dict['Price4'] = f'${price4:,.2f}'
        quantity3 = selected_sheet['C28'].value
        update_dict['Quantity4'] = str(quantity4)
        amount4 = price4 * quantity4
        update_dict['Amount4'] = f'${amount4:,.2f}'
        update_dict['Description4'] = product_descriptions[selected_sheet['B28'].value]
        total += amount4
    update_dict['Total'] = f'${total:,.2f}'
    return update_dict, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)
    
# Question 4
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data,directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)
# Question 5
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    gmail_user = 'jacksargent110@gmail.com'
    gmail_password = open('/Users/jacksargent/Documents/Advanced Analytics/17 - Project 3 - Robotic Process Automation/gmail_password.txt').read()
    body = f"""Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84602
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf', message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, name, 'jacksargent110@gmail.com', date, directory)
# Question 6
# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
print(filenames)
for file in filenames:
    order = 8515
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(order, file, directory)
        create_invoice(order, order_data, directory)
        send_email(order, name, email, date, directory)
        order += 1
