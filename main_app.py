import openpyxl
import datetime
from datetime import timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


sender_mail = 'sender@gmail.com'
password = 'your_password'

recipient_mail = 'recipient@gmail.com'





def find_task(file):
    # open xls file, find value
    # path parameter should point xls file
    xls = openpyxl.load_workbook(filename=file)
    sheet = xls.get_sheet_by_name('Arkusz1')
    delta = 5
    date_min = datetime.datetime.now() + timedelta(days=-delta)
    date_max = datetime.datetime.now() + timedelta(days=delta)
    for row_index in range(1, sheet.max_row):
        if sheet.cell(row=row_index, column=2).value >= date_min and sheet.cell(row=row_index, column=2).value <= date_max:
            if not sheet.cell(row=row_index, column=3).value == 1:
                task = sheet.cell(row=row_index, column=1).value
                print(task)
                send_message(task)
                sheet.cell(row=row_index, column=3).value = 1
    try:
        xls.save(filename=file)
    except PermissionError:
        print("I can't save it, check if file is closed")





def send_message(task):
    msg = MIMEMultipart()
    msg['From'] = sender_mail
    msg['To'] = recipient_mail
    msg['Subject'] = 'Task reminder'
    message = """
    Hi!
    I found that you have a new task:
    {}

    Thank You!
    """.format(task)
    msg.attach(MIMEText(message))

    mailserver = smtplib.SMTP('smtp.gmail.com', 587)
    # identify ourselves to smtp gmail client
    mailserver.ehlo()
    # secure our email with tls encryption
    mailserver.starttls()
    # re-identify ourselves as an encrypted connection
    mailserver.ehlo()
    mailserver.login(sender_mail, password)

    mailserver.sendmail(sender_mail, recipient_mail, msg.as_string())

    mailserver.quit()


path = "test.xlsx"
my_task = find_task(path)

