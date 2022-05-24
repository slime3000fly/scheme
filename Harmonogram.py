import xlrd
import smtplib, ssl
from datetime import date,timedelta

send_mail=0

workbook = xlrd.open_workbook(r"C:\Users\Piotr\Desktop\test.xls")
sheet = workbook.sheet_by_name("Gowno")

today = date.today()

today_2 = today + timedelta(days = 1)

col_a = sheet.col_values(0, start_rowx=1, end_rowx=None) #reading date from 1 column

count = 1

for xl_date in col_a:
    #changing xl format to python format date
    datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
    date_object = datetime_date.date()
    string_object= datetime_date.isoformat()
    if(date_object==today):
        item = sheet.cell_value(count, 1) #reading date from excel cell sgin to this date
        extra_message = "TODAY TEST of " + item + string_object
        send_mail = 1
        break;
    if(date_object==today_2):
        item = sheet.cell_value(count, 1)  # reading date from excel cell sgin to this date
        extra_message = "tommorow test of " + item + " " + string_object
        send_mail = 1
        break;
    count=count+1

if send_mail == 1:
    #date for email
    sender_email = "sendermail@gmail.com"
    receiver_email = "receiveremail@gmail.com"
    message = """\n
    Subject: Hi there

    """
    message = message + extra_message

    #sending email
    port = 465  # For SSL
    password = input("Type your password and press enter: ")

    # Create a secure SSL context
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login("bpm2914@gmail.com", password)
        server.sendmail(sender_email, receiver_email, message)
