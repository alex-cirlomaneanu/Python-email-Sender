# Alexandru Cîrlomăneanu 2021

# Python program to send email based on a certain format to a list of
# addresses from a Excel file.

# Imports
import pandas
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# My email and password
my_email = "bruce@wayneindustires.com"
my_password = "batmobile98"

# Reading the emails from spreadsheet
email_list = pandas.read_excel('C:\docs\emailsender\emailsList.xlsx')

# Getting the names and the emails from the list
names = email_list['NAME']
emails = email_list['EMAIL']

# Gmail connection
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(my_email, my_password)

# Send a message to every email address in list
for i in range(len(emails)):

    # For every record get the name and the email address

    name = names[i]
    email = emails[i]

    # The text to be send
    body = ("Dear " +
            name +
            ", \n\n My name is Bruce and I want you" +
            " to know I am Batman" +
            ".\n\n Best regards,\n The Dark Knight")

    # The name of the sender
    sender = "Bruce Wayne"

    # Create the message
    msg = MIMEMultipart()
    # Adding the subject, sender and reciever
    msg['Subject'] = "Test"
    msg['From'] = sender
    msg['To'] = email

    # Attach the text body
    msg.attach(MIMEText(body, 'plain'))

    message = msg.as_string()

    # Sending the email
    server.sendmail(my_email, [email], message)

# Closing the smtp server
server.close()
