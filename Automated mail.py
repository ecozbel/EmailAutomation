import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


'''
Change these to your credentials and name
'''
your_name = "Ceyhun Ozbel"
your_email = "ecozbel@gmail.com"
your_password = "ceyhun9678"

# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# Read the file
email_list = pd.read_excel("EmailList.xlsx")

# Get all the Names, Email Addreses, Subjects and Messages
all_names = email_list['Name']
all_emails = email_list['Email']
all_subjects = email_list['Subject']
all_messages = email_list['Message']

# Loop through the emails
for idx in range(len(all_emails)):

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    subject = all_subjects[idx]
    message = all_messages[idx]
    html = open('IIFF Haziran.html')



    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message + " " + name ))


    def py_mail(SUBJECT, BODY, TO, FROM):
        """With this function we send out our HTML email"""

        # Create message container - the correct MIME type is multipart/alternative here!
        MESSAGE = MIMEMultipart('alternative')
        MESSAGE['subject'] = all_subjects[idx]
        MESSAGE['To'] = all_messages[idx]
        MESSAGE['From'] = your_email
        MESSAGE.preamble = """
    Your mail reader does not support the report format.
    Please visit us online!"""
        html = open('IIFF Haziran.html')

        # Record the MIME type text/html.
        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        MESSAGE.attach(html)


    # In the email field, you can add multiple other emails if you want
    # all of them to receive the same text

    try:
        server.sendmail(your_email, [email], MESSAGE.as_string())
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

# Close the smtp server
server.close()