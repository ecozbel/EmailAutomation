import yagmail
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


email_list = pd.read_excel("EmailList_test.xlsx") #Add  path to excel email list here
receiver = "adress" #Add receiver here
body = "Hello there from Yagmail"
filename = "IIFF_Haziran.html"
yagmail.register('adress','password') #Add gmail login details here
all_names = email_list['Name']
all_emails = email_list['Email']
all_subjects = email_list['Subject']
all_messages = email_list['Message']


msg = MIMEMultipart('mixed')
report_file = open(filename)
html = report_file.read()

text = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"


part1 = MIMEText(html, 'html')
part2 = MIMEText(html, 'html')
charset="ISO-8859-1"
part1.set_charset(charset)
msg.attach(part1)
msg.attach(part2)


for idx in range(len(all_emails)):

    name = all_names[idx]
    email = all_emails[idx]
    subject1 = all_subjects[idx]
    body = '<font face="Times New Roman, monospace">' + "Dear " + name + ', \n' + "Istanbul Exporters’ Associations (İİB) would like to invite you to take part in Furniture Buyers Programme between 22-27 June 2021. During this mission, participants will visit Istanbul Furniture Fair and will be attending B2B matchmaking events to enhance trade relations, exchange of knowledge, networking and possible business co-operations in the furniture sector. The biggest furniture show throughout Europe with 23 Halls in 260.000 sqm; Istanbul Furniture Fair provides new business opportunities for its exhibitors and wide range of variety of products for visitors."+ '\n'+'\n'+"Don’t miss the opportunity to meet more than 400 furniture companies of Turkey which have a wide range of product variety ( Modern furniture for sophisticated comfort, inspiration for lifestyle-oriented interiors: suites, armchairs, divans, standalone sofas, convertible couches, tables, bedroom and dining room furniture, clever furniture for young lifestyles, furnitures for babies, kids and youths, office and call centre furnitures, home settings focuses on ready-to-go furniture: suites, armchairs, divans, mattresses and sleep systems, box-spring beds)."+"Please let us know if you would be interested in attending this buying mission." + '</font>'
    ending = '\n' +'\n' + "Kind regards, " + '\n' + 'Ceyhun Ozbel' + '\n'
    yag = yagmail.SMTP("ecozbel@gmail.com")
    yag.send(
        to=email,
        subject=subject1,
        contents=body + ending + msg.as_string(),
    )