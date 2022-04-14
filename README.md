# Automated-Birthday-Email-Sender
This program sends out automated emails. 

The program uses data that is saved in an Excel file (data.xlsx) in the format (Name, Birthday, Message, Year, Email, Mobile Number). It imports this infomation into an EmailMessage object and uses SMTP_SSL server to send out a custom birthday message to people whose information in located in the aformentioned Excel file.

You may have to register an App password to allow the script to access you GMAIL account. The following link is to a tutorial on how to do so.
https://support.google.com/accounts/answer/185833?hl=en

Make sure to goto the Python file and fill out the locations where the code says "YOUR EMAIL HERE" and "YOUR PASSWORD HERE"

NOTE:
This program is designed to function with GMAIL. Other services may not work with this script.

Modules used by this program:
pandas
datetime
time
smtplib
email.message
os


