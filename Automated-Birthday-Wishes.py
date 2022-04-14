## IMPORTING NEEDED LIBRARIES
import pandas as pd ## reading excel files
import datetime     ## for time and year
import smtplib      ## connection between gmail and python script
from email.message import EmailMessage ## sending emails
import os           ## interacting with the system
import time

## FUNCTION THAT SENDS EMAIL
"""
This function takes the receiver email, subject, and message as arguments
Then creates EmailMessage class populated with argument data
"""
def sendEmail(to, sub, message):
    print(f'Email to {to} \nsend with subject: {sub}\n message: {message}')
    ## CREATION OF EMAIL MESSAGE OBJECT
    email = EmailMessage()
    email['From'] = 'YOUR EMAIL HERE'
    email['To'] = f'{to}'
    email['Subject'] = f'{sub}'
    email.set_content(f'{message}')
    
    ## SET UP CONNECTION BETWEEN GMAIL AND PYTHON SCRIPT
    start = time.time()
    try:
        smtp_ssl = smtplib.SMTP_SSL(host='smtp.gmail.com', port=465)
    except Exception as e:
        print('ErrorType: {}, Error: {}'.format(type(e).__name__, e))
        smtp_ssl = None
    
    print("Connection Object: {}".format(smtp_ssl))
    print('Total Time Taken: {:,.2f} Seconds'.format(time.time() - start))

    print("\nLogging In.....")
    resp_code, response = smtp_ssl.login(user="YOUR EMAIL HERE", password="YOUR APP PASSWORD HERE")
    print("Response Code : {}".format(resp_code))
    print("Response      : {}".format(response.decode()))

    print("\nSending Mail.......")
    response = smtp_ssl.send_message(email)

    print("List of Failed Recipients : {}".format(response))

    print('\nLogging Out....')
    resp_code, response = smtp_ssl.quit()

    print("Response Code : {}".format(resp_code))
    print("Response      : {}".format(response))

if __name__ == "__main__":
  ## READING THE EXCEL FILE USING PANDAS
    df = pd.read_excel("data.xlsx")
    print(df)

    today = datetime.datetime.now() .strftime("%d-%m")
    
    update = []
    yearnow = datetime.datetime.now().strftime("%Y")

    for index, item in df.iterrows():
        ## FETCHING YEAR COLUMN FOR COMPARISON
        """
        This is to make sure the message is only sent once on someone's birthday.
        Using the datetime module we find the current year, and if it is equal to the
        value in the Year column of the Excel file then no email will be sent.
        """
        bday = item['Birthday'].strftime("%d-%m") ##wishing time  from excel file
       
        if(bday == today) and yearnow not in str(item["Year"]):  ## birthday data ==  today date and birthday year is not equal to current year
            sendEmail(item['Email'] ,"Happy Birthday "+item["Name"], item['Message']) ## pass arguments to the send email funciton and call it
            update.append(index) ## update the index by one ## we need to check the whole records
    for i in update:
        yr = df.loc[i, 'Year'] ## update the year by one
      
        df.loc[i,'Year'] = f"{yr}, {yearnow}"  

    df.to_excel("data.xlsx", index=False)  ## convert the df to excel file and save it

