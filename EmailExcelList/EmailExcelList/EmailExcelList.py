# Author: Keith Hill
# Date: 5/20/2017
#
# This program accepts a config file and a CSV file and generates emails to every email address with a configured messages with included variables

import os
import configparser 
import csv
import smtplib
import sys

myPath = os.path.dirname(os.path.abspath(__file__))


# check that the email file exists
try :
    open(myPath+"\emailList.csv")

except :
    print("Error: expected file emailList.csv could not be read \n")
    raise



print("email file found \n")

# load the config file
try :
    config = configparser.ConfigParser()
    config.read(myPath +"\config.ini")
    svr = config.get('email','svr')
    user = config.get('email','user')
    passwd = config.get('email','passwd')
    fromaddr = config.get('email','fromaddr')
    subject = config.get('email','subject')
    emailBody= config.get('email','body')
    emailBody = str(emailBody)
    


except :
    print("Error reading the config file \n")
    raise



print("config file loaded \n")

# open the file
with open(myPath+"\emailList.csv") as csvfile : 
    reader = csv.DictReader(csvfile)
    

    # count the number of variables in the file.  
    variableCount = 0
    done = False
    while not done :
        try :
            temp = reader.fieldnames[variableCount]
            variableCount +=1
        except :
            variableCount -=1
            done = True

    email = reader.fieldnames[0]
    
    # spin up email server
    server = smtplib.SMTP(svr)
    server.ehlo()
    server.starttls()
    server.login(user,passwd)

    # for each row, send email
    for row in reader:
        count = 1

        emailBody= config.get('email','body')
        emailBody = str(emailBody)
        
        # compile email body by inserting the variables
        while count <= variableCount :
            emailBody = emailBody.replace("||"+str(reader.fieldnames[count])+"||",str(row[reader.fieldnames[count]]))
            count +=1


        # compose the message
        msg = "\r\n".join([
            "From: " + fromaddr,
            "To: " + row[email],
            "Subject: " + subject,
            "MIME-Version: 1.0",
            "Content-type: text/html",
                "",
                emailBody
                ])
        
        # send the email
        print ("sending email to " + row[email])
        server.sendmail(fromaddr, row[email], msg)

    # close email server
    server.quit()
    print ("all emails have been sent")
    os.system("pause")

