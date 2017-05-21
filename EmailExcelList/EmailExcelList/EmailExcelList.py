import os
import configparser 
import csv
import smtplib

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
    message01 = config.get('email','message01')
    message02 = config.get('email','message02')

    message01 = message01.replace("\\n","\n")
    message02 = message02.replace("\\n","\n")

except :
    print("Error reading the config file \n")
    raise

print("config file loaded \n")

# open the file and go line by line
with open(myPath+"\emailList.csv") as csvfile : 
    reader = csv.DictReader(csvfile)
    
    email = reader.fieldnames[0]
    variable = reader.fieldnames[1]

    # spin up email server
    server = smtplib.SMTP(svr)
    server.ehlo()
    server.starttls()
    server.login(user,passwd)

    # for each row, send email
    for row in reader:

        # compose the message
        msg = "\r\n".join([
            "From: " + fromaddr,
            "To: " + row[email],
            "Subject: " + subject,
            "MIME-Version: 1.0",
            "Content-type: text/html",
                "",
                message01 + " "  + row[variable] + " " + message02
                ])
        
        # send the email
        print ("sending email to " + row[email])
        server.sendmail(fromaddr, row[email], msg)

    # close email server
    server.quit()
    print ("all emails have been sent")
    os.system("pause")

