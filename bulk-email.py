import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import os
import argparse
import re

def sendEmail(emailBody, resumePath, userEmail, userPassword, emailSubject, receiverEmail, smtpServer = '', smtpServerPort = ''):
  

    print("email", userEmail)
    print("password", userPassword)
    
    # Config
    smptp_server = smtpServer or 'smtp.gmail.com'
    smtp_port = smtpServerPort or 587
    sender_email = userEmail
    sender_pass = userPassword
    receiver_email = receiverEmail # receiverEmail

    # Creating the email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = emailSubject

    msg.attach(MIMEText(emailBody, 'plain'))

    # attach pdf
    with open(resumePath, 'rb') as pdf_file:
        pdf_attachment = MIMEApplication(pdf_file.read(), _subtype='pdf')
        pdf_attachment.add_header('Content-Disposition', 'attachment', filename=resumePath)
        msg.attach(pdf_attachment)

    # sending email
    try:
        with smtplib.SMTP(smptp_server, smtp_port) as server:
            server.starttls() # starting the server
            server.login(sender_email, sender_pass) # logging in to the server
            server.send_message(msg)
        pass
    except Exception as e:
        print(f"Failed to send email {e}")

def getEmailBody(path = ''):

    message = ""
    with open(path, 'r') as file:
        message = file.read()
    
    if len(message) == 0:
        raise Exception("Message cannot be empty, Please check the message body again")
    return message

def getTypeOfFile(path):
    fileExtension = os.path.splitext(path)
    if fileExtension[-1] not in [".xlsx", ".csv"]:
        raise Exception("Not a valid file type. Provide either .xlsx or .csv")

    return fileExtension[-1]

# used to validate the file path provided
def validateFilePaths(resumePath, sheetPath, messagePath):
    errors = []

    if not os.path.exists(resumePath):
        errors.append(resumePath)
        # raise Exception()

    if not os.path.exists(sheetPath):
        errors.append(sheetPath)
    
    if not os.path.exists(messagePath):
        errors.append(messagePath)

    
    if len(errors):
        raise Exception("Following Paths are not valid: ", errors)
    else:
        print("Valid Paths")

def validateEmail(email):
    valid = re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', email)

    if valid is None:
        raise Exception("Please Provide a valid email address: ", email)

def parseColumnName(df):
    df.columns = df.columns.str.lower()

    emailCol = False
    # Checking for column with "email" name
    for col in df.columns:
        if col == "email":
            emailCol = True

    if not emailCol:
        raise Exception("No \"email\" column present in the Excel Sheet. Please check and try again")


def readMessageAndSendEmail(resumePath, sheetPath, messagePath, userEmail, userPassword, emailSubject, smtpServer, smtpServerPort):
    
    typeOfFile = getTypeOfFile(sheetPath)

    df = None
    if typeOfFile.lower() == ".xlsx":
        df = pd.read_excel(sheetPath)
    elif typeOfFile.lower() == ".csv":
        df = pd.read_csv(sheetPath)

    emailBody = getEmailBody(messagePath)

    parseColumnName(df)

    for _, val in df.iterrows():
        msg = emailBody
        for col in df.columns:
            msg = msg.replace("{{" + col + "}}", str(val[col]))

        sendEmail(emailBody, resumePath, userEmail, userPassword, emailSubject, val["email"], smtpServer, smtpServerPort)
    

if __name__ == "__main__":
    print("Command Line Arguments")
    parser = argparse.ArgumentParser(description="Python program to help you bulk email people")

    # Path to different files
    parser.add_argument("-r", "--resume", required=True, help="Path to resume")
    parser.add_argument("-l", "--list", required=True, help="Path to xlsx/csv fille contaning list of people with their email")
    parser.add_argument("-m", "--message", required=True, help="Path to file containing the message")

    # Email subject
    parser.add_argument("--subject", help="Subject line of the email")
    
    # User credentaials
    parser.add_argument("-u", "--username", required=True, help="Your email address")
    parser.add_argument("-p", "--password", required=True, help="Your password")

    # SMTP server config
    parser.add_argument("-s", "--server",help="SMTP server")
    parser.add_argument("-po", "--port", help="SMTP Server port")


    args = parser.parse_args()

    resumePath = args.resume # path to resume
    sheetPath = args.list # path to excel sheet
    messagePath = args.message # path to message body

    email = args.username # User email
    password = args. password # User password
    
    smtpServer = args.server # SMTP server
    smtpPort = args.port # SMTP Server port

    subject = args.subject # email subject

    # Validations
    validateFilePaths(resumePath, sheetPath, messagePath)
    validateEmail(email)


    readMessageAndSendEmail(resumePath, sheetPath, messagePath, email, password, subject, smtpServer, smtpPort)
    

    
    