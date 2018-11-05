"""
Funzione da far runnare all'avvio del pc
Controlla se ci sono compleanni oggi e nel caso invia la mail
- Per controllare di non reinviare la stessa mail lo stesso giorno
check MeisterTask
"""

# Send an HTML email with an embedded image and a plain text
# message for email clients that don't want to display the HTML

import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import smtplib
import random
import json
import time
from pathlib import Path

from user import User


'''
FUNCTIONS
'''


def sendEmail():
    # Define these once; use them twice!
    strFrom = 'lorenzo.30000@gmail.com'
    strTo = 'lorenzo.30000@gmail.com'

    # Create the root message and fill in the from, to, and subject headers
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = 'Test message'
    msgRoot['From'] = strFrom
    msgRoot['To'] = strTo
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    msgText = MIMEText('This is the alternative plain text message.')
    msgAlternative.attach(msgText)

    # We reference the image in the IMG SRC attribute by the ID we give it below
    emailBody = open("emailBody.html", "r").read()

    msgText = MIMEText(emailBody, 'html')
    msgAlternative.attach(msgText)

    # This example assumes the image is in the current directory
    fp = open('test.JPG', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    # msgRoot.attach(msgImage)

    # Send the email (this example assumes SMTP authentication is required)

    server = smtplib.SMTP('smtp.gmail.com:587')

    server.ehlo()
    server.starttls()
    server.login('lorenzo.30000@gmail.com', 'Kawawa24')
    server.sendmail(strFrom, strTo, msgRoot.as_string())
    server.quit()


def createGroup(group):
    # group can be 2 or 3 long
    print(
        f"{group[0].firstName} is gonna have lunch with {group[1].firstName}", end='')
    if (len(group) == 3):
        print(f" and with {group[2].firstName}")
    print("")


def createGroups():
    # EXCEL FILE

    # Assign spreadsheet filename to `file`
    file = 'excelData.xlsx'

    # Load spreadsheet
    xl = pd.ExcelFile(file)

    # Load a sheet into a DataFrame by name: df1
    df1 = xl.parse('Sheet1')
    # print(df1)

    dfTable = df1.iloc[:, 1:]
    # print(dfTable)

    users = []
    for _, userRow in dfTable.iterrows():
        users.append(User(
            userRow['Nome'],
            userRow['Cognome'],
            userRow['Email']))

    while True:
        # check if it is the last group
        if (len(users) <= 3):
            createGroup(users)
            print(f"\nThe end\n")
            break

        # get two random indexes
        firstIndex = random.randint(0, len(users) - 1)
        while True:
            secondIndex = random.randint(0, len(users) - 1)
            if (secondIndex != firstIndex):
                break

        firstUser = users[firstIndex]
        secondUser = users[secondIndex]

        users.remove(firstUser)
        users.remove(secondUser)

        group = [firstUser, secondUser]
        createGroup(group)


def checkLastRun():
    # if file does not exist
    my_file = Path("data_file.json")
    if not my_file.is_file():

        ts = int(time.time() * 1000)

        data = {"timestamp": ts}

        with open("data_file.json", "w") as write_file:
            json.dump(data, write_file)

    # read last run ts
    with open("data_file.json", "r") as read_file:
        data = json.load(read_file)

    prev_ts = data['timestamp']

    # timestamp of one day in ms
    ts_day = 24 * 60 * 60 * 1000

    # timestamp from 00:00 of today in ms
    delta = int(time.time() * 1000) % ts_day

    # timestamp of tomorrow at 00:00 in ms
    ts_tomorrow = int(time.time() * 1000) - delta + ts_day

    # TODO è dipendente da quando lei accende il pc

    if (prev_ts < ts_tomorrow):
        return False

    data = {"timestamp": ts}

    with open("data_file.json", "w") as write_file:
        json.dump(data, write_file)

    return True


def main():
    print("Checking if run or not")
    if (checkLastRun()):
        createGroups()


if __name__ == "__main__":
    main()
