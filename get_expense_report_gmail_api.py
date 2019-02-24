from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import datetime
import calendar
import base64
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pytz
from tzlocal import get_localzone

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
EPOCH_DATE_STR = "01/01/1970"
EPOCH_MS = 86400000
SECONDS_IN_A_DAY= 86400

def ListMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except Exception as e:
    print('An error occurred: ' + e)
    
def extract_amount_string(message):
    #matchObj = re.search(r'(₹|Rs|INR)(\s|\.)*[0-9]+,*[0-9]*\.*[0-9]*', re.UNICODE, message)
    matchObj = re.search(r'(&#x20B9;|Rs|INR)(&nbsp;|\s|\.)*[0-9]+,*[0-9]*\.*[0-9]*\s+', message)
    if matchObj is not None and matchObj.group(0):
        return matchObj.group(0)
    else:
        return None
    
def extract_amount(message):
    matchObj = re.findall(r'([0-9]+,*[0-9]*\.*[0-9]*)', message)
    groupSize = len(matchObj)
    
    if matchObj is not None and matchObj[0]:
        if groupSize > 1:
            return matchObj[groupSize - 1]
        else:
            return matchObj[0]
    else:
        return None

def get_amount_spent(msg):
    payload = msg['payload'] # get payload of the message 
    if 'data' in payload['body']:
        messageBody = payload['body']['data']
        decodedMessageBody = base64.urlsafe_b64decode(messageBody).decode('utf-8')
        amountStr = extract_amount_string(decodedMessageBody)
        amount = extract_amount(amountStr)
        print(amount)
        if amount is not None:
            return float(amount.replace(",", ""))

def send_mail(total_spendings, typeOfReport, timeVsDebitTransaction):
    gmail_user = 'sender@gmail.com'  
    gmail_password = 'sender_pass'
    
    sent_from = gmail_user  
    to = "reciever@gmail.com" 
    subject = 'Daily Expense Summary'  
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = gmail_user
    msg['To'] = to
    
    text = "Hi!!!, \nYou can only view the expense report in html format.\nThanks,\nShofees Expense Tracker Bot"
    html = """<!DOCTYPE html>
            <html lang="en">
            <head>
              <title>Expense Tracker</title>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1">
              <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css">
              <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
              <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/js/bootstrap.min.js"></script>
              <style>
              table {
              font-family: arial, sans-serif;
              border-collapse: collapse;
              width: 40%;
            }
            
            th {
              background-color: #dddddd;
              border: 1px solid #000000;
              text-align: left;
              padding: 8px;
            }
            td {
              border: 1px solid #000000;
              text-align: left;
              padding: 8px;
            }
              </style>
            </head>
            <body>
            <div>
            <h3>Dear Shofee,</h3>
            </div>
            <h4>Your """ + typeOfReport + """ spending is : """ + str(total_spendings) + """</h4>
            <h5>Detailed analysis of """ + typeOfReport + """ expense is as below:</h5>
            <table>
            <thead >
            <tr><th>Time</th>
            <th>Amount</th>
            </tr>
            </thead>
            <tbody>"""
            
    for key, value in timeVsDebitTransaction.items():
        html += "<tr><td>" + str(key) + "</td><td>" + str(value) + "</td></tr>"
        
    html += "<tr><th>Total</th><th>Rs." + str(total_spendings) + "</th></tr>"
    html += """</tbody>
                </table>
                <br/>
                Thanks,
                <br/>
                Shofees Expense Tracker Bot
                </body>
                </html>
            """
    
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    
    msg.attach(part1)
    msg.attach(part2)
    
    try:  
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, msg.as_string())
        server.quit()
    
        print('Email sent!')
    except Exception as e: 
        print('Something went wrong...')
        print(e)  

def process_messages(messages, service):
    todaysDate = datetime.datetime.now().date()
    currentDayOfWeek = datetime.datetime.today().weekday()
    currentDayOfMonth = datetime.datetime.today().day
    lastDayOfPreviousMonth = calendar.monthrange(datetime.datetime.today().year, datetime.datetime.today().month - 1)[1]
    typeOfReport = "Daily"
    timeVsDebitTransaction = dict()
    total_spendings = 0.0
    
    if not messages:
        print("No messages found.")
    else:
        print("Message snippets:")
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            epochDate = datetime.datetime.strptime(EPOCH_DATE_STR, "%m/%d/%Y")
            currentDate = epochDate + datetime.timedelta(days=int(msg['internalDate'])/EPOCH_MS)
            
            #for local testing
            tz = pytz.timezone("Asia/Kolkata")
            if tz == get_localzone():
                currentDate += datetime.timedelta(seconds=19800)

            dateDifference = todaysDate - currentDate.date()
            if dateDifference.days == 0: #daily alerts
                amount = get_amount_spent(msg)
            elif currentDayOfWeek == 6 and dateDifference.days <= 6: #weekly alerts
                amount = get_amount_spent(msg)
                typeOfReport = "Weekly"
            elif currentDayOfMonth == 1 and dateDifference.days <= (lastDayOfPreviousMonth - 1): #monthly alerts
                amount = get_amount_spent(msg)
                typeOfReport = "Monthly"
            else:
                break
            if amount is not None:
                timeVsDebitTransaction[currentDate] = amount
                total_spendings += amount
        total_spendings = format(total_spendings, '.2f')
        print(total_spendings)
        send_mail(total_spendings, typeOfReport, timeVsDebitTransaction)
        if typeOfReport != "Daily":
            print("Retrieving daily record...")
            timeVsDebitTransactionDaily = dict()
            total_spendings = 0.0
            for currTime, amount in timeVsDebitTransaction.items():
                diff = todaysDate - currTime.date()
                if diff.days == 0:
                    timeVsDebitTransactionDaily[currTime] = amount
                    total_spendings += amount
            send_mail(total_spendings, "Daily", timeVsDebitTransactionDaily)
   
def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
            #creds = flow.run_console()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    #Gmail service
    service = build('gmail', 'v1', credentials=creds)
    
    messages = ListMessagesMatchingQuery(service, 'me', "{subject:debited subject:transaction}")
    
    process_messages(messages, service)
            
if __name__ == '__main__':
    main()

