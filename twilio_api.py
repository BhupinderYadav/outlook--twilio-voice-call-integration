import json
from twilio.rest import Client
from twilio.twiml.voice_response import VoiceResponse
from email_reader import outlook
import datetime

# importing twilio api auth details from config.json
with open("config.json", 'r') as p:
    param = json.load(p)

account_sid = param['account_sid']
auth_token = param['auth_token']
to_number = param['to_number']
from_number = param['from_number']
url = param['url']
p.close()


# reading the email using email_reader module and outlook class
outlook = outlook()
data = outlook.email_reader()
alert_msg = data['subject']
alert_time = data['time_stamp']
critical = data['critical']

# time calculation for the email alert
current_sys_time = datetime.datetime.now()
time_delta = current_sys_time - datetime.datetime.strptime(alert_time[0:16], "%Y-%m-%d %H:%M")
last_alert = time_delta.seconds/60

# call function
def voice_call():
    # resp = VoiceResponse()
    # resp.say("application server is down, please check you email")
    client = Client(account_sid, auth_token)
    call = client.calls.create(method='POST', url=str(url), to=str(to_number), from_=str(from_number))
    return call

# call alert logic
if alert_msg == critical and last_alert <= 5:
    voice_call()
    print(str(voice_call().sid))


