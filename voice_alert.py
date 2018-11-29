from email_reader import outlook
import pyttsx3
import time
import datetime

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

# main logic for voice alert

if alert_msg == critical and last_alert <= 5:
    engine = pyttsx3.init()
    volume = engine.getProperty('volume')
    engine.setProperty('volume', volume - 100)

    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - 50)

    for i in range(3):
        engine.say('Application server is down, for more details please check the mail box')
        engine.runAndWait()
        time.sleep(2)



