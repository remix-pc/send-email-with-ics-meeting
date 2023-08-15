import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from pathlib import Path
from email import encoders
from icalendar import Calendar, Event, vCalAddress, vText #You need install icalendar



#Start and End, this is the format for the ICS archive
start = "2023-08-12 12:00:00" 
end = "2023-08-12 12:00:00"



#E-mail configurations

bodyEmail = f"Text body here.. This is the start date:{start} \n This is the end date:{end}"
sender = f"sender@domain.com"
recipients = ["recipient1@domain.com", "recipient2@domain.com"]
password = "password"
organizer = vCalAddress(f"MAILTO:{sender}")
attendee = vCalAddress(f"MAILTO:{recipients}")
attendee.params['cn'] = vText("Email person name here")


#Create meeting and ics archive method
def createMeeting(organizer, bodyEmail, start, end, attendee):
    cal = Calendar()
    event = Event()
    event['organizer'] = organizer
    event.add("summary", f"SUMMARY HERE")
    event.add("location", "location here")
    event.add("description", bodyEmail)
    event.add("dtstart", start)
    event.add("dtend", end)
    event.add("attendee", attendee, encode=0)
    cal.add_component(event)
    with open("meeting.ics", "wb") as f:
        f.write(cal.to_ical())
    print("Meeting created")


#Send e-mail with ics archive.
def sendEmail(bodyEmail, sender, recipient, password):
    msg = MIMEMultipart()
    msg['Subject'] = f"Subject Here"
    msg["From"] = sender
    msg["To"] = recipient

    msg.attach(MIMEText(bodyEmail))
    path = "meeting.ics"
    part = MIMEBase('application', 'octet-stream')
    with open(path, 'rb') as file:
        part.set_payload(file.read())

    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment; filename={}".format(Path(path).name))
    msg.attach(part)
    txt = msg.as_string()


    host = smtplib.SMTP_SSL("smtp-server", 111)
    host.login(sender, password)
    host.sendmail(sendEmail, recipient, txt)
    host.quit


createMeeting(organizer, bodyEmail, start, end, attendee)

for email in recipients:
    sendEmail(bodyEmail, sender, email, password)
