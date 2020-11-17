import win32com.client
import pandas as pd
outlook = win32com.client.Dispatch("Outlook.Application")

def sendEmail():
    Msg = outlook.CreateItem(0) # Email
    Msg.To = "jfulbright@network21.com" # you can add multiple emails with the ; as delimiter. E.g. test@test.com; test2@test.com;
    Msg.CC = "ajf73130@uga.edu"
    Msg.Subject = "Test"
    Msg.Body = "This is a Test"
    Msg.Send()

def sendRecurringMeeting(rec, title, date, prop, freq):    
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.Start = date # yyyy-MM-dd hh:mm
    appt.Subject = title
    appt.Duration = 1 # In minutes (60 Minutes)
    appt.Location = prop
    appt.MeetingStatus = 1
  # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added

    appt.Recipients.Add(rec) # Don't end ; as delimiter

  # Set Pattern, to recur every day, for the next 5 days
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = 2
    pattern.Interval = freq
    pattern.Occurrences = "5"

    appt.Save()
    appt.Send()

#setupAppointment("2020-11-28 10:10", "apt", "ajf73130@uga.edu", "my prop", 2)
#sendRecurringMeeting("ajf73130@uga.edu","meeting","2020-11-28 10:10","my house",2)

def test():
    print("hello")



