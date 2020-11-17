from reminder_Create import *
import pandas as pd
import datetime

#9/1/2020
recp = input("Enter Recipient(email) of Reminders: ")
xlsF = input("Enter excel document name: ")
#yyyy-MM-dd hh:mm
aptdf = pd.read_excel(xlsF)
#Creates reminders for two years
for index, row in aptdf.iterrows():
    location = row["PhysicalName"]
    subject = row["Reminder"]
    dates = format_date(row["Completed"])
    inter = row["FreqMonths"]
    occur = calcOccur(inter)
    sendRecurringMeeting(recp, subject, dates, location, inter, occur)
#newdate = format_date()
