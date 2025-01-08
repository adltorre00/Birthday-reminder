# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# Messages list

messages_list = [
    "Today is the event: {event}.",  # Index 0
    "{days} day(s) ago it was the event: {event}.",  # Index 1
    "In {days} day(s), it will be the event: {event}.",  # Index 2
    "No events within the next 30 days or the past 10 days.",  # Index 3
    "Today is {name}'s Birthday.{age_part}",  # Index 4
    "{days} day(s) ago it was {name}'s Birthday.{age_part}",  # Index 5
    "In {days} day(s), it will be {name}'s Birthday.{age_part}",  # Index 6
    "No Birthdays within the next 30 days or past 10 days!"  # Index 7
]

# Initialize a flag to indicate whether the data loading was successful
data_loaded = False

# Load and clean data
current_dir = os.path.dirname(os.path.abspath(__file__))
DatesFile = os.path.join(current_dir, "Dates.xlsx")

# Assuming 'DatesFile' is the path to your Excel file
try:
    # Load the data from the Excel sheets
    data = pd.read_excel(DatesFile, sheet_name="Birthdates")
    data2 = pd.read_excel(DatesFile, sheet_name="Events")

    # Drop rows with missing essential data
    data.dropna(subset=['Name', 'Day', 'Month'], inplace=True)
    data2.dropna(subset=['Event', 'Day', 'Month'], inplace=True)

    # Convert day, month, year, and age columns to numeric, handling errors by coercing them to NaN, then filling with 0
    for col in ['Day', 'Month', 'Year', 'Age']:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).astype(int)
    
    for col in ['Day', 'Month']:
        if col in data2.columns:
            data2[col] = pd.to_numeric(data2[col], errors='coerce').fillna(0).astype(int)

    # Filter out rows with invalid day or month values
    data = data[(data['Day'] > 0) & (data['Day'] <= 31) & (data['Month'] > 0) & (data['Month'] <= 12)]
    data2 = data2[(data2['Day'] > 0) & (data2['Day'] <= 31) & (data2['Month'] > 0) & (data2['Month'] <= 12)]

    # Create 'Birthday' column, ensuring that datetime creation is handled correctly
    def create_birthday(row):
        try:
            if row['Year'] > 0:
                return pd.to_datetime(f"{row['Year']}-{row['Month']}-{row['Day']}", format="%Y-%m-%d")
            else:
                return pd.to_datetime(f"{row['Month']}-{row['Day']}", format="%m-%d", errors='coerce')
        except ValueError:
            return pd.NaT

    data['Birthday'] = data.apply(create_birthday, axis=1)

    # Validate that all entries in 'Birthday' are proper datetime objects or NaT
    if not all(isinstance(val, pd.Timestamp) or val is pd.NaT for val in data['Birthday']):
        raise ValueError("Not all 'Birthday' entries are datetime objects or NaT.")

    data_loaded = True
    print("Data loaded and cleaned successfully.")
except Exception as e:
    print(f"Error loading or processing Excel file: {e}")
    data_loaded = False

if data_loaded:   
    def check_events(df):
        """Check upcoming or recent events and return messages based on the day and month."""
        messages = []
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
        for _, row in df.iterrows():
            # Handle event date for the current year
            try:
                event_date = datetime(today.year, row['Month'], row['Day'])
            except ValueError:
                # Adjust to the last valid day of the month
                event_date = datetime(today.year, row['Month'] + 1, 1) - timedelta(days=1)
    
            days_diff = (event_date - today).days
    
            # Handle year-end transition for past events
            if event_date < today:
                if today.month == 1 and event_date.month == 12:
                    # Event was in December of the previous year
                    if -5 <= days_diff:
                        # Event is within the past 5 days
                        msg = messages_list[1].format(days=-days_diff, event=row['Event'])
                        messages.append(msg)
                elif -5 <= days_diff <= 0:
                    # Recent past event (not related to year-end transition)
                    msg = messages_list[1].format(days=-days_diff, event=row['Event'])
                    messages.append(msg)
            else:
                # Future event handling
                if 0 <= days_diff <= 30:
                    if days_diff == 0:
                        msg = messages_list[0].format(event=row['Event'])
                    else:
                        msg = messages_list[2].format(days=days_diff, event=row['Event'])
                    messages.append(msg)
    
        if not messages:
            msg = messages_list[3]
            messages.append(msg)
    
        return messages

    
    def check_birthdays(df):
        """Check birthdays and return messages."""
        messages = []
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    
        for _, row in df.iterrows():
            age_now = row['Age']
    
            # Handle invalid birthday dates by adjusting to the last valid day of the month.
            try:
                birthday_this_year = datetime(today.year, row['Month'], row['Day'])
            except ValueError:
                birthday_this_year = datetime(today.year, row['Month'] + 1, 1) - timedelta(days=1)
    
            # Calculate the difference in days considering the year transition.
            if birthday_this_year < today:
                # The birthday has already occurred this year.
                days_diff = (birthday_this_year - today).days
            else:
                # The birthday is yet to come this year or is today.
                days_diff = (birthday_this_year - today).days
    
            # Adjust age if the birthday has not yet occurred this year.
            if days_diff > 0:
                age_now += 1
    
            # Construct the age part of the message.
            age_part = f" He/She will be {age_now} years old!" if age_now > 0 else ""
    
            # Determine the appropriate message based on the days difference.
            if -10 <= days_diff <= 30:
                if days_diff == 0:
                    messages.append(messages_list[4].format(name=row['Name'], age_part=age_part))
                elif days_diff < 0:
                    # For past events, the days_diff is negative; adjust the message accordingly.
                    messages.append(messages_list[5].format(days=-days_diff, name=row['Name'], age_part=age_part))
                else:
                    messages.append(messages_list[6].format(days=days_diff, name=row['Name'], age_part=age_part))
    
        # Add a default message if there are no birthday messages.
        if not messages:
            messages.append(messages_list[7])
    
        return messages


    def send_daily_email_report(birthday_text, event_text, sender_address, sender_pass, receiver_address):
        if not birthday_text and not event_text:
            return  # No birthdays or events today, so no email to send
        
        message = MIMEMultipart("alternative")
        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = 'Dates Reminder'
        
        # HTML email body
        email_body = """
        <html>
        <head>
        <style>
          .heading {font-weight: bold; font-size: 16px;}
          .content {font-size: 14px;}
        </style>
        </head>
        <body>
        <p>Dear Reader,</p>
        <p>Here are the upcoming dates you shouldn't forget about :):</p>
        """
    
        if birthday_text:
            email_body += '<p class="heading">Upcoming Birthdays:</p><p class="content">' + '<br>'.join(birthday_text) + '</p>'
        if event_text:
            email_body += '<p class="heading">Upcoming Events:</p><p class="content">' + '<br>'.join(event_text) + '</p>'
    
        email_body += "<p>Best regards,<br><br>Antonio de la Torre</p></body></html>"
        
        message.attach(MIMEText(email_body, 'html'))
        
        try:
            session = smtplib.SMTP('smtp.gmail.com', 587)
            session.starttls()
            session.login(sender_address, sender_pass)
            session.sendmail(sender_address, receiver_address, message.as_string())
            session.quit()
            print('Mail Sent')
        except Exception as e:
            print(f"Failed to send email: {e}")
            print(email_body)
    
    # Prepare the email content
    text_to_send1 = check_birthdays(data)  # Assume this returns a list of birthday reminders
    text_to_send2 = check_events(data2)    # Assume this returns a list of event reminders
    
    # Send the email
    send_daily_email_report(text_to_send1, text_to_send2, 'adelatorreprz@gmail.com', 'uwhx rvfx pajk nomg', 'adelatorreprz@gmail.com')

else:
    print("Data was not loaded successfully. Unable to proceed with birthday checks or email notifications.")