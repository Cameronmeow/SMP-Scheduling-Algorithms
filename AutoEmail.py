import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "smpwebnominee@gmail.com"  # Replace with sender's email
SENDER_PASSWORD = "rrci cbbi ffep nzji"  # Replace with the sender's app password


# Load the interview schedule
schedule_file = "./Schedule/weekly_interview_schedule.csv"
df = pd.read_csv(schedule_file)

# Convert the interview date column to datetime
df['Interview Date'] = pd.to_datetime(df['Interview Date'], dayfirst=True)

# Get today's date and find emails for interviews happening 2 days later
reminder_date = datetime.today().date() + timedelta(days=2)
reminders = df[df['Interview Date'].dt.date == reminder_date]

def send_email(to_email, subject, body):
    """Function to send an email"""
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Connect to Gmail SMTP server
        context = ssl.create_default_context()
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.set_debuglevel(1)  # <-- This will print SMTP logs for debugging
        server.starttls(context=context)
        server.login(SENDER_EMAIL, SENDER_PASSWORD)

        # Send email
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()

        print(f"Reminder sent to {to_email}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {e}")


# Loop through the reminders and send emails
for _, row in reminders.iterrows():
    email_body = f"""
    Hello {row['Interviewee First Name']} {row['Interviewee Last Name']},

    This is a reminder for your upcoming interview.

    📅 Date: {row['Interview Date'].strftime('%A, %d %B %Y')}
    ⏰ Time: {row['Interview Time']}
    📍 Venue: Online / In-person (Specify if needed)

    If you have any questions, feel free to reach out.

    Best Regards,
    Interview Scheduling Team
    """

    send_email(row["Interviewee Email ID"], "Interview Reminder", email_body)

print("All reminders have been sent.")
