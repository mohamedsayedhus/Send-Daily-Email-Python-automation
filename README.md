# 📨 Daily Attendance Email Automation
This project automates the process of sending daily attendance reports via Microsoft Outlook. Built using Python, it utilizes the pywin32 library to interface with Outlook and handles date adjustments to send the correct report from the previous workday.

# 📌 Features
Automatically determines the correct attendance date (adjusted for weekends).
Uses the Outlook COM interface to generate and send emails.
Customizes email content based on the adjusted date.
Can be scheduled as a daily task via Windows Task Scheduler.

# Libraries Used
win32com.client (from pywin32) — for Outlook automation
datetime — for date manipulation
pythoncom — to initialize COM threading
os — for any file path handling (if included)

# 🎯 Goals
Automate the process of sending daily attendance emails without manual intervention.
Reduce the possibility of human error when selecting dates or formatting messages.
Ensure timely communication of attendance data to relevant stakeholders.
Create a reusable and adaptable solution that can work with Microsoft Outlook.

# 🌟 Benefits
Time-Saving: Eliminates the need for daily manual email preparation.
Accuracy: Automatically selects the correct date and formats it properly, even after weekends.
Consistency: Standardizes the email format and ensures it is sent on time every day.
Scalability: Can be easily extended to include attachments, integrate with databases, or support more recipients.
Easy Integration: Works seamlessly with Outlook using the pywin32 library.
