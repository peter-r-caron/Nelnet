'''
State unemployment insurance agencies require that claimants keep records of weekly job contacts.
In order to justify a weekly unemployment claim, the claimant must prepare a report for submittal
to the state unemployment insurance agency.  This is a time consuming and error prone manual task.
Generally, when a job application is submitted via an internet website, the prospective employer
immediately sends a confirmation email to the applicant.  The job application information is readily
available in the confirmation emails sent by the prospective employers. The goal of this automation
task will be to populate a relational database table with job-related contact information retrieved
from Outlook email messages.  The extracted information will be stored in a Microsoft SQL Server
database.
'''

from datetime import datetime
import sys
import pyodbc
import win32com.client

def _insert_job_contacts(the_date, the_sender, the_subject, the_body, cursor):
    sql_insert =  "insert into dbo.JobContacts (DateOfContact, EmailAddressOfSender, SubjectOfEmail, BodyOfEmail) "
    sql_insert += "select ?, convert(nvarchar(max),?), convert(nvarchar(max),?), convert(nvarchar(max),?) "
    sql_insert += "where not exists (select * from dbo.JobContacts "
    sql_insert += "where DateOfContact = ? "
    sql_insert += "and EmailAddressOfSender = convert(nvarchar(max),?) "
    sql_insert += "and SubjectOfEmail = convert(nvarchar(max),?) "
    sql_insert += "and BodyOfEmail = convert(nvarchar(max),?))"
    data_values = [the_date, the_sender, the_subject, the_body, the_date, the_sender, the_subject, the_body]
    cursor.execute(sql_insert,data_values)

def _print_job_contacts(cur, start, end):
    query = "select DateOfContact,EmailAddressOfSender,SubjectOfEmail,BodyOfEmail "
    query += "from dbo.JobContacts "
    query += "where DateOfContact between '" + start + "' and '" + end + "' "
    query += "order by DateOfContact desc"
    i = 1
    print("------------------------------------------------------------------------------------------------------------")
    for row in cur.execute(query):
        print(row.DateOfContact + " - " + row.EmailAddressOfSender + "\n")
        print(row.SubjectOfEmail + "\n")
        print(row.BodyOfEmail + "\n")
        print("------------------------------------------------------------------------------------------------------------")
        command = input()
        if command == 'q':
            break

# Check for command line arguments
ARG_COUNT = len(sys.argv)

if ARG_COUNT > 2:
    message_start_date = sys.argv[1]
    message_end_date = sys.argv[2]
else:
# Prompt user for start and end dates.
    print('Enter Start Date "mm/dd/yyyy hh:mm": ')
    message_start_date = input()
    print('Enter End Date "mm/dd/yyyy hh:mm": ')
    message_end_date = input()

message_start_date = datetime.strftime(datetime.strptime(message_start_date, '%m/%d/%Y %H:%M'), '%m/%d/%Y %H:%M')
message_end_date = datetime.strftime(datetime.strptime(message_end_date, '%m/%d/%Y %H:%M'), '%m/%d/%Y %H:%M')

# Connect to database.
conn = pyodbc.connect('Driver={SQL Server};Server=DESKTOP-RBLHC9P\SQLEXPRESS;Database=Nelnet;Trusted_Connection=yes;')
cursor = conn.cursor()
count = 0
command = ''

# Instantiate Outlook email client and prepare to read Inbox folder.
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# The Inbox is default folder number 6.
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items

# Setup filters for the messages to be selected by date range.
messages = messages.Restrict("[ReceivedTime] >= '" + message_start_date + "'")
messages = messages.Restrict("[ReceivedTime] <= '" + message_end_date + "'")

# Cycle through email Inbox searching for messages that fall within the start and end dates.
# When message meets criteria, display the date, sender, and subject of the message on the terminal.
for msg in list(messages):
    print("------------------------------------------------------------------------------------------------------------")
    print(datetime.strftime(msg.ReceivedTime, '%m/%d/%Y %H:%M') + " | " + msg.SenderEmailAddress + " | " + msg.Subject)
    print("------------------------------------------------------------------------------------------------------------")

# Prompt the user to (l)oad, (s)kip, or (q)uit.
    print("(l)oad, (s)kip, or (q)uit")
    command = input().lower()

# If log is selected, insert a row into the JobContacts table where the row does not already exist.
    if command in ('l', 'list'):
        _insert_job_contacts(datetime.strftime(msg.ReceivedTime, '%m/%d/%Y %H:%M'), msg.SenderEmailAddress, msg.Subject, msg.Body, cursor)
        count += 1

# If skip is selected, move on to the next email message that meets the search criteria.
    if command in ('s', 'skip'):
        continue

# If quit is selected, display the rows that were inserted and quit the program.
    if command in ('q', 'quit'):
        break

# Commit changes to the database.
conn.commit()

# Print the job contacts entered for the date range.
print(str(count) + ' contacts loaded into JobContacts table.')
print("View contacts for date range? y/n:")
if input().lower() == 'y': _print_job_contacts(cursor, message_start_date, message_end_date)

# Cleanup and exit
conn.close()
sys.exit()
