import win32com.client
from datetime import date
from configparser import ConfigParser
import os
import sys
import platform
import win32file


# Get current work week as int.
work_week = date.today().isocalendar()[1]
config_file_name = "team_document_mailer_config.txt"

# Get directory of the script
def get_script_path():
    # cwd didn't work when run as executable. Had to use sys.argv[0] in code below.
    return os.path.dirname(os.path.realpath(sys.argv[0]))


try:
    config_dir = get_script_path()
    config_path = os.path.realpath(os.path.join(config_dir, config_file_name))
    print("Config path: ", config_path)

except Exception as e:
    print("Could not find config file.")
    print(e)

#   read config.ini file.
cparser = ConfigParser()
cparser.read(config_path)

# Retrieveing configs below configs from team_document_mailer_config.txt in script
# folder.
documents = cparser.get('General', 'documents').split(', ')
print("config file documents: ", documents)
address = cparser.get('General', 'address')
print("config file address: ", address)
subject = cparser.get('General', 'subject')
print("config file subject: ", subject)
body = cparser.get('General', 'body')
print("config file body: ", body)

# Start or find existing instance of outlook.
o = win32com.client.Dispatch("Outlook.Application")

# Create email
Msg = o.CreateItem(0)

# Set email addresses in outlook format.
Msg.To = address

# Msg.CC = "more email addresses here"
# Msg.BCC = "more email addresses here"

# Set message subject with WW prepended
Msg.Subject = "Work week {}: {}".format(work_week, subject)

# Set message body.
Msg.Body = body

# Attach all documents specified in config file.
for document in documents:
    try:
        Msg.Attachments.Add(document)
    except Exception as e:
        print("Could not attach document: ", document)
        print(e)

# Send message
Msg.Send()