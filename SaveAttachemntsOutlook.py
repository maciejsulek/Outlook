import win32com.client
import os
import datetime as dt

# Initial Settings ---------------------------------------

# save to...
server = "//darfp02/3D_BIM/2) Projects - Post Win/Olympia/008 - Reports/" 

# launch Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# print out current time
date = dt.datetime.now()
print (date)

# setup range for outlook to search emails (so we don't go through the entire inbox)
lastWeekDateTime = dt.datetime.now() - dt.timedelta(days = 1)
lastWeekDateTime = lastWeekDateTime.strftime('%m/%d/%Y %H:%M %p')

# Select main Inbox and/or subfolder
inbox = outlook.GetDefaultFolder(6).Folders["ASITE (REPORTS)"]

messages = inbox.Items

# Only search emails in the time range above:
messages = messages.Restrict("[ReceivedTime] >= '" + lastWeekDateTime +"'")

print ('Reading Inbox, including Inbox Subfolders...')

# Download a select attachment ---------------------------------------

# Create a folder to capture attachments.
Myfolder = server + 'Asite Reports/'
if not os.path.exists(Myfolder): os.makedirs(Myfolder)

try:
    for message in list(messages):
        try:
            s = message.sender
            s = str(s)
            print('Sender:' , message.sender)
            for att in message.Attachments:
                # Give each attachment a path and filename
                outfile_name1 = Myfolder + att.FileName[:-47] + date.strftime('%Y-%m-%d') + '.xlsx'
                # save file 
                att.SaveASFile(outfile_name1)
                print('Saved file:', outfile_name1)

        except Exception as e:
            print("type error: " + str(e))

except Exception as e:
    print("type error: " + str(e))

#Purge unused file types (like .png)-----------------------------------------

directory = os.listdir(Myfolder)

for item in directory:
    if item.endswith(".png"):
        os.remove(os.path.join(Myfolder, item))