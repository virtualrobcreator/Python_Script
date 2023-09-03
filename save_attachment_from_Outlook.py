print('----------------------------------SRART-----------------------------------')
print('----------------------------------SRART-----------------------------------')
print('----------------------------------SRART-----------------------------------')
import win32com.client
import datetime
import os
today = datetime.date.today()
print(today)
outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
inbox = outlook.Folders('Enter Your Mail Box Id').Folders('Inbox')
messages = inbox.Items
i=0
print(a)
for msg in messages:
    if msg.subject == 'Enter Subject' and msg.unread:# specify subject and status of mail
        for atch in msg.attachments: # for mail which contains attachments
            #if atch.file_extension == pdf :
            # print(msg.Subject+"      "+str(msg.ReceivedTime)+"     "+str(msg.attachments)+"    "+str(atch.FileName))
            atch.SaveAsFile(os.getcwd()+'\\abs\\'+atch.FileName)     # save attachment 
            msg.unread = False # change status of mail
            msg.move(outlook.Folders('Enter Your Mail Box Id').Folders('Enter another Folder Name'))# move mail to another folder
            # print(i)
            # print((os.getcwd()+'\\abcs\\'+atch.FileName))
            i=i+1
print(totalmail ,"=",i)
print('----------------------------------END-----------------------------------')
print('----------------------------------END-----------------------------------')
