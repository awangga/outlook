# Python Outlook(Microsoft email service) Library
Python Library to read email from live, hotmail, outlook or any microsoft email service, just dowload to yout python script folder. This library using Imaplib python to read email with IMAP protocol.
## Prerequisite Library
Please make sure you have this library installed on your system first before your running this code
 * email
 * imaplib
 * smtplib
 * datetime

then rename config.py.sample to config.py and edit comment in config.py file
 
## Example
### To get latest Unread Message in inbox :
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.inbox()
print mail.unread()
```

### To get latest Unread Message in Junk :
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.junk()
print mail.unread()
```

### Retrive email element :
```py
print mail.mailbody()
print mail.mailsubject()
print mail.mailfrom()
print mail.mailto()
```

### To send Message :
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.sendEmail('recipient@email.com','subject','message body')
```

### To check Credentials :
```py
import outlook
mail = outlook.Outlook()
mail.checkLogin()
```
### Reading e-mails from Outlook with Python through MAPI and get email with word 'skype id'
```py
import Skype4Py
import outlook
import time
import config
import parser

skype = Skype4Py.Skype()
skype.Attach()

def checkingFolder(folder):
	mail = outlook.Outlook()
	mail.login(config.outlook_email,config.outlook_password)
	mail.readOnly(folder)
	print "  Looking Up "+folder
	try:
		unread_ids_today = mail.unreadIdsToday()
		print "   unread email ids today : "
		print unread_ids_today
		unread_ids_with_word = mail.getIdswithWord(unread_ids_today,'skype id')
		print "   unread email ids with word Skype ID today : "
		print unread_ids_with_word
	except:
		print config.nomail
	#fetch Inbox folder
	mail = outlook.Outlook()
	mail.login(config.outlook_email,config.outlook_password)
	mail.select(folder)
	try:
		for id_w_word in unread_ids_with_word:
			mail.getEmail(id_w_word)
			subject = mail.mailsubject()
			message = mail.mailbody()
			skypeidarr = parser.getSkype(message)
			print subject
			print skypeidarr
			i = 0
			while i < len(skypeidarr):
				skype.SendMessage(skypeidarr[i],config.intromsg+subject+"\r\n with Content : \r\n"+message)
				i += 1
			config.success()
			print "  sending reply message..."
			print "  to :"+mail.mailfrom().split('>')[0].split('<')[1]
			print "  subject : "+subject
			print "  content : "+config.replymessage
			mail.sendEmail(mail.mailfrom().split('>')[0].split('<')[1],"Re : "+subject,config.replymessage)
			time.sleep(10)
	except:
		print config.noword
		time.sleep(10)
		
while True:
	#checking ids in Inbox Folder
	print config.checkinbox
	checkingFolder('Inbox')
	#checking Junk Folder
	print config.checkjunk
	checkingFolder('Junk')
	
```
