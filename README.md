# Python Outlook (Microsoft email service) Library
Python Library to read email from live, hotmail, outlook or any microsoft email service, just dowload to yout python script folder. This library using Imaplib python to read email with IMAP protocol.
## Prerequisite Libraries
Please make sure you have these libraries installed on your system first before running this code:
 * email
 * imaplib
 * smtplib
 * datetime

then rename config.py.sample to config.py and edit comment in config.py file
 
## Examples
### To get latest Unread Message in inbox:
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.inbox()
print mail.unread()
```

### To get latest Unread Message in Junk:
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.junk()
print mail.unread()
```

Use `mail.select(folder)` to switch to folders other than `inbox, `junk`.
### Retrive email element:
```py
print mail.mailbody()
print mail.mailsubject()
print mail.mailfrom()
print mail.mailto()
```

### To send Message:
```py
import outlook
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.sendEmail('recipient@email.com','subject','message body')
```

### To check Credentials:
```py
import outlook
mail = outlook.Outlook()
mail.checkLogin()
```

### Reading e-mails from Outlook with Python through MAPI and get email with word 'skype id':
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

### Forward zoom recording email to recipients a mailing list:
Use `fwd_zoom.py`. The class `OutlookMailForwarder` defined there has the capability to filter messages based on the
received time (a time window from now -- in hours), and to do string search in email subject and body. Additionaly,
a filter can be defined to change the body before sending the email. For example, in `fwd_zoom.py`:

```py
def filter_zoom_mailbody(mailbody):
    ''' Returns the link to share. This filters out other info in the email such as the host-only link'''
    m = re.search(r'Share recording with viewers:<br>\s*(.*)\b', mailbody)
    return m.group(1)
```

To run, you can either define your `email[space]password` in `.cred` or giveemail/password in stdin upon prompt.
NOTE: Do not forget to `chmod 400 ./.cred` if the former method is used.

Example 1: With existing `.cred` with contents `test@example.com mypassword` and a small time window:
```
./fwd_zoom.py
How many hours to llok back?1
(' > Signed in as test@example.com', ['LOGIN completed.'])
looking up pattern in 3/3 most recent emails in folder zoom
1 items match subject_pattern
1 items match subject_pattern and body_pattern
skipping email_id 213 because its timedelta 6:31:25.556269 is greater than 1 hours
```

Example 2: without a `.cred` file and a large-enough time window:
```
./fwd_zoom.py
Outlook email:test@example.com
Outlook Password:
How many hours to look back?10
(' > Signed in as test@example.com', ['LOGIN completed.'])
looking up pattern in 3/3 most recent emails in folder zoom
1 items match subject_pattern
1 items match subject_pattern and body_pattern
email_id 213 is within range (6:45:22.572372 < 10:00:00)
maillsubject to send: Cloud Recording - Zoom Meeting is now available Tue, 12 Nov 2019 16:51:48 +0000
mailbody to send: https://test.zoom.us/recording/share/4abdcsefkHergre45grdgDdafdefMWd
   Sending email...
   email sent.
```
