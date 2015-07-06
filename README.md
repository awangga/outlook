# Python Outlook(Microsoft email service) Library
Python Library to read email from live, hotmail, outlook or any microsoft email service, just dowload to yout python script folder. This library using Imaplib python to read email with IMAP protocol.

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