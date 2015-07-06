# Python Outlook(Microsoft email service) Library
Python Library to read email from live, hotmail, outlook or any microsoft email service

# Example
To read Unread Message in inbox :
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.inbox()
mail.unread()

To read Unread Message in Junk :
mail = outlook.Outlook()
mail.login('emailaccount@live.com','yourpassword')
mail.junk()
mail.unread()