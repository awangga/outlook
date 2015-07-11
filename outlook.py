import imaplib
import email
import smtplib
import datetime

class Outlook():
	def __init__(self):
		mydate = datetime.datetime.now()
		self.today = mydate.strftime("%d-%b-%Y")
		self.imap = imaplib.IMAP4_SSL('imap-mail.outlook.com')
		self.smtp = smtplib.SMTP('smtp-mail.outlook.com')
		
	def checkLogin(self):
		username = raw_input("account : ")
		password = raw_input("password : ")
		try:
			r, d = self.imap.login(username, password)
			assert r == 'OK', 'login failed'
		except:
			print "Invalid Login"
		else:
			print "Valid Login",d

	def login(self,username,password):
		self.username = username
	    self.password = password
	    while True:
			r, d = self.imap.login(username, password)
			assert r == 'OK', 'login failed'
			try:
				print "Connected as ",d
			except SocketError as e:
				print "not connected"
				continue
			break

	def sendEmail(self,recipient,subject,message):
		headers = "\r\n".join(["from: " + "sms@kitaklik.com","subject: " + subject,"to: " + recipient,"mime-version: 1.0","content-type: text/html"])
		content = headers + "\r\n\r\n" + message
		try:
			self.smtp.ehlo()
			self.smtp.starttls()
			self.smtp.login(self.username, self.password)
			self.smtp.sendmail(self.username, recipient, content)
			print "email replied"
		except smtplib.SMTPException:
			print "Error: unable to send email"
			
	def list(self):
		#self.login()
		return self.imap.list()
	
	def select(self,str):
		return self.imap.select(str)
		
	def inbox(self):
		return self.imap.select("Inbox")
	
	def junk(self):
		return self.imap.select("Junk")
	
	def logout(self):
		return self.imap.logout()
		
	def today(self):
		mydate = datetime.datetime.now()
		return mydate.strftime("%d-%b-%Y")
		
	def unreadIdsToday(self):
		r, d = self.imap.search(None,'(SINCE "'+self.today+'")', 'UNSEEN')
		list = d[0].split(' ')
		return list
		
	def getIdswithWord(self,ids,word):
		stack = []
		for id in ids:
			self.getEmail(id)
			curr_mailmsg = self.mailbody()
			if word in self.mailbody().lower():
				stack.append(id)
		return stack
		
	def unreadIds(self):
		r, d = self.imap.search(None, "UNSEEN")
		list = d[0].split(' ')
		return list
		
	def readIdsToday(self):
		r, d = self.imap.search(None,'(SINCE "'+self.today+'")', 'SEEN')
		list = d[0].split(' ')
		return list
		
	def readIds(self):
		r, d = self.imap.search(None, "SEEN")
		list = d[0].split(' ')
		return list
		
	def getEmail(self,id):
		r, d = self.imap.fetch(id, "(RFC822)")
		self.raw_email = d[0][1]
		self.email_message = email.message_from_string(self.raw_email)
		return self.email_message
		
	def unread(self):
		list = self.unreadIds()
		latest_id = list[-1]
		return self.getEmail(latest_id)
	
	def read(self):
		list = self.readIds()
		latest_id = list[-1]
		return self.getEmail(latest_id)
		
	def readToday(self):
		list = self.readIdsToday()
		latest_id = list[-1]
		return self.getEmail(latest_id)
	
	def unreadToday(self):
		list = self.unreadIdsToday()
		latest_id = list[-1]
		return self.getEmail(latest_id)
		
	def readOnly(self,folder):
		return self.imap.select(folder,readonly=True)
	
	def writeEnable(self,folder):
		return self.imap.select(folder,readonly=False)
				
	def rawRead(self):
		list = self.readIds()
		latest_id = list[-1]
		r, d = self.imap.fetch(latest_id, "(RFC822)")
		self.raw_email = d[0][1]
		return self.raw_email
		
	def mailbody(self):
		if self.email_message.is_multipart():
			for payload in self.email_message.get_payload():
				# if payload.is_multipart(): ...
				body = payload.get_payload().split(self.email_message['from'])[0].split('\r\n\r\n2015')[0]
				return body
		else:
			body = self.email_message.get_payload().split(self.email_message['from'])[0].split('\r\n\r\n2015')[0]
			return body

	def mailsubject(self):
		return self.email_message['Subject']		
		
	def mailfrom(self):
		return self.email_message['from']
		
	def mailto(self):
		return self.email_message['to']	