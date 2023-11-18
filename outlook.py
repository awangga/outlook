import email
import imaplib
import smtplib
import datetime
import email.mime.multipart
import config
import base64
import quopri

class Outlook():
    def __init__(self):
        pass
        # self.imap = imaplib.IMAP4_SSL('imap-mail.outlook.com')
        # self.smtp = smtplib.SMTP('smtp-mail.outlook.com')

    def login(self, username, password):
        self.username = username
        self.password = password
        login_attempts = 0
        while True:
            try:
                self.imap = imaplib.IMAP4_SSL(config.imap_server,config.imap_port)
                r, d = self.imap.login(username, password)
                assert r == 'OK', 'login failed: %s' % str (r)
                print(" > Signed in as %s" % self.username, d[0].decode('utf-8'))
                return
            except Exception as err:
                print(" > Sign in error: %s" % str(err))
                login_attempts = login_attempts + 1
                if login_attempts < 3:
                    continue
                assert False, 'login failed'

    def sendEmailMIME(self, recipient, subject, message):
        msg = email.mime.multipart.MIMEMultipart()
        msg['to'] = recipient
        msg['from'] = self.username
        msg['subject'] = subject
        msg.add_header('reply-to', self.username)
        # headers = "\r\n".join(["from: " + "sms@kitaklik.com","subject: " + subject,"to: " + recipient,"mime-version: 1.0","content-type: text/html"])
        # content = headers + "\r\n\r\n" + message
        try:
            self.smtp = smtplib.SMTP(config.smtp_server, config.smtp_port)
            self.smtp.ehlo()
            self.smtp.starttls()
            self.smtp.login(self.username, self.password)
            self.smtp.sendmail(msg['from'], [msg['to']], msg.as_string())
            print("   email replied")
        except smtplib.SMTPException:
            print("Error: unable to send email")

    def sendEmail(self, recipient, subject, message):
        headers = "\r\n".join([
            "from: " + self.username,
            "subject: " + subject,
            "to: " + recipient,
            "mime-version: 1.0",
            "content-type: text/html"
        ])
        content = headers + "\r\n\r\n" + message
        attempts = 0
        while True:
            try:
                self.smtp = smtplib.SMTP(config.smtp_server, config.smtp_port)
                self.smtp.ehlo()
                self.smtp.starttls()
                self.smtp.login(self.username, self.password)
                self.smtp.sendmail(self.username, recipient, content)
                print("   email sent.")
                return
            except Exception as err:
                print("   Sending email failed: %s" % str(err))
                attempts = attempts + 1
                if attempts < 3:
                    continue
                raise Exception("Send failed. Check the recipient email address")
    def list_folders(self):
        resp, data = self.imap.list()
        return [i.decode().split(' "/" ')[1].replace('"','') for i in data]
    def list(self):
        # self.login()
        return self.imap.list()

    def select(self, str):
        return self.imap.select(str)

    def inbox(self):
        return self.imap.select("Inbox")

    def junk(self):
        return self.imap.select("Junk")

    def logout(self):
        return self.imap.logout()

    def since_date(self, days):
        mydate = datetime.datetime.now() - datetime.timedelta(days=days)
        return mydate.strftime("%d-%b-%Y")

    def allIdsSince(self, days):
        r, d = self.imap.search(None, '(SINCE "'+self.since_date(days)+'")', 'ALL')
        list = d[0].decode('utf-8').split(' ')
        return list

    def allIdsToday(self):
        return self.allIdsSince(1)

    def readIdsSince(self, days):
        r, d = self.imap.search(None, '(SINCE "'+self.date_since(days)+'")', 'SEEN')
        list = d[0].decode('utf-8').split(' ')
        return list

    def readIdsToday(self):
        return self.readIdsSince(1)

    def unreadIdsSince(self, days):
        r, d = self.imap.search(None, '(SINCE "'+self.since_date(days)+'")', 'UNSEEN')
        list = d[0].decode('utf-8').split(' ')
        return list

    def unreadIdsToday(self):
        return self.unreadIdsSince(1)

    def allIds(self):
        r, d = self.imap.search(None, "ALL")
        list = d[0].decode('utf-8').split(' ')
        return list

    def readIds(self):
        r, d = self.imap.search(None, "SEEN")
        list = d[0].decode('utf-8').split(' ')
        return list

    def unreadIds(self):
        r, d = self.imap.search(None, "UNSEEN")
        list = d[0].decode('utf-8').split(' ')
        return list

    def hasUnread(self):
        list = self.unreadIds()
        return list != ['']

    def getIdswithWord(self, ids, word):
        stack = []
        for id in ids:
            self.getEmail(id)
            if word in self.mailbody().lower():
                stack.append(id)
        return stack

    def getEmail(self, id):
        r, d = self.imap.fetch(bytes(str(id),'utf-8'), "(RFC822)")
        self.raw_email = quopri.decodestring(d[0][1]).decode('utf-8')
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

    def readOnly(self, folder):
        return self.imap.select(folder, readonly=True)

    def writeEnable(self, folder):
        return self.imap.select(folder, readonly=False)

    def rawRead(self):
        list = self.readIds()
        latest_id = list[-1]
        r, d = self.imap.fetch(bytes(latest_id,'utf-8'), "(RFC822)")
        self.raw_email = d[0][1].decode('utf-8')
        return self.raw_email

    def mailbody(self):
        if self.email_message.is_multipart():
            for payload in self.email_message.get_payload():
                # if payload.is_multipart(): ...
                body = (
                    payload.get_payload()
                    .split(self.email_message['from'])[0]
                    .split('\r\n\r\n2015')[0]
                )
                return body
        else:
            body = (
                self.email_message.get_payload()
                .split(self.email_message['from'])[0]
                .split('\r\n\r\n2015')[0]
            )
            return body

    def mailsubject(self):
        return self.email_message['Subject']

    def mailfrom(self):
        return self.email_message['from']

    def mailto(self):
        return self.email_message['to']

    def maildate(self):
        return self.email_message['date']

    def mailreturnpath(self):
        return self.email_message['Return-Path']

    def mailreplyto(self):
        return self.email_message['Reply-To']

    def mailall(self):
        return self.email_message

    def mailbodydecoded(self):
        return base64.urlsafe_b64decode(self.mailbody())
