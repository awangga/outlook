#!/usr/bin/env python2

import re
import getpass
import outlook
import datetime
from pytz import timezone

class OutlookMailForwarder:
    def __init__(self, email_addr, email_passwd, window_hours=24, folder_list=None, \
                 mailing_list=None, subject_pattern='', body_pattern='', filter_body=None):
        self.mail = outlook.Outlook()
        self.mail.login(email_addr, email_passwd)
        self.window_hours = window_hours
        self.window_days = (self.window_hours + self.window_hours % 24) / 24
        self.folder_list = folder_list
        self.mailing_list = mailing_list
        self.subject_pattern = subject_pattern
        self.body_pattern = body_pattern
        self.filter_body = filter_body

    def send_email(self, mailsubject, mailbody):
        if self.mailing_list is None:
            return

        if mailsubject is None:
            return

        for recipient in self.mailing_list:
            try:
                self.mail.sendEmail(recipient, mailsubject, mailbody)
                print('sent mail to %s' % recipient)
            except Exception as err:
                print('error sending mail to %s: %s' % (recipient, str(err)))

    def prepare_email(self, email_id):
        self.mail.getEmail(email_id)
        mailsubject = self.mail.mailsubject()
        maildate = self.mail.maildate()
        # filter out any date below now - window_hour
        # because outlook searche is limited to the beginning of the day of the now - window_days
        maildatetime = datetime.datetime.strptime(maildate[:-6], '%a, %d %b %Y %H:%M:%S')
        maildatetime.replace(tzinfo=timezone('UTC'))
        timedelta = datetime.datetime.utcnow() - maildatetime
        if timedelta > datetime.timedelta(0, 0, 0, 0, 0, self.window_hours):
            raise ValueError('skipping email_id %s because its timedelta %s is greater than %d hours' % \
                             (email_id, str(timedelta), self.window_hours))
        else:
            print('email_id %s is within range (%s < %s)' % (email_id, str(timedelta), str(datetime.timedelta(0, 0, 0, 0, 0, self.window_hours))))
        mailsubject = mailsubject + ' ' + self.mail.maildate()
        mailbody = self.mail.mailbody()
        if self.filter_body is not None:
            mailbody = self.filter_body(mailbody)
        print('maillsubject to send: %s' % mailsubject)
        print('mailbody to send: %s' % mailbody)
        return (mailsubject, mailbody)

    def lookup_pattern(self):
        if self.folder_list is None:
            return

        for folder in self.folder_list:
            try:
                self.mail.select(folder)
                all_ids = self.mail.allIdsSince(self.window_days)
                max_num = 100 if len(all_ids) > 100 else len(all_ids)
                print('looking up pattern in %d/%d most recent emails in folder %s' % (max_num, len(all_ids), folder))
                emails_with_subject_pattern = self.mail.getIdswithWord(all_ids[:max_num], self.subject_pattern)
                print('%d items match subject_pattern' % len(emails_with_subject_pattern))
                emails_match = self.mail.getIdswithWord(emails_with_subject_pattern, self.body_pattern)
                print('%d items match subject_pattern and body_pattern' % len(emails_match))
            except Exception as err:
                print('error looking up pattern in folder %s: %s' % (folder, str(err)))
                continue

            try:
                for email_id in emails_match:
                    try:
     		        (mailsubject, mailbody) = self.prepare_email(email_id)
 		        self.send_email(mailsubject, mailbody)
                    except ValueError as err:
                        print('%s' % str(err))
                        continue
            except Exception as err:
                print('error processing matched emails in folder %s: %s' % (folder, str(err)))
                continue

def filter_zoom_mailbody(mailbody):
    ''' Returns the link to share. This filters out other info in the email such as the host-only link'''
    m = re.search(r'Share recording with viewers:<br>\s*(.*)\b', mailbody)
    return m.group(1)

def main(_user, _pass, win_hours):
    zoom_forwarder = OutlookMailForwarder(_user, _pass, win_hours, folder_list=['zoom'], \
                                          mailing_list=['fwd@example.com'], \
                                          subject_pattern='cloud recording', \
                                          body_pattern='share recording with viewers:', \
                                          filter_body=filter_zoom_mailbody)
    zoom_forwarder.lookup_pattern()

if __name__ == '__main__':
    _user = raw_input('Outlook email:')
    _pass = getpass.getpass('Outlook Password:')
    win_hours = int(raw_input('how many hours?'))
    main(_user, _pass, win_hours)
