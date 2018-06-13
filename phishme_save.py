#!/usr/bin/env python3

# v1 Artem (c)
# https://github.com/ecederstrand/exchangelib
# https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook#properties

import os, unicodedata, re, time
from exchangelib import Credentials, Account, Configuration, DELEGATE, ItemAttachment, Message, EWSTimeZone, EWSDateTime
from datetime import timedelta

user=''
password=''
ews_server='outlook.office365.com'
last_x_minutes = 10 #check mails for received in last 10 minutes

tz = EWSTimeZone.localzone()

# check OS
if os.path.exists('C:/Temp/phish/'):
    SaveLocation = 'C:/Temp/phish/'
else:
    SaveLocation = '/mails/'

# TBD
def clean_mailbox(account, age_in_days):
    for item in account.inbox.all().filter():
        pass

def main():
    config = Configuration(
        server=ews_server,
        credentials = Credentials(username=user, password=password)
        )
    account = Account(
        primary_smtp_address=user,
        config=config,
        autodiscover=False,
        access_type=DELEGATE
        )
    save_attachments(account)

def save_attachments(account):
    for item in account.inbox.all():
        print('[INFO]', item.datetime_received, item.subject, item.sender)
        for attachment in item.attachments:
            if isinstance(attachment, ItemAttachment):
                if isinstance(attachment.item, Message):
                    # filename = slugify(attachment.name)
                    localtime = item.datetime_received.astimezone(tz)
                    filename = localtime.strftime('%Y-%m-%d_%H-%M-%S') + '_' + slugify(attachment.name) + '_' + slugify(item.sender.email_address)
                    local_path = os.path.join(SaveLocation, filename + ".eml")

                    with open(local_path, 'wb') as f:
                        f.write(attachment.item.mime_content)
                        print('[INFO] Attachment "' + filename + '" is saved')
        item.move_to_trash()  #move mail to trash. Could be replaced with .delete()

def slugify(value, allow_unicode=False):
    """
    Convert to ASCII if 'allow_unicode' is False. Convert spaces to hyphens.
    Remove characters that aren't alphanumerics, underscores, or hyphens.
    Convert to lowercase. Also strip leading and trailing whitespace.
    """
    value = str(value)
    if allow_unicode:
        value = unicodedata.normalize('NFKC', value)
    else:
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub(r'[^\w\s-]', '', value).strip().lower()
    return re.sub(r'[-\s]+', '-', value)

if __name__ == '__main__':  
    main()
