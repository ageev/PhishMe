#!/usr/bin/env python3

# v1 Artem (c)
# https://github.com/ecederstrand/exchangelib
# https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook#properties

import os, unicodedata, re, time, logging
from exchangelib import Credentials, Account, Configuration, DELEGATE, ItemAttachment, Message, EWSTimeZone, EWSDateTime
from datetime import datetime

user='phishmailbox@domain.com'
password='Pa55w.rd' #for those who likes searching github for passwords
ews_server='outlook.office365.com'
remove_mails = True #this trigger is used for testing. False == do not remove mails

tz = EWSTimeZone.localzone() # set local timezone

# check OS and set path accordingly. Needed only for easy development 
if os.path.exists('C:/Windows'):
    MAIL_DIR = 'C:/Temp/phish/mails/'
    LOG_DIR = 'C:/Temp/phish/'
    DATA_DIR = 'C:/Temp/phish/'
    remove_mails = False  #during tests do not remove phishing mails
else:
    MAIL_DIR = '/media/nas01/Controls/PhishMe reports/mails/'
    LOG_DIR = '/media/nas01/Controls/PhishMe reports/'
    DATA_DIR = '/media/nas01/Controls/PhishMe reports/'


logging.basicConfig(filename= LOG_DIR + 'phishme.log', level=logging.INFO)

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
        localtime = item.datetime_received.astimezone(tz)
        with open(DATA_DIR + 'data.csv', "a") as f:
            f.write(localtime.strftime("%Y-%m-%d %H:%M:%S") + ';' 
                + slugify(item.subject) + ';'
                + item.sender.email_address + '\n')
        logging.info(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' New message received at ' + item.datetime_received.strftime("%Y-%m-%d %H:%M:%S")  + ' from ' + item.sender.email_address)
        for attachment in item.attachments:
            if isinstance(attachment, ItemAttachment):
                if isinstance(attachment.item, Message):
                    if len(attachment.name) > 80:
                        attachment.name = attachment.name[0:80]

                    # I want to make sure that if file is not saved for any reason - proper log generated
                    try: 
                        filename = localtime.strftime('%Y-%m-%d_%H-%M-%S') + '_' + slugify(attachment.name) + '_' + item.sender.email_address
                        local_path = os.path.join(MAIL_DIR, filename + ".eml")
                        with open(local_path, 'wb') as f:
                            f.write(attachment.item.mime_content)
                        logging.info(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' Message saved to ' + filename)
                    except:
                        logging.error(datetime.now().strftime("%Y-%m-%d %H:%M:%S") \
                            + ' Fail to write mail from ' + slugify(item.sender.email_address))

        if remove_mails: 
            item.move_to_trash()  #move mail to trash (!). Could be replaced with .delete() to permanently delete. 
            logging.info(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' Message from ' + item.sender.email_address + ' was removed from the mailbox')

def slugify(value, allow_unicode=False, allow_spaces=False):  #code copied from stackoverflow
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

    if allow_spaces:
        return value
    else:
        return re.sub(r'[-\s]+', '-', value)

if __name__ == '__main__':  
    main()
