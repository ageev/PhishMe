#!/usr/bin/env python3

# v1 Artem (c)
# https://github.com/ecederstrand/exchangelib
# https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook#properties

import os, unicodedata, re, time
from exchangelib import Credentials, Account, Configuration, NTLM, DELEGATE, ItemAttachment, Message

user=''
password=''
ews_server='outlook.office365.com'
#SaveLocation = 'C:/Temp/phish/'
SaveLocation = ''
SubjectSearch = ''

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

    for item in account.inbox.all()[:10]:
        print('[INFO] Mail subject: ', item.subject, item.sender, item.datetime_received)
    #if SubjectSearch in item.subject:
        for attachment in item.attachments:
            if isinstance(attachment, ItemAttachment):
                if isinstance(attachment.item, Message):
                    # filename = slugify(attachment.name)
                    filename = item.datetime_received.strftime('%Y-%m-%d_%H-%M-%S') + '_' + slugify(attachment.name)
                    local_path = os.path.join(SaveLocation, filename + ".eml")

                    if os.path.exists(local_path):
                        local_path = os.path.join(SaveLocation, filename + '-' + slugify(attachment.item.message_id) +".eml")

                    with open(local_path, 'wb') as f:
                        f.write(attachment.item.mime_content)
                        print('[INFO] Attachment "' + filename + '" is saved')

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
