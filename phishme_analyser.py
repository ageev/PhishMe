#!/usr/bin/env python

# this script will anylise phishing mails
# v1 (c) Artyom Ageyev

# https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook#properties

import ExtractMsg
import os, csv, operator
from html.parser import HTMLParser

MSG_DIR = "r:\\mails\\"
DATA_PATH = "r:\\data.csv"
URLS_PATH = "r:\\urls.csv"

class URLParser(HTMLParser):
    def __init__(self, output_list=None):
        HTMLParser.__init__(self)
        if output_list is None:
            self.output_list = []
        else:
            self.output_list = output_list
    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            self.output_list.append(dict(attrs).get('href'))

def main():
    mails_details = []
    urls = []

#    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for subdir, dirs, files in os.walk(MSG_DIR):
        for file in files:
            filename = os.path.join(subdir, file)
            msg = ExtractMsg.Message(filename)
            url = URLParser()
            url.feed(msg.HTMLBody)
            
            mail_details = []
            mail_details.extend((msg.SentOn, msg.SenderEmailAddress, msg.To, msg.CC, msg.Subject, url.output_list))
            mails_details.append(mail_details)
            urls.extend(url.output_list)

    mails_details = sorted(mails_details, key=operator.itemgetter(0), reverse=False)
    urls = filter(None, urls)

    write_csv(DATA_PATH, mails_details)
    write_urls(URLS_PATH, urls)

def write_csv(path, data):
    with open(path, "w", encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerows(data)

def write_urls(path, data):
    with open(path, "w", encoding='utf-8', newline='') as f:
        f.writelines(["%s\n" % item  for item in data])


if __name__ == '__main__':  
    main()
