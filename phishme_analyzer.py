#!/usr/bin/env python3

# v1 Artem (c)
# https://github.com/ecederstrand/exchangelib
# https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook#properties
import email, os, base64, re, configparser, time, logging, sys, unicodedata
from email.header import Header, decode_header, make_header
from datetime import datetime, timedelta
from exchangelib import ItemAttachment, Message, EWSTimeZone, EWSDateTime

# check OS and set path accordingly. Needed only for easy development 
if os.path.exists('C:/Windows'):
    MAIL_DIR = 'C:/Temp/phish/mails/'
    LOG_DIR = 'C:/Temp/phish/'
    DATA_DIR = 'C:/Temp/phish/'
    INFECTED_DIR = 'C:/Temp/phish/infected/'
    METADATA_FILE = 'C:/Temp/phish/phishme_metadata.csv'
else:
    MAIL_DIR = '/media/nas01/Controls/PhishMe reports/mails/'
    LOG_DIR = '/media/nas01/Controls/PhishMe reports/'
    DATA_DIR = '/media/nas01/Controls/PhishMe reports/'
    INFECTED_DIR = '/media/nas01/Controls/PhishMe reports/INFECTED/'
    METADATA_FILE = '/media/nas01/Controls/PhishMe reports/phishme_metadata.csv'

#using local time zone
tz = EWSTimeZone.localzone() 
logging.basicConfig(filename= LOG_DIR + 'phishme.log', level=logging.INFO)

def main():
    # get last file date from config
    config = configparser.ConfigParser()
    config.readfp(open(DATA_DIR + 'phishme.cfg'))
    LAST_CHECKED_MAIL_DATE = config.get('Attachments', 'LAST_CHECKED_MAIL_DATE' )
    last_date = datetime.strptime(LAST_CHECKED_MAIL_DATE, '%Y-%m-%d_%H-%M-%S') #convert string to date object
    next_day = datetime.strftime(last_date + timedelta(days=1), '%Y-%m-%d')  # get next day to string
    new_last_date = last_date # defaul value

    # iterate through all files started with this date or later and pick fresh file(s)
    for file in os.listdir(MAIL_DIR):
        try:
            file_date = datetime.strptime(file[:19], '%Y-%m-%d_%H-%M-%S')
        except:
            logging.error(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' File ' + file + ' is not in correct format')
            continue
        if file_date > last_date:
            # put all actions here
            try:
                save_attachment(MAIL_DIR, file)
            except:
                logging.error(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Error opening file: " +  file)
            try:
                save_mail_metadata(MAIL_DIR + file)
            except:
                logging.error(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Error getting metadata. File: " +  file)
            if file_date > new_last_date:
                new_last_date = file_date
    # write last date to config file
    config.set('Attachments', 'LAST_CHECKED_MAIL_DATE', new_last_date.strftime('%Y-%m-%d_%H-%M-%S'))
    logging.info(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' Attachments saved till ' + new_last_date.strftime('%Y-%m-%d_%H-%M-%S'))
    with open(DATA_DIR + 'phishme.cfg', 'w') as configfile:
        config.write(configfile)

def save_attachment(directory, file):
    emlfile = os.path.join(directory, file)
    eml = email.message_from_file(open(emlfile))
    date_string = file[:19]
    if eml['X-MS-Has-Attach']:
        for part in eml.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            if part.get('Content-Disposition') is None:
                continue

            # sometimes enclosed attachemnt is another message, so special case needed
            if part.get_content_maintype() == 'message':
                filename = date_string + '_' + 'message.eml'
                with open(os.path.join(INFECTED_DIR, filename), 'w') as fp: 
                    fp.write(part.get_payload()[0].as_string())
                continue

            filename = part.get_filename()
            if not(filename): filename = "sample" + '.txt'
            if 'utf-8' or 'windows-125' in filename:
                filename = decode_strange_header(filename)

            filename = date_string + '_' + filename.replace('\n', ' ')   #remove new line symbol from file names. broke script few times before..

            try:
                with open(os.path.join(INFECTED_DIR, filename), 'wb') as fp: 
                    fp.write(part.get_payload(decode=True))
            except: 
                logging.error(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Error writing file. EML from " + date_string)

def decode_strange_header(header):
    h = make_header(decode_header(header))
    return str(h)

def save_mail_metadata(emlfile):
    eml = email.message_from_file(open(emlfile))
    result = ''

    # check for "NoneType" values and decode strange headers. windows-125 should catch windows-1251, 1252
    if not eml['Date']:
        eml['Date'] = '-'

    if not eml['From']:
        eml['From'] = '-'
    elif 'utf-8' or 'windows-125' in eml['From']:
        eml['From'] = decode_strange_header(eml['From'])

    if not eml['To']:
        eml['To'] = '-'
    elif 'utf-8' or 'windows-125' in eml['To']:
        eml['To'] = decode_strange_header(eml['To'])

    if not eml['Subject']:
        eml['Subject'] = '-'
    elif 'utf-8' or 'windows-125' in eml['Subject']:
        eml['Subject'] = decode_strange_header(eml['Subject'])

    # build result string
    result = ";".join([eml['Date'], eml['From'], eml['To'][:80], slugify(eml['Subject'], False, True)]) + ";"

    # add processed "Received" metadata
    for i in eml.items():
        if i[0] == 'Received':
            from_ip = i[1].replace('\n',' ').split(' by')[0].replace('from ', '')  # get host and IP only
            if 'prod.outlook.com' not in from_ip:    # don't store o365 servers
                result += from_ip + ';'

    # write metadata file
    with open(METADATA_FILE, "a") as f:
        f.write(result.replace('\n',' ') + "\n")

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