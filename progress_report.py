from __future__ import print_function
from datetime import datetime
import gmail
import httplib2
import logging
import os
import pytz
import urllib

from bs4 import BeautifulSoup
from apiclient import discovery
from apiclient import errors
from apiclient.http import MediaFileUpload
import oauth2client
from oauth2client import client
from oauth2client import tools
from pushbullet import Pushbullet

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

logging.basicConfig(filename='progress_report.log',
                    level=logging.INFO,
                    format='%(asctime)s.%(msecs)d::%(levelname)s::%(module)s::%(funcName)s:: %(message)s',
                    datefmt="%Y-%m-%d %H:%M:%S")
SCOPES = 'https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Progress Report Extractor'

email_dict = {}


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'progress_report_extractor.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def parseURLs():
    href_dict = {}
    mem_moments = {}
    grab_row = False
    for edate, values in email_dict.iteritems():
        for ehtml in values.itervalues():
            soup = BeautifulSoup(ehtml, 'html.parser')
            for link in soup.findAll('a', href=True, text='Save Picture'):
                href_dict[edate] = {}
                href_dict[edate] = link['href']
            for row in soup.findAll('tr'):
                for r in row.findAll('td'):
                    if grab_row:
                        mem_moments[edate] = {}
                        mem_moments[edate] = r.string
                        grab_row = False
                    elif r.string == 'Memorable Moment':
                        grab_row = True

    for edate, href in href_dict.iteritems():
        email_dict[edate]['url'] = {}
        email_dict[edate]['url'] = href

    for edate, mm in mem_moments.iteritems():
        email_dict[edate]['mem_moment'] = {}
        email_dict[edate]['mem_moment'] = mm


def grabHtml(emails):
    for email in emails:
        email.fetch()
        sent_timezone = email.sent_at.replace(tzinfo=pytz.utc).astimezone(pytz.timezone("US/Mountain"))
        logging.info('Processing email from %s' % str(sent_timezone))
        email_dict[str(sent_timezone)] = {}
        email_dict[str(sent_timezone)]['html'] = email.html
        email.read()
#        email.archive()


def grabImages():
    image_names = {}
    for edate, values in email_dict.iteritems():
        for key, parsed in values.iteritems():
            if key == 'url':
                temp_filename = edate.replace(' ', '').replace('-', '')
                save_filename = temp_filename.replace(':', '') + '.jpg'
                urllib.urlretrieve(parsed, '/tmp/' + save_filename)
                image_names[edate] = {}
                image_names[edate] = save_filename

    for edate, iname in image_names.iteritems():
        email_dict[edate]['image'] = {}
        email_dict[edate]['image'] = '/tmp/' + iname


def uploadImages():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v2', http=http)
    for edate in email_dict.iterkeys():
        logging.info('Uploading image: %s' % email_dict[edate]['image'])
        media_body = MediaFileUpload(email_dict[edate]['image'],
                                     mimetype='image/jpeg',
                                     resumable=True)
        body = {'title': edate.replace(' ', '').replace('-', '').replace(':', ''),
                'description': email_dict[edate]['mem_moment'],
                'mimeType': 'image/jpeg',
                'parents': [{'id': '<folder id>'}],
                }
        try:
            file = service.files().insert(
                body=body,
                media_body=media_body).execute()
            logging.info('File ID: %s' % file['id'])
        except errors.HttpError, error:
            logging.error('An error occured: %s' % error)

        try:
            os.remove(email_dict[edate]['image'])
            logging.info('Deleting file %s' % email_dict[edate]['image'])
        except OSError:
            pass


def sendBullet():
    pb = Pushbullet('<pushbullet id>')
    for key, values in email_dict.iteritems():
        logging.info('Sending PushBullet notification')
        push = pb.push_note('Processed email for %s' % key,
                            'Image upload complete: %s' % values['image'])


def main():
    #TODO: implement oauth authentication for email; was left to password to allow for different accounts between email and Drive.
    g = gmail.login('<email>', '<password>')
    emails = g.mailbox('<label name>').mail(unread=True)
    if emails:
        logging.info('Number of emails to process: %s' % len(emails))
        grabHtml(emails)
        parseURLs()
        grabImages()
        uploadImages()
        sendBullet()
    else:
        logging.info('No emails to process.  Exiting.')

if __name__ == '__main__':
    main()
