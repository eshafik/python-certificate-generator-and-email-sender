import httplib2
import os
import oauth2client
from oauth2client import client, tools
from oauth2client import file
import base64

# needed for attachment
import mimetypes
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from apiclient import errors, discovery  # needed for gmail service


SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'cred.json'
APPLICATION_NAME = 'Auto email sender'


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'gmail-python-email-send.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


# Get creds, prepare message and send it
def create_message_and_send(sender, to, subject, message_text_plain, attached_file):
    credentials = get_credentials()

    # Create an httplib2.Http object to handle our HTTP requests, and authorize it using credentials.authorize()
    http = httplib2.Http()

    # http is the authorized httplib2.Http()
    http = credentials.authorize(http)  # or: http = credentials.authorize(httplib2.Http())

    service = discovery.build('gmail', 'v1', http=http)

    if not attached_file:
        message_without_attachment = create_message_without_attachment(sender, to, subject, message_text_plain)
        send_Message_without_attachment(service, "me", message_without_attachment, message_text_plain)
        return True

    # with attachment
    message_with_attachment = create_Message_with_attachment(sender, to, subject, message_text_plain, attached_file)
    send_Message_with_attachment(service, "me", message_with_attachment, message_text_plain, attached_file)
    return True


def create_message_without_attachment(sender, to, subject, message_text_plain):
    # Create message container
    message = MIMEMultipart('alternative')  # needed for both plain & HTML (the MIME type is multipart/alternative)
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = to

    # Create the body of the message (a plain-text and an HTML version)
    message.attach(MIMEText(message_text_plain, 'plain'))

    raw_message_no_attachment = base64.urlsafe_b64encode(message.as_bytes())
    raw_message_no_attachment = raw_message_no_attachment.decode()
    body = {'raw': raw_message_no_attachment}
    return body


def create_Message_with_attachment(sender, to, subject, message_text_plain, attached_file):
    """Create a message for an email.

    message_text: The text of the email message.
    attached_file: The path to the file to be attached.

    Returns:
    An object containing a base64url encoded email object.
    """

    # Part 1
    message = MIMEMultipart()  # when alternative: no attach, but only plain_text
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    message.attach(MIMEText(message_text_plain, 'plain'))

    my_mimetype, encoding = mimetypes.guess_type(attached_file)

    # If the extension is not recognized it will return: (None, None)
    # If it's an .mp3, it will return: (audio/mp3, None) (None is for the encoding)
    # for unrecognized extension it set my_mimetypes to  'application/octet-stream' (so it won't return None again).
    if my_mimetype is None or encoding is not None:
        my_mimetype = 'application/octet-stream'

    main_type, sub_type = my_mimetype.split('/', 1)  # split only at the first '/'
    # if my_mimetype is audio/mp3: main_type=audio sub_type=mp3

    # -----3.2  creating the attachment
    # you don't really "attach" the file but you attach a variable that contains the "binary content" of the file you want to attach

    # option 1: use MIMEBase for all my_mimetype (cf below)  - this is the easiest one to understand
    # option 2: use the specific MIME (ex for .mp3 = MIMEAudio)   - it's a shorcut version of MIMEBase

    # this part is used to tell how the file should be read and stored (r, or rb, etc.)
    if main_type == 'text':
        print("text")
        temp = open(attached_file, 'r')  # 'rb' will send this error: 'bytes' object has no attribute 'encode'
        attachment = MIMEText(temp.read(), _subtype=sub_type)
        temp.close()

    elif main_type == 'image':
        print("image")
        temp = open(attached_file, 'rb')
        attachment = MIMEImage(temp.read(), _subtype=sub_type)
        temp.close()

    elif main_type == 'audio':
        print("audio")
        temp = open(attached_file, 'rb')
        attachment = MIMEAudio(temp.read(), _subtype=sub_type)
        temp.close()

    elif main_type == 'application' and sub_type == 'pdf':
        temp = open(attached_file, 'rb')
        attachment = MIMEApplication(temp.read(), _subtype=sub_type)
        temp.close()

    else:
        attachment = MIMEBase(main_type, sub_type)
        temp = open(attached_file, 'rb')
        attachment.set_payload(temp.read())
        temp.close()

    filename = os.path.basename(attached_file)
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)  # name preview in email
    message.attach(attachment)

    # Part 4 encode the message (the message should be in bytes)
    message_as_bytes = message.as_bytes()  # the message should converted from string to bytes.
    message_as_base64 = base64.urlsafe_b64encode(message_as_bytes)  # encode in base64 (printable letters coding)
    raw = message_as_base64.decode()  # need to JSON serializable (no idea what does it means)
    return {'raw': raw}


def send_Message_without_attachment(service, user_id, body, message_text_plain):
    try:
        message_sent = (service.users().messages().send(userId=user_id, body=body).execute())
        message_id = message_sent['id']
        # print(attached_file)
        print(f'Message sent (without attachment) \n\n Message Id: {message_id}\n\n Message:\n\n {message_text_plain}')
        # return body
    except errors.HttpError as error:
        print(f'An error occurred: {error}')


def send_Message_with_attachment(service, user_id, message_with_attachment, message_text_plain, attached_file):
    """Send an email message.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me" can be used to indicate the authenticated user.
    message: Message to be sent.

    Returns:
    Sent Message.
    """
    try:
        message_sent = (service.users().messages().send(userId=user_id, body=message_with_attachment).execute())
        message_id = message_sent['id']
        # print(attached_file)

        # return message_sent
    except errors.HttpError as error:
        print(f'An error occurred: {error}')


def send_email(send_to, name, subject, body, signature, attachment=None):
    to = send_to
    sender = ""
    subject = subject
    signature = signature.replace(",", ",\n")
    message_text_plain = f"Dear {name},\n\n\n{body}\n\n{signature}"
    attached_file = attachment
    create_message_and_send(sender, to, subject, message_text_plain, attached_file)
