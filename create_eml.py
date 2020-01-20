# create_eml.py
# This script creates an eml mail file with attachments
# This is used for testing eml,mbox and pst converters.
# Can be opend with Thunderbird email program
# Robert Fearn (c) 2020
# Working with Windows10/Python3.7
# This files are hardcoded and not argparsed in for simplicity 

import os
import sys

# For guessing MIME type based on file name extension
import mimetypes

from argparse import ArgumentParser
from email import generator
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

eml_file1 = 'test9.eml' #file to write

file2attach1="Screenshot.png"
file2attach2="Cert.pdf"
file2attach3="test.docx"

def create_attachment(filepath):
    #taken from https://docs.python.org/2/library/email-examples.html
    if os.path.isfile(filepath):
    # Guess the content type based on the file's extension.  Encoding
    # will be ignored, although we should check for simple things like
    # gzip'd or compressed files.
        ctype, encoding = mimetypes.guess_type(filepath)
        if ctype is None or encoding is not None:
            # No guess could be made, or the file is encoded (compressed), so
            # use a generic bag-of-bits type.
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        if maintype == 'text':
            with open(filepath) as fp:
                # Note: we should handle calculating the charset
                msg = MIMEText(fp.read(), _subtype=subtype)
        elif maintype == 'image':
            with open(filepath, 'rb') as fp:
                msg = MIMEImage(fp.read(), _subtype=subtype)
        elif maintype == 'audio':
            with open(filepath, 'rb') as fp:
                msg = MIMEAudio(fp.read(), _subtype=subtype)
        else:
            with open(filepath, 'rb') as fp:
                msg = MIMEBase(maintype, subtype)
                msg.set_payload(fp.read())
            # Encode the payload using Base64
            encoders.encode_base64(msg)
            # Set the filename parameter
        filename=os.path.basename(filepath)
        msg.add_header('Content-Disposition', 'attachment', filename=filename)
        
    return msg

#Create the email structure
outer = MIMEMultipart()
outer['Subject'] = 'Test Subject'
outer['To'] = 'to@here.com'
outer['From'] = 'from@here.com'
outer['Cc'] = "cc@here.com"
outer['Bcc'] = "bcc@here.com"
outer['Date'] = "Sat, 30 Apr 2018 19:28:29 -0300"
html_data = """\
<html>
    <head></head>
    <body>
        <p> Test Email 2 </p>
    </body>
</html>
"""

part = MIMEText(html_data, 'html')
outer.attach(part)
#Now attach files to the email
jpg2attach=create_attachment(file2attach1)
outer.attach(jpg2attach)

pdf2attach=create_attachment(file2attach2)
outer.attach(pdf2attach)

doc2attach=create_attachment(file2attach3)
outer.attach(doc2attach)


#We can write the file 2 ways. By flattening or writing the structure out to string

with open(eml_file1, 'w') as outfile:
    gen = generator.Generator(outfile)
    gen.flatten(outer)

#Second way
# composed = outer.as_string()
# with open(eml_file2, 'w') as fp:
#    fp.write(composed)
