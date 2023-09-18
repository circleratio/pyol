import win32com.client
import argparse
import os

#
# Parse arguments
#
parser = argparse.ArgumentParser(description='Create a mail draft.')

parser.add_argument('--to', help='Recipients, separated by semi-colon')
parser.add_argument('--cc', help='CC recipients, separated by semi-colon')
parser.add_argument('--bcc', help='BCC recipients, separated by semi-colon')
parser.add_argument('--attach', help='Attachment file, separated by comma')
parser.add_argument('subject', help='Subject')
parser.add_argument('text', help='A text file for mail body.')

args = parser.parse_args()

#
# Invoke Outlook and make a draft
#
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

if args.to:
    mail.to = args.to
if args.cc:
    mail.cc = args.cc
if args.bcc:
    mail.bcc = args.bcc
if args.attach:
    for i in args.attach.split(','):
        mail.Attachments.Add(os.path.abspath(i)) # full-path needed

mail.subject = args.subject

with open(args.text, 'r') as f:
    mail.body = f.read()

mail.bodyFormat = 1

mail.display(True)
