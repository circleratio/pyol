import argparse

import pyol

if __name__ == '__main__':
    agenda = {
        'subject': '',
        'body': '',
        'location': '',
        'allday': False,
    }

    parser = argparse.ArgumentParser(description='Add a meeting to Outlook Schedule.')

    parser.add_argument('subject', help='Subject')
    parser.add_argument('start', help='Date (and time) the meeting starts')
    parser.add_argument('-d', '--duration', help='duration of the meeting in minutes', type=int)
    parser.add_argument('-c', '--categories', help='category of the meeting')
    parser.add_argument('-r', '--recipients', help='recipients')
    
    args = parser.parse_args()
    
    agenda['subject'] = args.subject
    print('Subject: ' + args.subject)
    
    agenda['start'] = pyol.parse_datetime(args.start)
    print('Start: ' + args.start)
    
    if args.categories:
        print('Categories: ' + args.categories)
        agenda['categories'] = args.categories
        
    if args.recipients:
        print('Recipients: ' + args.recipients)
        agenda['recipients'] = args.recipients

    if args.duration:
        print('Duration: ' + str(args.duration))
        agenda['duration'] = args.duration
        agenda['allday'] = False
    else:
        print('No duration. Fallback to an All-day event.')
        agenda['allday'] = True
        
    pyol.add_item(agenda)
