import datetime
import re
import argparse

import pyol

def parse_datetime(s):
    m = re.match('\d{4}-\d{1,2}-\d{1,2} +\d{1,2}:\d{1,2}', s)
    if m:
        return(datetime.datetime.strptime(s, '%Y-%m-%d %H:%M'))
    
    m = re.match('\d{4}-\d{1,2}-\d{1,2}', s)
    if m:
        return(datetime.datetime.strptime(s, '%Y-%m-%d'))
    
    m = re.match('(mo|tu|we|th|fr|sa|su)', s)
    if m:
        print(m.group(1))
        dow_arg = int('motuwethfrsasu'.find(m.group(1)) / 2)
        dow_now = datetime.datetime.today().weekday()
        diff = dow_arg - dow_now
        if diff <= 0:
            diff += 7
        return(datetime.datetime.today() + datetime.timedelta(days=diff))
    
    print('Invalid datetime format: ' + s)
    exit(1)
        
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
    
    agenda['start'] = parse_datetime(args.start)
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
