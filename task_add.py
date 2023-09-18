import pyol
import datetime
import argparse
import re

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
    task = {
        'Subject': '',
        'Body': '',
        'StartDate': '',
        'DueDate': '',
    }

    parser = argparse.ArgumentParser(description='Add a task to Outlook TO-DO.')

    parser.add_argument('subject', help='Subject')
    parser.add_argument('-b', '--body', help='Body')
    parser.add_argument('-d', '--duedate', help='Due date of the task')
    
    args = parser.parse_args()
    
    task['Subject'] = args.subject
    print('Subject: ' + args.subject)

    if args.body:
        task['Body'] = args.body
        print('Body: ' + args.body)
        
    task['StartDate'] = datetime.datetime.today()
    print('Start: ' + args.start)
    
    if args.duedate:
        task['DueDate'] = parse_datetime(args.duedate)
        print('Due Date: ' + args.duedate)
    
    pyol.add_item(task)
    
