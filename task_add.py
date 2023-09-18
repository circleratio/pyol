import pyol
import datetime
import argparse
    
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
    print('Start: {}'.format(task['StartDate']))
    
    if args.duedate:
        task['DueDate'] = pyol.parse_datetime(args.duedate)
        print('Due Date: ' + args.duedate)
    
    pyol.add_todo_item(task)
