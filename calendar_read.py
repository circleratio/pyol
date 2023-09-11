import datetime
import re

import pyol

def fmt_subject(subject):
    replace_patterns = [
        ['BEFORE', 'AFTER'],
    ]
    for reg, rep in replace_patterns:
        subject = re.sub(reg, rep, subject)
    return(subject)

def print_outlook(start, end):
    ignore_patterns = [
        'IGNORED',
    ]

    ignore_categories = [
        'CATEGORY',
    ]
    
    item_list = pyol.read_outlook(start, end, True, False, ignore_patterns, ignore_categories)
    
    #keys = pyol.get_key_date_with_subject(item_list, 'KEY')
    keys = pyol.get_key_date(item_list)
    for k in keys:
        print(k)
        for agenda in pyol.get_values_date(k, item_list):
            if agenda.subject == 'KEY':
                continue
            print('  {}-{}: {}'.format(agenda.start_time, agenda.end_time, fmt_subject(agenda.subject)))
        
def main():
    today = datetime.date.today()
    week_start = today + datetime.timedelta(0 - today.weekday()) # this Monday
    week_end = today + datetime.timedelta(6 - today.weekday()) # this Saturday
    
    print_outlook(week_start, week_end)
    
if __name__ == "__main__":
    main()
