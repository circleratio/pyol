import win32com.client
import pytz
import re

#
# Read Calendar
#
class Agenda:
    subject = ''
    start_date = ''
    start_time = ''
    end_date = ''
    end_time = ''
    categories = []

    def __init__(self, subject, start_date, start_time, end_date, end_time, categories):
        self.subject = subject
        self.start_date = start_date
        self.end_date = end_date
        self.categories = categories

        h, m = start_time.split(':')
        self.start_time = '{:0>2}:{:0>2}'.format(int(h), int(m))
        
        h, m = end_time.split(':')
        self.end_time = '{:0>2}:{:0>2}'.format(int(h), int(m))
        
    def __str__(self):
        return('Agenda[{},{},{},{},{},{}]'.format(self.subject, self.start_date, self.start_time, self.end_date, self.end_time, self.categories))
        
def filter_subject(subject, ignore_patterns):
    for i in ignore_patterns:
        m = re.search(i, subject) 
        if m:
            return(True)
        
    return(False)

def filter_category(item, ignore_categories):
    for cat in ignore_categories:
        cats = re.split(', *', item)
        if cat in cats:
            return(True)
    return(False)

def read_outlook(start_date, end_date, read_private, read_tentative, ignore_patterns, ignore_categories):
    Outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    CalendarItems = Outlook.GetDefaultFolder(9).Items

    # 定期的な予定の二番目以降の予定を検索に含める
    CalendarItems.IncludeRecurrences = True
    
    # 開始時間でソート
    CalendarItems.Sort("[Start]")

    # "mm/dd/yyyy HH:MM AM"の形式に変換し、フィルター文字列を作成
    strStart = start_date.strftime('%m/%d/%Y %H:%M %p')
    strEnd = end_date.strftime('%m/%d/%Y %H:%M %p')
    sFilter = f"[Start] >= '{strStart}' And [End] <= '{strEnd}'"

    # フィルターを適用
    result = []
    FilteredItems = CalendarItems.Restrict(sFilter)
    for item in FilteredItems:
        if filter_subject(item.Subject, ignore_patterns):
            continue

        if filter_category(item.Categories, ignore_categories):
            continue
        
        if item.Sensitivity == 2 and read_private == False: # olPrivate
            continue
        
        if item.BusyStatus == 1 and read_tentative == False: # olTentative
            continue
        
        mon_s = int(item.Start.Format("%m"))
        day_s = int(item.Start.Format("%d"))
        h_s = int(item.Start.Format("%H"))
        m_s = int(item.Start.Format("%M"))
        
        mon_e = int(item.End.Format("%m"))
        day_e = int(item.End.Format("%d"))
        h_e = int(item.End.Format("%H"))
        m_e = int(item.End.Format("%M"))
        
        dow_s = ['日', '月', '火', '水', '木', '金', '土', '日'][int(item.Start.Format("%w"))]
        dow_e = ['日', '月', '火', '水', '木', '金', '土', '日'][int(item.End.Format("%w"))]
        
        d_s = '{}月{}日({})'.format(mon_s, day_s, dow_s)
        d_e = '{}月{}日({})'.format(mon_e, day_e, dow_e)
        
        t_s = '{}:{}'.format(h_s, m_s)
        t_e = '{}:{}'.format(h_e, m_e)
        
        s = str(item.Subject)

        agenda = Agenda(s, d_s, t_s, d_e, t_e, item.categories)
        #result += [[s, d_s, t_s, d_e, t_e, item.Categories]]
        result.append(agenda)
    return(result)

def get_key_date(ol_list):
    result = []
    for i in ol_list:
        result.append(i.start_date)
    return(sorted(set(result)))

def get_key_date_with_subject(ol_list, subject):
    result = []
    for i in ol_list:
        if i.subject == subject:
            result.append(i.start_date)
    return(sorted(set(result)))

def get_values_date(key, ol_list):
    result = []
    for i in ol_list:
        if key == i.start_date:
            result.append(i)
    return(sorted(result, key=lambda a: a.start_time))

#
# Add Calendar
#
def add_item(agenda):
    outlook = win32com.client.Dispatch("Outlook.Application")
    item = outlook.CreateItem(1)

    item.reminderSet = False
    
    item.subject = agenda['subject']
    item.body = agenda['body']
    item.location = agenda['location']
    
    if 'categories' in agenda:
        item.categories = agenda['categories']

    if 'recipients' in agenda:
        for n in re.split(', *', agenda['recipients']):
            item.Recipients.Add(n)

    if agenda['allday']:
        item.allDayEvent = True
        item.start = agenda['start'].strftime('%Y-%m-%d')
        item.Save()
        print(item.start.strftime('%Y-%m-%d') + ":" + item.subject + " is added")
    else:
        item.allDayEvent = False
        item.start = agenda['start'].strftime('%Y-%m-%d %H:%M')
        item.duration = agenda['duration']
        item.Save()
        print(item.start.strftime('%Y-%m-%d %H:%M') + ":" + item.subject + " is added")

#
# Add Calendar
#
def add_todo_item(task):
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    item = outlook.CreateItem(3)
    item.Subject = task['Subject']
    item.Body = task['Body']
    item.StartDate = pytz.utc.localize(task['StartDate'])
    item.DueDate = pytz.utc.localize(task['DueDate'])
    item.ReminderSet = True
    item.Save()
