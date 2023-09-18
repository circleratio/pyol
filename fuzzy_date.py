import sys
import re
import datetime
import calendar

def fuzzy_datetime(arg):
    dt_now = datetime.date.today()
    m = re.match('([^,]+), *([^,]+)', arg)
    if m:
        dt_orig = fuzzy_datetime_parse(m.group(2), dt_now)
        return(fuzzy_datetime_parse(m.group(1), dt_orig))
    else:
        return(fuzzy_datetime_parse(arg, dt_now))
    
def conv_weekday(s):
    s = s.lower()
    
    s = re.sub('monday', 'mo', s)
    s = re.sub('tuesday', 'tu', s)
    s = re.sub('wednesday', 'we', s)
    s = re.sub('thursday', 'th', s)
    s = re.sub('friday', 'fr', s)
    s = re.sub('saturday', 'sa', s)
    s = re.sub('sunday', 'su', s)

    s = re.sub('mon', 'mo', s)
    s = re.sub('tue', 'tu', s)
    s = re.sub('wed', 'we', s)
    s = re.sub('thu', 'th', s)
    s = re.sub('fri', 'fr', s)
    s = re.sub('sat', 'sa', s)
    s = re.sub('sun', 'su', s)
    
    return(s)
    
def fuzzy_datetime_parse(arg, dt_orig):
    arg = conv_weekday(arg)
    
    m = re.match('([0-9]+)/([0-9]+)/([0-9]+)$', arg)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        day = int(m.group(3))
        dt = datetime.datetime(year, month, day)
        return(dt)

    m = re.match('([0-9]+)/([0-9]+)$', arg)
    if m:
        year = int(m.group(1))
        month = int(m.group(2))
        dt = datetime.datetime(year, month, 1)
        return(dt)

    m = re.match('([0-9]+)/([0-9]+)$', arg)
    if m:
        month = int(m.group(1))
        day = int(m.group(2))
        dt = datetime.datetime(dt_orig.year, month, day)
        return(dt)

    m = re.match('([Mm]o|[Tt]u|[Ww]e|[Tt]h|[Ff]r|[Ss]a|[Ss]u)$', arg)
    if m:
        k = m.group(1).lower()
        pos = 'motuwethfrsasu'.find(k) // 2
        dow = dt_orig.weekday()
        day_diff = pos - dow
        if day_diff <= 0:
            day_diff += 7
        dt = dt_orig + datetime.timedelta(days=day_diff)
        return(dt)

    m = re.match('([0-9]+) day$', arg)
    if m:
        dt = dt_orig + datetime.timedelta(days=int(m.group(1)))
        return(dt)
        
    m = re.match('([0-9]+) days? ago$', arg)
    if m:
        dt = dt_orig + datetime.timedelta(days=-int(m.group(1)))
        return(dt)
    
    if arg == 'yesterday':
        dt = dt_orig + datetime.timedelta(days=-1)
        return(dt)
    
    if arg == 'tomorrow':
        dt = dt_orig + datetime.timedelta(days=1)
        return(dt)
    
    if arg == 'week':
        dt = dt_orig + datetime.timedelta(days=7)
        return(dt)
    
    m = re.match('([0-9]+) weeks?$', arg)
    if m:
        n = int(m.group(1)) * 7
        dt = dt_orig + datetime.timedelta(days=n)
        return(dt)
    
    if arg == 'week ago':
        dt = dt_orig + datetime.timedelta(days=-7)
        return(dt)
    
    m = re.match('([0-9]+) weeks? ago$', arg)
    if m:
        n = - int(m.group(1)) * 7
        dt = dt_orig + datetime.timedelta(days=n)
        return(dt)
    
    if arg == 'end of month':
        c = calendar.monthrange(dt_orig.year, dt_orig.month)
        dt = datetime.datetime(dt_orig.year, dt_orig.month, c[1])
        return(dt)
    
    print('Invalid argument: ' + arg)
    return(None)
        
def main():
    dt = fuzzy_datetime(sys.argv[1])
    print(dt)

if __name__ == "__main__":
    main()
