import re
import pandas as pd
from ics import Calendar, Event
from datetime import date, timedelta

# Special dates (TODO: don't hard-code)
# Third entry is the pattern: what day-of-week it corresponds to. See iter_meeting_dates.
special_dates = [
    ['2022-09-05', 'Labor Day', None],
    ['2022-10-10', 'Fall Break', None],
    ['2022-10-11', 'Fall Break', None],
    ['2022-11-01', 'Advising', None],
    ['2022-11-02', 'Advising', None],
    ['2022-11-23', 'Thanksgiving', None],
    ['2022-11-24', 'Thanksgiving', None],
    ['2022-11-25', 'Thanksgiving', None],
    ['2022-12-08', 'Study', -1]
]
special_dates = pd.DataFrame(special_dates, columns=["date", "desc", "pattern"])
special_dates['date'] = pd.to_datetime(special_dates['date'])

# Read the input file.
data = pd.read_excel('/Users/ka37/Dropbox/Mac/Downloads/View_My_Teaching_Schedule.xlsx')

# Merge shadow reservations (multiple locations for the same course section and time)
data['Location'] = data.groupby(['Course Section', 'Meeting Time'])['Location'].transform(lambda x: ', '.join(x))
data = data.drop_duplicates(['Course Section', 'Meeting Time'])

# Parse the "meeting time" field.
parsed = pd.concat([data, data['Meeting Time'].str.extract(r'^(?P<days>\w+) \| (?P<time>[^|]+)')], axis=1)

# Use single letters for each date ("R" instead of "TH" for Thursday)
parsed['days'] = parsed['days'].str.replace('TH', 'R')


def iter_meeting_dates(start_date: date, end_date: date, pattern: str, special_dates):
    '''Yield all meeting times for the given class, given a meeting pattern.'''
    one_day = timedelta(days=1)
    days = ['MTWRFSU'.index(d) for d in pattern]
    cur = start_date
    while cur <= end_date:
        effective_date = cur.weekday()
        for special in special_dates.itertuples():
            if cur == special.date:
                effective_date = special.pattern
        if effective_date in days:
            yield cur
        if effective_date == -1:
            break
        cur += one_day

def parse_time(x):
    """Parse 1:00 PM into hour=13, min=0."""
    hour, min, meridian = re.match(r'^(\d+):(\d+) (AM|PM)', x).groups()
    hour = int(hour)
    min = int(min)
    if meridian == 'PM':
        hour += 12
    return dict(hour=hour, minute=min)


cal = Calendar()
for i in range(len(data)):
    start_time, end_time = parsed['time'].iloc[i].split(' - ')
    start_time_p = parse_time(start_time)
    end_time_p = parse_time(end_time)
    for meeting_date in iter_meeting_dates(
        parsed['Start Date'].iloc[i].date(),
        parsed['End Date'].iloc[i].date(),
        parsed['days'].iloc[i],
        special_dates
    ):
        evt = Event()
        evt.name = data['Course Section'].iloc[i]
        evt.begin = pd.Timestamp(meeting_date).replace(**start_time_p).tz_localize('US/Eastern')
        evt.end = pd.Timestamp(meeting_date).replace(**end_time_p).tz_localize('US/Eastern')
        evt.location = data['Location'].iloc[i]
        cal.events.add(evt)



open('out.ics', 'w').write(cal.serialize())
