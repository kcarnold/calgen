import datetime
from uuid import uuid4

header = '''\
BEGIN:VCALENDAR
PRODID:-//Ken Arnold//Workday to ICS//EN
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VTIMEZONE
TZID:America/Detroit
X-LIC-LOCATION:America/Detroit
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
'''

footer = 'END:VCALENDAR\n'



def generate_uid():
    return str(uuid4())

def ics_date(date: datetime.date):
    return date.strftime("%Y%m%d")
    

def ics_datetime(date: str, time):
    return f"{ics_date(date)}T{time['hour']:02d}{time['minute']:02d}00"


def all_day_event(date: datetime.date, summary):
    assert '\n' not in summary
    assert date is not None
    return f'''BEGIN:VEVENT
DTSTART;VALUE=DATE:{ics_date(date)}
SUMMARY:{summary}
UID:{generate_uid()}
END:VEVENT
'''

def recurring_event(first_date: str, last_date: str, summary: str, location: str,
    start_time_p, end_time_p, meeting_pattern, exceptions):
    pattern = 'MO,WE,FR' # FIXME!
    exceptions_str = '\n'.join('EXDATE;TZID=America/Detroit:' + ics_datetime(exdate, start_time_p) for exdate in exceptions)
    return f'''\
BEGIN:VEVENT
DTSTAMP:20220827T120000Z
SUMMARY:{summary}
LOCATION:{location}
DTSTART;TZID=America/Detroit:{ics_datetime(first_date, start_time_p)}
DTEND;TZID=America/Detroit:{ics_datetime(first_date, end_time_p)}
UID:{generate_uid()}
RRULE:FREQ=WEEKLY;INTERVAL=1;BYDAY={pattern};UNTIL={ics_datetime(last_date, {'hour': 23, 'minute': 59})}Z
{exceptions_str}
END:VEVENT
'''

def write_ics(events):
    result = header + '\n' + '\n'.join(events) + '\n' + footer
    return result.replace('\n', '\r\n')
