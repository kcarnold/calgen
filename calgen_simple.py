from dataclasses import dataclass
import datetime
import re
import warnings
from collections import Counter
from typing import List

import pandas as pd

from ical_writer import all_day_event, recurring_event, write_ics

# Ignore warnings about missing default styles in openpyxl
# openpyxl/styles/stylesheet.py:226: UserWarning: Workbook contains no default style, apply openpyxl's default
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def letter_to_day(d):
    return 'MTWRFSU'.index(d)

# Special dates (TODO: don't hard-code)
# Third entry is the pattern: what day-of-week it corresponds to. See iter_meeting_dates.

class SpecialDate:
    def __init__(self, date, name: str, pattern):
        if isinstance(date, str):
            date =  datetime.date.fromisoformat(date)
        self.date = date
        self.name = name
        if isinstance(pattern, str):
            pattern = letter_to_day(pattern)
        self.pattern = pattern

special_dates = [
    # Fall 2022
    ['2022-09-05', 'Labor Day', None],
    ['2022-10-10', 'Fall Break', None],
    ['2022-10-11', 'Fall Break', None],
    ['2022-11-01', 'Advising', None],
    ['2022-11-02', 'Advising', None],
    ['2022-11-23', 'Thanksgiving', None],
    ['2022-11-24', 'Thanksgiving', None],
    ['2022-11-25', 'Thanksgiving', None],
    ['2022-12-08', 'Study', -1],
    # Spring 2023
    ['2023-02-27', 'Spring Break', None],
    ['2023-02-28', 'Spring Break', None],
    ['2023-03-01', 'Spring Break', None],
    ['2023-03-02', 'Spring Break', None],
    ['2023-03-03', 'Spring Break', None],
    ['2023-03-21', 'Advising', None],
    ['2023-03-22', 'Advising', None],
    ['2023-04-07', 'Good Friday', None],
    ['2023-04-10', 'Easter Monday', None],
    ['2023-04-20', 'Thursday with Friday schedule', "F"],
    ['2023-04-21', 'Exams Start', -1],
    # Fall 2023
    ['2023-09-04', 'Labor Day', None],
    ['2023-10-13', 'Fall Break', None],
    ['2023-10-16', 'Fall Break', None],
    ['2023-10-17', 'Advising', None],
    ['2023-10-18', 'Advising', None],
    ['2023-11-22', 'Thanksgiving', None],
    ['2023-11-23', 'Thanksgiving', None],
    ['2023-11-24', 'Thanksgiving', None],
    ['2023-12-08', 'Study', -1],
    # Spring 2024
    ['2024-01-15', 'MLK Day', None],
    ['2024-03-04', 'Spring Break', None],
    ['2024-03-05', 'Spring Break', None],
    ['2024-03-06', 'Spring Break', None],
    ['2024-03-07', 'Spring Break', None],
    ['2024-03-08', 'Spring Break', None],
    ['2024-03-20', 'Advising', None],
    ['2024-03-21', 'Advising', None],
    ['2024-03-29', 'Good Friday', None],
    ['2024-04-01', 'Easter Monday', None],
    ['2024-04-27', 'Exams Start', -1],
    # Summer 2024
    ['2024-05-24', 'Memorial Day', None],
    ['2024-05-27', 'Memorial Day', None],
    ['2023-06-19', 'Juneteenth', None],
    ['2024-07-03', 'Independence Day', None],
    ['2024-07-04', 'Independence Day', None],
    ['2024-07-05', 'Independence Day', None],
    ['2024-08-17', 'Exams Start', -1],
    # Fall 2024
    ['2024-09-02', 'Labor Day', None],
    ['2024-10-18', 'Fall Break', None],
    ['2024-10-21', 'Fall Break', None],
    ['2024-10-22', 'Advising', None],
    ['2024-10-23', 'Advising', None],
    ['2024-11-27', 'Thanksgiving', None],
    ['2024-11-28', 'Thanksgiving', None],
    ['2024-11-29', 'Thanksgiving', None],
    ['2024-12-13', 'Study', -1],

]
special_dates = [SpecialDate(*d) for d in special_dates]

duplicated_dates = [d for d, c in Counter([d.date for d in special_dates]).items() if c > 1]
if duplicated_dates:
    st.warning(f"Warning: duplicated dates: {duplicated_dates}")


@dataclass
class AcademicEvent:
    """
    An event that follows the academic calendar.

    Fields:
    - pattern: the day-of-week pattern, like "MTWRFSU"
    - name: the name of the event
    - location: the location of the event
    - meeting_time: the time of the event, like "8:30 AM - 9:20 AM"
    - start_date: the first day of the event
    - end_date: the last day of the event
    """
    pattern: str
    name: str
    location: str
    meeting_time: str
    start_date: datetime.datetime
    end_date: datetime.datetime


def iter_meeting_dates(start_date: datetime.date, end_date: datetime.date, pattern: str, special_dates):
    '''Yield all meeting times for the given class, given a meeting pattern.'''
    one_day = datetime.timedelta(days=1)
    days = [letter_to_day(d) for d in pattern]
    cur = start_date
    semester_ended = False
    while cur <= end_date:
        effective_date = true_date = cur.weekday()
        for special in special_dates:
            if cur == special.date:
                effective_date = special.pattern
        normally_meets_today = true_date in days
        meets_today = effective_date in days and not semester_ended
        is_exception = normally_meets_today and not meets_today
        is_abnormal_meeting = not normally_meets_today and meets_today
        yield cur, meets_today, is_exception, is_abnormal_meeting
        if effective_date == -1:
            semester_ended = True
        cur += one_day


def parse_time(x):
    """Parse time strings like 1:00 PM into hour=13, min=0.
    
    >>> parse_time("9:55 AM")
    {'hour': 9, 'minute': 55}
    >>> parse_time("12:15 PM")
    {'hour': 12, 'minute': 15}
    >>> parse_time("1:00 PM")
    {'hour': 13, 'minute': 0}
    >>> parse_time("12:05 AM")
    {'hour': 0, 'minute': 5}
    """
    hour, min, meridian = re.match(r'^(\d+):(\d+) (AM|PM)', x).groups()
    hour = int(hour)
    min = int(min)
    if hour == 12:
        hour = 0
    if meridian == 'PM':
        hour += 12
    return dict(hour=hour, minute=min)

#import doctest
#doctest.run_docstring_examples(parse_time, globals())

def generate_ics(data):
    # Merge shadow reservations (multiple locations for the same course section and time)
    data['Location'] = data.groupby(['Course Section', 'Meeting Time'])['Location'].transform(lambda x: ', '.join(x))
    data = data.drop_duplicates(['Course Section', 'Meeting Time'])

    # Parse the "meeting time" field.
    parsed = pd.concat([data, data['Meeting Time'].str.extract(r'^(?P<days>\w+) \| (?P<time>[^|]+)')], axis=1)

    # Use single letters for each date ("R" instead of "TH" for Thursday)
    parsed['days'] = parsed['days'].str.replace('TH', 'R')

    parsed_internal = parsed.rename(
            columns={"days": "pattern", "Course Section": "name", "Location": "location", "time": "meeting_time", "Start Date": "start_date", "End Date": "end_date"}
        )[['pattern', 'name', 'location', 'meeting_time', 'start_date', 'end_date']]
    
    edited = parsed_internal

    academic_events = [
        AcademicEvent(**e._asdict())
        for e in edited.itertuples(index=False, name="AcademicEvent")
    ]
    print(len(edited))

    earliest_date = min(parsed['Start Date']).date()
    latest_date = max(parsed['End Date']).date()

    recurring_events = []
    for i, academic_event in enumerate(academic_events):
        section_name = academic_event.name
        meeting_time = academic_event.meeting_time

        if not isinstance(meeting_time, str):
            st.warning(f"Skipping {section_name} because no meeting times.")
            continue

        location = academic_event.location
        meeting_pattern = academic_event.pattern
        start_time, end_time = meeting_time.split(' - ')
        start_time_p = parse_time(start_time)
        end_time_p = parse_time(end_time)
        exceptions = []

        has_occurred = False
        occurrences = list(iter_meeting_dates(
            academic_event.start_date.date(),
            academic_event.end_date.date(),
            meeting_pattern,
            special_dates
        ))

        actual_occurrences = [occur for occur in occurrences if occur[1]]
        first_meeting_date = actual_occurrences[0][0]
        last_meeting_date = actual_occurrences[-1][0]

        exceptions_dates = [
            meeting_date
            for meeting_date, meets_today, is_exception, is_abnormal_meeting in occurrences
            if is_exception and meeting_date <= last_meeting_date
        ]

        recurring_events.append(
            recurring_event(
                first_date=first_meeting_date,
                last_date=last_meeting_date,
                summary=section_name,
                location=location,
                start_time_p=start_time_p,
                end_time_p=end_time_p,
                meeting_pattern=meeting_pattern,
                exceptions=exceptions_dates)
        )

        # HACK: Add new events for each abnormal meeting.
        for meeting_date, meets_today, is_exception, is_abnormal_meeting in occurrences:
            if not is_abnormal_meeting:
                continue
            recurring_events.append(
                recurring_event(
                    first_date=meeting_date,
                    last_date=None,
                    summary=section_name,
                    location=location,
                    start_time_p=start_time_p,
                    end_time_p=end_time_p,
                    meeting_pattern=None,
                    exceptions=[])
            )

    relevant_special_dates = [
        special
        for special in special_dates
        if earliest_date <= special.date <= latest_date
    ]    

    if len(relevant_special_dates) > 0 and st.checkbox("Include special dates? ({})".format(
        ', '.join(special.name for special in relevant_special_dates)
    ), value=True):
        all_day_events = [
            all_day_event(special.date, special.name)
            for special in relevant_special_dates
        ]
    else:
        all_day_events = []

    # Actually returns an ICS file as a string, not writing to a file.
    ics_string = write_ics(
        all_day_events + recurring_events
    )

    return ics_string
