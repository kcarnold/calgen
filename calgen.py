from dataclasses import dataclass
import datetime
import re
import warnings
from collections import Counter
from typing import List

import icalendar
import pandas as pd
import recurring_ical_events
import streamlit as st
from calendar_view.calendar import Calendar
from calendar_view.core import data as calendar_view_data
from calendar_view.core.calendar_events import CalendarEvents
from calendar_view.core.calendar_grid import CalendarGrid
from calendar_view.core.event import Event as CVEvent
from calendar_view.core.event import EventStyles

from ical_writer import all_day_event, recurring_event, write_ics

import csv

# Ignore warnings about missing default styles in openpyxl
# openpyxl/styles/stylesheet.py:226: UserWarning: Workbook contains no default style, apply openpyxl's default
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

#st.set_page_config(layout="wide")

def letter_to_day(d):
    return 'MTWRFSU'.index(d)

# Special dates
# Third entry is the pattern: what day-of-week it corresponds to. See iter_meeting_dates.

class SpecialDate:
    def __init__(self, date, name: str, pattern):
        if isinstance(date, str):
            date =  datetime.date.fromisoformat(date)
        self.date = date
        self.name = name
        if isinstance(pattern, str) and len(pattern) == 1:
            pattern = letter_to_day(pattern)
        elif pattern == 'END_OF_SEMESTER':
            pattern = pattern
        elif pattern == '': # which means no class
            pattern = ""
        else:
            raise ValueError(f"Invalid pattern {pattern}")
        self.pattern = pattern

# Load special dates from CSV file
def load_special_dates(file_path):
    special_dates = []
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            date = row['date']
            name = row['name']
            pattern = row['pattern']
            special_dates.append([date, name, pattern])
    return special_dates

special_dates = load_special_dates('special_dates.csv')
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
        if effective_date == 'END_OF_SEMESTER':
            semester_ended = True
        cur += one_day


def get_sample_week_events(pattern: str, sample_week_start: datetime.date, start_time, end_time, title: str):
    '''Get a sample week of events for the given class.'''
    start_hour = start_time['hour']
    start_min = start_time['minute']
    end_hour = end_time['hour']
    end_min = end_time['minute']
    days = [letter_to_day(d) for d in pattern]
    # Find a Monday after the sample week start.
    cur = sample_week_start
    while cur.weekday() != 0:
        cur += datetime.timedelta(days=1)
    # Collect events for the week.
    sample_week_events = []
    for i in range(7):
        if i in days:
            sample_week_events.append(
                CVEvent(day=cur, start=f"{start_hour:02d}:{start_min:02d}", end=f"{end_hour:02d}:{end_min:02d}", title=title))
        cur += datetime.timedelta(days=1)
    return sample_week_events


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

st.title("Workday Schedule Converter")
st.write("by Ken Arnold (CS and Data Science) [Source code](https://github.com/kcarnold/calgen) Updated through Fall 2024.")
st.write("""
To use:

1. Go to your Teaching Schedule or Current Classes in Workday
2. Click the button in the top right of the table to export it to Excel format.
3. Drag and drop the resulting file to the box below.
4. Click the download button that will soon appear to save the calendar file.
5. Double-click or drag-and-drop the file into your calendar. (For Outlook, Google Calendar, macOS Calendar, etc.)
6. To make calendar events into Teams meetings, edit it in Outlook calendar view.

If you encounter any problems, please email your Excel file to ka37@calvin.edu.

<details><summary>Changelog</summary>

- 2024-06-17: Update for Fall 2024
- 2023-08-20: Add weekly calendar view and table of meetings
- 2023-05-31: Update through Summer 2024. (Does not include Summer 2023.)
- 2023-02-23: Fix duplicated Spring Break date.
- 2023-01-10: Add note about how to make a Teams meeting (thanks, Mark Muyskens)
- 2022-11-28: Load Spring 2023 schedule, fix bugs.
- 2022-09-10: Make recurring events instead of individual occurrences.

</details>
""", unsafe_allow_html=True)

st.header("Upload!")
uploaded_file = st.file_uploader("The Excel file exported from Workday goes here.")


def get_shortnames(items):
    shortnames = {
        loc: st.text_input("Short name for " + loc, loc, help="Type the short name here and press Enter").strip()
        for loc in sorted(set(items)) if loc
    }
    return [shortnames.get(it, it) for it in items]


expected_columns = ['Course Section', 'Meeting Time', 'Location', 'Start Date', 'End Date']
def load_file(uploaded_file):
    # Read the input file.
    def read_file(**kwargs):
        return pd.read_excel(uploaded_file, na_filter=False, dtype={"Location": str}, **kwargs)

    data = read_file()

    if data.columns[0] == 'My Enrolled Courses':
        # This is a student schedule.
        for idx in range(len(data)):
            if 'Start Date' in data.iloc[idx]:
                break
        else:
            st.error("The file is missing some data we expect. Please email the file to ka37@calvin.edu.")
            print("Error")
            print(data)
            st.stop()
        data = read_file(skiprows=idx + 1)
        data.rename({"Section": "Course Section"}, inplace=True)
        # Gotta parse out "Meeting Pattern"s...
        return data

    if data.columns[0].startswith("View My"):
        # There are two Excel export buttons on the Workday page. The one on the top includes some header data.
        # Skip header rows until we get to "Course Section"
        for idx, entry in enumerate(data.iloc[:, 0]):
            if entry in expected_columns:
                break
        else:
            st.error("The file is missing some data we expect. Please email the file to ka37@calvin.edu.")
            st.write(data)
            st.stop()
        data = read_file(skiprows=idx + 1)

    if 'Status' in data.columns:
        data = data[data['Status'] != 'Canceled']
    return data


if uploaded_file is not None:
    st.header("Download!")

    data = load_file(uploaded_file)
    if not all(col in data.columns for col in expected_columns):
        st.error("The file is missing some columns we expect. Please email the file to ka37@calvin.edu.")
        st.stop()

    # Merge shadow reservations (multiple locations for the same course section and time)
    data['Location'] = data.groupby(['Course Section', 'Meeting Time'])['Location'].transform(lambda x: ', '.join(x))
    data = data.drop_duplicates(['Course Section', 'Meeting Time'])

    # Parse the "meeting time" field.
    parsed = pd.concat([data, data['Meeting Time'].str.extract(r'^(?P<days>\w+) \| (?P<time>[^|]+)')], axis=1)

    # Use single letters for each date ("R" instead of "TH" for Thursday)
    parsed['days'] = parsed['days'].str.replace('TH', 'R')

    with st.expander(label = "Use abbreviations for names and locations? (Recommended!)", expanded=True):
        st.subheader("Sections")
        parsed['Course Section'] = get_shortnames(parsed['Course Section'])
        st.subheader("Locations")
        parsed['Location'] = get_shortnames(parsed['Location'])

    parsed_internal = parsed.rename(
            columns={"days": "pattern", "Course Section": "name", "Location": "location", "time": "meeting_time", "Start Date": "start_date", "End Date": "end_date"}
        )[['pattern', 'name', 'location', 'meeting_time', 'start_date', 'end_date']]
    
    edited = st.data_editor(parsed_internal, num_rows='dynamic', hide_index=True)

    academic_events = [
        AcademicEvent(**e._asdict())
        for e in edited.itertuples(index=False, name="AcademicEvent")
    ]
    print(len(edited))

    earliest_date = min(parsed['Start Date']).date()
    latest_date = max(parsed['End Date']).date()

    week_events = []
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

        week_events.extend(
            get_sample_week_events(
                pattern=meeting_pattern,
                sample_week_start=first_meeting_date,
                start_time = start_time_p,
                end_time = end_time_p,
                title=section_name,
            ))

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

    special_date_counter = Counter(special.name for special in relevant_special_dates)
    if len(relevant_special_dates) > 0 and st.checkbox("Include special dates? ({})".format(
        ', '.join(f"{name} ({count})" for name, count in special_date_counter.items())
    ), value=True):
        all_day_events = [
            all_day_event(special.date, special.name)
            for special in relevant_special_dates
        ]
    else:
        all_day_events = []

    ics_string = write_ics(
        all_day_events + recurring_events
    )

    st.download_button(
        label=":calendar: :floppy_disk: Download .ics file",
        data=ics_string,
        file_name="teaching_schedule.ics",
        mime="text/calendar"
    )

    st.write("""I recommend importing this into an unused calendar first, to test it.""")

    # Monkey-patch the calendar view to use US-locale day names
    def _get_day_title(self, day: datetime.date) -> str:
        return day.strftime("%a")
    
    CalendarGrid._get_day_title = _get_day_title
    CalendarEvents._get_day_title = _get_day_title

    beginning_of_week = earliest_date
    end_of_week = earliest_date + datetime.timedelta(days=4)
    
    cal_view = Calendar.build(calendar_view_data.CalendarConfig(
        lang = "en",
        dates = f"{beginning_of_week} - {end_of_week}",
        hours = "8 - 22",
    ))


    cal_view.add_events(week_events)
    cal_view.events.group_cascade_events()
    cal_view._build_image()
    st.image(cal_view.full_image)


    if st.checkbox("Show meeting calendar"):
        calendar = icalendar.Calendar.from_ical(ics_string)

        raw_events = recurring_ical_events.of(calendar).between(earliest_date, latest_date)

        cal_events = []
        for evt in raw_events:
            begin = evt['DTSTART'].dt
            start = begin.strftime("%I:%M %p")
            end = evt['DTEND'].dt.strftime("%I:%M %p")
            cal_events.append({
                "name": evt.decoded('SUMMARY').decode('utf-8'),
                "location": evt.decoded("LOCATION").decode('utf-8') if 'LOCATION' in evt else None,
                "begin": begin,
                "end": evt['DTEND'].dt,
                "day": begin.strftime("%a %b %d"),
                "day_of_week": begin.strftime("%a"),
                "time": f"{start} - {end}"
            })


        cal_table = pd.DataFrame(cal_events)
        cal_table['begin'] = pd.to_datetime(cal_table['begin'], errors='raise', utc=True)
        cal_table['short_name'] = cal_table['name'].str.extract(r'^(\w+\s*\d+)').fillna('')

        classes_to_include = []
        for class_name in cal_table['short_name'].unique():
            if class_name == '' or st.checkbox(f"Include {class_name}?", value=True):
                classes_to_include.append(class_name)
        if len(classes_to_include) > 0:
            cal_table = cal_table[cal_table['short_name'].isin(classes_to_include)]

        include_times = st.checkbox("Include times?", value=False)

        # Reorganize the dataframe into a row per week and a column for each day of the week when the class meets.
        cal_table['week'] = cal_table['begin'].dt.isocalendar().week
        cal_table['week'] = cal_table['week'] - st.number_input("Shift week numbers by", value=cal_table['week'].min() - 1, step=1, min_value=-52, max_value=52)


        days_to_include = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        # Make sure that days are in the right order.
        days_to_include = [day for day in days_to_include if day in cal_table['day_of_week'].unique()]
        cal_table['day_of_week'] = pd.Categorical(cal_table['day_of_week'], categories=days_to_include, ordered=True)

        # Get the first meeting day of each week.
        cal_table['first_meeting'] = cal_table.groupby('week', dropna=False)['begin'].transform('min').dt.strftime("%b %d")

        # For each day in the week, concatenate the names and times of all of the events that occur on that day.
        # For all-day events, include the name of the event.
        row_per_day = cal_table.groupby(['week', 'first_meeting', 'day_of_week'], dropna=False).apply(lambda rows: '; '.join([
            (f"{row['time']}: " if include_times else '') + f"{row['name']}" if row['short_name'] != '' else row['name']
            for _, row in rows.iterrows()
        ]))

        # Unstack one level of the index to get a column for each day of the week.
        row_per_day = row_per_day.unstack().fillna('')
        row_per_day.columns.name = None
        st.write(row_per_day.to_html(), unsafe_allow_html=True)
        if st.checkbox("Show as Markdown"):
            st.code(row_per_day.reset_index().to_markdown(index=False))

        if False:
            for week, data in cal_table.groupby(cal_table['begin'].dt.week, dropna=False):
                st.markdown(f"## Week {week}")
                st.write(data[['day', 'time', 'name', 'location']].style.hide_index().to_html(), unsafe_allow_html=True)

            for title, data in cal_table.groupby('short_name', dropna=False):
                col_names = ['day', 'time', 'name', 'location']
                if pd.isna(title):
                    title = 'Special Events'
                    col_names = ['day', 'name']

                st.markdown("## " + title)
                data = data.sort_values('begin')
                st.write(data[col_names].style.hide(axis="index").to_html(), unsafe_allow_html=True)
