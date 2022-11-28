from typing import List
import streamlit as st
import re
import pandas as pd
import datetime

from ical_writer import all_day_event, recurring_event, write_ics

#st.set_page_config(layout="wide")

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
    ['2023-03-03', 'Spring Break', None],
    ['2023-03-21', 'Advising', None],
    ['2023-03-22', 'Advising', None],
    ['2023-04-07', 'Good Friday', None],
    ['2023-04-10', 'Easter Monday', None],
    ['2023-04-20', 'Thursday with Friday schedule', "F"],
    ['2023-04-21', 'Exams Start', -1],
]
special_dates = [SpecialDate(*d) for d in special_dates]


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

st.title("Workday Schedule Converter (Fall 2022 only)")
st.write("by Ken Arnold (CS and Data Science) [Source code](https://github.com/kcarnold/calgen)")
st.write("""
To use:

1. Go to your Teaching Schedule or Current Classes in Workday
2. Click the button in the top right of the table to export it to Excel format.
3. Drag and drop the resulting file to the box below.
4. Click the download button that will soon appear to save the calendar file.
5. Double-click or drag-and-drop the file into your calendar. (For Outlook, Google Calendar, macOS Calendar, etc.)

If you encounter any problems, please email your Excel file to ka37@calvin.edu.

- 2022-11-28: Load Spring 2023 schedule, fix bugs.
- 2022-09-10: Make recurring events instead of individual occurrences.
""")

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

    earliest_date = min(parsed['Start Date']).date()
    latest_date = max(parsed['End Date']).date()

    recurring_events = []
    for i in range(len(parsed)):
        section_name = parsed['Course Section'].iloc[i]
        meeting_time = parsed['time'].iloc[i]

        if not isinstance(meeting_time, str):
            st.warning(f"Skipping {section_name} because no meeting times.")
            continue

        location = parsed['Location'].iloc[i]
        meeting_pattern = parsed['days'].iloc[i]
        start_time, end_time = meeting_time.split(' - ')
        start_time_p = parse_time(start_time)
        end_time_p = parse_time(end_time)
        exceptions = []

        has_occurred = False
        occurrences = list(iter_meeting_dates(
            parsed['Start Date'].iloc[i].date(),
            parsed['End Date'].iloc[i].date(),
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

    all_day_events = [
        all_day_event(special.date, special.name)
        for special in special_dates
        if earliest_date <= special.date <= latest_date
    ]

    ics_string = write_ics(
        all_day_events + recurring_events
    )

    st.download_button(
        label="Download .ics file",
        data=ics_string,
        file_name="teaching_schedule.ics",
        mime="text/calendar"
    )

    st.write("""I recommend importing this into an unused calendar first, to test it.""")

    if st.checkbox("Show all events (debugging) (may have the incorrect time zone)"):
        import icalendar
        import recurring_ical_events

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
                "day": begin.strftime("%a %b %d"),
                "time": f"{start} - {end}"
            })

        cal_table = pd.DataFrame(cal_events)
        cal_table['short_name'] = cal_table['name'].str.extract(r'^(\w+ \d+)')

        for title, data in cal_table.groupby('short_name', dropna=False):
            col_names = ['day', 'time', 'name', 'location']
            if pd.isna(title):
                title = 'Special Events'
                col_names = ['day', 'name']

            st.markdown("## " + title)
            data = data.sort_values('begin')
            st.write(data[col_names].style.hide_index().to_html(), unsafe_allow_html=True)
