import streamlit as st
import re
import pandas as pd
from ics import Calendar, Event
from datetime import date, timedelta

#st.set_page_config(layout="wide")

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
special_dates['date'] = pd.to_datetime(special_dates['date']).dt.tz_localize("US/Eastern")


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

st.title("Teaching Schedule Converter (Fall 2022 only)")
st.write("by Ken Arnold (CS and Data Science) [Source code](https://github.com/kcarnold/calgen)")
st.write("""
To use:

1. Go to your Teaching Schedule in Workday
2. Click the button in the top right to export it to Excel format.
3. Drag and drop the resulting file to the box below.
4. Click the download button that will soon appear to save the calendar file.
5. Double-click or drag-and-drop the file into your calendar. (For Outlook, Google Calendar, macOS Calendar, etc.)

If you encounter any problems, please email your Excel file to ka37@calvin.edu.
""")

st.header("Upload!")
uploaded_file = st.file_uploader("The Excel file exported from Workday goes here.")

cal = Calendar()
for i in range(len(special_dates)):
    evt = Event()
    evt.name = special_dates['desc'].iloc[i]
    evt.begin = special_dates['date'].iloc[i]
    evt.end = evt.begin
    evt.make_all_day()
    cal.events.add(evt)

def get_shortnames(items):
    shortnames = {loc: loc for loc in sorted(set(items))}
    for it in shortnames:
        shortnames[it] = st.text_input(it, it).strip()
    return [shortnames[it] for it in items]


if uploaded_file is not None:
    st.header("Download!")

    # Read the input file.
    def read_file(**kwargs):
        return pd.read_excel(uploaded_file, na_filter=False, dtype={"Location": str}, **kwargs)
    data = read_file()
    expected_columns = ['Course Section', 'Meeting Time', 'Location', 'Start Date', 'End Date']
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

    with st.expander(label = "Use abbreviations for names and locations? (Recommended!)", expanded=False):
        st.subheader("Sections")
        parsed['Course Section'] = get_shortnames(parsed['Course Section'])
        st.subheader("Locations")
        parsed['Location'] = get_shortnames(parsed['Location'])

    for i in range(len(parsed)):
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
            evt.name = parsed['Course Section'].iloc[i]
            evt.begin = pd.Timestamp(meeting_date).replace(**start_time_p).tz_localize('US/Eastern')
            evt.end = pd.Timestamp(meeting_date).replace(**end_time_p).tz_localize('US/Eastern')
            evt.location = parsed['Location'].iloc[i]
            cal.events.add(evt)


    st.download_button(
        label="Download .ics file",
        data=cal.serialize(),
        file_name="fall_2022_teaching.ics",
        mime="text/calendar"
    )

    st.write("""I recommend importing this into an unused calendar first, to test it.""")

    if st.checkbox("Show all events (debugging) (may have the incorrect time zone)"):
        cal_events = []
        for evt in cal.events:
            start = evt.begin.strftime("%I:%M %p")
            end = evt.end.strftime("%I:%M %p")
            cal_events.append({
                "name": evt.name,
                "location": evt.location,
                "begin": evt.begin,
                "day": evt.begin.strftime("%a %b %d"),
                "times": f"{start} - {end}"
            })

        cal_table = pd.DataFrame(cal_events)
        cal_table['short_name'] = cal_table['name'].str.extract(r'^(\w+ \d+)')

        for title, data in cal_table.groupby('short_name', dropna=False):
            col_names = ['day', 'times', 'name', 'location']
            if pd.isna(title):
                title = 'Special Events'
                col_names = ['day', 'name']

            st.markdown("## " + title)
            data = data.sort_values('begin')
            st.write(data[col_names].style.hide_index().to_html(), unsafe_allow_html=True)
