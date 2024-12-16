# Description:
Convert an excel weekly timetable with merged cell to ical.

# Motive:
Nothing justifiable, just felt really stubborn 1 night and want to do things via code.

# Source:
icsconverter shamelessly stolen from
https://github.com/n8henrie/icsConverter/blob/master/icsconverter.py

include slight modification to get things running
- removed some easygui mandatory pop up because tkinter tcl have issue with standalone python exec from uv.
- remove tzinfo reference from within icalendar because deprecated.

# How to use:
1. populate `weekly timetable template.xlsx`
2. run `uv run hello.py` or `python hello.py`