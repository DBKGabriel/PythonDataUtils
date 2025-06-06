# PythonDataUtils
Holds python scripts I've written that don't fit perfectly into other repos.

**exceltosqlite.py**
Converts a .xlsx file into a sqlite database. Ran via command prompt. Opens a gui that lets the user select excel files to convert. Select multiple files by holding ctrl while you click them.

**time_utils.py**
Converts UTC to Eastern time and vice-versa (with daylight saving taken into account, like a good citizen). Accepts epoch timestamps (in seconds or milliseconds) or ISO-8601 datetimes, and is Thanos-approved. Includes a small CLI because I got tired of writing the same 3 lines over and over again. Falls back to pytz if zoneinfo isn't available, so it plays nice with older Pythons too. That's necessary because I have some programs that can't work with anything newer than 3.9