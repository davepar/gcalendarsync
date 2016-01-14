# gcalendarsync
Apps Script for syncing a Google Spreadsheet with Google Calendar. Events can be entered in a
spreadsheet, and this script will add, update, and remove events in a Google Calendar to match
the entries.

## Limitations

Any changes made manually in the calendar will be wiped out.
The spreadsheet is considered the definitive source for events. Calendar entries
will be added, updated, and removed to make the calendar match the spreadsheet.

Recurring events are not currenty supported.

The user's current time zone is used when none is specified.

## Set Up

Part 1. Set up the calendar:
* Create a new Google Calendar (in the dropdown next to "My calendars" in the left sidebar
  of Calendar).
* Give the calendar a name and change other fields as desired, i.e. set up sharing.
* Open the new calendar's settings ("Calendar settings" in the dropdown next to the calendar name).
* Copy the "Calendar ID" from the Calendar Address section. It should look like an email address.

You have 2 options at this point. Copy an example spreadsheet, or set up your own. I highly
recommend the first, because it's known to work. Creating or using your own spreadsheet may
take some extra fiddling with formats.

Part 2a. Copy and modify the example spreadsheet:
* Make a copy of
  [this spreadsheet](https://docs.google.com/spreadsheets/d/1b0BBnmoDT4uDbN0pYsH--mpasFR45QlgNMTwUH-7MqU)
  (use File -> Make a copy).
* In the Tools menu, select Script Editor.
* For the "calendarId" value in the script, paste in the Calendar ID from above.
* Save the script.

Part 2b. If instead you want to create a new spreadsheet from scratch, or use one you already have:
* Create or open a Google Spreadsheet at http://drive.google.com.
* Create columns with these names (can be in any order):
  * Title - event title
  * Description - event description
  * Start Time - start date and time for the event, e.g. "1/27/2016 5:25pm". Should be just a date
    for all-day events. Set the format of this column to "Date time".
  * End Time - end date and time for the event. Ignored for all-day events.  Set the format of this
    column to "Date time".
  * All Day Event - true/false value.
  * Location - optional event location.
  * Id - used for syncing with calendar. This column could be hidden to prevent accidental edits.
* In the Tools menu, select Script Editor.
* Give the project a name.
* Paste in the code from
  [gcalendarsync.js](https://raw.githubusercontent.com/Davepar/gcalendarsync/master/gcalendarsync.js).
* For the "calendarId" value in the script, paste in the Calendar ID from above.
* Save the script.

That's it. Start entering and modifying events. You can add extra columns, and they'll be ignored.

Multi-day all-day events must have one entry for each day. This is a limitation of Calendar. An
alternative is to create a regular event that spans multiple days, e.g. 3/1/2016 00:00 to
3/3/2016 23:59.

## How to Sync

In the spreadsheet, look for a "Calendar Sync" menu and choose "Update Calendar". The first
time you run it, a dialog will pop up asking for authorization to manage the calendar and spreadsheet.
Depending on the number of changes, the script runs in a couple seconds to a couple of minutes.

## Troubleshooting

When an error occurs, the sync can generally be tried again without any bad side effects.

There are some data checks in the script for correctly formatted dates and times. If you see a
pop-up dialog, it will tell you which event has the error. Fix the error and run it again. See the
[test spreadsheet](https://docs.google.com/spreadsheets/d/1b0BBnmoDT4uDbN0pYsH--mpasFR45QlgNMTwUH-7MqU)
for examples of correct and incorrect date/times.

If the script runs more than several minutes, it will be stopped. Try adding or updating
events in smaller batches--maybe a few hundred at a time? Or you can try reducing the number in the
"Utilities.sleep()" calls in the script.

If you get an error about too many Calendar calls in a short amount of time, try increasing the
values in Utilities.sleep().

Other issues? Contact me or file an issue in GitHub.
