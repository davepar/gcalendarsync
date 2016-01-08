# gcalendarsync
Apps Script for syncing a Google Spreadsheet with Google Calendar.

This script will add, update, and remove events in a Google Calendar to match entries in a
Google Spreadsheet.

## Limitations

The spreadsheet is considered the definitive source for events. Calendar entries
will be added, updated, and removed to make the calendar match the spreadsheet. Any changes made
manually in the calendar will be wiped out.

Recurring events are not currenty supported.

The user's current time zone is used when none is specified.

## Set Up

* Create a new Google Calendar (in the dropdown next to "My calendars" in the left sidebar).
* Give the Calendar a name and change other fields as desired.
* Open the new Calendar's settings ("Calendar settings" in the dropdown next to the calendar name).
* Copy the "Calendar ID" from the Calendar Address section. It should look like an email address.
* Create a Google Spreadsheet at http://drive.google.com.
* In the Tools menu, select Script Editor.
* Give the project a name.
* Paste in the code from gcalendarsync.js.
* Paste in the Calendar ID from above in the "calendarId" value in the script.
* Save the script.

Now set up the spreadsheet. You can either copy
[this spreadsheet](https://docs.google.com/spreadsheets/d/1vRMycgL3wHSdYaww8Ony0_6ajZZN_FpVvKaefPJg7gI)
(use File -> Make a copy), or create one from scratch following these steps:

* Switch back to the spreadsheet, and create columns with these names (can be in any order):
  * Title - event title
  * Description - event description
  * Start Time - start date and time for the event, e.g. "1/27/2016 5:25pm". Can be just a date
    for all-day events.
  * End Time - end date and time for the event. Ignored for all-day events.
  * All Day Event - true/false value.
  * Location - optional event location.
  * Id - used for syncing with calendar. This column could be hidden to prevent accidental edits.

That's it. Start entering and modifying events. You can add extra columns, and they'll be ignored.

Multi-day all-day events must have one entry for each day. This is a limitation of Calendar.

## Use

In the spreadsheet, look for a "Calendar Sync" menu and choose "Update Calendar".
Depending on the number of changes, the script runs in a couple seconds to a couple of minutes.
If it runs more than several minutes, the script will be stopped. Try adding or updating events
in smaller batches. Maybe 200 at a time?
