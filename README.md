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

Follow these steps to set up the calendar, spreadsheet, and script:
1 Create a new Google Calendar (in the dropdown next to "My calendars" in the left sidebar
  of Calendar).
1 Give the calendar a name and change other fields as desired, i.e. set up sharing.
1 Open the new calendar's settings ("Calendar settings" in the dropdown next to the calendar name).
1 Copy the "Calendar ID" from the Calendar Address section. It should look like an email address.
1 Create a Google Spreadsheet at http://drive.google.com.
1 In the Tools menu, select Script Editor.
1 Give the project a name.
1 Paste in the code from gcalendarsync.js.
1 For the "calendarId" value in the script, paste in the Calendar ID from above.
1 Save the script.

Now set up the spreadsheet. You can either copy
[this spreadsheet](https://docs.google.com/spreadsheets/d/1vRMycgL3wHSdYaww8Ony0_6ajZZN_FpVvKaefPJg7gI)
(use File -> Make a copy), or create one from scratch following these steps:

* Switch back to the spreadsheet, and create columns with these names (can be in any order):
  * Title - event title
  * Description - event description
  * Start Time - start date and time for the event, e.g. "1/27/2016 5:25pm". Should be just a date
    for all-day events.
  * End Time - end date and time for the event. Ignored for all-day events.
  * All Day Event - true/false value.
  * Location - optional event location.
  * Id - used for syncing with calendar. This column could be hidden to prevent accidental edits.

That's it. Start entering and modifying events. You can add extra columns, and they'll be ignored.

Multi-day all-day events must have one entry for each day. This is a limitation of Calendar.

## How to Sync

In the spreadsheet, look for a "Calendar Sync" menu and choose "Update Calendar". The first
time you run it, a dialog will pop up asking for permissions to edit the calendar and spreadsheet.
Depending on the number of changes, the script runs in a couple seconds to a couple of minutes.

## Troubleshooting

When an error occurs, the sync can generally be tried again without any bad side effects.

If the script runs more than several minutes, it will be stopped. Try adding or updating
events in smaller batches--maybe a few hundred at a time? Or you can try reducing the number in the
"Utilities.sleep()" calls in the script.

If you get an error about too many Calendar calls in a short amount of time, try increasing the
values in Utilities.sleep().

Other issues? Contact me or file an issue in GitHub.
