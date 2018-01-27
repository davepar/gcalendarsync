# gcalendarsync
Apps Script for syncing a Google Spreadsheet with Google Calendar. Two commands are added in a
"Calendar Sync" menu. One copies events from a calendar into a spreadsheet. The other commands
goes in the opposite direction, spreadsheet to calendar.

I started this project for creating and updating a quarterly calendar for my swimming group.
It's much easier to type a work out schedule into a spreadsheet than Google Calendar. I'm no
longer actively using the script, but happy to work on bugs or features occasionally. Also
feel free to make fixes yourself and send me pull requests. (Oct 2017)

## Limitations

WARNING: Events may be removed! If you're copying to the calendar, any event not found in the
spreadsheet will be deleted! Likewise copying to the spreadsheet will delete any rows not found
in the calendar. It's a good idea to try copying into a fresh spreadsheet tab as an experiment
the first time you run it. "Undo" may also save you.

Recurring events are not currenty supported.

Google Calendar doesn't support multi-day "all day" events in the API. This causes the "all day"
indicator to disappear when syncing multi-day events to the spreadsheet and back to the calendar.
Leave the end time blank in the spreadsheet to create an all day event for one day.

## Set Up

Part 1. Set up the calendar:
* Create a new Google Calendar (in the dropdown next to "My calendars" in the left sidebar
  of Calendar).
* Give the calendar a name and change other fields as desired, i.e. set up sharing.
* Open the new calendar's settings ("Settings and sharing" in the dropdown next to the calendar name).
* Scroll down to the "Integrate calendar" section. Copy the "Calendar ID". It should look like an
  email address.

You have 2 options at this point. Copy an example spreadsheet that already has the script set up,
or use your own spreadsheet and add the script to it. The first option is a little easier. Using
your own spreadsheet will just take a little extra work to set up the column headers correctly.

Part 2a. Copy and modify the example spreadsheet:
* Make a copy of
  [this spreadsheet](https://docs.google.com/spreadsheets/d/1b0BBnmoDT4uDbN0pYsH--mpasFR45QlgNMTwUH-7MqU)
  (use File -> Make a copy).
* In the Tools menu, select Script Editor.
* Replace the "calendarId" value near the beginning of the script with the Calendar ID from Part 1,
  above.
* Set the correct time zone in File, Spreadsheet settings.
* Save the script.

Part 2b. If instead you want to create a new spreadsheet from scratch, or use one you already have:
* Create or open a Google Spreadsheet at http://drive.google.com.
* Create columns with these names (can be in any order):
  * Title - event title
  * Description - event description (optional)
  * Start Time - start date and time for the event, e.g. "1/27/2016 5:25pm". Should be just a date
    for all-day events. Set the format of this column to "Date time".
  * End Time - end date and time for the event. Set to blank for a one day all-day event.  Set the
    format of this column to "Date time".
  * Location - event location. (optional)
  * Guests - comma separated list of guest email addresses. (optional)
  * Color - a number from 1 to 11 that represents a color to set on the event. See the
    [list of colors](https://developers.google.com/apps-script/reference/calendar/event-color).
    (optional)
  * Id - used for syncing with calendar. This column could be hidden to prevent accidental edits.
* In the Tools menu, select Script Editor.
* Give the project a name.
* Paste in the code from
  [gcalendarsync.js](https://raw.githubusercontent.com/Davepar/gcalendarsync/master/gcalendarsync.js).
* For the "calendarId" value in the script, paste in the Calendar ID from above.
* Save the script.

That's it. Start entering and modifying events. You can add extra columns, and they'll be ignored.
See the "How to Sync" section below for how to run the script.

## Configuration options

There are two variables near the top of the script that can be modified:
* years - Set a range of years to synchronize. Defaults to 1970 to 2030.
* dateFormat - The date/time format to use when setting up the spreadsheet.

## Custom column names

Custom column names are now supported. In the script, find the "titleRowMap" variable. Change the
second entry on each line to match your column names. If you're not using one of the
optional fields, just leave it in titleRowMap.

## Time zones

There doesn't seem to be a way to enter a time zone for individual events into Google Spreadsheet.
The only option is to change the timezone for the entire spreadsheet. Look in the File menu,
Spreadsheet settings ([more info](https://support.google.com/docs/answer/58515?hl=en)).

## How to Sync

In the spreadsheet, select the desired sheet (tabs at the bottom of the spreadsheet) and then look
for a "Calendar Sync" menu. Choose "Update from Calendar" or "Update to Calendar" depending on the
direction you want to sync. The first time you do this, an "Authorization Required" dialog will pop
up. Click "Continue" and select your account in the next dialog. Then a somewhat scary dialog
will appear. Google is trying to protect against some malicious Apps scripts that were floating
around. I haven't jumped through the steps yet to become a trusted developer. By publishing the
source for this script, everybody can verify for themselves that it isn't doing anything nefarious.
If you trust that, click on "Advanced" and then the link at the bottom ending in "unsafe". Finally,
click "Allow" in the next dialog.

Depending on the number of changes, the script runs in a few seconds to a few minutes.

## Automatically syncing

You can also set up the script to automatically update a calendar whenever the sheet is
updated. This is really handy if your sheet is associated with a form for adding events.

Use the Run -> Run function menu to execute the "createSpreadsheetEditTrigger" function
one time. You will need to approve some special permissions. A popup dialog will say
"This app isn't verified". This is because the spreadsheet will be modifying the calendar
even when you aren't logged in. You can get around this by clicking "Advanced" in the
dialog and then clicking on your spreadsheet name. Approve the permissions in the next
dialog. You can modify the trigger by reading the
[documentation](https://developers.google.com/apps-script/guides/triggers/events).

To remove the trigger, use the same menu command to run the deleteTrigger function.

IMPORTANT: Be careful who has permissions to edit the spreadsheet and script. Once you
set up the trigger to run, someone else could modify the script maliciously.

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
