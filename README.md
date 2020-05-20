# gcalendarsync
A Google Sheet add-on for syncing events with Google Calendar. Want to easily enter a
bunch of events into Google Calendar? Enter them into a spreadsheet instead and use this script
to copy them to the calendar. Have a lot of events in a calendar, but need to modify them?
This add-on will make that easy as well. Sync from calendar to spreadsheet, make your changes,
and then sync back to the calendar.

There is also a [website](http://www.ballardsoftwarefoundry.com/gcalendarsync.html) for the add on.

## Install

Find the add-on in the GSuite Marketplace and install it from there.
(Note: Waiting on approval as of May 20, 2020.)

Two commands are added in a "GCalendar Sync" menu. One copies events from a calendar into
a spreadsheet. The other commands goes in the opposite direction, spreadsheet to calendar.

## Limitations

1. **WARNING:** Events may be removed! If you're copying to a calendar with existing events, those
  events will be deleted unless they are in the spreadsheet. Likewise copying to the spreadsheet
  will delete spreadsheet rows not found in the calendar. It's a good idea to try syncing with a
  fresh spreadsheet as an experiment the first time you run it. Or work with a copy of your
  spreadsheet (File -> Make a copy). "Undo" may save you, but it may not. Be careful!

2. Recurring events are not currenty supported.
  ([File an issue](https://github.com/Davepar/gcalendarsync/issues) if you're interested in this.)

## Set Up

The set up involves preparing a Google Calendar and Google Sheet.

**Calendar** Set up the calendar:
* Use an existing Google Calendar or create a new one (click the plus sign next to "Other calendars"
  in the left sidebar of Calendar).
* Give a new calendar a name and change other fields as desired, e.g. time zone. Exit settings
  and you should see the new calendar in the left sidebar.
* Open the calendar's settings by clicking on it's name.
* Scroll down to the "Integrate calendar" section. Copy the "Calendar ID". It should look like an
  email address.

**Sheet** You have 3 options for the Google Sheet: let the add-on create the spreadsheet columns
for you, copy an example spreadsheet, or use your own
existing spreadsheet. The first two options are easier. Using your own spreadsheet will just
take a little extra attention when setting up the column headers.

Step 1. All of the options below require the following:
* Install the GCalendar Sync add on.
* Open the GCalendar Sync Settings and paste in the calendar ID copied above.
* Go to File, Spreadsheet settings, and make sure the time zone matches the time zone you have set
  in your Calendar settings.

Option 2a. Let the add-on create the columns. This is a great option for an existing calendar and
an empty spreadsheet.
* Follow the steps in Step 1 above.
* Select Update from Calendar. This will create the required column headers for you.

Option 2b. Copy and modify the example spreadsheet.
* Make a copy of
  [this spreadsheet](https://docs.google.com/spreadsheets/d/1ap_PZXjgPtW1VA5LOQNCIWAGbIGyDfx8npnG_AAnE0c)
  (use File -> Make a copy).
* Follow the steps in Step 1 above.

Option 2c. Create a new spreadsheet yourself, or use one you already have.
* Create or open a Google Spreadsheet at http://drive.google.com.
* Follow the steps in Step 1 above.
* Create columns with these exact names (can be in any order and capitalization isn't significant):
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
  * All Day - when "use column" is selected in Settings, the value TRUE or FALSE will indicate
    whether an event is all day. This column is best formatted as a checkbox. Select all the cells
    in the column and then use the menu Insert -> Checkbox. (optional)
  * Id - used for syncing with calendar. This column should be hidden to prevent accidental edits.

That's it. Start entering and modifying events. Any extra columns that you add with
other names will be ignored. See the "How to Sync" section below for how to run the script.

## Settings

The following settings affect how events are synced:
* Start date, end date - Set these to sync up a smaller range of dates. When syncing *from* calendar,
  any rows outside the range will be removed from the sheet. Syncing *to* calendar simply ignores any
  events outside the range.
* Send email invites - Controls whether to send invites to guests when adding or updating events and
  syncing *to* calendar.
* Skip blank rows - Whether to show a warning dialog for blank rows in the data.
* All day events - Allows events by default to have a start/end time, or be all day. The third option
  requires a column called "All Day" that should be either the value TRUE or FALSE to indicate all day.

## How to Sync

In the spreadsheet, select the desired sheet (tabs at the bottom of the spreadsheet) and then look
for a "GCalendar Sync" menu. Choose "Update from Calendar" or "Update to Calendar" depending on the
direction you want to sync. The first time you do this, an "Authorization Required" dialog will pop
up. Click "Continue" and select your account in the next dialog.

Depending on the number of changes, the script runs in a few seconds to a few minutes.

## Troubleshooting

When an error occurs, the sync can generally be tried again without any bad side effects.

There are some data checks in the script for correctly formatted dates and times. If you see a
pop-up dialog, it will tell you which event has the error. Fix the error and run it again.

If the script runs more than several minutes, it will run out of time and be stopped. You should be
able to run it again and it will do the next batch of changes. About 900 calendar operations can be
done in one run, where an operation is updating one event field, adding an event, or removing an
event.

If you get an error about too many Calendar events being added or removed in a short amount of time,
try setting start/end date in settings to reduce the amount of events being synced. Or file an issue.

## Contributing

The project started as a simple script to publish swim team practice times on a Google Calendar
and has evolved from there. Fixes and improvements are done by volunteers, so please be patient.

You're also welcome to dive into the code and send suggested changes and fixes as a pull request
or as text. To try your own changes, you'll need to install Clasp in order to compile the Typescript
and push it to a project. Follow the instructions
[here](https://developers.google.com/apps-script/guides/typescript).

Running the tests requires running a special pre/post script to modify imports:
    $ ./pretest
    $ npm test
    $ ./posttest

Other issues? [Contact me](http://www.ballardsoftwarefoundry.com/gcalendarsync.html) or file an
issue in GitHub.
