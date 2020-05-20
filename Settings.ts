/*% export %*/ const Settings = {
  // TODO: Move the comments from removed settings to the docs.

  // Set this value to match your calendar!!!
  // Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
  //const CALENDAR_ID = '<your-calendar-id>@group.calendar.google.com';
  //CALENDAR_ID: '3icu4ffi1iuh935ep5ffubgo3s@group.calendar.google.com',

  SETTINGS_VERSION: 'v1',


  // Set the beginning and end dates that should be synced. BEGIN_DATE can be set to Date() to use
  // today. The numbers are year, month, date, where month is 0 for Jan through 11 for Dec.
  //BEGIN_DATE: new Date(1970, 0, 1),  // Default to Jan 1, 1970
  //END_DATE: new Date(2500, 0, 1),  // Default to Jan 1, 2500

  // Set this value to true if all events are "all day" events, or false if none of them are. Or
  // set to null to use a column called "All Day" in the spreadsheet.
  //ALL_DAY_DEFAULT: null,

  // This controls whether email invites are sent to guests when the event is created in the
  // calendar. Note that any changes to the event will cause email invites to be resent.
  //SEND_EMAIL_INVITES: false,

  // Setting this to true will silently skip rows that have a blank start and end time
  // instead of popping up an error dialog.
  //SKIP_BLANK_ROWS: false,

  // Updating too many events in a short time period triggers an error. These values
  // were successfully used for deleting and adding 240 events. Values in milliseconds.
  THROTTLE_SLEEP_TIME: 200,
  MAX_RUN_TIME: 5.75 * 60 * 1000,
};

// Values for the "all day" setting. The enum names must match the IDs in the HTML radio
// button definitions.
/*% export %*/ enum AllDayValue {
  use_column = 'USE_COLUMN',
  always_all_day = 'ALWAYS_ALL_DAY',
  never_all_day = 'NEVER_ALL_DAY',
}
