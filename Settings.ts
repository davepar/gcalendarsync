/*% import {Util} from './Util'; %*/

/*% export %*/ const Settings = {
  SETTINGS_VERSION: 'v1',

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

// Defines the fields for the user settings as stored in the Properties service. The
// property names must match the code for the settings dialog in SettingsDialog.html.
interface UserSettings {
  begin_date: string;
  end_date: string;
  send_email_invites: boolean;
  skip_blank_rows: boolean;
  all_day_events: AllDayValue;
}

// Defines the fields for user settings after parsing the dates.
interface ParsedUserSettings {
  begin_date: Date;
  end_date: Date;
  send_email_invites: boolean;
  skip_blank_rows: boolean;
  all_day_events: AllDayValue;
}

// Show modal dialog for sync settings.
/*% export %*/ function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog');
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// Retrieves settings from storage.
/*% export %*/ function getUserSettings(): ParsedUserSettings {
  let savedSettings = JSON.parse(PropertiesService.getDocumentProperties().getProperty(Settings.SETTINGS_VERSION));
  if (savedSettings) {
    // The JSON parser won't correctly parse dates, so manually do it
    savedSettings.begin_date = new Date(savedSettings.begin_date);
    savedSettings.end_date = new Date(savedSettings.end_date);
  }
  return savedSettings;
}

// Formats a date for display in the settings dialog, e.g. 2020-3-1.
function dateString(datestr: Date): string {
  return `${datestr.getFullYear()}-${datestr.getMonth() + 1}-${datestr.getDate()}`;
}

// Called by HTML script to get saved settings in a format compatible with the form.
function getUserSettingsForForm() {
  const savedSettings = getUserSettings();
  let result = {
    begin_date: '1970-1-1',
    end_date: '2500-1-1',
    send_email_invites: false,
    skip_blank_rows: false,
    all_day_events: AllDayValue.never_all_day.toLowerCase(),
  }
  if (savedSettings) {
    result.begin_date = dateString(savedSettings.begin_date);
    result.end_date = dateString(savedSettings.end_date);
    result.send_email_invites = savedSettings.send_email_invites;
    result.skip_blank_rows = savedSettings.skip_blank_rows;
    result.all_day_events = savedSettings.all_day_events.toLowerCase();
  }
  return result;
}

// Save user settings entered in modal dialog.
function saveUserSettings(formValues: UserSettings) {
  // Check that dates are valid.
  if (!Util.isValidDate(formValues.begin_date)) {
    throw('Invalid start date');
  }
  if (!Util.isValidDate(formValues.end_date)) {
    throw('Invalid end date');
  }
  PropertiesService.getDocumentProperties().setProperty(Settings.SETTINGS_VERSION, JSON.stringify(formValues));
  return true;
}

// This can be used during debugging from the Script Editor to remove all user settings.
function killUserSettings() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}
