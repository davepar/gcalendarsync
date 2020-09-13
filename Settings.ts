/*% import {Util} from './Util'; %*/

// Values for the "all day" setting. The enum names must match the IDs in the HTML radio
// button definitions.
/*% export %*/ enum AllDayValue {
  use_column = 'USE_COLUMN',
  always_all_day = 'ALWAYS_ALL_DAY',
  never_all_day = 'NEVER_ALL_DAY',
}

// Defines the fields for user settings in the dialog and PropertiesService.
interface SettingsBaseType {
  begin_date: string;
  end_date: string;
  send_email_invites: boolean;
  skip_blank_rows: boolean;
  all_day_events: string;
}

/*% export %*/ class Settings {

  static readonly SETTINGS_VERSION = 'v1';

  // Updating too many events in a short time period triggers an error. These values
  // were successfully used for deleting and adding 240 events. Values in milliseconds.
  static readonly THROTTLE_SLEEP_TIME = 200;
  static readonly MAX_RUN_TIME = 5.75 * 60 * 1000;

  // These member names must match the names used in SettingsDialog.html.
  constructor(
    public begin_date: Date,
    public end_date: Date,
    public send_email_invites: boolean,
    public skip_blank_rows: boolean,
    public all_day_events: AllDayValue,
  ) { }

  // Show modal dialog for sync settings.
  static showSettingsDialog() {
    const html = HtmlService.createHtmlOutputFromFile('SettingsDialog');
    SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
  }

  static getDefaultSettings(): Settings {
    return new Settings(new Date(1970, 0, 1), new Date(2500, 0, 1), false, false, AllDayValue.use_column);
  }

  // Retrieves settings from storage.
  static loadFromPropertyService(propertiesService = PropertiesService) {
    const storedSettingsJson = propertiesService.getDocumentProperties().getProperty(Settings.SETTINGS_VERSION);
    const storedSettings = JSON.parse(storedSettingsJson) as SettingsBaseType;
    if (!storedSettings) {
      // Get defaults and store them.
      const defaultSettings = Settings.getDefaultSettings();
      Settings.saveToPropertyService(defaultSettings.convertToBaseTypes(), propertiesService);
      return defaultSettings;
    }
    // The JSON parser won't correctly parse dates, so manually do it
    const begin_date = new Date(storedSettings.begin_date);
    const end_date = new Date(storedSettings.end_date);
    return new Settings(begin_date, end_date, storedSettings.send_email_invites,
      storedSettings.skip_blank_rows, storedSettings.all_day_events as AllDayValue);
  }

  // Save user settings entered in modal dialog.
  static saveToPropertyService(formValues: SettingsBaseType, propertiesService = PropertiesService) {
    // Check that dates are valid.
    if (!Util.isValidDate(formValues.begin_date)) {
      throw ('Invalid start date');
    }
    if (!Util.isValidDate(formValues.end_date)) {
      throw ('Invalid end date');
    }
    propertiesService.getDocumentProperties().setProperty(Settings.SETTINGS_VERSION, JSON.stringify(formValues));
    return true;
  }

  // Convert to base type
  convertToBaseTypes(): SettingsBaseType {
    return {
      begin_date: Settings.convertDateToString(this.begin_date),
      end_date: Settings.convertDateToString(this.end_date),
      send_email_invites: this.send_email_invites,
      skip_blank_rows: this.skip_blank_rows,
      all_day_events: this.all_day_events,
    };
  }

  // Called by HTML script to get saved settings in a format compatible with the form.
  static convertForDialog(): SettingsBaseType {
    const userSettings = Settings.loadFromPropertyService();
    const convertedSettings = userSettings.convertToBaseTypes();
    convertedSettings.all_day_events = convertedSettings.all_day_events.toLowerCase();
    return convertedSettings;
  }

  // Formats a date for display in the settings dialog, e.g. 2020-3-1.
  static convertDateToString(datestr: Date): string {
    return `${datestr.getFullYear()}-${datestr.getMonth() + 1}-${datestr.getDate()}`;
  }
}

// This can be used during debugging from the Script Editor to remove all user settings.
function killUserSettings() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}

// Make these two static methods available to the dialog JS.
function convertForDialog() {
  return Settings.convertForDialog();
}
function saveToPropertyService(formValues: SettingsBaseType) {
  return Settings.saveToPropertyService(formValues);
}
