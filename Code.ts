// Script to synchronize a calendar to a spreadsheet and vice versa.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//
// All settings are now located in a pop-dialog or in the Settings.ts file.

// These imports are only used for testing. Run pretest and posttest scripts to automatically
// uncomment and re-comment these lines.
/*% import {Settings, AllDayValue} from './Settings'; %*/
/*% import {Util} from './Util'; %*/
/*% import {GenericEvent} from './GenericEvent'; %*/

// Defines the fields for the user settings. The property names must match the code for the
// settings dialog in SettingsDialog.html.
interface UserSettings {
  calendar_id: string;
  begin_date: Date;
  end_date: Date;
  send_email_invites: boolean;
  skip_blank_rows: boolean;
  all_day_events: AllDayValue;
}

// Create the add-on menu.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('GCalendar Sync')
    .addItem('Update from Calendar', 'syncFromCalendar')
    .addItem('Update to Calendar', 'syncToCalendar')
    .addItem('Settings', 'showSettingsDialog')
    .addToUi();
}

// Set up formats and hide ID column for empty spreadsheet
function setUpSheet(sheet, fieldKeys) {
  // Date format to use in the spreadsheet. Meaning of letters defined at
  // https://developers.google.com/sheets/api/guides/formats
  const dateFormat = 'M/d/yyyy H:mm';
  sheet.getRange(1, fieldKeys.indexOf('starttime') + 1, 999).setNumberFormat(dateFormat);
  sheet.getRange(1, fieldKeys.indexOf('endtime') + 1, 999).setNumberFormat(dateFormat);
  sheet.hideColumns(fieldKeys.indexOf('id') + 1);
  // TODO: Is there a way to set up checkbox data validation?
  // TODO: Add checkbox instructions to README.
}

// Synchronize from calendar to spreadsheet.
function syncFromCalendar() {
  let userSettings = getUserSettings();
  if (!userSettings || !userSettings.calendar_id) {
    showSettingsDialog();
    return;
  }

  console.info('Starting sync from calendar');
  // Get calendar events
  let calendar = CalendarApp.getCalendarById(userSettings.calendar_id);
  if (!calendar) {
    Util.errorAlert('Cannot find calendar. Enter calendar ID in Settings. See Help for more info.');
    return;
  }
  const calEvents = calendar.getEvents(userSettings.begin_date, userSettings.end_date);

  // Get spreadsheet and data
  const spreadsheet:GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  let range = sheet.getDataRange();
  let data = range.getValues();
  let eventFound = new Array(data.length);

  // Check if spreadsheet is empty and add a title row
  const titleRow = Util.TITLE_ROW_KEYS.map(key => Util.TITLE_ROW_MAP[key]);
  if (data.length < 1) {
    data.push(titleRow);
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, Util.TITLE_ROW_KEYS);
  }

  if (data.length == 1 && data[0].length == 1 && data[0][0] === '') {
    data[0] = titleRow;
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, Util.TITLE_ROW_KEYS);
  }

  // Map spreadsheet headers to indices
  const idxMap = Util.createIdxMap(data[0]);
  const idIdx = idxMap.indexOf('id');

  // Verify header has all required fields
  const includeAllDay = userSettings.all_day_events === AllDayValue.use_column;
  let missingFields = Util.missingRequiredFields(idxMap, includeAllDay);
  if (missingFields.length > 0) {
    const reqFieldNames = missingFields.map(x => Util.TITLE_ROW_MAP[x]).join(', ');
    Util.errorAlert('Spreadsheet is missing ' + reqFieldNames + ' columns. See Help for more info.');
    return;
  }

  // Array of IDs in the spreadsheet
  const sheetEventIds = data.slice(1).map(row => row[idIdx]);

  // Loop through calendar events and put them in the spreadsheet data
  for (let calEvent of calEvents) {
    const calEventId = calEvent.getId();

    let ridx = sheetEventIds.indexOf(calEventId) + 1;
    if (ridx < 1) {
      // Event not found, create it
      ridx = data.length;
      let newRow = [];
      let rowSize = idxMap.length;
      while (rowSize--) newRow.push('');
      data.push(newRow);
    } else {
      eventFound[ridx] = true;
    }
    // Update event in spreadsheet data
    GenericEvent.fromCalendarEvent(calEvent).toSpreadsheetRow(idxMap, data[ridx]);
  }

  // Remove any data rows not found in the calendar
  let rowsDeleted = 0;
  for (let idx = eventFound.length - 1; idx > 0; idx--) {
    //event doesn't exists and has an event id
    if (!eventFound[idx] && sheetEventIds[idx - 1]) {
      data.splice(idx, 1);
      rowsDeleted++;
    }
  }

  // Save spreadsheet changes
  range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  if (rowsDeleted > 0) {
    sheet.deleteRows(data.length + 1, rowsDeleted);
  }
}

// Synchronize from spreadsheet to calendar.
function syncToCalendar() {
  let userSettings = getUserSettings();
  if (!userSettings || !userSettings.calendar_id) {
    showSettingsDialog();
    return;
  }

  console.info('Starting sync to calendar');
  let scriptStart = Date.now();
  // Get calendar and events
  let calendar = CalendarApp.getCalendarById(userSettings.calendar_id);
  if (!calendar) {
    Util.errorAlert('Cannot find calendar. Enter calendar ID in Settings. See Help for more info.');
    return;
  }
  let calEvents = calendar.getEvents(userSettings.begin_date, userSettings.end_date);
  let calEventIds = calEvents.map(val => val.getId());

  // Get spreadsheet and data
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getActiveSheet();
  let range = sheet.getDataRange();
  let data = range.getValues();
  if (data.length < 2) {
    Util.errorAlert('Spreadsheet must have a title row and at least one data row.');
    return;
  }

  // Map headers to indices
  let idxMap = Util.createIdxMap(data[0]);
  let idIdx = idxMap.indexOf('id');
  let idRange = range.offset(0, idIdx, data.length, 1);
  let idData = idRange.getValues()

  // Verify header has all required fields
  const includeAllDay = userSettings.all_day_events === AllDayValue.use_column;
  let missingFields = Util.missingRequiredFields(idxMap, includeAllDay);
  if (missingFields.length > 0) {
    let reqFieldNames = missingFields.map(x => Util.TITLE_ROW_MAP[x]).join(', ');
    Util.errorAlert('Spreadsheet is missing ' + reqFieldNames + ' columns. See help for more info.');
    return;
  }

  let keysToAdd = Util.missingFields(idxMap);

  // Loop through spreadsheet rows
  let numAdded = 0;
  let numUpdates = 0;
  let eventsAdded = false;
  for (let ridx = 1; ridx < data.length; ridx++) {
    let sheetEvent = GenericEvent.fromSpreadsheetRow(data[ridx], idxMap, keysToAdd, userSettings.all_day_events);

    // If enabled, skip rows with blank/invalid start and end times
    if (userSettings.skip_blank_rows && !(sheetEvent.starttime instanceof Date) &&
        !(sheetEvent.endtime instanceof Date)) {
      continue;
    }

    // Do some error checking first
    if (!sheetEvent.title) {
      Util.errorAlert('must have title', sheetEvent, ridx);
      continue;
    }
    if (!(sheetEvent.starttime instanceof Date)) {
      Util.errorAlert('start time must be a date/time', sheetEvent, ridx);
      continue;
    }
    if (!(sheetEvent.endtime instanceof Date)) {
      Util.errorAlert('end time must be a date/time', sheetEvent, ridx);
      continue;
    }
    if (sheetEvent.endtime < sheetEvent.starttime) {
      Util.errorAlert('end time must be after start time', sheetEvent, ridx);
      continue;
    }

    // Ignore events outside of the begin/end range desired.
    if (sheetEvent.starttime > userSettings.end_date) {
      continue;
    }
    if (sheetEvent.endtime < userSettings.begin_date) {
      continue;
    }

    // Determine if spreadsheet event is already in calendar and matches
    let addEvent = true;
    if (sheetEvent.id) {
      let eventIdx = calEventIds.indexOf(sheetEvent.id);
      if (eventIdx >= 0) {
        calEventIds[eventIdx] = null;  // Prevents removing event below
        addEvent = false;
        let calEvent = calEvents[eventIdx];
        let calGenericEvent = GenericEvent.fromCalendarEvent(calEvent);
        let eventDiffs = calGenericEvent.eventDifferences(sheetEvent);
        if (eventDiffs > 0) {
          // When there are only 1 or 2 event differences, it's quicker to
          // update the event. For more event diffs, delete and re-add the event.
          if (eventDiffs < 3) {
            numUpdates += calGenericEvent.updateEvent(sheetEvent, calEvent);
          } else {
            addEvent = true;
            calEventIds[eventIdx] = sheetEvent.id;
          }
        }
      }
    }
    console.info('%d updates, time: %d msecs', numUpdates, Date.now() - scriptStart);

    if (addEvent) {
      const eventOptions = {
        description: sheetEvent.description,
        location: sheetEvent.location,
        guests: sheetEvent.guests,
        sendInvites: userSettings.send_email_invites,
      }
      let newEvent: GoogleAppsScript.Calendar.CalendarEvent;
      if (sheetEvent.allday) {
        if (sheetEvent.endtime.getHours() === 23 && sheetEvent.endtime.getMinutes() === 59) {
          sheetEvent.endtime.setSeconds(sheetEvent.endtime.getSeconds() + 1);
        }
        newEvent = calendar.createAllDayEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent.endtime, eventOptions);
      } else {
        newEvent = calendar.createEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent.endtime, eventOptions);
      }
      // Put event ID back into spreadsheet
      idData[ridx][0] = newEvent.getId();
      eventsAdded = true;

      // Set event color
      const numericColor = parseInt(sheetEvent.color);
      if (numericColor > 0 && numericColor < 12) {
        newEvent.setColor(sheetEvent.color);
      }

      // Throttle updates.
      numAdded++;
      Utilities.sleep(Settings.THROTTLE_SLEEP_TIME);
      if (numAdded % 10 === 0) {
        console.info('%d events added, time: %d msecs', numAdded, Date.now() - scriptStart);
      }
    }
    // If the script is getting close to timing out, save the event IDs added so far to avoid lots
    // of duplicate events.
    if ((Date.now() - scriptStart) > Settings.MAX_RUN_TIME) {
      idRange.setValues(idData);
    }
  }

  // Save spreadsheet changes
  if (eventsAdded) {
    idRange.setValues(idData);
  }

  // Remove any calendar events not found in the spreadsheet
  let numToRemove = calEventIds.reduce((prevVal, curVal) => {
    if (curVal !== null) {
      prevVal++;
    }
    return prevVal;
  }, 0);
  if (numToRemove > 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Delete ' + numToRemove + ' calendar event(s) not found in spreadsheet?',
          ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      let numRemoved = 0;
      calEventIds.forEach((id, idx) => {
        if (id != null) {
          calEvents[idx].deleteEvent();
          Utilities.sleep(Settings.THROTTLE_SLEEP_TIME);
          numRemoved++;
          if (numRemoved % 10 === 0) {
            console.info('%d events removed, time: %d msecs', numRemoved, Date.now() - scriptStart);
          }
        }
      });
    }
  }
}

// Show modal dialog for sync settings.
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsDialog');
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

// Retrieves settings from storage.
function getUserSettings(): UserSettings {
  let savedSettings = JSON.parse(PropertiesService.getDocumentProperties().getProperty(Settings.SETTINGS_VERSION));
  if (savedSettings) {
    // The JSON parser won't correctly parse dates, so manually do it
    savedSettings.begin_date = new Date(savedSettings.begin_date);
    savedSettings.end_date = new Date(savedSettings.end_date);
  }
  Logger.log('in getUserSettings');
  Logger.log(savedSettings);
  return savedSettings;
}

// Formats a date for display in the settings dialog, e.g. 2020-3-1.
function dateString(datestr): string {
  return `${datestr.getFullYear()}-${datestr.getMonth() + 1}-${datestr.getDate()}`;
}

// Called by HTML script to get saved settings in a format compatible with the form.
function getUserSettingsForForm() {
  const savedSettings = getUserSettings();
  let result = {
    calendar_id: '',
    begin_date: '1970-1-1',
    end_date: '2500-1-1',
    send_email_invites: false,
    skip_blank_rows: false,
    all_day_events: AllDayValue.never_all_day.toLowerCase(),
  }
  if (savedSettings) {
    result.calendar_id = savedSettings.calendar_id || '';
    result.begin_date = dateString(savedSettings.begin_date);
    result.end_date = dateString(savedSettings.end_date);
    result.send_email_invites = savedSettings.send_email_invites;
    result.skip_blank_rows = savedSettings.skip_blank_rows;
    result.all_day_events = savedSettings.all_day_events.toLowerCase();
  }
  Logger.log('in getUserSettingsForForm');
  Logger.log(result);
  return result;
}

// Save user settings entered in modal dialog.
function saveUserSettings(formValues) {
  Logger.log('in saveUserSettings');
  Logger.log(formValues);
  if (formValues.calendar_id.indexOf('@') === -1) {
    formValues.calendar_id = formValues.calendar_id + '@group.calendar.google.com';
  }
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

// Set up a trigger to automatically update the calendar when the spreadsheet is
// modified. See the instructions for how to use this.
function createSpreadsheetEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('syncToCalendar')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
}

// Delete the trigger. Use this to stop automatically updating the calendar.
function deleteTrigger() {
  // Loop over all triggers.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'syncToCalendar') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

// Simple function to test syntax of this script, since otherwise it's not exercised until the
// code is uploaded via Clasp and run in Sheets.
export function exerciseSyntax() {
  return true;
}
