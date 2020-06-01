// Script to synchronize a calendar to a spreadsheet and vice versa.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//
// All settings are now located in a pop-dialog or in the Settings.ts file.

// These imports are only used for testing. Run pretest and posttest scripts to automatically
// uncomment and re-comment these lines.
/*% import {Settings, AllDayValue, showSettingsDialog, getUserSettings} from './Settings'; %*/
/*% import {Util} from './Util'; %*/
/*% import {EventColor, GenericEvent, GenericEventKey} from './GenericEvent'; %*/

// Create the add-on menu.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('GCalendar Sync')
    .addItem('Update from Calendar', 'syncFromCalendar')
    .addItem('Update to Calendar', 'syncToCalendar')
    .addItem('Settings', 'showSettingsDialog')
    .addToUi();
}

// Synchronize from calendar to spreadsheet.
function syncFromCalendar() {
  let userSettings = getUserSettings();
  if (!userSettings) {
    showSettingsDialog();
    return;
  }

  //Logger.log('Starting sync from calendar');
  // Loop through all sheets
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let calendarIdsFound: string[] = [];
  for (let sheet of allSheets) {
    const sheetName = sheet.getName();
    // Get sheet data and pull calendar ID from first row
    let data = sheet.getDataRange().getValues();
    if (data.length <= Util.CALENDAR_ID_ROW ||
      !data[Util.CALENDAR_ID_ROW][0].replace(/\s/g, '').toLowerCase().startsWith('calendarid')) {
      // Only sync sheets that start with "Calendar ID" in cell A1
      continue;
    }
  
    // Get calendar events
    const calendarId = data[Util.CALENDAR_ID_ROW][1];
    if (calendarIdsFound.indexOf(calendarId) >= 0) {
      if (Util.errorAlertHalt(`Calendar ID ${calendarId} is in more than one sheet. This can have unpredictable results.`)) {
        return;
      }
    }
    calendarIdsFound.push(calendarId);
    let calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Util.errorAlert(`Could not find calendar with ID ${calendarId} from sheet ${sheetName}.`);
      continue;
    }
    const calEvents = calendar.getEvents(userSettings.begin_date, userSettings.end_date);

    // Check if spreadsheet needs a title row added
    if (data.length <= Util.TITLE_ROW ||
      (data.length == Util.TITLE_ROW + 1 && data[Util.TITLE_ROW].length == 1 && data[Util.TITLE_ROW][0] === '')) {
      Util.setUpSheet(sheet, calEvents.length);
      // Refresh data from first two rows
      data = sheet.getDataRange().getValues().slice(0, Util.FIRST_DATA_ROW);
    }

    // Map spreadsheet column titles to indices
    const idxMap = Util.createIdxMap(data[Util.TITLE_ROW]);
    const idIdx = idxMap.indexOf('id');
    const startTimeIdx = idxMap.indexOf('starttime');

    // Verify title row has all required fields
    const includeAllDay = userSettings.all_day_events === AllDayValue.use_column;
    let missingFields = Util.missingRequiredFields(idxMap, includeAllDay);
    if (missingFields.length > 0) {
      Util.displayMissingFields(missingFields, sheetName);
      continue;
    }

    // Get all of the event IDs from the sheet
    const sheetEventIds = data.map(row => Util.generateUniqueId(row[idIdx],row[startTimeIdx]));

    // Loop through calendar events and update or add to sheet data
    let eventFound = new Array(data.length);
    for (let calEvent of calEvents) {
      const calEventId = calEvent.getId();
	  const calEventStartTime = calEvent.getStartTime();
      let rowIdx = sheetEventIds.indexOf(Util.generateUniqueId(calEventId,calEventStartTime));
      if (rowIdx < Util.FIRST_DATA_ROW) {
        // Event not found, create it
        rowIdx = data.length;
        let newRow = Array(idxMap.length).fill('');
        data.push(newRow);
      } else {
        eventFound[rowIdx] = true;
      }
      // Update event in spreadsheet data
      GenericEvent.fromCalendarEvent(calEvent).toSpreadsheetRow(idxMap, data[rowIdx]);
    }

    // Remove any data rows not found in the calendar from the bottom up
    let rowsDeleted = 0;
    for (let idx = eventFound.length - 1; idx > Util.TITLE_ROW; idx--) {
      if (!eventFound[idx] && sheetEventIds[idx]) {
        data.splice(idx, 1);
        rowsDeleted++;
      }
    }

    // Save spreadsheet changes
    let range = sheet.getRange(1, 1, data.length, data[Util.TITLE_ROW].length);
    range.setValues(data);
    if (rowsDeleted > 0) {
      sheet.deleteRows(data.length + 1, rowsDeleted);
    }
  }
  if (calendarIdsFound.length === 0) {
    Util.errorAlert('Could not find any calendar IDs in sheets. See Help for setup instructions.');
  }
}

// Synchronize from spreadsheet to calendar.
function syncToCalendar() {
  let userSettings = getUserSettings();
  if (!userSettings) {
    showSettingsDialog();
    return;
  }

  //Logger.log('Starting sync to calendar');
  let scriptStart = Date.now();

  // Loop through all sheets
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let calendarIdsFound: string[] = [];
  for (let sheet of allSheets) {
    const sheetName = sheet.getName();
    // Get sheet data and pull calendar ID from first row
    let range = sheet.getDataRange();
    let data = range.getValues();
    if (!data[Util.CALENDAR_ID_ROW][0].replace(/\s/g, '').toLowerCase().startsWith('calendarid')) {
      // Only sync sheets that have a calendar ID
      continue;
    }
    if (data.length < Util.FIRST_DATA_ROW + 1) {
      Util.errorAlert(`Sheet ${sheetName} must have a title row and at least one data row.`);
      continue;
    }
      
    // Get calendar events
    const calendarId = data[Util.CALENDAR_ID_ROW][1];
    if (calendarIdsFound.indexOf(calendarId) >= 0) {
      if (Util.errorAlertHalt(`Calendar ID ${calendarId} is in more than one sheet. This can have unpredictable results.`)) {
        return;
      }
    }
    calendarIdsFound.push(calendarId);
    let calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Util.errorAlert(`Could not find calendar with ID ${calendarId} from sheet ${sheetName}.`);
      continue;
    }
    const calEvents = calendar.getEvents(userSettings.begin_date, userSettings.end_date);
    let calEventIds = calEvents.map(val => val.getId());

    // Map column headers to indices
    let idxMap = Util.createIdxMap(data[Util.TITLE_ROW]);
    let idIdx = idxMap.indexOf('id');
    let idRange = range.offset(0, idIdx, data.length, 1);
    let idData = idRange.getValues()

    // Verify title row has all required fields
    const includeAllDay = userSettings.all_day_events === AllDayValue.use_column;
    let missingFields = Util.missingRequiredFields(idxMap, includeAllDay);
    if (missingFields.length > 0) {
      Util.displayMissingFields(missingFields, sheetName);
      continue;
    }
    let keysToAdd = Util.missingFields(idxMap);

    // Loop through sheet rows
    let numAdded = 0;
    let numUpdates = 0;
    let eventsAdded = false;
    for (let ridx = Util.FIRST_DATA_ROW; ridx < data.length; ridx++) {
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
      //Logger.log(`${sheetEvent.title} ${numUpdates} updates, time: ${Date.now() - scriptStart} msecs`);

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
          //Logger.log('%d events added, time: %d msecs', numAdded, Date.now() - scriptStart);
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
    const countNonNull = (prevVal: number, curVal: string) => curVal === null ? prevVal : prevVal + 1;
    const numToRemove = calEventIds.reduce(countNonNull, 0);
    if (numToRemove > 0) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(`Delete ${numToRemove} calendar event(s) not found in spreadsheet?`,
        ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
        let numRemoved = 0;
        calEventIds.forEach((id, idx) => {
          if (id != null) {
            calEvents[idx].deleteEvent();
            Utilities.sleep(Settings.THROTTLE_SLEEP_TIME);
            numRemoved++;
            // if (numRemoved % 10 === 0) {
            //   Logger.log('%d events removed, time: %d msecs', numRemoved, Date.now() - scriptStart);
            // }
          }
        });
      }
    }
  }
  if (calendarIdsFound.length === 0) {
    Util.errorAlert('Could not find any calendar IDs in sheets. See Help for setup instructions.');
  }
}

// Simple function to test syntax of this script, since otherwise it's not exercised until the
// code is uploaded via Clasp and run in Sheets.
export function exerciseSyntax() {
  return true;
}
