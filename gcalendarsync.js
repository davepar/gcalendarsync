// Script to synchronize a calendar to a spreadsheet and vice versa.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//

// Set this value to match your calendar!!!
// Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
var calendarId = '<your-calendar-id>@group.calendar.google.com';

// Set the beginning and end dates that should be synced. beginDate can be set to Date() to use
// today. The numbers are year, month, date, where month is 0 for Jan through 11 for Dec.
// When no time is given, it will be set at 00:00:00
// (optional) Time can be set like new Date(1970, 0, 3, 22, 15) -> Jan 3, 1970, 22:15.
var beginDate = new Date(1970, 0, 1);  // Default to Jan 1, 1970, 00:00:00
var endDate = new Date(2500, 0, 1);  // Default to Jan 1, 2500, 00:00:00

// Date format to use in the spreadsheet.
var dateFormat = 'M/d/yyyy H:mm';

var titleRowMap = {
  'title': 'Title',
  'description': 'Description',
  'location': 'Location',
  'starttime': 'Start Time',
  'endtime': 'End Time',
  'guests': 'Guests',
  'color': 'Color',
  'id': 'Id'
};
var titleRowKeys = ['title', 'description', 'location', 'starttime', 'endtime', 'guests', 'color', 'id'];
var requiredFields = ['id', 'title', 'starttime', 'endtime'];

// This controls whether email invites are sent to guests when the event is created in the
// calendar. Note that any changes to the event will cause email invites to be resent.
var SEND_EMAIL_INVITES = false;

// Setting this to true will silently skip rows that have a blank start and end time
// instead of popping up an error dialog.
var SKIP_BLANK_ROWS = false;

// Updating too many events in a short time period triggers an error. These values
// were successfully used for deleting and adding 240 events. Values in milliseconds.
var THROTTLE_SLEEP_TIME = 200;
var MAX_RUN_TIME = 5.75 * 60 * 1000;

// Special flag value. Don't change.
var EVENT_DIFFS_WITH_GUESTS = 999;

// Adds the custom menu to the active spreadsheet.
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {
      name: "Update from Calendar",
      functionName: "syncFromCalendar"
    }, {
      name: "Update to Calendar",
      functionName: "syncToCalendar"
    }
  ];
  spreadsheet.addMenu('Calendar Sync', menuEntries);
}

// Creates a mapping array between spreadsheet column and event field name
function createIdxMap(row) {
  var idxMap = [];
  for (var idx = 0; idx < row.length; idx++) {
    var fieldFromHdr = row[idx];
    for (var titleKey in titleRowMap) {
      if (titleRowMap[titleKey] == fieldFromHdr) {
        idxMap.push(titleKey);
        break;
      }
    }
    if (idxMap.length <= idx) {
      // Header field not in map, so add null
      idxMap.push(null);
    }
  }
  return idxMap;
}

// Converts a spreadsheet row into an object containing event-related fields
function reformatEvent(row, idxMap, keysToAdd) {
  var reformatted = row.reduce(function(event, value, idx) {
    if (idxMap[idx] != null) {
      event[idxMap[idx]] = value;
    }
    return event;
  }, {});
  for (var k in keysToAdd) {
    reformatted[keysToAdd[k]] = '';
  }
  return reformatted;
}

// Converts a calendar event to a psuedo-sheet event.
function convertCalEvent(calEvent) {
  convertedEvent = {
    'id': calEvent.getId(),
    'title': calEvent.getTitle(),
    'description': calEvent.getDescription(),
    'location': calEvent.getLocation(),
    'guests': calEvent.getGuestList().map(function(x) {return x.getEmail();}).join(','),
    'color': calEvent.getColor()
  };
  if (calEvent.isAllDayEvent()) {
    convertedEvent.starttime = calEvent.getAllDayStartDate();
    var endtime = calEvent.getAllDayEndDate();
    if (endtime - convertedEvent.starttime === 24 * 3600 * 1000) {
      convertedEvent.endtime = '';
    } else {
      convertedEvent.endtime = endtime;
      if (endtime.getHours() === 0 && endtime.getMinutes() == 0) {
        convertedEvent.endtime.setSeconds(endtime.getSeconds() - 1);
      }
    }
  } else {
    convertedEvent.starttime = calEvent.getStartTime();
    convertedEvent.endtime = calEvent.getEndTime();
  }
  return convertedEvent;
}

// Converts calendar event into spreadsheet data row
function calEventToSheet(calEvent, idxMap, dataRow) {
  convertedEvent = convertCalEvent(calEvent);

  for (var idx = 0; idx < idxMap.length; idx++) {
    if (idxMap[idx] !== null) {
      dataRow[idx] = convertedEvent[idxMap[idx]];
    }
  }
}

// Returns empty string or time in milliseconds for Date object
function getEndTime(ev) {
  return ev.endtime === '' ? '' : ev.endtime.getTime();
}

// Determines the number of field differences between a calendar event and
// a spreadsheet event
function eventDifferences(convertedCalEvent, sev) {
  var eventDiffs = 0 + (convertedCalEvent.title !== sev.title) +
    (convertedCalEvent.description !== sev.description) +
    (convertedCalEvent.location !== sev.location) +
    (convertedCalEvent.starttime.toString() !== sev.starttime.toString()) +
    (getEndTime(convertedCalEvent) !== getEndTime(sev)) +
    (convertedCalEvent.guests !== sev.guests) +
    (convertedCalEvent.color !== ('' + sev.color));
  if (eventDiffs > 0 && convertedCalEvent.guests) {
    // Use a special flag value if an event changed, but it has guests.
    eventDiffs = EVENT_DIFFS_WITH_GUESTS;
  }
  return eventDiffs;
}

// Determine whether required fields are missing
function areRequiredFieldsMissing(idxMap) {
  return requiredFields.some(function(val) {
    return idxMap.indexOf(val) < 0;
  });
}

// Returns list of fields that aren't in spreadsheet
function missingFields(idxMap) {
  return titleRowKeys.filter(function(val) {
    return idxMap.indexOf(val) < 0;
  });
}

// Set up formats and hide ID column for empty spreadsheet
function setUpSheet(sheet, fieldKeys) {
  sheet.getRange(1, fieldKeys.indexOf('starttime') + 1, 999).setNumberFormat(dateFormat);
  sheet.getRange(1, fieldKeys.indexOf('endtime') + 1, 999).setNumberFormat(dateFormat);
  sheet.hideColumns(fieldKeys.indexOf('id') + 1);
}

// Display error alert
function errorAlert(msg, evt, ridx) {
  var ui = SpreadsheetApp.getUi();
  if (evt) {
    ui.alert('Skipping row: ' + msg + ' in event "' + evt.title + '", row ' + (ridx + 1));
  } else {
    ui.alert(msg);
  }
}

// Updates a calendar event from a sheet event.
function updateEvent(calEvent, convertedCalEvent, sheetEvent){
  var numChanges = 0;
  sheetEvent.sendInvites = SEND_EMAIL_INVITES;
  if (convertedCalEvent.starttime.toString() !== sheetEvent.starttime.toString() ||
      getEndTime(convertedCalEvent) !== getEndTime(sheetEvent)) {
    if (sheetEvent.endtime === '') {
      calEvent.setAllDayDate(sheetEvent.starttime);
    } else {
      calEvent.setTime(sheetEvent.starttime, sheetEvent.endtime);
    }
    numChanges++;
  }
  if (convertedCalEvent.title !== sheetEvent.title) {
    calEvent.setTitle(sheetEvent.title);
    numChanges++;
  }
  if (convertedCalEvent.description !== sheetEvent.description) {
    calEvent.setDescription(sheetEvent.description);
    numChanges++;
  }
  if (convertedCalEvent.location !== sheetEvent.location) {
    calEvent.setLocation(sheetEvent.location);
    numChanges++;
  }
  if (convertedCalEvent.color !== ('' + sheetEvent.color)) {
    if (sheetEvent.color > 0 && sheetEvent.color < 12) {
      calEvent.setColor('' + sheetEvent.color);
      numChanges++;
    }
  }
  if (convertedCalEvent.guests !== sheetEvent.guests) {
    var guestCal = calEvent.getGuestList().map(function (x) {
      return {
        email: x.getEmail(),
        added: false
      };
    });
    var sheetGuests = sheetEvent.guests || '';
    var guests = sheetGuests.split(',').map(function (x) {
      return x ? x.trim() : '';
    });
    // Check guests that are already invited.
    for (var gIx = 0; gIx < guestCal.length; gIx++) {
      var index = guests.indexOf(guestCal[gIx].email);
      if (index >= 0) {
        guestCal[gIx].added = true;
        guests.splice(index, 1);
      }
    }
    guests.forEach(function (guest) {
      if (guest) {
        calEvent.addGuest(guest);
        numChanges++;
      }
    });
    guestCal.forEach(function (guest) {
      if (!guest.added) {
        calEvent.removeGuest(guest.email);
        numChanges++;
      }
    });
  }
  // Throttle updates.
  Utilities.sleep(THROTTLE_SLEEP_TIME * numChanges);
  return numChanges;
}

// Synchronize from calendar to spreadsheet.
function syncFromCalendar() {
  console.info('Starting sync from calendar');
  // Get calendar and events
  var calendar = CalendarApp.getCalendarById(calendarId);
  var calEvents = calendar.getEvents(beginDate, endDate);

  // Get spreadsheet and data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();
  var data = range.getValues();
  var eventFound = new Array(data.length);

  // Check if spreadsheet is empty and add a title row
  var titleRow = [];
  for (var idx = 0; idx < titleRowKeys.length; idx++) {
    titleRow.push(titleRowMap[titleRowKeys[idx]]);
  }
  if (data.length < 1) {
    data.push(titleRow);
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }

  if (data.length == 1 && data[0].length == 1 && data[0][0] === '') {
    data[0] = titleRow;
    range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    setUpSheet(sheet, titleRowKeys);
  }

  // Map spreadsheet headers to indices
  var idxMap = createIdxMap(data[0]);
  var idIdx = idxMap.indexOf('id');

  // Verify header has all required fields
  if (areRequiredFieldsMissing(idxMap)) {
    var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
    errorAlert('Spreadsheet must have ' + reqFieldNames + ' columns');
    return;
  }

  // Array of IDs in the spreadsheet
  var sheetEventIds = data.slice(1).map(function(row) {return row[idIdx];});

  // Loop through calendar events
  for (var cidx = 0; cidx < calEvents.length; cidx++) {
    var calEvent = calEvents[cidx];
    var calEventId = calEvent.getId();

    var ridx = sheetEventIds.indexOf(calEventId) + 1;
    if (ridx < 1) {
      // Event not found, create it
      ridx = data.length;
      var newRow = [];
      var rowSize = idxMap.length;
      while (rowSize--) newRow.push('');
      data.push(newRow);
    } else {
      eventFound[ridx] = true;
    }
    // Update event in spreadsheet data
    calEventToSheet(calEvent, idxMap, data[ridx]);
  }

  // Remove any data rows not found in the calendar
  var rowsDeleted = 0;
  for (var idx = eventFound.length - 1; idx > 0; idx--) {
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
  console.info('Starting sync to calendar');
  var scriptStart = Date.now();
  // Get calendar and events
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    errorAlert('Cannot find calendar. Check instructions for set up.');
  }
  var calEvents = calendar.getEvents(beginDate, endDate);
  var calEventIds = calEvents.map(function(val) {return val.getId();});

  // Get spreadsheet and data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getDataRange();
  var data = range.getValues();
  if (data.length < 2) {
    errorAlert('Spreadsheet must have a title row and at least one data row');
    return;
  }

  // Map headers to indices
  var idxMap = createIdxMap(data[0]);
  var idIdx = idxMap.indexOf('id');
  var idRange = range.offset(0, idIdx, data.length, 1);
  var idData = idRange.getValues()

  // Verify header has all required fields
  if (areRequiredFieldsMissing(idxMap)) {
    var reqFieldNames = requiredFields.map(function(x) {return titleRowMap[x];}).join(', ');
    errorAlert('Spreadsheet must have ' + reqFieldNames + ' columns');
    return;
  }

  var keysToAdd = missingFields(idxMap);

  // Loop through spreadsheet rows
  var numAdded = 0;
  var numUpdates = 0;
  var eventsAdded = false;
  for (var ridx = 1; ridx < data.length; ridx++) {
    var sheetEvent = reformatEvent(data[ridx], idxMap, keysToAdd);

    // If enabled, skip rows with blank/invalid start and end times
    if (SKIP_BLANK_ROWS && !(sheetEvent.starttime instanceof Date) &&
        !(sheetEvent.endtime instanceof Date)) {
      continue;
    }

    // Do some error checking first
    if (!sheetEvent.title) {
      errorAlert('must have title', sheetEvent, ridx);
      continue;
    }
    if (!(sheetEvent.starttime instanceof Date)) {
      errorAlert('start time must be a date/time', sheetEvent, ridx);
      continue;
    }
    if (sheetEvent.endtime !== '') {
      if (!(sheetEvent.endtime instanceof Date)) {
        errorAlert('end time must be empty or a date/time', sheetEvent, ridx);
        continue;
      }
      if (sheetEvent.endtime < sheetEvent.starttime) {
        errorAlert('end time must be after start time for event', sheetEvent, ridx);
        continue;
      }
    }

    // Ignore events outside of the begin/end range desired.
    if (sheetEvent.starttime > endDate) {
      continue;
    }
    if (sheetEvent.endtime === '') {
      if (sheetEvent.starttime < beginDate) {
        continue;
      }
    } else {
      if (sheetEvent.endtime < beginDate) {
        continue;
      }
    }

    // Determine if spreadsheet event is already in calendar and matches
    var addEvent = true;
    if (sheetEvent.id) {
      var eventIdx = calEventIds.indexOf(sheetEvent.id);
      if (eventIdx >= 0) {
        calEventIds[eventIdx] = null;  // Prevents removing event below
        addEvent = false;
        var calEvent = calEvents[eventIdx];
        var convertedCalEvent = convertCalEvent(calEvent);
        var eventDiffs = eventDifferences(convertedCalEvent, sheetEvent);
        if (eventDiffs > 0) {
          // When there are only 1 or 2 event differences, it's quicker to
          // update the event. For more event diffs, delete and re-add the event. The one
          // exception is if the event has guests (eventDiffs=99). We don't
          // want to force guests to re-confirm, so go through the slow update
          // process instead.
          if (eventDiffs < 3 && eventDiffs !== EVENT_DIFFS_WITH_GUESTS) {
            numUpdates += updateEvent(calEvent, convertedCalEvent, sheetEvent);
          } else {
            addEvent = true;
            calEventIds[eventIdx] = sheetEvent.id;
          }
        }
      }
    }
    console.info('%d updates, time: %d msecs', numUpdates, Date.now() - scriptStart);

    if (addEvent) {
      var newEvent;
      sheetEvent.sendInvites = SEND_EMAIL_INVITES;
      if (sheetEvent.endtime === '') {
        newEvent = calendar.createAllDayEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent);
      } else {
        newEvent = calendar.createEvent(sheetEvent.title, sheetEvent.starttime, sheetEvent.endtime, sheetEvent);
      }
      // Put event ID back into spreadsheet
      idData[ridx][0] = newEvent.getId();
      eventsAdded = true;

      // Set event color
      if (sheetEvent.color > 0 && sheetEvent.color < 12) {
        newEvent.setColor('' + sheetEvent.color);
      }

      // Throttle updates.
      numAdded++;
      Utilities.sleep(THROTTLE_SLEEP_TIME);
      if (numAdded % 10 === 0) {
        console.info('%d events added, time: %d msecs', numAdded, Date.now() - scriptStart);
      }
    }
    // If the script is getting close to timing out, save the event IDs added so far to avoid lots
    // of duplicate events.
    if ((Date.now() - scriptStart) > MAX_RUN_TIME) {
      idRange.setValues(idData);
    }
  }

  // Save spreadsheet changes
  if (eventsAdded) {
    idRange.setValues(idData);
  }

  // Remove any calendar events not found in the spreadsheet
  var numToRemove = calEventIds.reduce(function(prevVal, curVal) {
    if (curVal !== null) {
      prevVal++;
    }
    return prevVal;
  }, 0);
  if (numToRemove > 0) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Delete ' + numToRemove + ' calendar event(s) not found in spreadsheet?',
          ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      var numRemoved = 0;
      calEventIds.forEach(function(id, idx) {
        if (id != null) {
          calEvents[idx].deleteEvent();
          Utilities.sleep(THROTTLE_SLEEP_TIME);
          numRemoved++;
          if (numRemoved % 10 === 0) {
            console.info('%d events removed, time: %d msecs', numRemoved, Date.now() - scriptStart);
          }
        }
      });
    }
  }
}

// Set up a trigger to automatically update the calendar when the spreadsheet is
// modified. See the instructions for how to use this.
function createSpreadsheetEditTrigger() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('syncToCalendar')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
}

// Delete the trigger. Use this to stop automatically updating the calendar.
function deleteTrigger() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var idx = 0; idx < allTriggers.length; idx++) {
    if (allTriggers[idx].getHandlerFunction() === 'syncToCalendar') {
      ScriptApp.deleteTrigger(allTriggers[idx]);
    }
  }
}
