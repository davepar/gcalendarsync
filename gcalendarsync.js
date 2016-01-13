// Script to synchronize a calendar to a spreadsheet.
//
// See https://github.com/Davepar/gcalendarsync for instructions on setting this up.
//

// Set these two values to match your calendar.
// Calendar ID can be found in the "Calendar Address" section of the Calendar Settings.
var calendarId = 'YOUR CALENDAR ID HERE'
var sheetName = 'Sheet1'

var fields = ['title', 'description', 'starttime', 'endtime', 'alldayevent', 'location', 'id'];

// Adds a custom menu to the active spreadsheet.
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{
    name: "Update Calendar",
    functionName: "syncCalendar"
  }];
  spreadsheet.addMenu('Calendar Sync', menuEntries);
}

// Creates a mapping between spreadsheet column and event field name
function createIdxMap(row) {
  var idxMap = [];
  for (var idx = 0; idx < row.length; idx++) {
    var fieldFromHdr = row[idx].toLowerCase().replace(/ /g, '');
    if (fields.indexOf(fieldFromHdr) > -1) {
      idxMap.push(fieldFromHdr);
    } else {
      idxMap.push(null);
    }
  }
  return idxMap;
}

// Converts a spreadsheet row into an object containing event-related fields
function reformatEvent(row, idxMap) {
  return row.reduce(function(event, value, idx) {
    if (idxMap[idx] != null) {
      event[idxMap[idx]] = value;
    }
    return event;
  }, {});
}

// Tests whether calendar event matches spreadsheet event
function eventMatches(cev, sev) {
  return cev.getTitle() == sev.title &&
    cev.getDescription() == sev.description &&
      cev.getStartTime().getTime() == sev.starttime.getTime() &&
        (!sev.endtime || cev.getEndTime().getTime() == sev.endtime.getTime()) &&
          cev.isAllDayEvent() == sev.alldayevent &&
            cev.getLocation() == sev.location;
}

// Display error alert
function errorAlert(msg, evt, ridx) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Skipping row: ' + msg + ' in event "' + evt.title + '", row ' + (ridx + 1));
}

// Synchronize spreadsheet to calendar.
function syncCalendar() {
  // Get calendar and events
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(new Date('1/1/1970'), new Date('1/1/2030'));
  var eventIds = events.map(function(val) {return val.getId()});

  // Get spreadsheet and data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var data = range.getValues();

  // Map headers to indices
  var idxMap = createIdxMap(data[0]);
  var idIdx = idxMap.indexOf('id');

  // Loop through spreadsheet rows
  var numUpdates = 0;
  var changesMade = false;
  for (var ridx = 1; ridx < data.length; ridx++) {
    var eventFromSheet = reformatEvent(data[ridx], idxMap);

    // Do some error checking first
    if (!eventFromSheet.title) {
      errorAlert('must have title', eventFromSheet, ridx);
      continue;
    }
    if (!(eventFromSheet.starttime instanceof Date)) {
      errorAlert('start time must be a date/time', eventFromSheet, ridx);
      continue;
    }
    if (!eventFromSheet.alldayevent && !(eventFromSheet.endtime instanceof Date)) {
      errorAlert('end time must be a date/time', eventFromSheet, ridx);
      continue;
    }
    if (!eventFromSheet.alldayevent && eventFromSheet.endtime < eventFromSheet.starttime) {
      errorAlert('end time must be after start time for event', eventFromSheet, ridx);
      continue;
    }

    var addEvent = true;
    if (eventFromSheet.id) {
      var eventIdx = eventIds.indexOf(eventFromSheet.id);
      if (eventIdx >= 0) {
        eventIds[eventIdx] = null;  // Prevents removing event below
        var event = events[eventIdx];
        if (eventMatches(event, eventFromSheet)) {
          addEvent = false;
        } else {
          // Delete and re-create event. It's easier than updating in place.
          event.deleteEvent();
        }
      }
    }
    if (addEvent) {
      var newEvent;
      if (eventFromSheet.alldayevent) {
        newEvent = calendar.createAllDayEvent(eventFromSheet.title, eventFromSheet.starttime, eventFromSheet);
      } else {
        newEvent = calendar.createEvent(eventFromSheet.title, eventFromSheet.starttime, eventFromSheet.endtime, eventFromSheet);
      }
      // Put event ID back into spreadsheet
      data[ridx][idIdx] = newEvent.getId();
      changesMade = true;

      // Updating too many calendar events in a short time interval triggers an error. Still experimenting with
      // the exact values to use here, but this works for updating about 40 events.
      numUpdates++;
      if (numUpdates > 10) {
        Utilities.sleep(50);
      }
    }
  }
  // Save spreadsheet changes
  if (changesMade) {
    range.setValues(data);
  }
  // Remove any calendar events not found in the spreadsheet
  eventIds.forEach(function(id, idx) {
    if (id != null) {
      events[idx].deleteEvent();
      Utilities.sleep(10);
    }
  });
}
