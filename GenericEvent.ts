// These imports are only used for testing. Run pretest and posttest scripts to automatically
// uncomment and re-comment these lines.
/*% import {Settings, AllDayValue} from './Settings'; %*/

/*% export %*/ enum EventColor { PALE_BLUE, PALE_GREEN, MAUVE, PALE_RED, YELLOW, ORANGE, CYAN, GRAY, BLUE, GREEN, RED }

/*% export %*/ type GenericEventKey = 'id' | 'title' | 'description' | 'location' |
  'guests' | 'color' | 'allday' | 'starttime' | 'endtime';

/*% export %*/ class GenericEvent {
  constructor(
    public id: string,
    public title: string,
    public description: string,
    public location: string,
    public guests: string,
    public color: string,  // Stored as color number
    public allday: boolean,
    public starttime: Date,
    public endtime: Date
  ) {}

  static fromArray(params: any[]) {
    let [id, title, description, location, guests, color, allday, starttime, endtime] =
      params;
    let convertedColor = color ? (EventColor[color] + 1).toString() : '';
    return new GenericEvent(id, title, description, location, guests, convertedColor,
      allday, starttime, endtime);
  }

  // Converts a GCalendar event to an instance of this class.
  static fromCalendarEvent(calEvent: GoogleAppsScript.Calendar.CalendarEvent) {
    const allday = calEvent.isAllDayEvent();
    let starttime: Date, endtime: Date;
    if (allday) {
      starttime = calEvent.getAllDayStartDate();
      endtime = calEvent.getAllDayEndDate();
      if (endtime.getHours() === 0 && endtime.getMinutes() === 0 && endtime.getSeconds() === 0) {
        endtime.setDate(endtime.getDate() - 1);
      }
    } else {
      starttime = calEvent.getStartTime();
      endtime = calEvent.getEndTime();
    }
    return new GenericEvent(
      calEvent.getId(),
      calEvent.getTitle(),
      calEvent.getDescription(),
      calEvent.getLocation(),
      calEvent.getGuestList().map(x => x.getEmail()).join(','),
      calEvent.getColor(),
      allday,
      starttime,
      endtime);
  }

  // Convert a spreadsheet row to an instance of this class.
  static fromSpreadsheetRow(row: any[], idxMap: GenericEventKey[], keysToAdd: string[],
      all_day_events: AllDayValue, timeZone = 'America/Los_Angeles') {
    const eventObject = row.reduce((event, value, idx) => {
      const field = idxMap[idx];
      if (field != null) {
        if (field === 'starttime' || field === 'endtime') {
          event[field] = (isNaN(value) || value == 0) ? null : value;
        } else if (field === 'allday') {
          event[field] = (value === true);
        } else if (field === 'color') {
          if (value) {
            value = EventColor[value];
            event[field] = isNaN(value) ? '' : (value + 1).toString();
          } else {
            event[field] = '';
          }
        } else {
          event[field] = value;
        }
      }
      return event;
    }, {});
    for (let keyToAdd of keysToAdd) {
      eventObject[keyToAdd] = (keyToAdd === 'starttime' || keyToAdd === 'endTime') ? null : '';
    }
    let {id, title, description, location, guests, color, allday, starttime, endtime} =
      eventObject;
    if (all_day_events !== AllDayValue.use_column) {
      allday = (all_day_events === AllDayValue.always_all_day);
    }
    // Adjust allday events to correct time zone and add 1 day to end time
    if (allday) {
      starttime = GenericEvent.convertToScriptTimeZone(starttime, timeZone);
      if (endtime) {
        endtime = GenericEvent.convertToScriptTimeZone(endtime, timeZone, 1);
      }
    }
    return new GenericEvent(id, title, description, location, guests, color, allday,
      starttime, endtime);
  }

  toSpreadsheetRow(idxMap: GenericEventKey[], spreadsheetRow: any[]) {
    for (let idx = 0; idx < idxMap.length; idx++) {
      if (idxMap[idx] !== null) {
        const label = idxMap[idx];
        const value = this[label] as any;
        if (label === 'allday') {
          spreadsheetRow[idx] = !!value;
        } else if (label === 'color') {
          spreadsheetRow[idx] = value ? EventColor[Number(value) - 1] : '';
        } else if (this.allday && (label === 'starttime' || label === 'endtime')) {
          // Convert to a string to get around time zone issues
          spreadsheetRow[idx] = value.toLocaleDateString('en-US');
        } else {
          spreadsheetRow[idx] = value;
        }
      }
    }
  }

  // Determines the number of field differences between this and a another event
  eventDifferences(other: GenericEvent) {
    let eventDiffs = 0;
    if (this.title !== other.title) eventDiffs += 1;
    if (this.description !== other.description) eventDiffs += 1;
    if (this.location !== other.location) eventDiffs += 1;
    if (this.starttime.toString() !== other.starttime.toString()) eventDiffs += 1;
    if (this.endtime.toString() !== other.endtime.toString()) eventDiffs += 1;
    if (this.guests !== other.guests) eventDiffs += 1;
    if (this.color !== other.color) eventDiffs += 1;
    if (this.allday !== other.allday) eventDiffs += 1;
    if (eventDiffs > 0 && this.guests) {
      // When an event changes and it has guests, set the diffs to one. This will force the
      // calling function to update all of the fields instead of deleting and re-adding the event,
      // which would force every guest to re-confirm the event.
      eventDiffs = 1;
    }
    return eventDiffs;
  }

  // Updates a calendar event from a sheet event.
  updateEvent(sheetEvent: GenericEvent, calEvent: GoogleAppsScript.Calendar.CalendarEvent) {
    let numChanges = 0;
    const isAllDayChanged = (this.allday !== sheetEvent.allday);
    if (this.starttime.toString() !== sheetEvent.starttime.toString() ||
        this.endtime.toString() !== sheetEvent.endtime.toString() || isAllDayChanged) {
      if (sheetEvent.allday) {
        calEvent.setAllDayDates(sheetEvent.starttime, sheetEvent.endtime);
      } else {
        calEvent.setTime(sheetEvent.starttime, sheetEvent.endtime);
      }
      numChanges++;
    }
    if (this.title !== sheetEvent.title) {
      calEvent.setTitle(sheetEvent.title);
      numChanges++;
    }
    if (this.description !== sheetEvent.description) {
      calEvent.setDescription(sheetEvent.description);
      numChanges++;
    }
    if (this.location !== sheetEvent.location) {
      calEvent.setLocation(sheetEvent.location);
      numChanges++;
    }
    if (this.color !== ('' + sheetEvent.color)) {
      const color = parseInt(sheetEvent.color);
      if (color > 0 && color < 12) {
        calEvent.setColor('' + color);
        numChanges++;
      }
    }
    if (this.guests !== sheetEvent.guests) {
      const guestCal = calEvent.getGuestList().map(x => ({email: x.getEmail(), added: false}));
      const sheetGuests = sheetEvent.guests || '';
      let guests = sheetGuests.split(',').map((x) => x ? x.trim() : '');
      // Check guests that are already invited.
      for (let gIx = 0; gIx < guestCal.length; gIx++) {
        const index = guests.indexOf(guestCal[gIx].email);
        if (index >= 0) {
          guestCal[gIx].added = true;
          guests.splice(index, 1);
        }
      }
      for (let guest of guests) {
        if (guest) {
          calEvent.addGuest(guest);
          numChanges++;
        }
      }
      for (let guest of guestCal) {
        if (!guest.added) {
          calEvent.removeGuest(guest.email);
          numChanges++;
        }
      }
    }
    // Throttle updates.
    Utilities.sleep(Settings.THROTTLE_SLEEP_TIME * numChanges);
    return numChanges;
  }

  // AppScripts seem to always run in Pacific time. This function will convert
  // a date to whatever time zone the script is running in.
  static convertToScriptTimeZone(d: Date, timeZone: string, daydelta = 0) {
    if (!(d instanceof Date)) {
      return null
    }
    const adjDate = d.toLocaleDateString('en-US', {timeZone});
    let [month, day, year] = adjDate.split('/').map(x => parseInt(x));
    return new Date(year, month - 1, day + daydelta)
  }

} // GenericEvent
