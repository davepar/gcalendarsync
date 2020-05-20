// These imports are only used for testing. Run pretest and posttest scripts to automatically
// uncomment and re-comment these lines.
/*% import {Settings, AllDayValue} from './Settings'; %*/

/*% export %*/ class GenericEvent {
  constructor(
    public id: string,
    public title: string,
    public description: string,
    public location: string,
    public guests: string,
    public color: string,
    public allday: boolean,
    public starttime: Date,
    public endtime: Date
  ) {}

  static fromArray(params: any[]) {
    const [id, title, description, location, guests, color, allday, starttime, endtime] =
      params;
    return new GenericEvent(id, title, description, location, guests, color,
      allday, starttime, endtime);
  }

  // Converts a GCalendar event to an instance of this class.
  static fromCalendarEvent(calEvent: GoogleAppsScript.Calendar.CalendarEvent) {
    const allday = calEvent.isAllDayEvent();
    let starttime: Date, endtime: Date;
    if (allday) {
      starttime = calEvent.getAllDayStartDate();
      endtime = calEvent.getAllDayEndDate();
      if (endtime.getHours() === 0 && endtime.getMinutes() === 0) {
        endtime.setSeconds(endtime.getSeconds() - 1);
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
  static fromSpreadsheetRow(row: any[], idxMap: string[], keysToAdd: string[],
      all_day_events: AllDayValue) {
    const eventObject = row.reduce((event, value, idx) => {
      const field = idxMap[idx];
      if (field != null) {
        if (field === 'starttime' || field === 'endtime') {
          event[field] = (isNaN(value) || value == 0) ? null : new Date(value);
        } else if (field === 'allday') {
          event[field] = (value === true);
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
    return new GenericEvent(id, title, description, location, guests, color, allday,
      starttime, endtime);
  }

  toSpreadsheetRow(idxMap: string[], spreadsheetRow: any[]) {
    for (let idx = 0; idx < idxMap.length; idx++) {
      if (idxMap[idx] !== null) {
        if (idxMap[idx] === 'allday') {
          spreadsheetRow[idx] = !!this[idxMap[idx]];
        } else {
          spreadsheetRow[idx] = this[idxMap[idx]];
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
        if (sheetEvent.endtime.getHours() === 23 && sheetEvent.endtime.getMinutes() === 59) {
          sheetEvent.endtime.setSeconds(sheetEvent.endtime.getSeconds() + 1);
        }
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

} // GenericEvent
