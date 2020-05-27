class FakeGuest implements GoogleAppsScript.Calendar.EventGuest {
  constructor(
   public email: string
  ) {}
  getEmail() {
   return this.email;
  }
  getAdditionalGuests(): number { throw "not implemented"; };
  getGuestStatus(): GoogleAppsScript.Calendar.GuestStatus { throw "not implemented"; };
  getName(): string { throw "not implemented"; };
  getStatus(): string { throw "not implemented"; };
}
  
export class FakeCalendarEvent implements GoogleAppsScript.Calendar.CalendarEvent {
  constructor(
    public id: string,
    public title: string,
    public description: string,
    public location: string,
    public guests: GoogleAppsScript.Calendar.EventGuest[],
    public color: string,
    public allday: boolean,
    public starttime: Date,
    public endtime: Date
  ) {}
  static fromArray(params: any[]) {
    const [id, title, description, location, guests, color, allday, starttime, endtime] =
      params;
    let guestList = guests ? guests.split(',').map((x) => new FakeGuest(x.trim())) : [];
    return new FakeCalendarEvent(id, title, description, location, guestList, color,
      allday, starttime, endtime);
  }
  getId() {
    return this.id;
  }
  getTitle() {
    return this.title;
  }
  getDescription() {
    return this.description;
  }
  getLocation() {
    return this.location;
  }
  getGuestList() {
    return this.guests;
  }
  getColor() {
    return this.color;
  }
  isAllDayEvent() {
    return this.allday;
  }
  getAllDayStartDate() {
    if (!this.allday) {
      throw "invalid call";
    }
    return this.starttime;
  }
  getAllDayEndDate() {
    if (!this.allday) {
      throw "invalid call";
    }
    return this.endtime;
  }
  getStartTime() {
    if (this.allday) {
      throw "invalid call";
    }
    return this.starttime;
  }
  getEndTime() {
    if (this.allday) {
      throw "invalid call";
    }
    return this.endtime;
  }
  addEmailReminder(minutesBefore: number): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  addGuest(email: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  addPopupReminder(minutesBefore: number): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  addSmsReminder(minutesBefore: number): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  anyoneCanAddSelf(): boolean { throw "not implemented"; };
  deleteEvent() { throw "not implemented"; };
  deleteTag(key: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  getAllTagKeys(): string[] { throw "not implemented"; };
  getCreators(): string[] { throw "not implemented"; };
  getDateCreated(): Date { throw "not implemented"; };
  getEmailReminders(): number[] { throw "not implemented"; };
  getEventSeries(): GoogleAppsScript.Calendar.CalendarEventSeries { throw "not implemented"; };
  getGuestByEmail(email: string): GoogleAppsScript.Calendar.EventGuest { throw "not implemented"; };
  getLastUpdated(): Date { throw "not implemented"; };
  getMyStatus(): GoogleAppsScript.Calendar.GuestStatus { throw "not implemented"; };
  getOriginalCalendarId(): string { throw "not implemented"; };
  getPopupReminders(): number[] { throw "not implemented"; };
  getSmsReminders(): number[] { throw "not implemented"; };
  getTag(key: string): string { throw "not implemented"; };
  getVisibility(): GoogleAppsScript.Calendar.Visibility { throw "not implemented"; };
  guestsCanInviteOthers(): boolean { throw "not implemented"; };
  guestsCanModify(): boolean { throw "not implemented"; };
  guestsCanSeeGuests(): boolean { throw "not implemented"; };
  isOwnedByMe(): boolean { throw "not implemented"; };
  isRecurringEvent(): boolean { throw "not implemented"; };
  removeAllReminders(): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  removeGuest(email: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  resetRemindersToDefault(): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setAllDayDate(date: Date): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setAllDayDates(startDate: Date, endDate: Date): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setAnyoneCanAddSelf(anyoneCanAddSelf: boolean): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setColor(color: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setDescription(description: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setGuestsCanInviteOthers(guestsCanInviteOthers: boolean): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setGuestsCanModify(guestsCanModify: boolean): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setGuestsCanSeeGuests(guestsCanSeeGuests: boolean): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setLocation(location: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setMyStatus(status: GoogleAppsScript.Calendar.GuestStatus): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setTag(key: string, value: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setTime(startTime: Date, endTime: Date): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setTitle(title: string): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
  setVisibility(visibility: GoogleAppsScript.Calendar.Visibility): GoogleAppsScript.Calendar.CalendarEvent { throw "not implemented"; };
}
