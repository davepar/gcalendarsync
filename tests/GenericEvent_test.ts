import {GenericEvent} from '../GenericEvent';
import {Util} from '../Util';
import {AllDayValue} from '../Settings';
import {FakeCalendarEvent} from './FakeCalendarEvent';

const DATE1 = new Date('1995-12-17T03:24:00');
const DATE2 = new Date('1995-12-18T04:56:00');
const DATE3 = new Date('1995-12-19T07:08:00');
const DATE4 = new Date('1995-12-20T00:00:00');
const DATE5 = new Date('1995-12-22T00:00:00');

const EVENT1_VALUES = ['testid1', 'Test Title 1', 'Test Description 1', 'Test Location 1',
  'guest1@example.com,guest2@example.com', 'ORANGE', false, DATE1, DATE2];
const EVENT2_VALUES = ['testid2', 'Test Title 2', 'Test Description 2', 'Test Location 2',
  'guest3@example.com,guest4@example.com', 'GRAY', false, DATE2, DATE3];
const EVENT_NOGUESTS_VALUES = ['testid3', 'Test Title 3', 'Test Description 3', 'Test Location 3',
  '', 'MAUVE', false, DATE2, DATE3];
const EVENT_ALLDAY_VALUES = ['testid4', 'Test Title 4', 'Test Description 4', 'Test Location 4',
  '', '', true, DATE4, DATE5];
const EVENT_BADDATES_VALUES = ['testid5', 'Test Title 5', 'Test Description 5', 'Test Location 5',
  '', '', false, 'abc', 0]
  const EVENT_BADCOLOR_VALUES = ['testid6', 'Test Title 6', 'Test Description 6', 'Test Location 6',
  '', 'foobar', false, DATE2, DATE3]

const IDX_MAP = Util.createIdxMap(['Id', 'Title', 'Description', 'Location', 'Guests', 'Color',
  'All Day', 'Start Time', 'End Time']);
const IDX_MAP_NO_GUESTS = Util.createIdxMap(['Id', 'Title', 'Description', 'Location', 'Color', 'All Day',
  'Start Time', 'End Time']);

describe('GenericEvent', () => {
  let event1, event2, event_noguests, event_allday: GenericEvent;
  beforeAll(() => {
    event1 = GenericEvent.fromArray(EVENT1_VALUES);
    event2 = GenericEvent.fromArray(EVENT2_VALUES);
    event_noguests = GenericEvent.fromArray(EVENT_NOGUESTS_VALUES);
    event_allday = GenericEvent.fromArray(EVENT_ALLDAY_VALUES);
  });
  it('instantiates correctly', () => {
    expect(event1.id).toBe('testid1');
    expect(event1.color).toBe('6');
  });

  describe('fromCalendarEvent', () => {
    it('initantiates correctly with all fields', () => {
      const fakeCalEvent = FakeCalendarEvent.fromArray(EVENT1_VALUES);
      expect(fakeCalEvent.getColor()).toEqual('6')
      const event1_fromcal = GenericEvent.fromCalendarEvent(fakeCalEvent);
      expect(event1_fromcal).toEqual(event1);
    });

    it('initantiates correctly for all day events', () => {
      const fakeCalEvent = FakeCalendarEvent.fromArray(EVENT_ALLDAY_VALUES);
      const event_allday_fromcal = GenericEvent.fromCalendarEvent(fakeCalEvent);
      expect(event_allday_fromcal).toEqual(event_allday);
    });
  });

  describe('fromSpreadsheetRow', () => {
    it('instantiates correctly with all fields', () => {
      const event1_fromsheet = GenericEvent.fromSpreadsheetRow(EVENT1_VALUES, IDX_MAP, [], AllDayValue.use_column);
      expect(event1_fromsheet).toEqual(event1);
    });
    it('instantiates all day event correctly', () => {
      const event_allday_fromsheet = GenericEvent.fromSpreadsheetRow(EVENT_ALLDAY_VALUES, IDX_MAP, [], AllDayValue.use_column);
      expect(event_allday_fromsheet).toEqual(event_allday);
    });
    it('instantiates correctly with a blank field added later', () => {
      const event_noguests_fromsheet = GenericEvent.fromSpreadsheetRow(
        EVENT_NOGUESTS_VALUES.filter(val => val !== ''), IDX_MAP_NO_GUESTS, ['guests'], AllDayValue.use_column);
      expect(event_noguests_fromsheet).toEqual(event_noguests);
    });
    it('turns bad dates into null', () => {
      const event_baddates = GenericEvent.fromSpreadsheetRow(EVENT_BADDATES_VALUES, IDX_MAP, [], AllDayValue.use_column);
      expect(event_baddates.starttime).toBeNull();
      expect(event_baddates.endtime).toBeNull();
    });
    it('handles always all-day event', () => {
      const event1_fromsheet = GenericEvent.fromSpreadsheetRow(EVENT1_VALUES, IDX_MAP, [], AllDayValue.always_all_day);
      expect(event1_fromsheet.allday).toBeTruthy();
    });
    it('handles never all-day event', () => {
      const event_allday_fromsheet = GenericEvent.fromSpreadsheetRow(EVENT_ALLDAY_VALUES, IDX_MAP, [], AllDayValue.never_all_day);
      expect(event_allday_fromsheet.allday).toBeFalsy();
    });
    it('skips bad color', () => {
      const event_allday_fromsheet = GenericEvent.fromSpreadsheetRow(EVENT_BADCOLOR_VALUES, IDX_MAP, [], AllDayValue.use_column);
      expect(event_allday_fromsheet.color).toEqual('');
    });
  });

  describe('toSpreadsheetRow', () => {
    it('translates to spreadsheet row', () => {
      const row = new Array(8);
      event1.toSpreadsheetRow(IDX_MAP, row)
      expect(row).toEqual(EVENT1_VALUES)
    });
  });

  describe('eventDifferences', () => {
    it('calculates correct number of diffs', () => {
      expect(event_noguests.eventDifferences(event1)).toBe(7);
    });
    it('calculates correct number when no diffs', () => {
      expect(event1.eventDifferences(event1)).toBe(0);
    });
    it('sets diffs to 1 when there are guests', () => {
      expect(event1.eventDifferences(event2)).toBe(1);
    });
  });
});
