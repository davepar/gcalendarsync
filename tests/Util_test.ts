import { Util } from "../Util";
import { GenericEventKey } from '../GenericEvent';

const DATE1 = new Date('1995-12-17T03:24:00');
const FAKE_IDX_MAP: GenericEventKey[] = ['id', null, 'title', 'description', 'color'];

describe('createIdxMap', () => {
  it('creates map', () => {
    const result = Util.createIdxMap(['Title', 'Color', 'unknown', 'Start Time']);
    expect(result).toEqual(['title', 'color', null, 'starttime']);
  });
});

describe('missingFields', () => {
  it('returns missing fields', () => {
    const result = Util.missingFields(FAKE_IDX_MAP);
    expect(result).toEqual(['location', 'starttime', 'endtime', 'guests', 'allday']);
  });
});

describe('missingRequiredFields', () => {
  it('returns missing required fields', () => {
    const result1 = Util.missingRequiredFields(FAKE_IDX_MAP, true);
    expect(result1).toEqual(['starttime', 'endtime', 'allday']);
    const result2 = Util.missingRequiredFields(FAKE_IDX_MAP, false);
    expect(result2).toEqual(['starttime', 'endtime']);
  });
});

describe('isValidDate', () => {
  it('identifies valid dates', () => {
    expect(Util.isValidDate('5/1/2020')).toBeTruthy();
    expect(Util.isValidDate('2020-5-1')).toBeTruthy();
    expect(Util.isValidDate('30/1/2020')).toBeFalsy();
    expect(Util.isValidDate('2020-5-89')).toBeFalsy();
  });
});
