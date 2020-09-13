import { Settings, AllDayValue } from '../Settings';

describe('Settings', () => {
  let fakePropertiesService: any;
  let fakeDocumentProperties: any;
  beforeEach(() => {
    fakeDocumentProperties = {
      getProperty: () => '',
      setProperty: (key: string, value: string): void => { },
    };
    fakePropertiesService = {
      getDocumentProperties: () => fakeDocumentProperties as any,
    };
  });

  it('creates defaults', () => {
    const settings = Settings.getDefaultSettings();
    expect(settings.begin_date).toEqual(new Date(1970, 0, 1));
  });

  it('loads from PropertiesService', () => {
    const jsonSettings = '{"begin_date":"1980-1-1","end_date":"2400-1-1","send_email_invites":true,"skip_blank_rows":true,"all_day_events":"USE_COLUMN"}';
    spyOn(fakeDocumentProperties, 'getProperty').and.returnValue(jsonSettings);

    const settings = Settings.loadFromPropertyService(fakePropertiesService as any);
    expect(settings).toEqual(new Settings(
      new Date(1980, 0, 1), new Date(2400, 0, 1), true, true, AllDayValue.use_column
    ));
  });

  it('loads defaults from PropertyService', () => {
    spyOn(fakeDocumentProperties, 'getProperty').and.returnValue(null);
    const setPropertySpy = spyOn(fakeDocumentProperties, 'setProperty');

    const settings = Settings.loadFromPropertyService(fakePropertiesService as any);
    expect(settings).toEqual(Settings.getDefaultSettings());
    expect(setPropertySpy).toHaveBeenCalled();
  });

  it('saves to PropertiesService', () => {
    const setPropertySpy = spyOn(fakeDocumentProperties, 'setProperty');
    const formValues = {
      begin_date: '1980-1-1',
      end_date: '2400-1-1',
      send_email_invites: true,
      skip_blank_rows: true,
      all_day_events: 'USE_COLUMN',
    }
    Settings.saveToPropertyService(formValues, fakePropertiesService);
    const expectedJson = '{"begin_date":"1980-1-1","end_date":"2400-1-1","send_email_invites":true,"skip_blank_rows":true,"all_day_events":"USE_COLUMN"}';
    expect(setPropertySpy).toHaveBeenCalledWith('v1', expectedJson);
  });

  it('converts to base types', () => {
    const settings = Settings.getDefaultSettings();
    expect(settings.convertToBaseTypes()).toEqual({
      begin_date: '1970-1-1',
      end_date: '2500-1-1',
      send_email_invites: false,
      skip_blank_rows: false,
      all_day_events: 'USE_COLUMN',
    })
  });

  it('converts for dialog', () => {
    const savedSettings = Settings.getDefaultSettings();
    spyOn(Settings, 'loadFromPropertyService').and.returnValue(savedSettings);
    expect(Settings.convertForDialog()).toEqual({
      begin_date: '1970-1-1',
      end_date: '2500-1-1',
      send_email_invites: false,
      skip_blank_rows: false,
      all_day_events: 'use_column',
    });
  });

  it('converts date to string', () => {
    const str = Settings.convertDateToString(new Date(1970, 0, 1));
    expect(str).toEqual('1970-1-1');
  });
});
