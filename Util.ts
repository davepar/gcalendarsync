// Utility functions

/*% import {EventColor, GenericEvent, GenericEventKey} from './GenericEvent'; %*/

const TITLE_ROW_MAP: Map<GenericEventKey, string> = new Map([
  ['title', 'Title'],
  ['description', 'Description'],
  ['location', 'Location'],
  ['starttime', 'Start Time'],
  ['endtime', 'End Time'],
  ['guests', 'Guests'],
  ['color', 'Color'],
  ['allday', 'All Day'],
  ['id', 'Id'],
]);

/*% export %*/ class Util {
  // Be aware that row indices are tricky, since Apps Script is one-based and Typescript arrays are zero-based.
  // These values are zero-based, so add one when using them with Apps Script API.
  static CALENDAR_ID_ROW = 0;
  static TITLE_ROW = 1;
  static FIRST_DATA_ROW = 2;
  static MAX_DATA_ROWS = 999;

  // Sets up an empty spreadsheet with a title row and suggested data formats
  static setUpSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet, numDataRows: number) {
    // Date format to use in the spreadsheet. Meaning of letters defined at
    // https://developers.google.com/sheets/api/guides/formats
    const dateFormat = 'M/d/yyyy H:mm a/p';

    // Add title row
    const titleRowValues = Array.from(TITLE_ROW_MAP.values());
    let range = sheet.getRange(Util.TITLE_ROW + 1, 1, 1, titleRowValues.length);
    range.setValues([titleRowValues]);

    // Set up date formats, checkbox for all day, and dropdown for color names
    const titleRowKeys = Array.from(TITLE_ROW_MAP.keys());
    const getRangeByFieldName =
      (fieldName: GenericEventKey, numRows: number) => sheet.getRange(Util.FIRST_DATA_ROW + 1, titleRowKeys.indexOf(fieldName) + 1, numRows);
    getRangeByFieldName('starttime', Util.MAX_DATA_ROWS).setNumberFormat(dateFormat);
    getRangeByFieldName('endtime', Util.MAX_DATA_ROWS).setNumberFormat(dateFormat);
    let checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    numDataRows = Math.max(numDataRows, 1);
    getRangeByFieldName('allday', numDataRows).setDataValidation(checkboxRule);
    const colorList = [];
    for (let enumColor in EventColor) {
      if (isNaN(parseInt(enumColor, 10))) {
        colorList.push(enumColor);
      }
    }
    let colorDropdownRule =
      SpreadsheetApp.newDataValidation().requireValueInList(colorList, true).build();
    getRangeByFieldName('color', numDataRows).setDataValidation(colorDropdownRule);

    // Hide ID column so people are less liken to modify it accidentally
    sheet.hideColumns(titleRowKeys.indexOf('id') + 1);
  }

  // Creates a mapping array between spreadsheet column and event field name
  static createIdxMap(row: any[]): GenericEventKey[] {
    let idxMap: GenericEventKey[] = [];
    for (let fieldFromHdr of row) {
      let found = false;
      for (let titleKey of Array.from(TITLE_ROW_MAP.keys())) {
        if (TITLE_ROW_MAP.get(titleKey) == fieldFromHdr) {
          idxMap.push(titleKey);
          found = true;
          break;
        }
      }
      if (!found) {
        // Header field not in map, so add null
        idxMap.push(null);
      }
    }
    return idxMap;
  }

  // Returns list of fields that aren't in spreadsheet
  static missingFields(idxMap: GenericEventKey[]): GenericEventKey[] {
    return Array.from(TITLE_ROW_MAP.keys()).filter(val => idxMap.indexOf(val) < 0);
  }

  // Return list of missing required fields.
  static missingRequiredFields(idxMap: GenericEventKey[], includeAllDay: boolean):
    GenericEventKey[] {
    let requiredFields: GenericEventKey[] = ['id', 'title', 'starttime', 'endtime'];
    if (includeAllDay) {
      requiredFields.push('allday');
    }
    return requiredFields.filter(val => idxMap.indexOf(val) < 0);
  }

  static displayMissingFields(missingFields: GenericEventKey[], sheetName: string) {
    const reqFieldNames = missingFields.map(x => `"${TITLE_ROW_MAP.get(x)}"`).join(', ');
    Util.errorAlert(`Sheet "${sheetName}" is missing ${reqFieldNames} columns. See Help for setup instructions.`);
  }

  // Display error alert
  static errorAlert(msg: string, evt: GenericEvent = null, ridx = 0) {
    const ui = SpreadsheetApp.getUi();
    if (evt) {
      ui.alert(`Skipping row: ${msg} in event "${evt.title}", row ${ridx + 1}`);
    } else {
      ui.alert(msg);
    }
  }

  static errorAlertHalt(msg: string) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(`${msg} Continue?`, ui.ButtonSet.YES_NO);
    return response === ui.Button.NO;
  }

  static isValidDate(d: string) {
    return isNaN(Date.parse(d)) === false;
  }
}
