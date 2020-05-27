// Utility functions

/*% import {GenericEvent, GenericEventKey} from './GenericEvent'; %*/

const title_row_map: Map<GenericEventKey, string> = new Map([
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
  static TITLE_ROW_MAP = title_row_map;

  // Creates a mapping array between spreadsheet column and event field name
  static createIdxMap(row:any[]): GenericEventKey[] {
    let idxMap: GenericEventKey[] = [];
    for (let fieldFromHdr of row) {
      let found = false;
      for (let titleKey of Array.from(title_row_map.keys())) {
        if (title_row_map.get(titleKey) == fieldFromHdr) {
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
  static missingFields(idxMap: GenericEventKey[]) {
    return Array.from(title_row_map.keys()).filter(val => idxMap.indexOf(val) < 0);
  }

  // Return list of missing required fields.
  static missingRequiredFields(idxMap: GenericEventKey[], includeAllDay: boolean) {
    let requiredFields: GenericEventKey[] = ['id', 'title', 'starttime', 'endtime'];
    if (includeAllDay) {
      requiredFields.push('allday');
    }
    return requiredFields.filter(val => idxMap.indexOf(val) < 0);
  }

  // Display error alert
  static errorAlert(msg: string, evt: GenericEvent=null, ridx=0) {
    const ui = SpreadsheetApp.getUi();
    if (evt) {
      ui.alert(`Skipping row: ${msg} in event "${evt.title}", row ${ridx + 1}`);
    } else {
      ui.alert(msg);
    }
  }

  static isValidDate(d: string) {
    return isNaN(Date.parse(d)) === false;
  }
}
