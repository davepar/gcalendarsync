// Utility functions

const title_row_map = {
  'title': 'Title',
  'description': 'Description',
  'location': 'Location',
  'starttime': 'Start Time',
  'endtime': 'End Time',
  'guests': 'Guests',
  'color': 'Color',
  'allday': 'All Day',
  'id': 'Id',
};

const title_row_keys = Object.keys(title_row_map);

/*% export %*/ class Util {
  static TITLE_ROW_MAP = title_row_map;
  static TITLE_ROW_KEYS = title_row_keys;

  // Creates a mapping array between spreadsheet column and event field name
  static createIdxMap(row:any[]): string[] {
    let idxMap = [];
    for (let fieldFromHdr of row) {
      let found = false;
      for (let titleKey in title_row_map) {
        if (title_row_map[titleKey] == fieldFromHdr) {
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
  static missingFields(idxMap: string[]) {
    return title_row_keys.filter(val => idxMap.indexOf(val) < 0);
  }

  // Return list of missing required fields.
  static missingRequiredFields(idxMap: string[], includeAllDay: boolean) {
    let requiredFields = ['id', 'title', 'starttime', 'endtime'];
    if (includeAllDay) {
      requiredFields.push('allday');
    }
    return requiredFields.filter(val => idxMap.indexOf(val) < 0);
  }

  // Display error alert
  static errorAlert(msg, evt=null, ridx=0) {
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
