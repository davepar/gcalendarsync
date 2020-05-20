/*% export %*/ const Settings = {
  SETTINGS_VERSION: 'v1',

  // Updating too many events in a short time period triggers an error. These values
  // were successfully used for deleting and adding 240 events. Values in milliseconds.
  THROTTLE_SLEEP_TIME: 200,
  MAX_RUN_TIME: 5.75 * 60 * 1000,
};

// Values for the "all day" setting. The enum names must match the IDs in the HTML radio
// button definitions.
/*% export %*/ enum AllDayValue {
  use_column = 'USE_COLUMN',
  always_all_day = 'ALWAYS_ALL_DAY',
  never_all_day = 'NEVER_ALL_DAY',
}
